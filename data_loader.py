# data_loader.py
import pandas as pd
import math

def normalize_id(val):
    """标准化工号为纯数字字符串，去前导零。无效值返回 None。"""
    if val is None:
        return None
    try:
        f = float(val)
        if math.isnan(f):
            return None
        return str(int(f))
    except (ValueError, TypeError):
        s = str(val).strip()
        if s.isdigit():
            return str(int(s))
        return None


def format_display_id(norm_id, length: int = 6) -> str:
    """将标准化工号补全前导零到 length 位（不足则补，超出则原样）。"""
    if not norm_id:
        return ''
    s = str(norm_id)
    return s.zfill(max(length, len(s)))


def load_prev_list(path: str, sheet_name: str = None) -> set:
    """
    从公示表 Excel 中提取"进入能力提升计划名单"那一节的工号集合（标准化，去前导零）。
    工号固定在第 5 列（E 列，index=4）。
    只取标题行含"进入"之后、下一个标题行之前的工号，排除"未出营"一节。
    找不到"进入"节时返回空集合（不退化为读全部，避免把未出营人员误带入）。
    sheet_name 指定时只读该 Sheet，否则遍历所有 Sheet。
    """
    ids = set()
    try:
        xl = pd.ExcelFile(path)
        sheets = [sheet_name] if sheet_name else xl.sheet_names
        for sn in sheets:
            df = pd.read_excel(xl, sheet_name=sn, header=None, dtype=str)
            if df.shape[1] < 5:
                continue

            # 找"进入能力提升"节的起始行：只看 A 列（合并标题行），排除"未出营"
            section_start = None
            for i, row in df.iterrows():
                a_val = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
                if '进入' in a_val and ('能力提升' in a_val or '提升' in a_val):
                    if '未出营' not in a_val:
                        section_start = i
                        break

            if section_start is None:
                # 找不到"进入"节，跳过该 Sheet，不退化读全部
                continue

            # 从节起始行的下一行开始，读到文件末或下一个 A 列标题行
            for i in range(section_start + 1, len(df)):
                row = df.iloc[i]
                e_val = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ''
                a_val = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
                # A 列含"名单"且 E 列不是工号 → 遇到新 section，停止
                if '名单' in a_val and normalize_id(e_val) is None:
                    break
                norm = normalize_id(e_val)
                if norm:
                    ids.add(norm)
    except Exception:
        pass
    return ids


def load_roster(path: str, sheet_name: str = '花名册最新') -> pd.DataFrame:
    """读取花名册，返回在职人员，工号标准化为字符串。"""
    df = pd.read_excel(path, sheet_name=sheet_name, header=0)
    df = df[df['雇佣状态'] == '在职'].copy()
    df['工号'] = df['工号'].apply(normalize_id)
    df = df[df['工号'].notna()].reset_index(drop=True)
    return df


def load_expert_vehicles(path: str, sheet_name: str = '专家底表') -> list:
    """读取专家底表车系列，返回去重后的非"全部"车系（按出现顺序）。"""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=1)
        seen, result = set(), []
        for v in df['车系'].dropna().astype(str):
            v = v.strip()
            if v and v != '全部' and v not in seen:
                seen.add(v)
                result.append(v)
        return result
    except Exception:
        return []


def load_expert_orders(path: str,
                       sheet_name: str = '专家底表',
                       exclude_month1: list = None,
                       exclude_month2: list = None) -> pd.DataFrame:
    """
    读取专家底表。

    口径定义（业务确认）：
      全量 = 车系为"全部"的那一行（就是所有车型的合计）
      绩效 = 全量 - 非绩效车型（即"全部"减去被剔除的车系）

    返回列: 专家工号, 专家姓名,
            月1净锁单, 月2净锁单          （绩效口径）
            月1净锁单_全量, 月2净锁单_全量  （全量口径 = "全部"行，不受剔除影响）
    """
    df = pd.read_excel(path, sheet_name=sheet_name, header=1)
    df = df.rename(columns={'净锁单': '月1净锁单', '净锁单.1': '月2净锁单'})
    df['专家工号'] = df['专家工号'].apply(normalize_id)
    df = df[df['专家工号'].notna()]
    df['月1净锁单'] = pd.to_numeric(df['月1净锁单'], errors='coerce').fillna(0)
    df['月2净锁单'] = pd.to_numeric(df['月2净锁单'], errors='coerce').fillna(0)

    df['_车系'] = df['车系'].astype(str).str.strip()
    df_all = df[df['_车系'] == '全部'].copy()
    df_det = df[df['_车系'] != '全部'].copy()

    # 全量 = "全部"汇总行；底表缺该行时按明细求和兜底
    if len(df_all) > 0:
        full = (df_all[['专家工号', '专家姓名', '月1净锁单', '月2净锁单']]
                .drop_duplicates(subset=['专家工号'], keep='first')
                .rename(columns={'月1净锁单': '月1净锁单_全量',
                                 '月2净锁单': '月2净锁单_全量'}))
    else:
        full = (df_det.groupby('专家工号', as_index=False)
                .agg(专家姓名=('专家姓名', 'first'),
                     月1净锁单_全量=('月1净锁单', 'sum'),
                     月2净锁单_全量=('月2净锁单', 'sum')))

    if not exclude_month1 and not exclude_month2:
        # 未剔除非绩效车型：绩效 = 全量
        perf = full[['专家工号', '专家姓名', '月1净锁单_全量', '月2净锁单_全量']].copy()
        perf['月1净锁单'] = perf['月1净锁单_全量']
        perf['月2净锁单'] = perf['月2净锁单_全量']
    else:
        # 绩效 = 全部 - 被剔除车系
        excl1 = set(exclude_month1 or [])
        excl2 = set(exclude_month2 or [])
        m1 = (df_det[df_det['_车系'].isin(excl1)]
              .groupby('专家工号')['月1净锁单'].sum().rename('_剔1'))
        m2 = (df_det[df_det['_车系'].isin(excl2)]
              .groupby('专家工号')['月2净锁单'].sum().rename('_剔2'))
        perf = full[['专家工号', '专家姓名', '月1净锁单_全量', '月2净锁单_全量']].copy()
        perf = perf.merge(m1, on='专家工号', how='left')
        perf = perf.merge(m2, on='专家工号', how='left')
        perf['月1净锁单'] = perf['月1净锁单_全量'] - perf['_剔1'].fillna(0)
        perf['月2净锁单'] = perf['月2净锁单_全量'] - perf['_剔2'].fillna(0)
        perf = perf[['专家工号', '专家姓名', '月1净锁单', '月2净锁单',
                     '月1净锁单_全量', '月2净锁单_全量']]

    return perf.reset_index(drop=True)


def load_manager_data(path: str) -> pd.DataFrame:
    """
    读取主管底表，过滤汇总行（工号不是纯数字的行）。
    返回原始月度数据，供 rules 计算季度达成率。
    """
    df = pd.read_excel(path, sheet_name='主管底表', header=0)
    df['工号'] = df['工号'].apply(normalize_id)
    df = df[df['工号'].notna()].copy()
    return df.reset_index(drop=True)


QUARTER_MONTHS = {
    'Q1': ['1月', '2月', '3月'],
    'Q2': ['4月', '5月', '6月'],
    'Q3': ['7月', '8月', '9月'],
    'Q4': ['10月', '11月', '12月'],
}

def find_manager_cols(columns, quarter: str, year_str: str):
    """
    从主管底表列名中找到对应季度的达成列和目标列。
    同时匹配两位年（如"26年"）和四位年（如"2026年"）。
    返回 (achieve_cols, target_cols)
    """
    year_str = str(year_str).strip()
    if len(year_str) == 4 and year_str.startswith('20'):
        year_str = year_str[2:]  # 自动规范化 "2026" → "26"
    months = QUARTER_MONTHS.get(quarter, [])
    year_prefixes = [f'{year_str}年', f'20{year_str}年']

    achieve_cols, target_cols = [], []
    for col in columns:
        col_str = str(col)
        for m in months:
            if any(f'{yp}{m}' in col_str for yp in year_prefixes):
                if '达成' in col_str and '率' not in col_str:
                    achieve_cols.append(col)
                elif '目标' in col_str:
                    target_cols.append(col)
    return achieve_cols, target_cols
