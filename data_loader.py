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


def load_roster(path: str) -> pd.DataFrame:
    """读取花名册，返回在职人员，工号标准化为字符串。"""
    df = pd.read_excel(path, sheet_name='花名册最新', header=0)
    df = df[df['雇佣状态'] == '在职'].copy()
    df['工号'] = df['工号'].apply(normalize_id)
    df = df[df['工号'].notna()].reset_index(drop=True)
    return df


def load_expert_orders(path: str) -> pd.DataFrame:
    """
    读取专家底表，只取车系='全部'的合计行。
    净锁单在 header=1 后为 '净锁单'（第一个月）和 '净锁单.1'（第二个月）。
    返回列: 专家工号, 专家姓名, 月1净锁单, 月2净锁单
    """
    df = pd.read_excel(path, sheet_name='专家底表', header=1)
    df = df[df['车系'] == '全部'].copy()
    df = df.rename(columns={'净锁单': '月1净锁单', '净锁单.1': '月2净锁单'})
    df['专家工号'] = df['专家工号'].apply(normalize_id)
    df = df[df['专家工号'].notna()]
    result = df[['专家工号', '专家姓名', '月1净锁单', '月2净锁单']].copy()
    result['月1净锁单'] = pd.to_numeric(result['月1净锁单'], errors='coerce').fillna(0)
    result['月2净锁单'] = pd.to_numeric(result['月2净锁单'], errors='coerce').fillna(0)
    return result.reset_index(drop=True)


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
