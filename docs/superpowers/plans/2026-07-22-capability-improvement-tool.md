# 能力提升计划识别工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个 Streamlit 本地工具，导入花名册+专家定单底表+主管底表，自动计算排名、识别进入能力提升计划的人员，输出与现有公示表格式完全一致的 Excel 文件，并自动记录历史以区分"未出营"和"新进入"。

**Architecture:** 四个模块各司其职：data_loader 负责读取和标准化输入 Excel；rules 负责排名计算和条件判断；history 负责本地历史持久化；exporter 负责生成格式化输出 Excel。app.py 是 Streamlit 入口，只做 UI 编排。

**Tech Stack:** Python 3.8+, Streamlit, pandas, openpyxl, pytest

---

## 文件结构

```
D:/dingyuan/Documents/能力提升计划工具/
├── app.py              # Streamlit UI 入口
├── data_loader.py      # Excel 读取 + 工号标准化 + find_manager_cols
├── rules.py            # 排名计算 + 能力提升识别逻辑
├── history.py          # history.json 读写
├── exporter.py         # 输出 Excel 格式化
├── requirements.txt    # 依赖声明
├── history.json        # 自动生成，记录历史名单
└── tests/
    ├── test_rules.py
    ├── test_data_loader.py
    └── test_history.py
```

---

## Task 1: 环境准备与 git 初始化

**Files:**
- Create: `D:/dingyuan/Documents/能力提升计划工具/requirements.txt`

- [ ] **Step 1: 写 requirements.txt**

```
streamlit>=1.30.0
pandas>=2.0.0
openpyxl>=3.1.0
pytest>=7.0.0
```

- [ ] **Step 2: 初始化 git 仓库**

```bash
cd "D:/dingyuan/Documents/能力提升计划工具"
git init
git add requirements.txt
git commit -m "chore: init project"
```

- [ ] **Step 3: 安装依赖**

```bash
pip install -r requirements.txt
```

Expected: 所有包安装成功，无报错

- [ ] **Step 4: 验证安装**

```bash
python -c "import streamlit, pandas, openpyxl; print('OK')"
```

Expected: 输出 `OK`

---

## Task 2: data_loader — 工号标准化、数据加载、find_manager_cols

**Files:**
- Create: `data_loader.py`
- Create: `tests/test_data_loader.py`

工号在不同表里格式不一致（带前导零 vs 不带、float vs string）。本模块统一标准化为纯数字字符串（去前导零）。`find_manager_cols` 也放在此模块，便于测试。

- [ ] **Step 1: 写 test_data_loader.py 中工号标准化测试**

```python
# tests/test_data_loader.py
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from data_loader import normalize_id

def test_normalize_strips_leading_zeros():
    assert normalize_id("084300") == "84300"

def test_normalize_handles_float():
    assert normalize_id(84300.0) == "84300"

def test_normalize_handles_int():
    assert normalize_id(84300) == "84300"

def test_normalize_handles_nan():
    assert normalize_id(float('nan')) is None

def test_normalize_handles_none():
    assert normalize_id(None) is None

def test_normalize_handles_non_numeric_string():
    assert normalize_id("区域达成") is None
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd "D:/dingyuan/Documents/能力提升计划工具"
python -m pytest tests/test_data_loader.py -v
```

Expected: FAILED - data_loader module not found

- [ ] **Step 3: 实现 data_loader.py — normalize_id**

```python
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
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
python -m pytest tests/test_data_loader.py -v
```

Expected: 所有 normalize_id 测试 PASSED

- [ ] **Step 5: 写 load_roster 测试**

```python
# 追加到 tests/test_data_loader.py
import pandas as pd
from data_loader import load_roster

def test_load_roster_returns_required_columns():
    df = pd.DataFrame({
        '工号': [84300, 52457],
        '员工姓名': ['丁琦', '付俊'],
        '雇佣状态': ['在职', '在职'],
        '三级部门': ['湖北二部', '湖北二部'],
        '四级部门': ['黄石港万达零售展厅', '武汉江夏永旺零售中心'],
        '新岗位名称': ['产品专家', '产品专家'],
        '新职级': [13, 13],
        '入职日期': pd.to_datetime(['2024-01-01', '2023-06-01']),
        '二级部门': ['湖北战区', '湖北战区'],
    })
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='花名册最新')
    result = load_roster(path)
    os.unlink(path)
    assert '工号' in result.columns
    assert '四级部门' in result.columns
    assert result['工号'].iloc[0] == '84300'

def test_load_roster_filters_active_only():
    df = pd.DataFrame({
        '工号': [1, 2],
        '员工姓名': ['A', 'B'],
        '雇佣状态': ['在职', '离职'],
        '三级部门': ['X', 'X'],
        '四级部门': ['门店A', '门店A'],
        '新岗位名称': ['产品专家', '产品专家'],
        '新职级': [13, 13],
        '入职日期': pd.to_datetime(['2024-01-01', '2023-01-01']),
        '二级部门': ['战区', '战区'],
    })
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='花名册最新')
    result = load_roster(path)
    os.unlink(path)
    assert len(result) == 1
    assert result.iloc[0]['员工姓名'] == 'A'
```

- [ ] **Step 6: 运行测试，确认失败**

```bash
python -m pytest tests/test_data_loader.py::test_load_roster_returns_required_columns -v
```

Expected: FAILED

- [ ] **Step 7: 实现 load_roster**

```python
# 追加到 data_loader.py

def load_roster(path: str) -> pd.DataFrame:
    """读取花名册，返回在职人员，工号标准化为字符串。"""
    df = pd.read_excel(path, sheet_name='花名册最新', header=0)
    df = df[df['雇佣状态'] == '在职'].copy()
    df['工号'] = df['工号'].apply(normalize_id)
    df = df[df['工号'].notna()].reset_index(drop=True)
    return df
```

- [ ] **Step 8: 运行测试**

```bash
python -m pytest tests/test_data_loader.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 9: 写 load_expert_orders 测试**

```python
# 追加到 tests/test_data_loader.py
from data_loader import load_expert_orders

def test_load_expert_orders_returns_aggregate_rows():
    """专家底表只取车系='全部'的行，即每人的合计行。"""
    import tempfile, os
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '专家底表'
    # Row1: 大分组标题（header=1 时跳过）
    ws.append(['区域','销售部','门店','专家姓名','专家工号','车系',
               '当前周期','','','','','','','','','','','','','','',
               '对比周期','','','','','','','','','','','','','','',
               '环比','','','',''])
    # Row2: 列名（这是 header=1 读到的行）
    ws.append(['区域','销售部','门店','专家姓名','专家工号','车系',
               '跟进量','目标','达成率%','线下线索','目标','达成率%',
               '新增排程','目标','达成率%','试驾量','目标','达成率%',
               '净锁单','目标','达成率%',
               '跟进量','目标','达成率%','线下线索','目标','达成率%',
               '新增排程','目标','达成率%','试驾量','目标','达成率%',
               '净锁单','目标','达成率%',
               '跟进量','线下线索','新增排程','试驾量','净锁单'])
    # 全部行
    ws.append(['湖北战区','','','丁琦','084300','全部',
               100,'','',10,'','',5,'','',3,'','',3,'','',
               90,'','',8,'','',4,'','',2,'','',8,'','','','','','',''])
    # 车系明细行（应被过滤）
    ws.append(['湖北战区','','','丁琦','084300','W04',
               50,'','',5,'','',3,'','',2,'','',3,'','',
               40,'','',4,'','',2,'','',1,'','',3,'','','','','','',''])
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    wb.save(path)
    result = load_expert_orders(path)
    os.unlink(path)
    assert len(result) == 1
    assert result.iloc[0]['专家工号'] == '84300'
    assert result.iloc[0]['月1净锁单'] == 3
    assert result.iloc[0]['月2净锁单'] == 8
```

- [ ] **Step 10: 运行测试，确认失败**

```bash
python -m pytest tests/test_data_loader.py::test_load_expert_orders_returns_aggregate_rows -v
```

Expected: FAILED

- [ ] **Step 11: 实现 load_expert_orders**

```python
# 追加到 data_loader.py

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
```

- [ ] **Step 12: 运行测试**

```bash
python -m pytest tests/test_data_loader.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 13: 写 load_manager_data 测试**

```python
# 追加到 tests/test_data_loader.py
from data_loader import load_manager_data

def test_load_manager_data_returns_clean_rows():
    df = pd.DataFrame({
        '工号': ['008210', '009846', '区域达成（绩效口径）'],
        '姓名': ['邓明君', '左飞', None],
        '26年5月达成': [20, 82, None],
        '26年5月目标（不含L9）': [20, 54, None],
        '26年6月达成': [15, 55, None],
        '26年6月目标（不含L8）': [15, 55, None],
    })
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='主管底表')
    result = load_manager_data(path)
    os.unlink(path)
    assert len(result) == 2
    assert result.iloc[0]['工号'] == '8210'
```

- [ ] **Step 14: 运行测试，确认失败**

```bash
python -m pytest tests/test_data_loader.py::test_load_manager_data_returns_clean_rows -v
```

Expected: FAILED

- [ ] **Step 15: 实现 load_manager_data**

```python
# 追加到 data_loader.py

def load_manager_data(path: str) -> pd.DataFrame:
    """
    读取主管底表，过滤汇总行（工号不是纯数字的行）。
    返回原始月度数据，供 rules 计算季度达成率。
    """
    df = pd.read_excel(path, sheet_name='主管底表', header=0)
    df['工号'] = df['工号'].apply(normalize_id)
    df = df[df['工号'].notna()].copy()
    return df.reset_index(drop=True)
```

- [ ] **Step 16: 写 find_manager_cols 测试**

```python
# 追加到 tests/test_data_loader.py
from data_loader import find_manager_cols

def test_find_manager_cols_q2_two_digit_year():
    cols = ['工号', '姓名', '26年4月达成', '26年4月目标（含i6）',
            '26年5月达成', '26年5月目标（不含L9）',
            '26年6月达成', '26年6月目标（不含L8）', '其他列']
    achieve, target = find_manager_cols(cols, 'Q2', '26')
    assert '26年4月达成' in achieve
    assert '26年5月达成' in achieve
    assert '26年6月达成' in achieve
    assert '26年4月目标（含i6）' in target
    assert '26年5月目标（不含L9）' in target
    assert '26年6月目标（不含L8）' in target

def test_find_manager_cols_q1_four_digit_year():
    cols = ['工号', '2026年1月达成', '2026年1月目标', '2026年2月达成',
            '2026年2月目标', '2026年3月达成', '2026年3月目标']
    achieve, target = find_manager_cols(cols, 'Q1', '26')
    assert len(achieve) == 3
    assert len(target) == 3

def test_find_manager_cols_returns_empty_on_mismatch():
    cols = ['工号', '姓名', '随便列']
    achieve, target = find_manager_cols(cols, 'Q2', '26')
    assert achieve == []
    assert target == []
```

- [ ] **Step 17: 运行测试，确认失败**

```bash
python -m pytest tests/test_data_loader.py::test_find_manager_cols_q2_two_digit_year -v
```

Expected: FAILED

- [ ] **Step 18: 实现 find_manager_cols**

```python
# 追加到 data_loader.py

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
```

- [ ] **Step 19: 运行全部 data_loader 测试**

```bash
python -m pytest tests/test_data_loader.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 20: commit**

```bash
git add data_loader.py tests/test_data_loader.py
git commit -m "feat: data_loader - normalize id, load roster/expert/manager data, find_manager_cols"
```

---

## Task 3: rules — 专家排名与识别

**Files:**
- Create: `rules.py`
- Create: `tests/test_rules.py`

关键设计决策（已与业务方确认）：
- 专家识别范围用 `岗位类别 == '产品专家'` 圈定，包含产品专员/产品专家/高级产品专家（不区分职级）
- 后15%阈值分母 = 参与排名的人数（非新员工），不含新员工
- 高级产品专家省区后50%的分母 = 省区内所有 `岗位类别=='产品专家'` 的参与人数（含产品专员/专家/高级专家），不含新员工
- 高级产品专家的判断条件：`新岗位名称 == '高级产品专家'`

- [ ] **Step 1: 写 identify_experts 测试**

```python
# tests/test_rules.py
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import pandas as pd
from rules import identify_experts

def _make_expert_df(rows):
    """rows: list of (工号, 姓名, 门店, 三级部门, 职级, 入职日期_str, 月1, 月2)"""
    return pd.DataFrame([{
        '工号': r[0], '员工姓名': r[1], '四级部门': r[2], '三级部门': r[3],
        '新职级': r[4], '入职日期': pd.Timestamp(r[5]),
        '月1净锁单': r[6], '月2净锁单': r[7],
        '新岗位名称': '高级产品专家' if r[4] == 14 else '产品专家',
        '岗位类别': '产品专家',  # 所有专家岗统一用此字段圈定
        '二级部门': '湖北战区',
    } for r in rows])

def test_bottom_15_percent_store_triggers():
    """门店5人，后15%=round(5*0.15)=1人，定单最少的应触发。"""
    df = _make_expert_df([
        ('1', 'A', '门店X', '省区A', 13, '2020-01-01', 10, 10),
        ('2', 'B', '门店X', '省区A', 13, '2020-01-01', 8, 8),
        ('3', 'C', '门店X', '省区A', 13, '2020-01-01', 6, 6),
        ('4', 'D', '门店X', '省区A', 13, '2020-01-01', 4, 4),
        ('5', 'E', '门店X', '省区A', 13, '2020-01-01', 2, 2),
    ])
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    triggered = result[result['触发识别'] == True]
    assert len(triggered) == 1
    assert triggered.iloc[0]['工号'] == '5'

def test_bottom_15_percent_with_ties():
    """门店5人，后2人定单相同，rank(method='min')使两人共享排名1，应都触发。"""
    df = _make_expert_df([
        ('1', 'A', '门店X', '省区A', 13, '2020-01-01', 10, 10),
        ('2', 'B', '门店X', '省区A', 13, '2020-01-01', 8, 8),
        ('3', 'C', '门店X', '省区A', 13, '2020-01-01', 6, 6),
        ('4', 'D', '门店X', '省区A', 13, '2020-01-01', 2, 2),  # 并列最低
        ('5', 'E', '门店X', '省区A', 13, '2020-01-01', 2, 2),  # 并列最低
    ])
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    triggered = result[result['触发识别'] == True]
    # 两人并列 rank=1，阈值=1，两人都应触发
    triggered_ids = set(triggered['工号'].tolist())
    assert '4' in triggered_ids
    assert '5' in triggered_ids

def test_store_3_or_less_uses_province_rank():
    """门店仅3人（≤3），改用省区后15%。省区共10人（剔新员工），后15%=2人。"""
    rows = []
    for i in range(7):
        rows.append((str(i+1), f'P{i+1}', '门店Y', '省区A', 13, '2020-01-01', (i+1)*2, (i+1)*2))
    rows += [
        ('8', 'S1', '门店X', '省区A', 13, '2020-01-01', 3, 3),
        ('9', 'S2', '门店X', '省区A', 13, '2020-01-01', 2, 2),
        ('10', 'S3', '门店X', '省区A', 13, '2020-01-01', 1, 1),
    ]
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    triggered = result[result['触发识别'] == True]
    # 省区10人后15%=round(10*0.15)=2人，工号9(合计4)和10(合计2)最低
    triggered_ids = set(triggered['工号'].tolist())
    assert '10' in triggered_ids
    assert '9' in triggered_ids

def test_senior_expert_province_50_percent_triggers():
    """
    L14高级专家省区后50%：分母=省区所有专家岗(L13+L14)参与排名人数。
    10人省区（3个L14，7个L13），后50%=5人。
    定单最低的L14应触发。
    """
    rows = []
    # L14: 定单 20, 15, 5（第3个最低，应触发）
    rows.append(('l14_1', 'High14', '门店A', '省区A', 14, '2020-01-01', 20, 20))
    rows.append(('l14_2', 'Mid14', '门店B', '省区A', 14, '2020-01-01', 15, 15))
    rows.append(('l14_3', 'Low14', '门店C', '省区A', 14, '2020-01-01', 2, 3))
    # L13: 定单 30,25,22,18,12,8,4
    for i, v in enumerate([30, 25, 22, 18, 12, 8, 4]):
        rows.append((f'l13_{i}', f'L13_{i}', f'门店{i+4}', '省区A', 13, '2020-01-01', v, v))
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    # 省区10人，后50%=5人。定单从低到高：4,5,8,12,18…
    # Low14(合计5)排名2，低于阈值5，应触发
    low14_row = result[result['工号'] == 'l14_3']
    assert low14_row.iloc[0]['触发识别'] == True
    # High14(合计40)排名10，高于阈值5，不触发
    high14_row = result[result['工号'] == 'l14_1']
    assert high14_row.iloc[0]['触发识别'] == False

def test_new_employee_excluded():
    """入职3个月内的新员工不参与排名，不触发。"""
    rows = [
        ('1', 'Old', '门店X', '省区A', 13, '2020-01-01', 1, 1),
        ('2', 'New', '门店X', '省区A', 13, '2026-05-01', 0, 0),  # 新员工
        ('3', 'C', '门店X', '省区A', 13, '2020-01-01', 5, 5),
        ('4', 'D', '门店X', '省区A', 13, '2020-01-01', 8, 8),
        ('5', 'E', '门店X', '省区A', 13, '2020-01-01', 10, 10),
    ]
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    new_row = result[result['工号'] == '2']
    assert new_row.iloc[0]['触发识别'] == False
    assert new_row.iloc[0]['是否新员工'] == True

def test_all_new_employees_no_crash():
    """全门店都是新员工时，不抛出异常，无人触发。"""
    rows = [
        ('1', 'New1', '门店X', '省区A', 13, '2026-05-01', 0, 0),
        ('2', 'New2', '门店X', '省区A', 13, '2026-06-01', 0, 0),
    ]
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    assert result['触发识别'].sum() == 0
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
python -m pytest tests/test_rules.py -v
```

Expected: FAILED - rules module not found

- [ ] **Step 3: 实现 identify_experts**

```python
# rules.py
import pandas as pd
import math

def _is_new_employee(hire_date, period_end: pd.Timestamp) -> bool:
    """入职日期距识别周期结束日 < 3个月视为新员工。"""
    if pd.isna(hire_date):
        return False
    cutoff = period_end - pd.DateOffset(months=3)
    return pd.Timestamp(hire_date) > cutoff


def identify_experts(df: pd.DataFrame, period_end: pd.Timestamp) -> pd.DataFrame:
    """
    识别产品专家/产品专员/高级产品专家是否进入能力提升。

    df 必须包含列: 工号, 员工姓名, 四级部门(门店), 三级部门(省区/城市部),
                   新职级, 入职日期, 月1净锁单, 月2净锁单, 新岗位名称, 岗位类别, 二级部门
    period_end: 识别周期结束日期

    后15%阈值分母 = 参与排名人数（非新员工），不含新员工。
    高级产品专家省区后50%分母 = 省区所有岗位类别=='产品专家'的参与人数（含专员/专家/高级）。
    """
    result = df.copy().reset_index(drop=True)
    result['双月合计'] = result['月1净锁单'] + result['月2净锁单']
    result['是否新员工'] = result['入职日期'].apply(lambda d: _is_new_employee(d, period_end))

    # 参与排名的子集（非新员工）
    active_mask = ~result['是否新员工']
    active = result[active_mask].copy()

    if len(active) == 0:
        result['省区排名'] = None
        result['省区参与人数'] = 0
        result['门店排名'] = None
        result['门店参与人数'] = 0
        result['门店总人数'] = result.groupby('四级部门')['工号'].transform('count')
        result['触发识别'] = False
        result['识别原因'] = ''
        return result

    # 省区排名（所有专家岗 L13+L14 合计，仅非新员工）
    province_rank = active.groupby('三级部门')['双月合计'].rank(method='min', ascending=True)
    province_count = active.groupby('三级部门')['双月合计'].transform('count')

    result['省区排名'] = None
    result['省区参与人数'] = 0
    result.loc[active_mask, '省区排名'] = province_rank.values
    result.loc[active_mask, '省区参与人数'] = province_count.values

    # 门店总人数（含新员工，用于判断是否≤3）
    result['门店总人数'] = result.groupby('四级部门')['工号'].transform('count')

    # 门店排名（仅非新员工参与）
    store_rank = active.groupby('四级部门')['双月合计'].rank(method='min', ascending=True)
    store_count = active.groupby('四级部门')['双月合计'].transform('count')

    result['门店排名'] = None
    result['门店参与人数'] = 0
    result.loc[active_mask, '门店排名'] = store_rank.values
    result.loc[active_mask, '门店参与人数'] = store_count.values

    triggered = []
    reasons = []

    for _, row in result.iterrows():
        if row['是否新员工']:
            triggered.append(False)
            reasons.append('')
            continue

        level = row['新职级']
        store_total = row['门店总人数']
        store_part = row['门店参与人数']
        province_part = row['省区参与人数']
        store_rank_val = row['门店排名']
        province_rank_val = row['省区排名']

        hit = False
        reason = ''

        # 门店≤3时改用省区后15%（阈值基于省区参与人数）
        if store_total <= 3:
            thresh = max(1, round(province_part * 0.15))
            if province_rank_val is not None and not (isinstance(province_rank_val, float) and math.isnan(province_rank_val)):
                if province_rank_val <= thresh:
                    hit = True
                    reason = '双月专家定单量排名在省区后15%'
        else:
            # 门店后15%（阈值基于门店参与人数）
            thresh = max(1, round(store_part * 0.15))
            if store_rank_val is not None and not (isinstance(store_rank_val, float) and math.isnan(store_rank_val)):
                if store_rank_val <= thresh:
                    hit = True
                    reason = '双月专家定单量排名在所在门店的后15%'

        # L14额外判断：省区所有专家岗后50%（分母=省区所有岗位类别==产品专家的参与人数，含专员/专家/高级）
        if row.get('新岗位名称') == '高级产品专家' and not hit:
            thresh50 = max(1, round(province_part * 0.50))
            if province_rank_val is not None and not (isinstance(province_rank_val, float) and math.isnan(province_rank_val)):
                if province_rank_val <= thresh50:
                    hit = True
                    reason = '双月高级专家定单量排名在省区专家岗位后50%'

        triggered.append(hit)
        reasons.append(reason)

    result['触发识别'] = triggered
    result['识别原因'] = reasons
    return result
```

- [ ] **Step 4: 运行测试**

```bash
python -m pytest tests/test_rules.py -v
```

Expected: 所有 identify_experts 测试 PASSED

- [ ] **Step 5: 写 identify_managers 测试**

```python
# 追加到 tests/test_rules.py
from rules import identify_managers

def _make_manager_df(rows, month_cols):
    records = []
    for r in rows:
        rec = {'工号': r[0], '姓名': r[1], '四级部门': r[2], '三级部门': r[3],
               '入职日期': pd.Timestamp(r[4]), '二级部门': '湖北战区',
               '新岗位名称': '零售主管', '新职级': 14}
        for j, (col_d, col_t) in enumerate(month_cols):
            rec[col_d] = r[5 + j*2]
            rec[col_t] = r[5 + j*2 + 1]
        records.append(rec)
    return pd.DataFrame(records)

def test_manager_bottom_10_percent_triggers():
    """10名主管，后10%=1人，达成率最低的应触发。"""
    month_cols = [
        ('26年4月达成', '26年4月目标'),
        ('26年5月达成', '26年5月目标'),
        ('26年6月达成', '26年6月目标'),
    ]
    rows = []
    for i in range(10):
        achieve = (i + 1) * 10
        rows.append((str(i+1), f'M{i+1}', f'门店{i+1}', '省区A', '2020-01-01',
                     achieve, 100, achieve, 100, achieve, 100))
    df = _make_manager_df(rows, month_cols)
    result = identify_managers(
        df,
        ['26年4月达成', '26年5月达成', '26年6月达成'],
        ['26年4月目标', '26年5月目标', '26年6月目标'],
        pd.Timestamp('2026-06-30')
    )
    triggered = result[result['触发识别'] == True]
    assert len(triggered) == 1
    assert triggered.iloc[0]['工号'] == '1'

def test_manager_new_employee_excluded():
    month_cols = [
        ('26年4月达成', '26年4月目标'),
        ('26年5月达成', '26年5月目标'),
        ('26年6月达成', '26年6月目标'),
    ]
    rows = [
        ('1', 'Old', '门店A', '省区A', '2020-01-01', 5, 100, 5, 100, 5, 100),
        ('2', 'New', '门店B', '省区A', '2026-05-01', 0, 100, 0, 100, 0, 100),
    ]
    df = _make_manager_df(rows, month_cols)
    result = identify_managers(
        df,
        ['26年4月达成', '26年5月达成', '26年6月达成'],
        ['26年4月目标', '26年5月目标', '26年6月目标'],
        pd.Timestamp('2026-06-30')
    )
    new_row = result[result['工号'] == '2']
    assert new_row.iloc[0]['触发识别'] == False
    assert new_row.iloc[0]['是否新员工'] == True
```

- [ ] **Step 6: 运行测试，确认失败**

```bash
python -m pytest tests/test_rules.py::test_manager_bottom_10_percent_triggers -v
```

Expected: FAILED

- [ ] **Step 7: 实现 identify_managers**

```python
# 追加到 rules.py

def identify_managers(df: pd.DataFrame,
                      achieve_cols: list,
                      target_cols: list,
                      period_end: pd.Timestamp) -> pd.DataFrame:
    """
    识别零售主管是否进入能力提升。
    achieve_cols: 季度对应月份达成列名列表
    target_cols:  季度对应月份目标列名列表
    period_end:   识别周期结束日
    """
    result = df.copy().reset_index(drop=True)
    result['是否新员工'] = result['入职日期'].apply(lambda d: _is_new_employee(d, period_end))

    result['季度达成'] = result[achieve_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    result['季度目标'] = result[target_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    result['季度达成率'] = result.apply(
        lambda r: r['季度达成'] / r['季度目标'] if r['季度目标'] > 0 else 0, axis=1
    )

    active_mask = ~result['是否新员工']
    active = result[active_mask].copy()

    if len(active) == 0:
        result['省区排名'] = None
        result['省区参与人数'] = 0
        result['触发识别'] = False
        result['识别原因'] = ''
        return result

    province_rank = active.groupby('三级部门')['季度达成率'].rank(method='min', ascending=True)
    province_count = active.groupby('三级部门')['季度达成率'].transform('count')

    result['省区排名'] = None
    result['省区参与人数'] = 0
    result.loc[active_mask, '省区排名'] = province_rank.values
    result.loc[active_mask, '省区参与人数'] = province_count.values

    triggered = []
    reasons = []
    for _, row in result.iterrows():
        if row['是否新员工']:
            triggered.append(False)
            reasons.append('')
            continue
        count = row['省区参与人数']
        thresh = max(1, round(count * 0.10))
        rank_val = row['省区排名']
        if rank_val is not None and not (isinstance(rank_val, float) and math.isnan(rank_val)):
            if rank_val <= thresh:
                triggered.append(True)
                reasons.append('季度小组定单达成率排名后10%')
                continue
        triggered.append(False)
        reasons.append('')

    result['触发识别'] = triggered
    result['识别原因'] = reasons
    return result
```

- [ ] **Step 8: 运行全部 rules 测试**

```bash
python -m pytest tests/test_rules.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 9: commit**

```bash
git add rules.py tests/test_rules.py
git commit -m "feat: rules - expert/manager ranking and identification logic"
```

---

## Task 4: history — 历史持久化

**Files:**
- Create: `history.py`
- Create: `tests/test_history.py`

- [ ] **Step 1: 写 history 测试**

```python
# tests/test_history.py
import sys, os, json, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from history import save_period, load_last_period, get_previous_ids

def _tmp_path():
    f = tempfile.NamedTemporaryFile(suffix='.json', delete=False)
    f.close()
    os.unlink(f.name)
    return f.name

def test_save_and_load_round_trip():
    path = _tmp_path()
    save_period(path, '2026-05/06', ['1001', '1002', '1003'])
    last = load_last_period(path)
    assert last['period'] == '2026-05/06'
    assert set(last['ids']) == {'1001', '1002', '1003'}
    os.unlink(path)

def test_load_returns_none_when_no_history():
    path = _tmp_path()
    last = load_last_period(path)
    assert last is None

def test_get_previous_ids_returns_set():
    path = _tmp_path()
    save_period(path, '2026-05/06', ['1001', '1002'])
    ids = get_previous_ids(path)
    assert ids == {'1001', '1002'}
    os.unlink(path)

def test_get_previous_ids_returns_empty_when_no_history():
    path = _tmp_path()
    ids = get_previous_ids(path)
    assert ids == set()

def test_save_appends_history():
    path = _tmp_path()
    save_period(path, '2026-03/04', ['900'])
    save_period(path, '2026-05/06', ['1001', '1002'])
    with open(path) as f:
        data = json.load(f)
    assert len(data['history']) == 2
    assert data['history'][-1]['period'] == '2026-05/06'
    os.unlink(path)
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
python -m pytest tests/test_history.py -v
```

Expected: FAILED

- [ ] **Step 3: 实现 history.py**

```python
# history.py
import json, os
from typing import List, Optional, Set

def _load_data(path: str) -> dict:
    if not os.path.exists(path):
        return {'history': []}
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_period(path: str, period: str, ids: List[str]) -> None:
    data = _load_data(path)
    data['history'].append({'period': period, 'ids': list(ids)})
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_last_period(path: str) -> Optional[dict]:
    data = _load_data(path)
    if not data['history']:
        return None
    return data['history'][-1]

def get_previous_ids(path: str) -> Set[str]:
    last = load_last_period(path)
    if last is None:
        return set()
    return set(last['ids'])
```

- [ ] **Step 4: 运行测试**

```bash
python -m pytest tests/test_history.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 5: commit**

```bash
git add history.py tests/test_history.py
git commit -m "feat: history - persist and load period records"
```

---

## Task 5: exporter — 输出 Excel

**Files:**
- Create: `exporter.py`
- Create: `tests/test_exporter.py`

- [ ] **Step 1: 写 exporter 测试**

```python
# tests/test_exporter.py
import sys, os, io
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import openpyxl
from exporter import build_output_excel

def _sample_records():
    return [
        {'工号': '1001', '员工姓名': '张三', '二级部门': '湖北战区',
         '三级部门': '湖北一部', '四级部门': '武汉龙阳大道零售中心',
         '新岗位名称': '产品专家', '新职级': 13,
         '识别原因': '双月专家定单量排名在所在门店的后15%'},
        {'工号': '1002', '员工姓名': '李四', '二级部门': '湖北战区',
         '三级部门': '湖北三部', '四级部门': '恩施金桂大道汽车城零售中心',
         '新岗位名称': '高级产品专家', '新职级': 14,
         '识别原因': '双月高级专家定单量排名在省区专家岗位后50%'},
    ]

def test_build_output_excel_returns_bytesio():
    buf = build_output_excel('7月', '7月-8月', [_sample_records()[0]], [_sample_records()[1]])
    assert isinstance(buf, io.BytesIO)

def test_output_excel_has_correct_sheet():
    buf = build_output_excel('7月', '7月-8月', [], _sample_records())
    wb = openpyxl.load_workbook(buf)
    assert '7月能力提升计划最终名单-公示' in wb.sheetnames

def test_output_contains_names():
    buf = build_output_excel('7月', '7月-8月', [], _sample_records())
    wb = openpyxl.load_workbook(buf)
    ws = wb['7月能力提升计划最终名单-公示']
    flat = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
    assert '张三' in flat
    assert '李四' in flat

def test_not_graduated_shows_no():
    buf = build_output_excel('7月', '7月-8月', _sample_records(), [])
    wb = openpyxl.load_workbook(buf)
    ws = wb['7月能力提升计划最终名单-公示']
    flat = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
    assert '否' in flat

def test_new_entries_shows_yes_and_period():
    buf = build_output_excel('7月', '7月-8月', [], _sample_records())
    wb = openpyxl.load_workbook(buf)
    ws = wb['7月能力提升计划最终名单-公示']
    flat = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
    assert '是' in flat
    assert '7月-8月' in flat
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
python -m pytest tests/test_exporter.py -v
```

Expected: FAILED

- [ ] **Step 3: 实现 exporter.py**

```python
# exporter.py
import io
from typing import List, Dict
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def _write_row(ws, row_idx, values):
    for col, val in enumerate(values, 1):
        ws.cell(row=row_idx, column=col, value=val)


def build_output_excel(period_label: str,
                       improve_period: str,
                       not_graduated: List[Dict],
                       new_entries: List[Dict]) -> io.BytesIO:
    """
    生成公示表 Excel，格式与现有公示表一致。
    period_label: 如 '7月'
    improve_period: 如 '7月-8月'
    not_graduated: 未出营人员记录列表
    new_entries: 新进入人员记录列表
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'{period_label}能力提升计划最终名单-公示'

    row = 1

    # === 第一部分：未出营名单 ===
    ws.cell(row=row, column=1, value=f'{period_label}未出营产品专家名单').font = Font(bold=True, size=12)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    row += 1

    headers_grad = ['序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级', '是否出营', '未出营原因']
    _write_row(ws, row, headers_grad)
    for cell in ws[row]:
        cell.font = Font(bold=True)
    row += 1

    for i, rec in enumerate(not_graduated, 1):
        _write_row(ws, row, [
            i, rec.get('二级部门', ''), rec.get('三级部门', ''), rec.get('四级部门', ''),
            rec.get('工号', ''), rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''), rec.get('新职级', ''),
            '否', rec.get('识别原因', ''),
        ])
        row += 1

    row += 1  # 空行

    # === 第二部分：进入能力提升计划名单 ===
    ws.cell(row=row, column=1, value=f'{period_label}进入能力提升计划人员名单').font = Font(bold=True, size=12)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    row += 1

    headers_entry = ['序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级',
                     '是否进入能力提升计划', '识别进入能力提升考核期原因', '改进周期', '备注']
    _write_row(ws, row, headers_entry)
    for cell in ws[row]:
        cell.font = Font(bold=True)
    row += 1

    for i, rec in enumerate(new_entries, 1):
        _write_row(ws, row, [
            i, rec.get('二级部门', ''), rec.get('三级部门', ''), rec.get('四级部门', ''),
            rec.get('工号', ''), rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''), rec.get('新职级', ''),
            '是', rec.get('识别原因', ''), improve_period, '',
        ])
        row += 1

    # 列宽自适应
    for col in ws.columns:
        max_len = max((len(str(cell.value or '')) for cell in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
```

- [ ] **Step 4: 运行测试**

```bash
python -m pytest tests/test_exporter.py -v
```

Expected: 所有测试 PASSED

- [ ] **Step 5: commit**

```bash
git add exporter.py tests/test_exporter.py
git commit -m "feat: exporter - generate formatted 公示表 Excel output"
```

---

## Task 6: app.py — Streamlit UI

**Files:**
- Create: `app.py`

注意：主管数据合并后需将 `姓名` rename 为 `员工姓名`，否则导出时姓名为空。

- [ ] **Step 1: 实现 app.py**

```python
# app.py
import streamlit as st
import pandas as pd
import os, tempfile, shutil
import openpyxl
from data_loader import load_roster, load_expert_orders, load_manager_data, find_manager_cols
from rules import identify_experts, identify_managers
from history import save_period, get_previous_ids
from exporter import build_output_excel

HISTORY_PATH = os.path.join(os.path.dirname(__file__), 'history.json')

st.set_page_config(page_title='能力提升计划识别工具', layout='wide')
st.title('能力提升计划识别工具')

# ── 侧边栏：参数输入 ──────────────────────────────────────
with st.sidebar:
    st.header('识别参数')

    st.subheader('专家识别（双月）')
    month1 = st.text_input('第一个月（如 5月）', value='5月')
    month2 = st.text_input('第二个月（如 6月）', value='6月')
    expert_period_end = st.date_input('识别周期结束日', value=pd.Timestamp('2026-06-30'))

    st.subheader('主管识别（季度）')
    quarter = st.selectbox('季度', ['Q1', 'Q2', 'Q3', 'Q4'], index=1)
    year_str = st.text_input('年份后两位（如 26）', value='26')

    st.subheader('输出参数')
    period_label = st.text_input('输出表期间标签（如 7月）', value='7月')
    improve_period = st.text_input('改进周期（如 7月-8月）', value='7月-8月')

# ── 主体：文件上传 ────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    roster_file = st.file_uploader('花名册（需含"花名册最新" Sheet）', type=['xlsx'])
with col2:
    expert_file = st.file_uploader('专家定单底表（需含"专家底表" Sheet）', type=['xlsx'])
with col3:
    manager_file = st.file_uploader('主管底表（需含"主管底表" Sheet）', type=['xlsx'])

# ── 生成按钮 ──────────────────────────────────────────────
if st.button('生成能力提升名单', type='primary'):
    if not (roster_file and expert_file and manager_file):
        st.error('请上传全部3个文件')
        st.stop()

    with st.spinner('计算中...'):
        tmp_dir = tempfile.mkdtemp()
        try:
            roster_path = os.path.join(tmp_dir, 'roster.xlsx')
            expert_path = os.path.join(tmp_dir, 'expert.xlsx')
            manager_path = os.path.join(tmp_dir, 'manager.xlsx')
            for src, dst in [(roster_file, roster_path), (expert_file, expert_path), (manager_file, manager_path)]:
                with open(dst, 'wb') as f:
                    f.write(src.read())

            try:
                roster = load_roster(roster_path)
            except Exception as e:
                st.error(f'花名册读取失败：{e}\n请确认文件包含"花名册最新" Sheet，且"雇佣状态"列存在。')
                st.stop()

            try:
                expert_orders = load_expert_orders(expert_path)
            except Exception as e:
                st.error(f'专家底表读取失败：{e}\n请确认文件包含"专家底表" Sheet。')
                st.stop()

            try:
                manager_raw = load_manager_data(manager_path)
            except Exception as e:
                st.error(f'主管底表读取失败：{e}\n请确认文件包含"主管底表" Sheet。')
                st.stop()

            # 合并专家数据（花名册补充门店/职级/入职日期），只取岗位类别=='产品专家'的人
            expert_merged = expert_orders.merge(
                roster[roster['岗位类别'] == '产品专家'][
                    ['工号', '员工姓名', '二级部门', '三级部门', '四级部门', '新岗位名称', '新职级', '入职日期', '岗位类别']
                ],
                left_on='专家工号', right_on='工号', how='inner'
            ).drop(columns=['专家工号'])

            # 合并主管数据，并将 '姓名' rename 为 '员工姓名'
            manager_merged = manager_raw.rename(columns={'姓名': '员工姓名'}).merge(
                roster[['工号', '二级部门', '三级部门', '四级部门', '新岗位名称', '新职级', '入职日期']],
                on='工号', how='inner'
            )

            # 识别专家
            period_end_ts = pd.Timestamp(expert_period_end)
            expert_result = identify_experts(expert_merged, period_end_ts)

            # 识别主管
            achieve_cols, target_cols = find_manager_cols(list(manager_merged.columns), quarter, year_str)
            if not achieve_cols or not target_cols:
                st.error(
                    f'未在主管底表中找到 {quarter} 对应的达成/目标列。\n'
                    f'请检查列名是否包含"{year_str}年X月"或"20{year_str}年X月"格式。'
                )
                st.stop()
            manager_result = identify_managers(manager_merged, achieve_cols, target_cols, period_end_ts)

            # 汇总触发人员
            expert_triggered = expert_result[expert_result['触发识别'] == True].to_dict('records')
            manager_triggered = manager_result[manager_result['触发识别'] == True].to_dict('records')
            all_triggered = expert_triggered + manager_triggered
            triggered_ids = {r['工号'] for r in all_triggered}

            # 区分未出营 vs 新进入
            prev_ids = get_previous_ids(HISTORY_PATH)
            not_graduated = [r for r in all_triggered if r['工号'] in prev_ids]
            new_entries = [r for r in all_triggered if r['工号'] not in prev_ids]

        finally:
            shutil.rmtree(tmp_dir)

    st.success(f'识别完成：未出营 {len(not_graduated)} 人，新进入 {len(new_entries)} 人')

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f'未出营名单（{len(not_graduated)} 人）')
        if not_graduated:
            import pandas as pd
            st.dataframe(pd.DataFrame(not_graduated)[['工号', '员工姓名', '四级部门', '新岗位名称', '新职级', '识别原因']])
        else:
            st.info('无未出营人员')

    with col_b:
        st.subheader(f'新进入名单（{len(new_entries)} 人）')
        if new_entries:
            import pandas as pd
            st.dataframe(pd.DataFrame(new_entries)[['工号', '员工姓名', '四级部门', '新岗位名称', '新职级', '识别原因']])
        else:
            st.info('无新进入人员')

    excel_buf = build_output_excel(period_label, improve_period, not_graduated, new_entries)
    st.download_button(
        label='下载公示表 Excel',
        data=excel_buf,
        file_name=f'{period_label}能力提升计划最终名单-公示.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    period_key = f'{month1}/{month2}'
    save_period(HISTORY_PATH, period_key, list(triggered_ids))
    st.caption(f'本期名单已保存至历史记录（周期：{period_key}）')

# ── 历史初始化（可选） ────────────────────────────────────
with st.expander('初始化历史记录（首次使用时导入上期名单）'):
    st.write('上传已有公示表 Excel，从中提取上期进入名单工号作为历史基准。')
    hist_file = st.file_uploader('上传历史公示表', type=['xlsx'], key='hist')
    hist_period = st.text_input('历史周期标签（如 5月/6月）', key='hist_period')
    if st.button('导入历史'):
        if hist_file and hist_period:
            import tempfile as _tmp
            with _tmp.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
                f.write(hist_file.read())
                tmp = f.name
            wb = openpyxl.load_workbook(tmp)
            os.unlink(tmp)
            ids = []
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row_data in ws.iter_rows(values_only=True):
                    val = row_data[4] if len(row_data) > 4 else None
                    if val is not None:
                        s = str(val).strip()
                        # 工号至少4位纯数字
                        if s.isdigit() and len(s) >= 4:
                            ids.append(str(int(s)))
            save_period(HISTORY_PATH, hist_period, list(set(ids)))
            st.success(f'已导入 {len(set(ids))} 条历史记录（周期：{hist_period}）')
        else:
            st.warning('请上传文件并填写周期标签')
```

- [ ] **Step 2: 运行 app**

```bash
cd "D:/dingyuan/Documents/能力提升计划工具"
streamlit run app.py
```

Expected: 浏览器自动打开，页面正常显示，无报错

- [ ] **Step 3: 首次初始化历史记录**

  展开"初始化历史记录"折叠区：
  - 上传 `能力提升计划名单-湖北战区26年7月.xlsx`
  - 历史周期标签填 `5月/6月`
  - 点击"导入历史"
  - 验证显示"已导入 X 条历史记录"

- [ ] **Step 4: 端到端测试（用真实数据）**

  **注意：** 三个文件都在同一个 Excel 里，分别上传同一文件即可（程序只读取各自的 Sheet）。

  参数设置：
  - 第一个月：`5月`，第二个月：`6月`
  - 识别周期结束日：`2026-06-30`
  - 季度：`Q2`，年份后两位：`26`
  - 输出标签：`7月`，改进周期：`7月-8月`

  验证结果对照现有公示表 `7月能力提升计划最终名单-公示` Sheet：
  - `喻佳（034244）` 在新进入名单，原因含"门店后15%"
  - `文莉莉（043853）` 在新进入名单，原因含"省区后50%"
  - `岑意（077071）` 在新进入名单，原因含"省区后15%"
  - 下载 Excel，确认格式与现有公示表一致

- [ ] **Step 5: commit**

```bash
git add app.py
git commit -m "feat: app.py - streamlit UI with upload, calculation, download, history init"
```

---

## Task 7: 收尾验证

- [ ] **Step 1: 运行全部测试**

```bash
python -m pytest tests/ -v
```

Expected: 所有测试 PASSED，无 FAILED

- [ ] **Step 2: 最终 commit**

```bash
git add .
git commit -m "feat: complete capability improvement identification tool v1.0"
```

---

## 快速启动说明

```bash
cd "D:/dingyuan/Documents/能力提升计划工具"
pip install -r requirements.txt   # 首次运行一次
streamlit run app.py               # 每次启动
```

浏览器会自动打开 `http://localhost:8501`。
