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
        '岗位类别': ['产品专家', '产品专家'],
        '新职级': [13, 13],
        '入职日期': pd.to_datetime(['2024-01-01', '2023-06-01']),
        '二级部门': ['湖北战区', '湖北战区'],
    })
    import tempfile, os as _os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='花名册最新')
    result = load_roster(path)
    _os.unlink(path)
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
        '岗位类别': ['产品专家', '产品专家'],
        '新职级': [13, 13],
        '入职日期': pd.to_datetime(['2024-01-01', '2023-01-01']),
        '二级部门': ['战区', '战区'],
    })
    import tempfile, os as _os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='花名册最新')
    result = load_roster(path)
    _os.unlink(path)
    assert len(result) == 1
    assert result.iloc[0]['员工姓名'] == 'A'

from data_loader import load_expert_orders

def test_load_expert_orders_returns_aggregate_rows():
    """专家底表只取车系='全部'的行，即每人的合计行。"""
    import tempfile, os as _os
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '专家底表'
    ws.append(['区域','销售部','门店','专家姓名','专家工号','车系',
               '当前周期','','','','','','','','','','','','','','',
               '对比周期','','','','','','','','','','','','','','',
               '环比','','','',''])
    ws.append(['区域','销售部','门店','专家姓名','专家工号','车系',
               '跟进量','目标','达成率%','线下线索','目标','达成率%',
               '新增排程','目标','达成率%','试驾量','目标','达成率%',
               '净锁单','目标','达成率%',
               '跟进量','目标','达成率%','线下线索','目标','达成率%',
               '新增排程','目标','达成率%','试驾量','目标','达成率%',
               '净锁单','目标','达成率%',
               '跟进量','线下线索','新增排程','试驾量','净锁单'])
    ws.append(['湖北战区','','','丁琦','084300','全部',
               100,'','',10,'','',5,'','',3,'','',3,'','',
               90,'','',8,'','',4,'','',2,'','',8,'','','','','','',''])
    ws.append(['湖北战区','','','丁琦','084300','W04',
               50,'','',5,'','',3,'','',2,'','',3,'','',
               40,'','',4,'','',2,'','',1,'','',3,'','','','','','',''])
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    wb.save(path)
    result = load_expert_orders(path)
    _os.unlink(path)
    assert len(result) == 1
    assert result.iloc[0]['专家工号'] == '84300'
    assert result.iloc[0]['月1净锁单'] == 3
    assert result.iloc[0]['月2净锁单'] == 8

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
    import tempfile, os as _os
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    df.to_excel(path, index=False, sheet_name='主管底表')
    result = load_manager_data(path)
    _os.unlink(path)
    assert len(result) == 2
    assert result.iloc[0]['工号'] == '8210'

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
