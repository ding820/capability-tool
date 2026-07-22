# tests/test_rules.py
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from rules import identify_experts, identify_managers


def _make_expert_df(rows):
    """rows: list of (工号, 姓名, 门店, 三级部门, 职级, 入职日期_str, 月1, 月2, 岗位名称=None)"""
    records = []
    for r in rows:
        岗位名称 = r[8] if len(r) > 8 else ('高级产品专家' if r[4] == 14 else '产品专家')
        records.append({
            '工号': r[0], '员工姓名': r[1], '四级部门': r[2], '三级部门': r[3],
            '新职级': r[4], '入职日期': pd.Timestamp(r[5]),
            '月1净锁单': r[6], '月2净锁单': r[7],
            '新岗位名称': 岗位名称,
            '岗位类别': '产品专家',
            '二级部门': '湖北战区',
        })
    return pd.DataFrame(records)


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
        ('4', 'D', '门店X', '省区A', 13, '2020-01-01', 2, 2),
        ('5', 'E', '门店X', '省区A', 13, '2020-01-01', 2, 2),
    ])
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    triggered = result[result['触发识别'] == True]
    triggered_ids = set(triggered['工号'].tolist())
    assert '4' in triggered_ids
    assert '5' in triggered_ids


def test_store_3_or_less_uses_province_rank():
    """门店仅3人（≤3），改用省区后15%。省区共10人，后15%=round(10*0.15)=2人。"""
    rows = []
    for i in range(7):
        rows.append((str(i + 1), f'P{i+1}', '门店Y', '省区A', 13, '2020-01-01', (i + 1) * 2, (i + 1) * 2))
    rows += [
        ('8', 'S1', '门店X', '省区A', 13, '2020-01-01', 3, 3),
        ('9', 'S2', '门店X', '省区A', 13, '2020-01-01', 2, 2),
        ('10', 'S3', '门店X', '省区A', 13, '2020-01-01', 1, 1),
    ]
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    triggered = result[result['触发识别'] == True]
    triggered_ids = set(triggered['工号'].tolist())
    assert '10' in triggered_ids
    assert '9' in triggered_ids


def test_senior_expert_province_50_percent_triggers():
    """
    高级产品专家省区后50%：分母=省区所有岗位类别=='产品专家'的非新员工（含L13+L14）。
    10人省区（3个高级专家，7个普通专家），后50%=5人。
    定单最低的高级专家应触发。
    """
    rows = []
    rows.append(('l14_1', 'High14', '门店A', '省区A', 14, '2020-01-01', 20, 20))
    rows.append(('l14_2', 'Mid14', '门店B', '省区A', 14, '2020-01-01', 15, 15))
    rows.append(('l14_3', 'Low14', '门店C', '省区A', 14, '2020-01-01', 2, 3))
    for i, v in enumerate([30, 25, 22, 18, 12, 8, 4]):
        rows.append((f'l13_{i}', f'L13_{i}', f'门店{i+4}', '省区A', 13, '2020-01-01', v, v))
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30'))
    # Low14 合计=5，省区10人后50%=5人，排名2，应触发
    low14_row = result[result['工号'] == 'l14_3']
    assert low14_row.iloc[0]['触发识别'] == True
    # High14 合计=40，排名10（最高），不触发
    high14_row = result[result['工号'] == 'l14_1']
    assert high14_row.iloc[0]['触发识别'] == False


def test_new_employee_excluded():
    """入职3个月内的新员工不参与排名，不触发。"""
    rows = [
        ('1', 'Old', '门店X', '省区A', 13, '2020-01-01', 1, 1),
        ('2', 'New', '门店X', '省区A', 13, '2026-05-01', 0, 0),
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


def _make_manager_df(rows, month_cols):
    records = []
    for r in rows:
        rec = {
            '工号': r[0], '姓名': r[1], '四级部门': r[2], '三级部门': r[3],
            '入职日期': pd.Timestamp(r[4]), '二级部门': '湖北战区',
            '新岗位名称': '零售主管', '新职级': 14,
        }
        for j, (col_d, col_t) in enumerate(month_cols):
            rec[col_d] = r[5 + j * 2]
            rec[col_t] = r[5 + j * 2 + 1]
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
        rows.append((str(i + 1), f'M{i+1}', f'门店{i+1}', '省区A', '2020-01-01',
                     achieve, 100, achieve, 100, achieve, 100))
    df = _make_manager_df(rows, month_cols)
    result = identify_managers(
        df,
        ['26年4月达成', '26年5月达成', '26年6月达成'],
        ['26年4月目标', '26年5月目标', '26年6月目标'],
        pd.Timestamp('2026-06-30'),
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
        pd.Timestamp('2026-06-30'),
    )
    new_row = result[result['工号'] == '2']
    assert new_row.iloc[0]['触发识别'] == False
    assert new_row.iloc[0]['是否新员工'] == True
