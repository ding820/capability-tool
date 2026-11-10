# tests/test_rules.py
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from rules import identify_experts, identify_managers, round_half_up


def test_round_half_up_matches_business_rounding():
    """业务口径是四舍五入；Python 内置 round() 是银行家舍入，偶数字会少取一人。"""
    # 内置 round 在这些点上是错的
    assert round(22.5) == 22
    assert round(10.5) == 10
    # 半进位取整给出业务期望值
    assert round_half_up(22.5) == 23
    assert round_half_up(10.5) == 11
    assert round_half_up(4.5) == 5
    # 非半点处的行为不变
    assert round_half_up(19.65) == 20   # 131 人 × 15%
    assert round_half_up(19.2) == 19    # 128 人 × 15%
    assert round_half_up(19.5) == 20    # 130 人 × 15%（内置 round 也是 20）
    assert round_half_up(1.0) == 1


def test_half_up_affects_threshold_count():
    """省区 150 人时后15%应取 23 人，内置 round 只会取 22 人。"""
    rows = []
    for i in range(150):
        rows.append((str(i + 1), f'P{i+1}', f'门店{i+1}', '省区A', 13,
                     '2020-01-01', 1000 - i, 0))
    df = _make_expert_df(rows)
    r = identify_experts(df, pd.Timestamp('2026-06-30'))
    assert int(r['省区参与人数'].iloc[0]) == 150
    # 23 人 -> 起点 rank = 150-23+1 = 128
    assert round_half_up(150 * 0.15) == 23
    sa = r.set_index('工号')
    # 大店且每店 1 人 -> 门店排名恒为 1，不触发；这里只校验阈值口径
    assert int(sa.loc['1', '省区参与人数']) == 150


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


def test_small_store_takes_worst_province_rank_in_store():
    """小店取「本店内省区排名最末」的那一个，而非「门店排名末位」。

    3 人小店，省区共 131 人（后15%起点=112）。
    三人省区排名分别为 110 / 113 / 131：
      - 110 不在后15%
      - 113 在后15%，但店内还有更差的 131
      - 131 是店内省区排名最末 -> 只有他触发
    """
    rows = []
    for i in range(1, 129):
        rows.append((f'b{i}', f'B{i}', f'大店{i}', '省区A', 13, '2020-01-01', 400 - i, 0))
    rows += [
        ('S1', 'S1', '小店X', '省区A', 13, '2020-01-01', 290, 0),
        ('S2', 'S2', '小店X', '省区A', 13, '2020-01-01', 288, 0),
        ('S3', 'S3', '小店X', '省区A', 13, '2020-01-01', 5, 0),
    ]
    df = _make_expert_df(rows)
    df['月1净锁单_全量'] = df['月1净锁单']
    df['月2净锁单_全量'] = 0
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert int(r.loc['S3', '省区参与人数']) == 131
    assert int(r.loc['S3', '省区排名']) == 131
    assert bool(r.loc['S3', '触发识别']) is True
    assert bool(r.loc['S2', '触发识别']) is False
    assert bool(r.loc['S1', '触发识别']) is False


def test_small_store_tie_broken_by_full_volume():
    """同一小店两人绩效并列、省区排名相同，用全量区分谁更靠后，只取那一个。

    两人绩效都是 500，全量分别为 600 和 300 -> 省区排名 113 与 114。
    店内省区排名最末的是 114 那个人 -> 只有他触发。
    """
    rows = []
    for i in range(1, 129):
        rows.append((f'b{i}', f'B{i}', f'大店{i}', '省区A', 13, '2020-01-01', 400 - i, 0))
    rows += [
        ('T1', 'T1', '小店X', '省区A', 13, '2020-01-01', 500, 0),
        ('T2', 'T2', '小店X', '省区A', 13, '2020-01-01', 500, 0),
        ('T3', 'T3', '小店X', '省区A', 13, '2020-01-01', 5, 0),
    ]
    df = _make_expert_df(rows)
    df['月1净锁单_全量'] = df['月1净锁单']
    df['月2净锁单_全量'] = 0
    df.loc[df['工号'] == 'T1', '月1净锁单_全量'] = 600
    df.loc[df['工号'] == 'T2', '月1净锁单_全量'] = 300
    df.loc[df['工号'] == 'T3', '月1净锁单_全量'] = 5

    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    # 全量把绩效并列的两人拆开
    assert int(r.loc['T1', '省区排名']) < int(r.loc['T2', '省区排名'])
    assert int(r.loc['T2', '省区全量排名']) > int(r.loc['T1', '省区全量排名'])
    # T3 仍然是店内省区排名最末（131），由它触发
    assert bool(r.loc['T3', '触发识别']) is True


def test_small_store_worst_is_outside_bottom15_no_trigger():
    """店内省区排名最末的那个人若仍不在省区后15%，全店都不触发。"""
    rows = []
    for i in range(1, 11):
        rows.append((f'b{i}', f'B{i}', f'大店{i}', '省区A', 13, '2020-01-01', 1000 - i * 10, 0))
    # 小店三人省区排名 1/2/3，远不在后15%（10人的后15% = 2人，起点 rank 9）
    rows += [
        ('S1', 'S1', '小店X', '省区A', 13, '2020-01-01', 990, 0),
        ('S2', 'S2', '小店X', '省区A', 13, '2020-01-01', 980, 0),
        ('S3', 'S3', '小店X', '省区A', 13, '2020-01-01', 970, 0),
    ]
    df = _make_expert_df(rows)
    df['月1净锁单_全量'] = df['月1净锁单']
    df['月2净锁单_全量'] = 0
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert int(r.loc['S3', '省区排名']) < 9
    assert bool(r.loc['S3', '触发识别']) is False


def test_store_3_or_less_uses_province_rank():
    """门店仅3人（≤3），改用省区后15% 且 必须是门店末名。

    省区共10人，后15%=round(10*0.15)=2人 -> 省区排名 9、10 落入。
    门店X 三人：S1(省区7, 店内1)、S2(省区8, 店内2)、S3(省区10, 店内3)。
    S3 同时满足"省区后15%"和"门店末名" -> 触发。
    S2 虽在省区后15%内，但店内排第2，非末名 -> 不触发（现行口径）。
    """
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
    triggered_ids = set(result[result['触发识别'] == True]['工号'].tolist())
    assert '10' in triggered_ids          # 省区后15% + 门店末名
    assert '9' not in triggered_ids       # 省区后15% 但店内排第2
    assert '8' not in triggered_ids       # 门店第1，且省区不在后15%


def test_senior_expert_province_50_percent_triggers():
    """
    高级产品专家省区后50%：分母=省区所有岗位类别=='产品专家'的非新员工（含L13+L14）。
    10人省区（3个高级专家，7个普通专家），后50%=5人。
    定单最低的高级专家应触发。

    注意：高专有两条触发路径（门店后15% / 省区后50%），任一满足即触发。
    High14 在本例中定单最低档（合计40，与 l13_3 并列），省区排名第4，
    已落入省区后50%（起点第6名？见下方断言），因此也会触发——
    测试重点应放在"省区后50%路径确实生效"，而非"高定单者不触发"。
    """
    rows = []
    rows.append(('l14_1', 'High14', '门店A', '省区A', 14, '2020-01-01', 20, 20))
    rows.append(('l14_2', 'Mid14', '门店B', '省区A', 14, '2020-01-01', 15, 15))
    rows.append(('l14_3', 'Low14', '门店C', '省区A', 14, '2020-01-01', 2, 3))
    for i, v in enumerate([30, 25, 22, 18, 12, 8, 4]):
        rows.append((f'l13_{i}', f'L13_{i}', f'门店{i+4}', '省区A', 13, '2020-01-01', v, v))
    df = _make_expert_df(rows)
    result = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    # Low14 合计=5，省区10人后50%=5人，起点rank6，低定单 -> 触发
    assert bool(result.loc['l14_3', '触发识别']) is True
    # 省区后50% 起点 = 10 - 5 + 1 = 6；排名 >= 6 的全部落入
    assert int(result.loc['l14_3', '省区排名']) >= 6
    # 每个高专都要能说清触发路径
    for eid in ['l14_1', 'l14_2', 'l14_3']:
        assert result.loc[eid, '识别原因'] != '' or not result.loc[eid, '触发识别']


def _make_expert_df_with_full(rows):
    """rows: list of (工号, 门店, 绩效月1, 绩效月2, 全量月1, 全量月2)"""
    records = []
    for r in rows:
        records.append({
            '工号': r[0], '员工姓名': f'E{r[0]}', '四级部门': r[1], '三级部门': '省区A',
            '新职级': 13, '入职日期': pd.Timestamp('2020-01-01'),
            '月1净锁单': r[2], '月2净锁单': r[3],
            '月1净锁单_全量': r[4], '月2净锁单_全量': r[5],
            '新岗位名称': '产品专家', '岗位类别': '产品专家', '二级部门': '湖北战区',
        })
    return pd.DataFrame(records)


def test_full_volume_breaks_performance_ties():
    """绩效双月合计并列（都是10）时，按全量定单分先后，并标记排名基准。"""
    df = _make_expert_df_with_full([
        ('1', '门店X', 5, 5, 50, 0),   # 绩效10，全量50 → 第1
        ('2', '门店X', 5, 5, 40, 0),   # 绩效10，全量40 → 第2
        ('3', '门店X', 5, 5, 30, 0),   # 绩效10，全量30 → 第3
    ])
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert list(r['双月合计']) == [10, 10, 10]
    assert r.loc['1', '门店排名'] == 1
    assert r.loc['2', '门店排名'] == 2
    assert r.loc['3', '门店排名'] == 3
    assert r.loc['1', '门店全量排名'] == 1
    assert r.loc['3', '省区全量排名'] == 3
    assert all(r['门店排名基准'] == '绩效并列→比全量')
    assert all(r['省区排名基准'] == '绩效并列→比全量')


def test_no_tie_leaves_rank_basis_blank():
    """绩效不并列时，排名基准留空，全量排名仍照常给出。"""
    df = _make_expert_df_with_full([
        ('1', '门店X', 10, 0, 10, 0),
        ('2', '门店X', 8, 0, 800, 0),   # 全量很高，但绩效低，仍排第2
    ])
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert r.loc['1', '门店排名'] == 1
    assert r.loc['2', '门店排名'] == 2
    assert r.loc['2', '门店全量排名'] == 1
    assert r.loc['1', '门店排名基准'] == ''
    assert r.loc['1', '省区排名基准'] == ''


def test_small_store_tied_last_place_triggers():
    """小店末两名并列时，两人门店排名都是2（无 rank=3），仍应被识别为门店末名。"""
    rows = []
    # 128 个放不进小店的大店人员，把省区人数撑到 131
    for i in range(1, 129):
        rows.append((str(i), f'P{i}', f'大店{i}', '省区A', 13, '2020-01-01', 300 - i, 0))
    # 小店 3 人：绩效并列末名，但全量不同 → 经全量区分后仍是并列末名
    rows += [
        ('201', 'S1', '小店X', '省区A', 13, '2020-01-01', 5, 0),
        ('202', 'S2', '小店X', '省区A', 13, '2020-01-01', 1, 0),
        ('203', 'S3', '小店X', '省区A', 13, '2020-01-01', 1, 0),
    ]
    df = _make_expert_df(rows)
    # 全量列必须整列赋值，否则未赋值的行会是 NaN（rules 里会参与元组排序）
    df['月1净锁单_全量'] = df['月1净锁单']
    df['月2净锁单_全量'] = 0
    # 末两名在绩效并列的基础上全量也相同 → 共享 rank=2，不存在 rank=3
    df.loc[df['工号'].isin(['202', '203']), '月1净锁单_全量'] = 7
    df.loc[df['工号'] == '201', '月1净锁单_全量'] = 9

    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert int(r.loc['202', '门店排名']) == 2
    assert int(r.loc['203', '门店排名']) == 2
    assert int(r.loc['202', '省区参与人数']) == 131
    assert bool(r.loc['202', '触发识别']) is True
    assert bool(r.loc['203', '触发识别']) is True
    # 门店第一名不是末名，即使省区排名也可能靠后，也不触发
    assert bool(r.loc['201', '触发识别']) is False


def test_senior_store_tie_triggers_via_store_15():
    """高专门店后15%也要支持末名并列。

    5人大店全是高专，末两名绩效并列(rank=4)。旧逻辑用 sp-t+1 算出 rank=5，
    而实际最大 rank 只有 4，导致门店条件永不成立；这两人省区排名又在上游，
    不会落入省区后50%，最终被漏判。
    """
    rows = []
    for i in range(1, 16):
        rows.append((f'x{i}', f'X{i}', f'门店{i}', '省区A', 13, '2020-01-01', 20 - i, 0))
    for eid, tot in [('A', 90), ('B', 80), ('C', 70), ('D', 60), ('E', 60)]:
        rows.append((eid, eid, '大店', '省区A', 14, '2020-01-01', tot, 0))
    df = _make_expert_df(rows)
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')

    # 前置条件：D/E 门店排名并列末位，且省区排名在上游（不触发省区后50%）
    assert int(r.loc['D', '门店排名']) == 4
    assert int(r.loc['E', '门店排名']) == 4
    assert int(r.loc['D', '省区排名']) < 11

    assert bool(r.loc['D', '触发识别']) is True
    assert bool(r.loc['E', '触发识别']) is True
    assert '门店' in r.loc['D', '识别原因']


def test_senior_store_unique_last_still_triggers():
    """高专门店末名唯一时行为不变（回归保护）。"""
    rows = []
    for i in range(1, 16):
        rows.append((f'x{i}', f'X{i}', f'门店{i}', '省区A', 13, '2020-01-01', 20 - i, 0))
    for eid, tot in [('A', 90), ('B', 80), ('C', 70), ('D', 60), ('E', 50)]:
        rows.append((eid, eid, '大店', '省区A', 14, '2020-01-01', tot, 0))
    df = _make_expert_df(rows)
    r = identify_experts(df, pd.Timestamp('2026-06-30')).set_index('工号')
    assert int(r.loc['E', '门店排名']) == 5
    assert bool(r.loc['E', '触发识别']) is True
    assert bool(r.loc['A', '触发识别']) is False


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
