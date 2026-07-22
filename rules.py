# rules.py
import math
import pandas as pd


def _is_new_employee(hire_date, period_end: pd.Timestamp) -> bool:
    """入职日期距识别周期结束日 < 3个月视为新员工。"""
    if pd.isna(hire_date):
        return False
    cutoff = period_end - pd.DateOffset(months=3)
    return pd.Timestamp(hire_date) > cutoff


def _rank_ok(v):
    return v is not None and not (isinstance(v, float) and math.isnan(v))


def identify_experts(df: pd.DataFrame, period_end: pd.Timestamp) -> pd.DataFrame:
    """
    识别产品专家/产品专员/高级产品专家是否进入能力提升。

    df 必须包含列: 工号, 员工姓名, 四级部门(门店), 三级部门(省区/城市部),
                   新职级, 入职日期, 月1净锁单, 月2净锁单, 新岗位名称, 岗位类别, 二级部门
    period_end: 识别周期结束日期

    后15%阈值分母 = 参与排名人数（非新员工）。
    高级产品专家省区后50%分母 = 省区所有岗位类别=='产品专家'的参与人数（含专员/专家/高级）。
    Precondition: df 必须包含所有相关省区的 岗位类别=='产品专家' 人员（含专员/专家/高级），
                  否则高级产品专家的后50%分母将不准确。
    """
    result = df.copy().reset_index(drop=True)
    result['双月合计'] = result['月1净锁单'] + result['月2净锁单']
    result['是否新员工'] = result['入职日期'].apply(lambda d: _is_new_employee(d, period_end))

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

    # 省区排名（所有岗位类别=='产品专家'，含L12/L13/L14）
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

        store_total = int(row['门店总人数'])
        store_part = int(row['门店参与人数'])
        province_part = int(row['省区参与人数'])
        store_rank_val = row['门店排名']
        province_rank_val = row['省区排名']

        hit = False
        reason = ''

        if store_total <= 3:
            # 门店≤3人，改用省区后15%
            thresh = max(1, round(province_part * 0.15))
            if _rank_ok(province_rank_val) and province_rank_val <= thresh:
                hit = True
                reason = '双月专家定单量排名在省区后15%'
        else:
            thresh = max(1, round(store_part * 0.15))
            if _rank_ok(store_rank_val) and store_rank_val <= thresh:
                hit = True
                reason = '双月专家定单量排名在所在门店的后15%'

        # 高级产品专家额外判断：省区所有专家岗后50%
        if row.get('新岗位名称') == '高级产品专家' and not hit:
            thresh50 = max(1, round(province_part * 0.50))
            if _rank_ok(province_rank_val) and province_rank_val <= thresh50:
                hit = True
                reason = '双月高级专家定单量排名在省区专家岗位后50%'

        triggered.append(hit)
        reasons.append(reason)

    result['触发识别'] = triggered
    result['识别原因'] = reasons
    return result


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
        count = int(row['省区参与人数'])
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
