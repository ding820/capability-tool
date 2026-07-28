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
    识别产品专家是否进入能力提升。

    排名规则（主键 DESC 绩效，次键 DESC 全量）：
      rank=1 = 最好；rank 最大 = 最差

    触发规则：
      L12/L13 产品专员/产品专家（非高专）
        大店（门店非新人 > 3）：门店后15%触发（≤9人取最后1名）
        小店（门店非新人 ≤ 3）：全省统一排名后15% 且 该门店省区排名最末的那人
      L14 高级产品专家（任一触发）
        条件B：门店后15%（无论大店/小店）
        条件A：全省统一排名后50%

    省区排名 = 全省所有非新人产品专家（132人剔新后128人）统一排名，不按城市部分组
    """
    result = df.copy().reset_index(drop=True)

    result['双月合计'] = result['月1净锁单'] + result['月2净锁单']
    if '月1净锁单_全量' in result.columns and '月2净锁单_全量' in result.columns:
        result['双月合计_全量'] = result['月1净锁单_全量'] + result['月2净锁单_全量']
    else:
        result['双月合计_全量'] = result['双月合计']
        result['月1净锁单_全量'] = result['月1净锁单']
        result['月2净锁单_全量'] = result['月2净锁单']

    result['是否新员工'] = result['入职日期'].apply(lambda d: _is_new_employee(d, period_end))
    active_mask = ~result['是否新员工']
    active = result[active_mask].copy()

    # ── 参与人数 ─────────────────────────────────────────────────────
    store_active_cnt = active.groupby('四级部门')['工号'].count()
    result['门店总人数']   = result.groupby('四级部门')['工号'].transform('count')
    result['门店参与人数'] = result['四级部门'].map(store_active_cnt).fillna(0).astype(int)

    small_store_ids = store_active_cnt[store_active_cnt <= 3].index

    # 省区参与人数 = 全省所有非新员工产品专家（大店+小店，全省统一）
    prov_active_cnt = len(active)
    result['省区参与人数']     = prov_active_cnt
    result['省区全员参与人数'] = prov_active_cnt

    result['门店排名']     = None
    result['省区排名']     = None   # 全省口径（小店15%、高专50%共用）
    result['省区全员排名'] = None   # 同上，兼容列名
    result['触发识别']     = False
    result['识别原因']     = ''

    if len(active) == 0:
        return result

    # ── 排名：绩效 DESC，全量 DESC ────────────────────────────────────
    def _rank(grp: pd.DataFrame) -> pd.Series:
        key = pd.Series(
            list(zip(grp['双月合计'], grp['双月合计_全量'])),
            index=grp.index,
        )
        return key.rank(method='min', ascending=False).astype(int)

    # 门店排名
    for _, grp in active.groupby('四级部门'):
        ranks = _rank(grp)
        result.loc[ranks.index, '门店排名'] = ranks

    # 省区排名：全省所有非新员工产品专家作为一个统一排名池
    prov_ranks = _rank(active)
    result.loc[prov_ranks.index, '省区排名']     = prov_ranks
    result.loc[prov_ranks.index, '省区全员排名'] = prov_ranks

    # ── 触发判断 ──────────────────────────────────────────────────────
    triggered = {i: False for i in result.index}
    reasons   = {i: ''    for i in result.index}

    for i, row in result.iterrows():
        if row['是否新员工']:
            continue

        sp       = int(row['门店参与人数'])
        pa       = int(row['省区参与人数'])
        store_rv = row['门店排名']
        prov_rv  = row['省区排名']
        is_senior = (row.get('新岗位名称') == '高级产品专家')
        is_small  = (row['四级部门'] in small_store_ids)

        if is_senior:
            # 条件B：门店后15%
            if sp > 0:
                t = max(1, round(sp * 0.15))
                if _rank_ok(store_rv) and int(store_rv) >= sp - t + 1:
                    triggered[i] = True
                    reasons[i]   = f'高级产品专家门店排名后15%'
            # 条件A：全省后50%
            if not triggered[i] and pa > 0:
                t = max(1, round(pa * 0.50))
                if _rank_ok(prov_rv) and int(prov_rv) >= pa - t + 1:
                    triggered[i] = True
                    reasons[i]   = f'高级产品专家省区排名后50%'
        else:
            if not is_small:
                # 大店：门店后15%
                if sp > 0:
                    t = max(1, round(sp * 0.15))
                    if _rank_ok(store_rv) and int(store_rv) >= sp - t + 1:
                        triggered[i] = True
                        reasons[i]   = f'绩效双月定单量门店排名后15%'
            else:
                # 小店：省区后15%，且是该门店中省区排名最末的那一个（同一排序键，等价于门店最后1名）
                if pa > 0 and sp > 0:
                    t_prov  = max(1, round(pa * 0.15))
                    hit_prov = _rank_ok(prov_rv) and int(prov_rv) >= pa - t_prov + 1
                    hit_last = _rank_ok(store_rv) and int(store_rv) == sp
                    if hit_prov and hit_last:
                        triggered[i] = True
                        reasons[i]   = f'绩效双月定单量省区排名后15%'

    result['触发识别'] = [triggered[i] for i in result.index]
    result['识别原因'] = [reasons[i]   for i in result.index]
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

    # 主管全员统一排名（不按省区分组）：1=达成率最高，后10%=排名最低的那几人
    total_active = len(active)
    global_rank = active['季度达成率'].rank(method='min', ascending=False)

    result['省区排名'] = None
    result['省区参与人数'] = total_active  # 全员人数，方便展示
    result.loc[active_mask, '省区排名'] = global_rank.values

    total_thresh = max(1, round(total_active * 0.10))
    bottom_thresh = total_active - total_thresh + 1  # 排名 >= 此值即为后10%

    triggered = []
    reasons = []
    for _, row in result.iterrows():
        if row['是否新员工']:
            triggered.append(False)
            reasons.append('')
            continue
        rank_val = row['省区排名']
        if rank_val is not None and not (isinstance(rank_val, float) and math.isnan(rank_val)):
            if rank_val >= bottom_thresh:
                triggered.append(True)
                reasons.append(f'季度主管达成率全员排名后10%')
                continue
        triggered.append(False)
        reasons.append('')

    result['触发识别'] = triggered
    result['识别原因'] = reasons
    return result
