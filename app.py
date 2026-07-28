# app.py
import streamlit as st
import pandas as pd
import os, tempfile, shutil
from data_loader import load_roster, load_expert_orders, load_expert_vehicles, load_manager_data, find_manager_cols, load_prev_list, format_display_id
from rules import identify_experts, identify_managers
from history import save_period, get_previous_ids
from exporter import build_output_excel

HISTORY_PATH = os.path.join(os.path.dirname(__file__), 'history.json')

st.set_page_config(page_title='能力提升计划识别工具', layout='wide')

st.markdown("""
<style>
/* 侧边栏背景：深绿 */
[data-testid="stSidebar"] {
    background-color: #002D28;
}
[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stDateInput input {
    background-color: #004a40;
    color: #FFFFFF;
    border-color: #CEA472;
}
/* Selectbox 触发框 — 覆盖 BaseWeb 所有嵌套层 */
[data-testid="stSidebar"] [data-baseweb="select"],
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] [data-baseweb="select"] > div > div,
[data-testid="stSidebar"] [data-baseweb="select"] > div > div > div,
[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"],
[data-testid="stSidebar"] [data-baseweb="select"] [data-testid="stSelectbox"],
[data-testid="stSidebar"] [class*="ValueContainer"],
[data-testid="stSidebar"] [class*="control"],
[data-testid="stSidebar"] [class*="singleValue"],
[data-testid="stSidebar"] [class*="Input"] {
    background-color: #004a40 !important;
    border-color: #CEA472 !important;
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] span,
[data-testid="stSidebar"] [data-baseweb="select"] div,
[data-testid="stSidebar"] [data-baseweb="select"] input {
    background-color: #004a40 !important;
    color: #FFFFFF !important;
}
/* stSelectbox wrapper — newer Streamlit versions */
[data-testid="stSidebar"] [data-testid="stSelectbox"] > div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div > div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] * {
    background-color: #004a40 !important;
    color: #FFFFFF !important;
    border-color: #CEA472 !important;
}
/* Catch-all: any div/span inside sidebar selectbox widget */
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] ~ div div,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] ~ div span,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] ~ div input {
    background-color: #004a40 !important;
    color: #FFFFFF !important;
}
/* 下拉弹出层 */
[data-baseweb="popover"] [data-baseweb="menu"] {
    background-color: #004a40 !important;
}
[data-baseweb="popover"] [role="option"] {
    background-color: #004a40 !important;
    color: #FFFFFF !important;
}
[data-baseweb="popover"] [role="option"]:hover,
[data-baseweb="popover"] [aria-selected="true"] {
    background-color: #CEA472 !important;
    color: #000000 !important;
}

/* 主标题 */
h1 {
    color: #002D28 !important;
    border-bottom: 3px solid #CEA472;
    padding-bottom: 10px;
}

/* 次级标题 */
h2, h3 {
    color: #002D28 !important;
}

/* 主按钮：沙金色 */
div.stButton > button[kind="primary"] {
    background-color: #CEA472;
    color: #000000;
    border: none;
    font-weight: 600;
    padding: 0.5rem 2rem;
}
div.stButton > button[kind="primary"]:hover {
    background-color: #b8905e;
    color: #000000;
}

/* 下载按钮：深绿 */
div.stDownloadButton > button {
    background-color: #002D28;
    color: #FFFFFF;
    border: none;
    font-weight: 600;
}
div.stDownloadButton > button:hover {
    background-color: #004a40;
    color: #FFFFFF;
}

/* 文件上传区 */
[data-testid="stFileUploader"] {
    border: 2px dashed #CEA472;
    border-radius: 8px;
    padding: 8px;
}

/* 成功提示 */
div[data-testid="stAlert"] {
    border-left: 4px solid #CEA472;
}
</style>
""", unsafe_allow_html=True)

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

# ── 主体：文件上传 + Sheet 选择 ───────────────────────────
data_file = st.file_uploader(
    '📂 上传数据文件（可包含全部内容：花名册 + 专家底表 + 主管底表 + 上期公示）',
    type=['xlsx'], key='data_file'
)
col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    roster_file = st.file_uploader(
        '📂 单独上传花名册（可选）',
        type=['xlsx'], key='roster_file'
    )
with col_opt2:
    orders_file = st.file_uploader(
        '📂 单独上传底表（可选；含专家底表、主管底表）',
        type=['xlsx'], key='orders_file'
    )

prev_sheet_name = None
roster_sheet_name = None   # 将由下面的 selectbox 赋值；None 表示未确定
orders_detail_sheet_name = None
exclude_month1 = []
exclude_month2 = []

# 底表来源：单独上传的底表文件 > 数据文件
_orders_source = orders_file if orders_file else data_file

def _get_sheets(file_obj) -> list:
    """安全读取 Excel 文件的 sheet 列表，不影响原始指针。"""
    import io as _io
    file_obj.seek(0)
    raw = file_obj.read()
    file_obj.seek(0)
    return pd.ExcelFile(_io.BytesIO(raw)).sheet_names

if roster_file:
    # 单独花名册：让用户选 sheet
    try:
        _all_sheets = _get_sheets(roster_file)
        _rc = [s for s in _all_sheets if '花名册' in s or '名册' in s]
        if not _rc:
            _rc = _all_sheets  # 找不到就列出所有 sheet
        roster_sheet_name = st.selectbox(
            '📋 选择花名册 Sheet（来自单独上传的花名册文件）',
            options=_rc,
            index=0,
        )
        st.caption(f'✅ 将读取 Sheet：**{roster_sheet_name}**')
    except Exception as _e:
        st.warning(f'读取花名册文件 Sheet 列表失败：{_e}')
elif data_file:
    # 从主文件里选花名册 sheet
    try:
        _all_sheets = _get_sheets(data_file)
        _rc = [s for s in _all_sheets if '花名册' in s or '名册' in s]
        if not _rc:
            _rc = _all_sheets
        roster_sheet_name = st.selectbox(
            '📋 选择花名册 Sheet（来自数据文件）',
            options=_rc,
            index=0,
        )
        st.caption(f'✅ 将读取 Sheet：**{roster_sheet_name}**')
    except Exception as _e:
        st.warning(f'读取数据文件 Sheet 列表失败：{_e}')


if data_file:
    try:
        _sheets = _get_sheets(data_file)

        # 上期公示 Sheet 选择
        _pub_sheets = [s for s in _sheets if '公示' in s or '能力提升' in s]
        if _pub_sheets:
            prev_sheet_name = st.selectbox(
                '📋 选择上期能力提升计划名单 Sheet（用于判断未出营）',
                options=['（不选，跳过未出营判断）'] + _pub_sheets,
                index=0,
            )
            if prev_sheet_name == '（不选，跳过未出营判断）':
                prev_sheet_name = None

        # 按月剔除车型（从底表来源预览车系）
        import tempfile as _tmpfile, shutil as _shutil, io as _io2
        _tmp2 = _tmpfile.mkdtemp()
        try:
            _dp2 = os.path.join(_tmp2, 'preview.xlsx')
            _src = orders_file if orders_file else data_file
            _src.seek(0)
            with open(_dp2, 'wb') as _f:
                _f.write(_src.read())
            _src.seek(0)
            _vehicles = load_expert_vehicles(_dp2, sheet_name=orders_detail_sheet_name or '专家底表')
        finally:
            _shutil.rmtree(_tmp2)

        if _vehicles:
            st.markdown('**🚗 剔除非绩效车型**（选中车型将从对应月份定单中剔除；不选任何车型则使用"全部"合计行）')
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                exclude_month1 = st.multiselect(
                    f'第一个月（{month1}）剔除车型',
                    options=_vehicles, default=[],
                    help=f'选中的车型在{month1}定单中被剔除后再汇总',
                )
            with col_m2:
                exclude_month2 = st.multiselect(
                    f'第二个月（{month2}）剔除车型',
                    options=_vehicles, default=[],
                    help=f'选中的车型在{month2}定单中被剔除后再汇总',
                )

    except Exception:
        pass

# ── 生成按钮 ──────────────────────────────────────────────
if st.button('生成能力提升名单', type='primary'):
    if not data_file:
        st.error('请上传数据文件（底表）')
        st.stop()
    if roster_sheet_name is None:
        st.error('请先上传花名册文件，并选择对应 Sheet。')
        st.stop()

    with st.spinner('计算中...'):
        not_graduated, new_entries, triggered_ids = [], [], set()
        expert_all, manager_all = [], []
        tmp_dir = tempfile.mkdtemp()
        try:
            data_path = os.path.join(tmp_dir, 'data.xlsx')
            data_file.seek(0)
            with open(data_path, 'wb') as f:
                f.write(data_file.read())

            # 花名册来源：单独上传的花名册文件 或 数据文件
            if roster_file:
                roster_path = os.path.join(tmp_dir, 'roster.xlsx')
                roster_file.seek(0)
                with open(roster_path, 'wb') as f:
                    f.write(roster_file.read())
            else:
                roster_path = data_path

            # 底表来源：单独上传的底表文件 或 数据文件
            if orders_file:
                orders_path = os.path.join(tmp_dir, 'orders.xlsx')
                orders_file.seek(0)
                with open(orders_path, 'wb') as f:
                    f.write(orders_file.read())
            else:
                orders_path = data_path

            # 读取上期名单工号
            prev_ids: set = set()
            if prev_sheet_name:
                prev_ids = load_prev_list(data_path, sheet_name=prev_sheet_name)
            else:
                prev_ids = get_previous_ids(HISTORY_PATH)

            try:
                roster = load_roster(roster_path, sheet_name=roster_sheet_name)
            except Exception as e:
                st.error(f'花名册读取失败：{e}\n请确认所选 Sheet 存在且"雇佣状态"列存在。')
                st.stop()

            try:
                expert_orders = load_expert_orders(
                    orders_path,
                    sheet_name=orders_detail_sheet_name or '专家底表',
                    exclude_month1=exclude_month1 or None,
                    exclude_month2=exclude_month2 or None,
                )
            except Exception as e:
                st.error(f'专家底表读取失败：{e}\n请确认文件包含所选 Sheet。')
                st.stop()

            try:
                manager_raw = load_manager_data(orders_path)
            except Exception as e:
                st.error(f'主管底表读取失败：{e}\n请确认文件包含"主管底表" Sheet。')
                st.stop()

            try:
                # 以花名册产品专家为基准（right join），无底表数据的人补 0
                roster_expert = roster[roster['岗位类别'] == '产品专家'][
                    ['工号', '员工姓名', '二级部门', '三级部门', '四级部门', '新岗位名称', '新职级', '入职日期', '岗位类别']
                ].copy()
                expert_merged = roster_expert.merge(
                    expert_orders.rename(columns={'专家工号': '工号'}),
                    on='工号', how='left'
                )
                for col in ['月1净锁单', '月2净锁单', '月1净锁单_全量', '月2净锁单_全量']:
                    if col in expert_merged.columns:
                        expert_merged[col] = expert_merged[col].fillna(0)
                    else:
                        expert_merged[col] = 0

                # 以花名册零售主管为基准（left join），无底表数据的人达成/目标补 0
                roster_mgr = roster[roster['岗位类别'] == '零售主管'][
                    ['工号', '二级部门', '三级部门', '四级部门', '新岗位名称', '新职级', '入职日期']
                ].copy()
                manager_merged = roster_mgr.merge(
                    manager_raw.rename(columns={'姓名': '员工姓名'}),
                    on='工号', how='left'
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

                # 汇总触发人员，工号补前导零
                expert_triggered = expert_result[expert_result['触发识别']].to_dict('records')
                manager_triggered = manager_result[manager_result['触发识别']].to_dict('records')
                all_triggered = expert_triggered + manager_triggered
                for r in all_triggered:
                    r['工号'] = format_display_id(r['工号'])
                triggered_ids = {r['工号'] for r in all_triggered}

                # 过程数据（全员，含未触发）：工号同样补零
                expert_all = expert_result.to_dict('records')
                manager_all = manager_result.to_dict('records')
                for r in expert_all + manager_all:
                    r['工号'] = format_display_id(r['工号'])

                # 未出营：本次触发 且 在上期名单中
                # 新进入：本次触发 且 不在上期名单中
                prev_ids_display = {format_display_id(i) for i in prev_ids}
                not_graduated = [r for r in all_triggered if r['工号'] in prev_ids_display]
                new_entries    = [r for r in all_triggered if r['工号'] not in prev_ids_display]

            except Exception as e:
                st.error(f'计算识别结果时出错：{e}')
                st.stop()

        finally:
            shutil.rmtree(tmp_dir)

    st.success(f'识别完成：未出营 {len(not_graduated)} 人，新进入 {len(new_entries)} 人')

    # ── 结果名单 ──────────────────────────────────────────
    DISPLAY_COLS = ['工号', '员工姓名', '四级部门', '新岗位名称', '新职级', '识别原因']

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f'未出营名单（{len(not_graduated)} 人）')
        if not_graduated:
            st.dataframe(pd.DataFrame(not_graduated)[DISPLAY_COLS], use_container_width=True)
        else:
            st.info('无未出营人员')

    with col_b:
        st.subheader(f'新进入名单（{len(new_entries)} 人）')
        if new_entries:
            st.dataframe(pd.DataFrame(new_entries)[DISPLAY_COLS], use_container_width=True)
        else:
            st.info('无新进入人员')

    # ── 过程数据-专家排名 ──────────────────────────────────
    st.divider()
    st.subheader('过程数据 — 专家业绩排名')
    st.caption('排名=1为绩效定单最高；绩效合计相同时看全量合计；🟡=触发识别，🔘=新员工（不参与排名）')

    if expert_all:
        import math as _math

        def _rank_ok(v):
            return v is not None and not (isinstance(v, float) and _math.isnan(v))

        rows = []
        for rec in expert_all:
            is_new      = rec.get('是否新员工', False)
            triggered   = rec.get('触发识别', False)
            store_part  = rec.get('门店参与人数') or 0
            prov_part   = rec.get('省区参与人数') or 0   # 小店口径（后15%用）
            prov_all    = rec.get('省区全员参与人数') or 0  # 省区全部产专（高专后50%用）
            store_rank  = rec.get('门店排名')
            prov_rank   = rec.get('省区排名')             # 小店口径排名
            prov_all_rank = rec.get('省区全员排名')       # 省区全产专排名
            is_small    = int(store_part) <= 3
            is_senior   = rec.get('新岗位名称') == '高级产品专家'

            if is_small:
                thresh_s_label = '（小店→看省区）'
                thresh_p_label = f'后{max(1,round(prov_part*0.15))}名' if prov_part else ''
            else:
                thresh_s_label = f'后{max(1,round(store_part*0.15))}名' if store_part else ''
                thresh_p_label = ''

            thresh50_label = f'后{max(1,round(prov_all*0.50))}名' if is_senior and prov_all else ''

            rows.append({
                '标记':                  '🟡触发' if triggered else ('🔘新员工' if is_new else ''),
                '工号':                  rec.get('工号', ''),
                '姓名':                  rec.get('员工姓名', ''),
                '门店':                  rec.get('四级部门', ''),
                '城市部':                rec.get('三级部门', ''),
                '岗位':                  rec.get('新岗位名称', ''),
                '职级':                  rec.get('新职级', ''),
                '入职日期':              str(rec.get('入职日期', ''))[:10] if rec.get('入职日期') else '',
                '绩效月1':               rec.get('月1净锁单', ''),
                '绩效月2':               rec.get('月2净锁单', ''),
                '绩效双月合计':          rec.get('双月合计', ''),
                '全量月1':               rec.get('月1净锁单_全量', ''),
                '全量月2':               rec.get('月2净锁单_全量', ''),
                '全量双月合计':          rec.get('双月合计_全量', ''),
                '门店参与(剔新)':        store_part or '',
                '门店排名':              store_rank if _rank_ok(store_rank) else '',
                '门店后15%阈值':         thresh_s_label,
                '省区小店参与(剔新)':    prov_part or '',
                '省区小店排名':          prov_rank if _rank_ok(prov_rank) else '',
                '省区小店后15%阈值':     thresh_p_label,
                '省区产专总人数(剔新)':  prov_all or '',
                '省区全员排名':          prov_all_rank if _rank_ok(prov_all_rank) else '',
                '省区后50%阈值(高专)':   thresh50_label,
                '识别原因':              rec.get('识别原因', ''),
                '_store':                rec.get('四级部门', ''),
                '_total':                rec.get('双月合计', 0) or 0,
            })

        df_expert = pd.DataFrame(rows)
        # 同一门店内按绩效定单量从高到低排列
        df_expert = df_expert.sort_values(['_store', '_total'], ascending=[True, False])
        df_expert = df_expert.drop(columns=['_store', '_total'])
        st.dataframe(
            df_expert,
            use_container_width=True,
            height=500,
            column_config={
                '绩效月1':      st.column_config.NumberColumn(format='%.0f'),
                '绩效月2':      st.column_config.NumberColumn(format='%.0f'),
                '绩效双月合计': st.column_config.NumberColumn(format='%.0f'),
                '全量月1':      st.column_config.NumberColumn(format='%.0f'),
                '全量月2':      st.column_config.NumberColumn(format='%.0f'),
                '全量双月合计': st.column_config.NumberColumn(format='%.0f'),
                '门店排名':     st.column_config.NumberColumn(format='%.0f'),
                '省区排名':     st.column_config.NumberColumn(format='%.0f'),
            },
        )

    # ── 过程数据-主管排名 ──────────────────────────────────
    st.divider()
    st.subheader('过程数据 — 主管季度达成率排名')

    if manager_all:
        rows_m = []
        for rec in manager_all:
            is_new    = rec.get('是否新员工', False)
            triggered = rec.get('触发识别', False)
            prov_part = rec.get('省区参与人数') or 0
            prov_rank = rec.get('省区排名')
            thresh10  = max(1, round(prov_part * 0.10)) if prov_part else ''

            rate = rec.get('季度达成率', '')
            try:
                rate = f'{float(rate):.1%}' if rate != '' and rate is not None else ''
            except Exception:
                rate = ''

            rows_m.append({
                '标记':          '🟡触发' if triggered else ('🔘新员工' if is_new else ''),
                '工号':          rec.get('工号', ''),
                '姓名':          rec.get('员工姓名', ''),
                '门店':          rec.get('四级部门', ''),
                '城市部':        rec.get('三级部门', ''),
                '职级':          rec.get('新职级', ''),
                '季度达成':      rec.get('季度达成', ''),
                '季度目标':      rec.get('季度目标', ''),
                '季度达成率':    rate,
                '参与人数(全员)': prov_part or '',
                '全员排名':      prov_rank if _rank_ok(prov_rank) else '',
                '后10%阈值':     thresh10,
                '识别原因':      rec.get('识别原因', ''),
            })

        df_mgr = pd.DataFrame(rows_m)
        st.dataframe(df_mgr, use_container_width=True, height=300)

    excel_buf = build_output_excel(period_label, improve_period, not_graduated, new_entries,
                                   expert_records=expert_all, manager_records=manager_all)
    st.download_button(
        label='下载公示表 Excel',
        data=excel_buf,
        file_name=f'{period_label}能力提升计划最终名单-公示.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    period_key = f'{month1}/{month2}'
    save_period(HISTORY_PATH, period_key, list(triggered_ids))
    st.caption(f'本期名单已自动保存（周期：{period_key}），下次使用时若不上传上期名单将以此为基准。')


