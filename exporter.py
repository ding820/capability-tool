# exporter.py
import io
import math
from typing import List, Dict, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 品牌色
_GREEN  = '002D28'
_GOLD   = 'CEA472'
_LGRAY  = 'F5F0E8'
_YELLOW = 'FFF2CC'  # 触发行高亮


def _write_row(ws, row_idx: int, values: list) -> None:
    for col, val in enumerate(values, 1):
        ws.cell(row=row_idx, column=col, value=val)


def _header_style(ws, row_idx: int, n_cols: int, bg: str = _GREEN) -> None:
    fill = PatternFill('solid', fgColor=bg)
    font = Font(bold=True, color='FFFFFF' if bg == _GREEN else '000000')
    for col in range(1, n_cols + 1):
        c = ws.cell(row=row_idx, column=col)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _auto_col_width(ws, max_width: int = 35) -> None:
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col if cell.value is not None),
            default=0,
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, max_width)


def _rank_ok(v) -> bool:
    return v is not None and not (isinstance(v, float) and math.isnan(v))


# ── 公示表 Sheet ────────────────────────────────────────────────────────────

def _build_notice_sheet(wb, period_label: str, improve_period: str,
                         not_graduated: List[Dict], new_entries: List[Dict]) -> None:
    ws = wb.active
    ws.title = f'{period_label}能力提升计划最终名单-公示'
    ws.freeze_panes = 'A3'

    row = 1

    # 第一节标题
    ws.cell(row=row, column=1, value=f'{period_label}未出营产品专家名单')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color=_GREEN)
    row += 1

    hdr_grad = ['序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级', '是否出营', '未出营原因']
    _write_row(ws, row, hdr_grad)
    _header_style(ws, row, len(hdr_grad))
    row += 1

    for i, rec in enumerate(not_graduated, 1):
        _write_row(ws, row, [
            i, rec.get('二级部门', ''), rec.get('三级部门', ''), rec.get('四级部门', ''),
            rec.get('工号', ''), rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''), rec.get('新职级', ''), '否', rec.get('识别原因', ''),
        ])
        row += 1

    row += 1  # 空行

    # 第二节标题
    ws.cell(row=row, column=1, value=f'{period_label}进入能力提升计划人员名单')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color=_GREEN)
    row += 1

    hdr_entry = [
        '序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级',
        '是否进入能力提升计划', '识别进入能力提升考核期原因', '改进周期', '备注',
    ]
    _write_row(ws, row, hdr_entry)
    _header_style(ws, row, len(hdr_entry))
    row += 1

    for i, rec in enumerate(new_entries, 1):
        _write_row(ws, row, [
            i, rec.get('二级部门', ''), rec.get('三级部门', ''), rec.get('四级部门', ''),
            rec.get('工号', ''), rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''), rec.get('新职级', ''),
            '是', rec.get('识别原因', ''), improve_period, '',
        ])
        row += 1

    _auto_col_width(ws)


# ── 过程数据-专家 Sheet ─────────────────────────────────────────────────────

def _build_expert_detail_sheet(wb, period_label: str,
                                expert_records: List[Dict]) -> None:
    ws = wb.create_sheet(title=f'{period_label}过程数据-专家')
    ws.freeze_panes = 'A2'

    headers = [
        '序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级',
        '入职日期', '是否新员工',
        '月1净锁单', '月2净锁单', '双月合计',
        '门店总人数', '门店参与人数(剔新员工)', '门店排名',
        '门店后15%阈值', '是否触发门店后15%',
        '省区参与人数(剔新员工)', '省区排名',
        '省区后15%阈值', '是否触发省区后15%',
        '省区后50%阈值(高级专家)', '是否触发省区后50%',
        '触发识别', '识别原因',
    ]
    _write_row(ws, 1, headers)
    _header_style(ws, 1, len(headers))

    gold_fill   = PatternFill('solid', fgColor=_GOLD)
    yellow_fill = PatternFill('solid', fgColor='FFF2CC')

    for i, rec in enumerate(expert_records, 1):
        is_new    = rec.get('是否新员工', False)
        triggered = rec.get('触发识别', False)

        store_total  = rec.get('门店总人数', 0) or 0
        store_part   = rec.get('门店参与人数', 0) or 0
        prov_part    = rec.get('省区参与人数', 0) or 0
        store_rank   = rec.get('门店排名')
        prov_rank    = rec.get('省区排名')
        is_senior    = rec.get('新岗位名称', '') == '高级产品专家'
        is_small     = store_part <= 3

        # 阈值计算（后15%/50% = 排名 >= total - thresh + 1）
        thresh15_store = max(1, round(store_part * 0.15)) if store_part else None
        thresh15_prov  = max(1, round(prov_part * 0.15)) if prov_part else None
        thresh50_prov  = max(1, round(prov_part * 0.50)) if is_senior and prov_part else None

        bottom15_store = store_part - thresh15_store + 1 if thresh15_store else None
        bottom15_prov  = prov_part  - thresh15_prov  + 1 if thresh15_prov  else None
        bottom50_prov  = prov_part  - thresh50_prov  + 1 if thresh50_prov  else None

        if is_small:
            hit_store15 = False
            # 小店：省区后15% 且 门店最后一名（两个条件同时满足）
            hit_prov15  = (_rank_ok(prov_rank) and bottom15_prov is not None
                           and int(prov_rank) >= bottom15_prov
                           and _rank_ok(store_rank) and int(store_rank) == store_part)
        else:
            hit_store15 = (_rank_ok(store_rank) and bottom15_store is not None
                           and int(store_rank) >= bottom15_store)
            hit_prov15  = False

        hit_store15_senior = (is_senior and _rank_ok(store_rank) and bottom15_store is not None
                              and int(store_rank) >= bottom15_store)
        hit_prov50 = (is_senior and _rank_ok(prov_rank) and bottom50_prov is not None
                      and int(prov_rank) >= bottom50_prov)

        row_data = [
            i,
            rec.get('二级部门', ''),
            rec.get('三级部门', ''),
            rec.get('四级部门', ''),
            rec.get('工号', ''),
            rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''),
            rec.get('新职级', ''),
            rec.get('入职日期', ''),
            '是' if is_new else '否',
            rec.get('月1净锁单', ''),
            rec.get('月2净锁单', ''),
            rec.get('双月合计', ''),
            store_total if store_total else '',
            store_part if store_part else '',
            store_rank if _rank_ok(store_rank) else '',
            f'后{thresh15_store}名' if not is_small and thresh15_store else ('→看省区' if is_small else ''),
            '✓' if hit_store15 else ('✓(高专)' if hit_store15_senior else ''),
            prov_part if prov_part else '',
            prov_rank if _rank_ok(prov_rank) else '',
            f'后{thresh15_prov}名' if thresh15_prov else '',
            '✓' if hit_prov15 else '',
            f'后{thresh50_prov}名' if thresh50_prov else '',
            '✓' if hit_prov50 else '',
            '是' if triggered else '否',
            rec.get('识别原因', ''),
        ]

        row_idx = i + 1
        _write_row(ws, row_idx, row_data)

        # 触发行：沙金高亮
        if triggered:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = gold_fill
        elif is_new:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = PatternFill('solid', fgColor=_LGRAY)

    _auto_col_width(ws, max_width=25)
    # 识别原因列宽加宽
    ws.column_dimensions[get_column_letter(len(headers))].width = 40


# ── 过程数据-主管 Sheet ─────────────────────────────────────────────────────

def _build_manager_detail_sheet(wb, period_label: str,
                                 manager_records: List[Dict]) -> None:
    ws = wb.create_sheet(title=f'{period_label}过程数据-主管')
    ws.freeze_panes = 'A2'

    headers = [
        '序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级',
        '入职日期', '是否新员工',
        '季度达成', '季度目标', '季度达成率',
        '省区参与人数(剔新员工)', '省区排名',
        '省区后10%阈值', '触发识别', '识别原因',
    ]
    _write_row(ws, 1, headers)
    _header_style(ws, 1, len(headers))

    gold_fill = PatternFill('solid', fgColor=_GOLD)
    gray_fill = PatternFill('solid', fgColor=_LGRAY)

    for i, rec in enumerate(manager_records, 1):
        is_new    = rec.get('是否新员工', False)
        triggered = rec.get('触发识别', False)
        prov_part = rec.get('省区参与人数', 0) or 0
        prov_rank = rec.get('省区排名')
        thresh10  = max(1, round(prov_part * 0.10)) if prov_part else None

        rate = rec.get('季度达成率', '')
        if rate != '' and rate is not None:
            try:
                rate = f'{float(rate):.1%}'
            except Exception:
                pass

        row_data = [
            i,
            rec.get('二级部门', ''),
            rec.get('三级部门', ''),
            rec.get('四级部门', ''),
            rec.get('工号', ''),
            rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''),
            rec.get('新职级', ''),
            rec.get('入职日期', ''),
            '是' if is_new else '否',
            rec.get('季度达成', ''),
            rec.get('季度目标', ''),
            rate,
            prov_part if prov_part else '',
            prov_rank if _rank_ok(prov_rank) else '',
            thresh10 if thresh10 else '',
            '是' if triggered else '否',
            rec.get('识别原因', ''),
        ]

        row_idx = i + 1
        _write_row(ws, row_idx, row_data)

        if triggered:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = gold_fill
        elif is_new:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = gray_fill

    _auto_col_width(ws, max_width=25)
    ws.column_dimensions[get_column_letter(len(headers))].width = 35


# ── 主入口 ──────────────────────────────────────────────────────────────────

def build_output_excel(
    period_label: str,
    improve_period: str,
    not_graduated: List[Dict],
    new_entries: List[Dict],
    expert_records: Optional[List[Dict]] = None,
    manager_records: Optional[List[Dict]] = None,
) -> io.BytesIO:
    """
    生成公示表 + 过程数据 Excel。

    Sheets:
      1. {period_label}能力提升计划最终名单-公示
      2. {period_label}过程数据-专家   （若提供 expert_records）
      3. {period_label}过程数据-主管   （若提供 manager_records）
    """
    wb = openpyxl.Workbook()

    _build_notice_sheet(wb, period_label, improve_period, not_graduated, new_entries)

    if expert_records:
        _build_expert_detail_sheet(wb, period_label, expert_records)

    if manager_records:
        _build_manager_detail_sheet(wb, period_label, manager_records)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
