# exporter.py
import io
from typing import List, Dict

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def _write_row(ws, row_idx: int, values: list) -> None:
    for col, val in enumerate(values, 1):
        ws.cell(row=row_idx, column=col, value=val)


def build_output_excel(
    period_label: str,
    improve_period: str,
    not_graduated: List[Dict],
    new_entries: List[Dict],
) -> io.BytesIO:
    """
    生成公示表 Excel，格式与现有公示表一致。

    Args:
        period_label:   期间标签，如 '7月'
        improve_period: 改进周期，如 '7月-8月'
        not_graduated:  未出营人员记录列表
        new_entries:    新进入人员记录列表

    Returns:
        BytesIO 对象，可直接传给 Streamlit 的 download_button。
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
            i,
            rec.get('二级部门', ''),
            rec.get('三级部门', ''),
            rec.get('四级部门', ''),
            rec.get('工号', ''),
            rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''),
            rec.get('新职级', ''),
            '否',
            rec.get('识别原因', ''),
        ])
        row += 1

    row += 1  # 空行

    # === 第二部分：进入能力提升计划名单 ===
    ws.cell(row=row, column=1, value=f'{period_label}进入能力提升计划人员名单').font = Font(bold=True, size=12)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    row += 1

    headers_entry = [
        '序号', '战区', '城市部', '门店', '工号', '姓名', '岗位类别', '职级',
        '是否进入能力提升计划', '识别进入能力提升考核期原因', '改进周期', '备注',
    ]
    _write_row(ws, row, headers_entry)
    for cell in ws[row]:
        cell.font = Font(bold=True)
    row += 1

    for i, rec in enumerate(new_entries, 1):
        _write_row(ws, row, [
            i,
            rec.get('二级部门', ''),
            rec.get('三级部门', ''),
            rec.get('四级部门', ''),
            rec.get('工号', ''),
            rec.get('员工姓名', ''),
            rec.get('新岗位名称', ''),
            rec.get('新职级', ''),
            '是',
            rec.get('识别原因', ''),
            improve_period,
            '',
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
