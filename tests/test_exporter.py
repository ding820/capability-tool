# tests/test_exporter.py
import sys
import os
import io

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from exporter import build_output_excel, sort_records


def _sample_records():
    return [
        {
            '工号': '1001', '员工姓名': '张三', '二级部门': '湖北战区',
            '三级部门': '湖北一部', '四级部门': '武汉龙阳大道零售中心',
            '新岗位名称': '产品专家', '新职级': 13,
            '识别原因': '双月专家定单量排名在所在门店的后15%',
        },
        {
            '工号': '1002', '员工姓名': '李四', '二级部门': '湖北战区',
            '三级部门': '湖北三部', '四级部门': '恩施金桂大道汽车城零售中心',
            '新岗位名称': '高级产品专家', '新职级': 14,
            '识别原因': '双月高级专家定单量排名在省区专家岗位后50%',
        },
    ]


def test_build_output_excel_returns_bytesio():
    buf = build_output_excel('7月', '7月-8月', [_sample_records()[0]], [_sample_records()[1]])
    assert isinstance(buf, io.BytesIO)
    assert buf.tell() == 0
    assert len(buf.getvalue()) > 0


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


def test_both_lists_empty_does_not_raise():
    buf = build_output_excel('7月', '7月-8月', [], [])
    wb = openpyxl.load_workbook(buf)
    ws = wb.active
    assert ws is not None


def test_sequence_numbers_start_at_1():
    buf = build_output_excel('7月', '7月-8月', _sample_records(), [])
    wb = openpyxl.load_workbook(buf)
    ws = wb.active
    # Row 1 = section title, row 2 = headers, row 3 = first data row
    assert ws.cell(row=3, column=1).value == 1
    assert ws.cell(row=4, column=1).value == 2


def _expert_detail_sheet(records):
    buf = build_output_excel('7月', '7月-8月', [], [], expert_records=records)
    wb = openpyxl.load_workbook(buf)
    ws = wb['7月过程数据-专家']
    headers = [c.value for c in ws[1]]
    return ws, headers, {h: i for i, h in enumerate(headers)}


def test_expert_sheet_has_full_volume_rank_columns():
    """过程数据-专家应包含全量定单与全量排名列，便于核对并列时的先后。"""
    _, _, idx = _expert_detail_sheet([{
        '工号': '1001', '员工姓名': '张三', '四级部门': '门店X',
        '月1净锁单': 5, '月2净锁单': 5, '双月合计': 10,
        '月1净锁单_全量': 50, '月2净锁单_全量': 0, '双月合计_全量': 50,
        '门店参与人数': 3, '门店排名': 2, '门店全量排名': 2,
        '省区参与人数': 131, '省区排名': 113, '省区全量排名': 113,
        '门店排名基准': '绩效并列→比全量', '省区排名基准': '绩效并列→比全量',
    }])
    for col in ['双月全量合计', '门店全量排名', '省区全量排名',
                '门店排名基准', '省区排名基准']:
        assert col in idx, f'缺少列: {col}'


def test_expert_sheet_marks_rank_basis_only_on_ties():
    """并列行标出排名基准，非并列行留空。"""
    ws, _, idx = _expert_detail_sheet([
        {'工号': '1001', '员工姓名': '张三', '四级部门': '门店X',
         '双月合计': 10, '双月合计_全量': 50,
         '门店参与人数': 3, '门店排名': 1, '门店全量排名': 1,
         '省区参与人数': 131, '省区排名': 112, '省区全量排名': 112,
         '门店排名基准': '绩效并列→比全量', '省区排名基准': '绩效并列→比全量'},
        {'工号': '1002', '员工姓名': '李四', '四级部门': '门店Y',
         '双月合计': 8, '双月合计_全量': 80,
         '门店参与人数': 5, '门店排名': 2, '门店全量排名': 1,
         '省区参与人数': 131, '省区排名': 120, '省区全量排名': 50,
         '门店排名基准': '', '省区排名基准': ''},
    ])
    assert ws.cell(row=2, column=idx['门店排名基准'] + 1).value == '绩效并列→比全量'
    assert ws.cell(row=3, column=idx['门店排名基准'] + 1).value in (None, '')


# ── 触发表记与 rules.py 一致性 ─────────────────────────────────────

def _store_tie_records():
    """5 人大店，末两名绩效并列且全量也并列 → 两人 rank 都是 4，无人 rank=5。"""
    return [
        {'工号': f'D{i}', '员工姓名': f'D{i}', '四级部门': '大店',
         '双月合计': p, '双月合计_全量': f,
         '门店参与人数': 5, '门店排名': r, '门店全量排名': r,
         '省区参与人数': 5, '省区排名': r, '省区全量排名': r,
         '新岗位名称': '产品专家', '触发识别': r >= 4}
        for i, (p, f, r) in enumerate(
            [(10, 10, 1), (8, 8, 2), (6, 6, 3), (2, 1, 4), (2, 1, 4)], 1)
    ]


def test_store_tie_marks_all_tied_last_place():
    """末名并列时，两人都应打✓（旧逻辑用 sp-t+1 会一个都不打）。"""
    ws, _, idx = _expert_detail_sheet(_store_tie_records())
    marks = [ws.cell(row=r, column=idx['是否触发门店后15%'] + 1).value
             for r in range(2, 7)]
    assert marks[3] == '✓'
    assert marks[4] == '✓'
    assert marks[0] in (None, '')


def test_new_employee_has_no_store_hit_mark():
    ws, _, idx = _expert_detail_sheet([
        {'工号': '1001', '员工姓名': '新人', '四级部门': '门店X', '是否新员工': True,
         '双月合计': 0, '门店参与人数': 5, '门店排名': None},
    ])
    assert ws.cell(row=2, column=idx['是否触发门店后15%'] + 1).value in (None, '')


# ── 名单排序：门店 → 双月合计降序 → 双月全量合计降序 ─────────────────

def _unsorted_entries():
    return [
        {'工号': 'E1', '员工姓名': '甲', '四级部门': '武汉A店',
         '双月合计': 10, '双月合计_全量': 100},
        {'工号': 'E2', '员工姓名': '乙', '四级部门': '恩施B店',
         '双月合计': 30, '双月合计_全量': 200},
        {'工号': 'E3', '员工姓名': '丙', '四级部门': '武汉A店',
         '双月合计': 12, '双月合计_全量': 80},
        {'工号': 'E4', '员工姓名': '丁', '四级部门': '武汉A店',
         '双月合计': 12, '双月合计_全量': 150},
        {'工号': 'E5', '员工姓名': '戊', '四级部门': '武汉A店',
         '双月合计': 5, '双月合计_全量': 500},
    ]


def test_sort_records_by_store_then_totals_desc():
    order = [r['工号'] for r in sort_records(_unsorted_entries())]
    # 门店升序：恩施B店 < 武汉A店；店内 双月合计 DESC，再按双月全量合计 DESC
    assert order == ['E2', 'E4', 'E3', 'E1', 'E5']


def test_sort_records_handles_missing_and_nan():
    recs = [
        {'工号': 'M1', '四级部门': '武汉A店', '双月合计': float('nan'),
         '双月合计_全量': None},
        {'工号': 'E1', '四级部门': '武汉A店', '双月合计': 1, '双月合计_全量': 1},
    ]
    assert [r['工号'] for r in sort_records(recs)] == ['E1', 'M1']


def test_notice_sheet_rows_are_sorted():
    buf = build_output_excel('7月', '7月-8月', [], _unsorted_entries())
    wb = openpyxl.load_workbook(buf)
    ws = wb['7月能力提升计划最终名单-公示']
    # 定位"进入能力提升计划人员名单"节：序号列从 1 开始的那一行
    start = next(r for r in range(1, ws.max_row + 1)
                 if ws.cell(row=r, column=1).value == 1)
    names = [ws.cell(row=r, column=6).value for r in range(start, start + 5)]
    assert names == ['乙', '丁', '丙', '甲', '戊']
    assert [ws.cell(row=r, column=1).value for r in range(start, start + 5)] == [1, 2, 3, 4, 5]
