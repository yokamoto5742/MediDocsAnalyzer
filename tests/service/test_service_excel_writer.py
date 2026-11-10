from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import (DocumentCounts, DocumentFilter, DocumentType, OverallSummary, StaffCounts,
                                           TrendCounts)
from service.service_excel_writer import save_results, save_staff_results, save_trend_results
from tests.conftest import TEMPLATE_DIR

TITLE = '医療文書作成件数'
STAFF_NAME = '花﨑'
DEPARTMENT = '内科'
DOCUMENT = '診断書'


def active_sheet(path: Path | str) -> Worksheet:
    sheet = load_workbook(path).active
    assert isinstance(sheet, Worksheet)
    return sheet


def sheet_values(sheet: Worksheet) -> list[list[object]]:
    return [list(row) for row in sheet.iter_rows(values_only=True)]


def row_values(sheet: Worksheet, row: int, last_column: int) -> list[object]:
    return [sheet.cell(row=row, column=column).value for column in range(1, last_column + 1)]


def assert_chart_kept(path: str) -> None:
    """テンプレートのグラフシートが残っている"""
    chart_sheet = load_workbook(path)['グラフ']
    assert isinstance(chart_sheet, Chartsheet)
    assert len(getattr(chart_sheet, '_charts')) == 1


def test_save_results_fills_template(tmp_path: Path) -> None:
    counts = DocumentCounts(
        by_staff_department={('植田', '内科'): 2, ('花﨑', '外科'): 1, ('外部', '内科'): 1, ('植田', '眼科'): 5},
        by_department={'内科': 3, '外科': 1, '眼科': 5},
        total=9,
        first_date=date(2025, 4, 1),
        last_date=date(2025, 5, 20),
    )
    # 「外部」はテンプレートに無いので出さない
    summary = OverallSummary(counts, ['植田', '花﨑', '外部'], ['内科', '外科', '眼科'], DocumentType.MEDICAL)

    output_path = save_results(summary, str(TEMPLATE_DIR), str(tmp_path / 'output'))

    assert Path(output_path) == tmp_path / 'output' / f'{TITLE}_医療文書_20250401-20250520.xlsx'
    sheet = active_sheet(output_path)
    assert sheet['A1'].value == f'{TITLE}_医療文書_20250401-20250520'
    assert row_values(sheet, 2, 4) == ['氏名', '合計', '内科', '外科']
    # 植田(3 行目)・花﨑(13 行目)・合計(16 行目)。眼科は H 列
    assert row_values(sheet, 3, 12) == ['植田', 7, 2, 0, 0, 0, 0, 5, 0, 0, 0, 0]
    assert row_values(sheet, 13, 4) == ['花﨑', 1, 0, 1]
    assert row_values(sheet, 16, 4) == ['合計', 9, 3, 1]
    # マスタに無い担当者・診療科は 0 件
    assert row_values(sheet, 4, 3) == ['沖野', 0, 0]
    assert sheet['E3'].value == 0
    assert_chart_kept(output_path)


def test_save_results_raises_when_template_is_missing(tmp_path: Path) -> None:
    counts = DocumentCounts({('植田', '内科'): 1}, {'内科': 1}, 1, date(2025, 4, 1), date(2025, 4, 1))

    with pytest.raises(FileNotFoundError):
        save_results(OverallSummary(counts, ['植田'], ['内科']), str(tmp_path), str(tmp_path / 'output'))


def test_save_staff_results(tmp_path: Path) -> None:
    staff_counts = StaffCounts(STAFF_NAME, 3, [(DEPARTMENT, 3)], [(DOCUMENT, 3)],
                               [('2025/04', 2), ('2025/05', 1)])

    output_path = save_staff_results(staff_counts, date(2025, 4, 1), date(2025, 5, 31), str(tmp_path / 'output'))

    assert Path(output_path).name == f'{TITLE}_{STAFF_NAME}_全書類_20250401-20250531.xlsx'
    sheet = active_sheet(output_path)
    assert sheet['A1'].value == f'{TITLE} {STAFF_NAME} 全書類 2025年04月01日-2025年05月31日'
    assert sheet['B2'].value == 3
    assert (sheet['D2'].value, sheet['E2'].value) == (None, None)
    values = sheet_values(sheet)
    assert values[4][:8] == [DEPARTMENT, 3, None, DOCUMENT, 3, None, '2025/04', 2]
    assert values[5][6:8] == ['2025/05', 1]
    # 既定の文字の大きさは 12
    assert (sheet['A2'].font.sz, sheet['D4'].font.sz, sheet['H6'].font.sz) == (12, 12, 12)
    assert getattr(sheet, '_charts') == []


def test_save_staff_results_uses_font_size(tmp_path: Path) -> None:
    staff_counts = StaffCounts(STAFF_NAME, 1, [(DEPARTMENT, 1)], [(DOCUMENT, 1)], [('2025/04', 1)])

    output_path = save_staff_results(staff_counts, date(2025, 4, 1), date(2025, 4, 30), str(tmp_path / 'output'), 20)

    sheet = active_sheet(output_path)
    assert (sheet['A2'].font.sz, sheet['G5'].font.sz) == (20, 20)
    assert sheet.row_dimensions[5].height == 31


def test_save_trend_results_fills_template_for_fiscal_years(tmp_path: Path) -> None:
    # テンプレートの年は 2025〜2027 年。2024 年度は出さない
    trend = TrendCounts(4, [(2024, [1] * 12), (2025, [2, 0, *[2] * 10]), (2026, [3, *[None] * 11])])

    output_path = save_trend_results(trend, DocumentFilter(DEPARTMENT, DocumentType.MEDICAL), STAFF_NAME,
                                     date(2024, 4, 1), date(2026, 4, 30), str(TEMPLATE_DIR), str(tmp_path / 'output'))

    title = f'{TITLE}推移_{STAFF_NAME}_医療文書_{DEPARTMENT}_20240401-20260430'
    assert Path(output_path).name == f'{title}.xlsx'
    sheet = active_sheet(output_path)
    assert sheet['A1'].value == title
    assert row_values(sheet, 2, 14) == ['作成件数', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月',
                                        '1月', '2月', '3月', '合計']
    # 0 件の月は空欄にしてグラフの線を途切れさせる。データの無い年は空にする
    assert row_values(sheet, 3, 14) == ['2025年度', 2, None, *[2] * 10, 22]
    assert row_values(sheet, 4, 14) == ['2026年度', 3, *[None] * 11, 3]
    assert row_values(sheet, 5, 14) == ['2027年度', *[None] * 13]
    assert_chart_kept(output_path)


def test_save_trend_results_without_staff_drops_name(tmp_path: Path) -> None:
    trend = TrendCounts(1, [(2025, [5] * 12)])

    output_path = save_trend_results(trend, DocumentFilter(), None, date(2025, 1, 1), date(2025, 12, 31),
                                     str(TEMPLATE_DIR), str(tmp_path / 'output'))

    sheet = active_sheet(output_path)
    assert sheet['A1'].value == f'{TITLE}推移_全書類_合計_20250101-20251231'
    assert row_values(sheet, 2, 3) == ['作成件数', '1月', '2月']
    assert row_values(sheet, 3, 14) == ['2025年', *[5] * 12, 60]
