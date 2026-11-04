from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import DocumentCounts, DocumentType, OverallSummary, StaffCounts, TrendCounts
from service.service_excel_writer import save_results, save_staff_results, save_trend_results

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


def test_save_results(tmp_path: Path) -> None:
    counts = DocumentCounts(
        by_staff_department={('植田', '内科'): 2, ('花﨑', '外科'): 1, ('外部', '内科'): 1, ('植田', '眼科'): 5},
        by_department={'内科': 3, '外科': 1, '眼科': 5},
        total=9,
        first_date=date(2025, 4, 1),
        last_date=date(2025, 5, 20),
    )

    output_path = save_results(OverallSummary(counts, ['植田', '花﨑'], ['内科', '外科'], DocumentType.MEDICAL),
                               str(tmp_path / 'output'))

    assert Path(output_path) == tmp_path / 'output' / '医療文書作成件数_医療文書_20250401-20250520.xlsx'
    sheet = active_sheet(output_path)
    assert sheet.title == '一覧'
    assert sheet_values(sheet) == [
        ['医療文書作成件数 医療文書 2025年04月01日-2025年05月20日', None, None, None],
        ['氏名', '合計', '内科', '外科'],
        ['植田', 2, 2, 0],
        ['花﨑', 1, 0, 1],
        ['合計', 9, 3, 1],
    ]
    assert sheet['A1'].font.b
    assert sheet['D5'].border.left.style == 'thin'
    assert sheet['D5'].alignment.horizontal == 'center'
    assert sheet.freeze_panes == 'B3'
    assert sheet.page_setup.orientation == 'landscape'
    assert sheet.print_area == "'一覧'!$A$1:$L$34"
    assert chart_types(sheet) == ['barChart']
    chart = getattr(sheet, '_charts')[0]
    assert (chart.type, chart.grouping, chart.overlap) == ('bar', 'stacked', 100)
    # 系列は合計列を除いた診療科、分類は合計行を除いた担当者
    assert [series.tx.strRef.f for series in chart.series] == ["'一覧'!C2", "'一覧'!D2"]
    assert chart.series[0].cat.numRef.f == "'一覧'!$A$3:$A$4"


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
    assert chart_types(sheet) == ['lineChart']


def chart_types(sheet: Worksheet) -> list[str]:
    return [chart.tagname for chart in getattr(sheet, '_charts')]


def test_save_trend_results(tmp_path: Path) -> None:
    trend = TrendCounts(4, [(2024, [1] * 12), (2025, [2, 0, *[2] * 10]), (2026, [3, *[None] * 11])])

    output_path = save_trend_results(trend, DEPARTMENT, date(2024, 4, 1), date(2026, 4, 30),
                                     str(tmp_path / 'output'))

    assert Path(output_path).name == f'{TITLE}推移_{DEPARTMENT}_20240401-20260430.xlsx'
    sheet = active_sheet(output_path)
    values = sheet_values(sheet)
    assert values[0][0] == f'{TITLE}推移(医師依頼日) {DEPARTMENT} 2024/04/01～2026/04/30'
    assert values[2] == ['件数', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月',
                         '3月', '合計']
    assert values[3:6] == [['2024年度', *[1] * 12, 12], ['2025年度', 2, 0, *[2] * 10, 22],
                           ['2026年度', 3, *[None] * 11, 3]]
    assert chart_types(sheet) == ['lineChart']
    chart = getattr(sheet, '_charts')[0]
    assert [series.marker.symbol for series in chart.series] == ['circle'] * 3
    # 凡例とタイトルをグラフ(月のラベル)に重ねない
    assert (chart.legend.position, chart.legend.overlay, chart.title.overlay) == ('b', False, False)
    # グラフの元は隠しシートで、0 件の月は空欄にして線を引かない
    chart_sheet = load_workbook(output_path)['グラフ用']
    assert chart_sheet.sheet_state == 'hidden'
    assert [cell.value for cell in chart_sheet[2]][:4] == ['2025年度', 2, None, 2]
