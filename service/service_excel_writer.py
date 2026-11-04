from datetime import date
from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.title import title_maker
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import TOTAL_LABEL, OverallSummary, StaffCounts, TrendCounts

TREND_CORNER_LABEL = '件数'


OVERALL_FONT_NAME = 'ＭＳ Ｐゴシック'
OVERALL_ROW_HEIGHT = 19.2
THIN_SIDE = Side(style='thin')
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
OVERALL_CHART_HEIGHT = 14
OVERALL_MIN_COLUMN_WIDTH = 8


def write_counts(sheet: Worksheet, summary: OverallSummary) -> None:
    """2 行目に見出し、3 行目から担当者ごとの件数と合計行を入れ、罫線で囲んで中央に揃える"""
    headers = [TOTAL_LABEL, *summary.departments]
    # 全角文字は 1 文字で幅 2 を使うため、見出しが切れないよう列幅を広げる
    for column_index, header in enumerate(headers, start=2):
        width = max(OVERALL_MIN_COLUMN_WIDTH, len(header) * 2 + 2)
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    table = [['氏名', *headers], *summary.rows]
    for row_index, row in enumerate(table, start=2):
        sheet.row_dimensions[row_index].height = OVERALL_ROW_HEIGHT
        for column_index, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_index, column=column_index, value=value)
            cell.font = Font(name=OVERALL_FONT_NAME, size=11)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal='center')


def add_overall_chart(sheet: Worksheet, summary: OverallSummary, title: str, anchor: str) -> None:
    """担当者ごとの件数を診療科で色分けした積み上げ横棒グラフ(全体タブの図と同じ)を置く"""
    chart = BarChart()
    chart.type = 'bar'
    chart.grouping = 'stacked'
    # 積み上げるには系列を 100% 重ねる
    chart.overlap = 100
    set_chart_frame(chart, title)
    # 合計行を除いた担当者の行と、合計列を除いた診療科の列を使う
    last_staff_row = 2 + len(summary.staff_members)
    chart.add_data(Reference(sheet, min_col=3, max_col=2 + len(summary.departments), min_row=2,
                             max_row=last_staff_row), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=3, max_row=last_staff_row))
    # 上から表と同じ順に並べる
    chart.x_axis.scaling.orientation = 'maxMin'
    chart.height = OVERALL_CHART_HEIGHT
    assert chart.legend is not None
    chart.legend.position = 'r'
    chart.legend.overlay = False
    sheet.add_chart(chart, anchor)


def save_results(summary: OverallSummary, output_dir: str) -> str:
    """集計結果の表と積み上げ横棒グラフを新しいブックに書いて保存し、保存先のパスを返す"""
    wb = Workbook()
    sheet = cast(Worksheet, wb.active)
    sheet.title = '一覧'

    counts = summary.counts
    document_type = summary.document_type.value
    title = (f'医療文書作成件数 {document_type} {counts.first_date:%Y年%m月%d日}'
             f'-{counts.last_date:%Y年%m月%d日}')
    sheet['A1'] = title
    sheet['A1'].font = Font(name=OVERALL_FONT_NAME, size=12, bold=True)
    sheet.row_dimensions[1].height = 20.4
    sheet.column_dimensions['A'].width = 10.66
    write_counts(sheet, summary)
    sheet.freeze_panes = 'B3'

    chart_row = len(summary.rows) + 4
    add_overall_chart(sheet, summary, f'{document_type} 担当者別件数', f'A{chart_row}')
    # 印刷範囲は指定しない(Excel はグラフまで含めて印刷し、A4 横 1 枚に縮める)
    set_print_layout(sheet)

    return save_workbook(wb, output_dir,
                         f'医療文書作成件数_{document_type}_{counts.first_date:%Y%m%d}-{counts.last_date:%Y%m%d}.xlsx')


def save_workbook(wb: Workbook, output_dir: str, file_name: str) -> str:
    """出力先フォルダに保存し、保存先のパスを返す"""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    output_file_path = output_path / file_name
    wb.save(output_file_path)
    return str(output_file_path)


def set_print_layout(sheet: Worksheet) -> None:
    """A4 横 1 枚に収めて印刷する"""
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1


def set_chart_frame(chart: BarChart | LineChart, title: str) -> None:
    """タイトル・大きさ・軸の表示をそろえる"""
    # overlay を明示しないと、Excel はタイトルをグラフに重ねて表示する
    chart_title = title_maker(title)
    chart_title.overlay = False
    chart.title = chart_title
    chart.width = 24
    chart.height = 10
    # openpyxl 3.1 は軸を既定で隠すため、明示的に表示する
    chart.x_axis.delete = False
    chart.y_axis.delete = False


def make_line_chart(title: str) -> LineChart:
    chart = LineChart()
    set_chart_frame(chart, title)
    return chart


def write_table(sheet: Worksheet, row: int, column: int, headers: tuple[str, str],
                rows: list[tuple[str, int]]) -> None:
    for offset, value in enumerate(headers):
        sheet.cell(row=row, column=column + offset, value=value).font = Font(bold=True)
    for row_offset, (label, count) in enumerate(rows, start=1):
        sheet.cell(row=row + row_offset, column=column, value=label)
        sheet.cell(row=row + row_offset, column=column + 1, value=count)


def save_staff_results(staff_counts: StaffCounts, start_date: date, end_date: date, output_dir: str) -> str:
    """1 人分の集計(診療科別・文書名別・月別と月別グラフ)を新しいブックに書いて保存する"""
    wb = Workbook()
    sheet = cast(Worksheet, wb.active)
    sheet.title = '個人'
    name = staff_counts.staff_name
    document_type = staff_counts.document_type.value
    sheet['A1'] = f'医療文書作成件数 {name} {document_type} {start_date:%Y年%m月%d日}-{end_date:%Y年%m月%d日}'
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A2'] = TOTAL_LABEL
    sheet['B2'] = staff_counts.total

    write_table(sheet, 4, 1, ('診療科', '件数'), staff_counts.by_department)
    write_table(sheet, 4, 4, ('文書名', '件数'), staff_counts.by_document)
    write_table(sheet, 4, 7, ('月', '件数'), staff_counts.by_month)
    for column, width in (('A', 14), ('D', 48), ('G', 10)):
        sheet.column_dimensions[column].width = width

    chart = make_line_chart(f'{name} {document_type} 月別件数')
    chart.legend = None
    chart.add_data(Reference(sheet, min_col=8, min_row=4, max_row=4 + len(staff_counts.by_month)),
                   titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=7, min_row=5, max_row=4 + len(staff_counts.by_month)))
    table_rows = max(len(staff_counts.by_department), len(staff_counts.by_document), len(staff_counts.by_month))
    sheet.add_chart(chart, f'A{table_rows + 7}')
    set_print_layout(sheet)

    return save_workbook(wb, output_dir,
                         f'医療文書作成件数_{name}_{document_type}_{start_date:%Y%m%d}-{end_date:%Y%m%d}.xlsx')


def save_trend_results(trend: TrendCounts, subject: str, start_date: date, end_date: date,
                       output_dir: str) -> str:
    """年(年度)ごとの月別件数の表と、1 年 1 本の折れ線グラフを新しいブックに書いて保存する。
    subject は「集計対象 診療科」(個人推移なら前に氏名)"""
    wb = Workbook()
    sheet = cast(Worksheet, wb.active)
    sheet.title = '推移'
    sheet['A1'] = f'医療文書作成件数推移(医師依頼日) {subject} {start_date:%Y/%m/%d}～{end_date:%Y/%m/%d}'
    sheet['A1'].font = Font(bold=True, size=14)

    headers = [TREND_CORNER_LABEL, *trend.month_labels, TOTAL_LABEL]
    for column_index, header in enumerate(headers, start=1):
        sheet.cell(row=3, column=column_index, value=header).font = Font(bold=True)
    rows = trend.table_rows
    for row_index, row in enumerate(rows, start=4):
        for column_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)
    sheet.column_dimensions['A'].width = 12

    # 0 件の月に線を引かないよう、グラフには 0 を空欄にした値を隠しシートに置いて使う
    chart_sheet = wb.create_sheet('グラフ用')
    chart_sheet.sheet_state = 'hidden'
    for year, counts in trend.by_year:
        chart_sheet.append([trend.year_label(year), *(count or None for count in counts)])

    chart = make_line_chart(f'{subject} 月別件数')
    chart.visible_cells_only = False
    # 1 行を 1 本の線にし、行の先頭(年の名前)を線の名前にする
    chart.add_data(Reference(chart_sheet, min_col=1, max_col=13, min_row=1, max_row=len(rows)),
                   titles_from_data=True, from_rows=True)
    chart.set_categories(Reference(sheet, min_col=2, max_col=13, min_row=3))
    # 前後が 0 件の月も見えるよう、点の印を付ける
    for series in chart.series:
        series.marker = Marker(symbol='circle', size=5)
    assert chart.legend is not None
    chart.legend.position = 'b'
    # overlay を明示しないと、Excel は凡例を月のラベルに重ねて表示する
    chart.legend.overlay = False
    sheet.add_chart(chart, f'A{len(rows) + 6}')
    set_print_layout(sheet)

    file_subject = subject.replace(' ', '_')
    return save_workbook(wb, output_dir,
                         f'医療文書作成件数推移_{file_subject}_{start_date:%Y%m%d}-{end_date:%Y%m%d}.xlsx')
