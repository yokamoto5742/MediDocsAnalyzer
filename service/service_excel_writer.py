import re
from datetime import date
from pathlib import Path
from typing import cast

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import TOTAL_LABEL, DocumentFilter, OverallSummary, StaffCounts, TrendCounts

OVERALL_TEMPLATE_NAME = '医療文書作成件数.xlsx'
TREND_TEMPLATE_NAME = '医療文書作成件数推移.xlsx'
# テンプレートの表は 2 行目が見出しで、3 行目から 1 行 1 件(担当者・年)が並ぶ
HEADER_ROW = 2
FIRST_DATA_ROW = 3
# 推移のテンプレートは B〜M 列が 12 か月、N 列が合計
FIRST_MONTH_COLUMN = 2
DEFAULT_STAFF_FONT_SIZE = 12
TITLE_FONT_SIZE = 16
# 文字が切れない行の高さ(pt)の、文字の大きさに対する割合
ROW_HEIGHT_RATIO = 1.55
TITLE_ROW_HEIGHT = 24
THIN_SIDE = Side(style='thin')
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def load_template(template_dir: str, file_name: str) -> tuple[Workbook, Worksheet]:
    """テンプレートを読み、ブックと表のシート(先頭のシート)を返す。グラフは別のシートにあり、件数を入れると Excel が描き直す"""
    wb = load_workbook(Path(template_dir) / file_name)
    return wb, wb.worksheets[0]


def fill_title(sheet: Worksheet, values: dict[str, str]) -> str:
    """A1 の {名前} を values で置き換え、置き換えた表題を返す。表題はファイル名にも使う"""
    title = str(sheet['A1'].value)
    for name, value in values.items():
        title = title.replace(f'{{{name}}}', value)
    sheet['A1'] = title
    return title


def period_text(start_date: date, end_date: date) -> str:
    return f'{start_date:%Y%m%d}-{end_date:%Y%m%d}'


def data_rows(sheet: Worksheet) -> list[int]:
    """A 列に名前がある、3 行目から続く行の番号"""
    rows = []
    row = FIRST_DATA_ROW
    while sheet.cell(row=row, column=1).value not in (None, ''):
        rows.append(row)
        row += 1
    return rows


def save_results(summary: OverallSummary, template_dir: str, output_dir: str) -> str:
    """全体のテンプレートの氏名・見出し(合計と診療科)の位置に件数を入れて保存し、保存先のパスを返す。
    テンプレートに無い担当者・診療科は出さない"""
    wb, sheet = load_template(template_dir, OVERALL_TEMPLATE_NAME)
    counts = summary.counts
    title = fill_title(sheet, {'集計対象': summary.document_type.value,
                               '対象期間': period_text(counts.first_date, counts.last_date)})

    headers = [TOTAL_LABEL, *summary.departments]
    counts_by_name = {str(row[0]): dict(zip(headers, row[1:])) for row in summary.rows}
    columns = range(2, sheet.max_column + 1)
    for row in data_rows(sheet):
        row_counts = counts_by_name.get(str(sheet.cell(row=row, column=1).value), {})
        for column in columns:
            header = sheet.cell(row=HEADER_ROW, column=column).value
            if header not in (None, ''):
                sheet.cell(row=row, column=column, value=row_counts.get(str(header), 0))

    return save_workbook(wb, output_dir, f'{title}.xlsx')


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


def write_table(sheet: Worksheet, row: int, column: int, headers: tuple[str, str],
                rows: list[tuple[str, int]], font_size: int) -> None:
    write_table_row(sheet, row, column, list(headers), font_size, bold=True)
    for row_offset, (label, count) in enumerate(rows, start=1):
        write_table_row(sheet, row + row_offset, column, [label, count], font_size)


def write_table_row(sheet: Worksheet, row: int, column: int, values: list[str | int | None], font_size: int,
                    bold: bool = False) -> None:
    """表の 1 行分を font_size の文字で書き、文字が切れないよう行の高さを広げる"""
    sheet.row_dimensions[row].height = font_size * ROW_HEIGHT_RATIO
    for offset, value in enumerate(values):
        cell = sheet.cell(row=row, column=column + offset, value=value)
        cell.font = Font(bold=bold, size=font_size)
        cell.border = THIN_BORDER


def save_staff_results(staff_counts: StaffCounts, start_date: date, end_date: date, output_dir: str,
                       font_size: int = DEFAULT_STAFF_FONT_SIZE) -> str:
    """1 人分の集計(診療科別・文書名別・月別)の表を新しいブックに書いて保存する"""
    wb = Workbook()
    sheet = cast(Worksheet, wb.active)
    sheet.title = '個人'
    name = staff_counts.staff_name
    document_type = staff_counts.document_type.value
    sheet['A1'] = f'医療文書作成件数 {name} {document_type} {start_date:%Y年%m月%d日}-{end_date:%Y年%m月%d日}'
    sheet['A1'].font = Font(bold=True, size=TITLE_FONT_SIZE)
    sheet.row_dimensions[1].height = TITLE_ROW_HEIGHT
    write_table_row(sheet, 2, 1, [TOTAL_LABEL, staff_counts.total], font_size)

    write_table(sheet, 4, 1, ('診療科', '件数'), staff_counts.by_department, font_size)
    write_table(sheet, 4, 4, ('文書名', '件数'), staff_counts.by_document, font_size)
    write_table(sheet, 4, 7, ('月', '件数'), staff_counts.by_month, font_size)
    for column, width in (('A', 18), ('D', 60), ('G', 13)):
        sheet.column_dimensions[column].width = width
    set_print_layout(sheet)

    return save_workbook(wb, output_dir,
                         f'医療文書作成件数_{name}_{document_type}_{start_date:%Y%m%d}-{end_date:%Y%m%d}.xlsx')


def template_year(label: object) -> int | None:
    """テンプレートの年の名前(「2025年」など)の先頭の数字"""
    matched = re.match(r'\d+', str(label))
    return int(matched.group()) if matched else None


def save_trend_results(trend: TrendCounts, document_filter: DocumentFilter, staff_name: str | None,
                       start_date: date, end_date: date, template_dir: str, output_dir: str) -> str:
    """推移のテンプレートの年の行に、月別件数と合計を入れて保存し、保存先のパスを返す。
    月の見出しと年の名前は年の始まりの月に合わせて書き換える。テンプレートに無い年は出さない。
    0 件の月は空欄にして、グラフの線を途切れさせる"""
    wb, sheet = load_template(template_dir, TREND_TEMPLATE_NAME)
    if staff_name is None:
        sheet['A1'] = str(sheet['A1'].value).replace('_{氏名}', '')
    title = fill_title(sheet, {'氏名': staff_name or '', '集計対象': document_filter.document_type.value,
                               '診療科': document_filter.department or TOTAL_LABEL,
                               '対象期間': period_text(start_date, end_date)})

    for offset, label in enumerate(trend.month_labels):
        sheet.cell(row=HEADER_ROW, column=FIRST_MONTH_COLUMN + offset, value=label)
    counts_by_year = dict(trend.by_year)
    for row in data_rows(sheet):
        year = template_year(sheet.cell(row=row, column=1).value)
        if year is None:
            continue
        sheet.cell(row=row, column=1, value=trend.year_label(year))
        counts = counts_by_year.get(year)
        months = [count or None for count in counts] if counts is not None else [None] * 12
        total = sum(count for count in counts if count is not None) if counts is not None else None
        for offset, value in enumerate([*months, total]):
            sheet.cell(row=row, column=FIRST_MONTH_COLUMN + offset, value=value)

    return save_workbook(wb, output_dir, f'{title}.xlsx')
