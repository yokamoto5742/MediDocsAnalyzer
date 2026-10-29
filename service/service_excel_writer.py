import os
import shutil
from pathlib import Path

import polars as pl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import DocumentCounts
from service.service_excel_reader import COLUMN_COUNT, PATIENT_ID_COLUMN

CENTER_COLUMNS = (1, 2, 5, 6, 7, 8)  # A, B, E, F, G, H列
TOTAL_LABEL = '合計'


def to_database_value(header: str, text: str) -> str | int | None:
    if text == '':
        return None
    if header == PATIENT_ID_COLUMN and text.isdigit():
        return int(text)
    return text


def apply_cell_formats(sheet: Worksheet, last_row: int) -> None:
    for row in range(2, last_row + 1):
        for column in range(1, COLUMN_COUNT + 1):
            if column in CENTER_COLUMNS:
                alignment = Alignment(horizontal='center', vertical='center')
            else:
                alignment = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)
            sheet.cell(row=row, column=column).alignment = alignment


def write_database(df: pl.DataFrame, file_path: str) -> None:
    """DB を df の内容で書き換える。既存ファイルなら書式を残して値だけ入れ替える"""
    wb = load_workbook(file_path) if Path(file_path).exists() else Workbook()
    sheet = wb.active
    assert sheet is not None

    for row in sheet.iter_rows(max_col=COLUMN_COUNT):
        for cell in row:
            cell.value = None

    for column_index, header in enumerate(df.columns, start=1):
        sheet.cell(row=1, column=column_index, value=header)
    for row_index, row in enumerate(df.iter_rows(), start=2):
        for column_index, (header, text) in enumerate(zip(df.columns, row), start=1):
            sheet.cell(row=row_index, column=column_index, value=to_database_value(header, text))

    apply_cell_formats(sheet, df.height + 1)
    wb.save(file_path)


def backup_file(file_path: str, backup_dir: str) -> str:
    backup_dir_path = Path(backup_dir)
    backup_dir_path.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir_path / f'backup_{Path(file_path).name}'
    shutil.copy2(file_path, backup_path)
    return str(backup_path)


def write_counts(sheet: Worksheet, counts: DocumentCounts, staff_members: list[str],
                 departments: list[str]) -> None:
    """2 行目に診療科、A列に担当者を並べ、件数を入れる。'合計' 列には行の合計を入れる"""
    sheet['A2'] = '氏名'
    for column_index, department in enumerate(departments, start=2):
        sheet.cell(row=2, column=column_index, value=department)

    total_row = len(staff_members) + 3
    for row_index, staff in enumerate(staff_members, start=3):
        sheet.cell(row=row_index, column=1, value=staff)
    sheet.cell(row=total_row, column=1, value=TOTAL_LABEL)

    for column_index, department in enumerate(departments, start=2):
        if department == TOTAL_LABEL:
            continue
        for row_index, staff in enumerate(staff_members, start=3):
            sheet.cell(row=row_index, column=column_index,
                       value=counts.by_staff_department.get((staff, department), 0))
        sheet.cell(row=total_row, column=column_index, value=counts.by_department.get(department, 0))

    if TOTAL_LABEL in departments:
        total_column = departments.index(TOTAL_LABEL) + 2
        listed_departments = [department for department in departments if department != TOTAL_LABEL]
        for row_index, staff in enumerate(staff_members, start=3):
            staff_total = sum(counts.by_staff_department.get((staff, department), 0)
                              for department in listed_departments)
            sheet.cell(row=row_index, column=total_column, value=staff_total)
        # 一覧にない担当者・診療科の件数も含めた全体の件数
        sheet.cell(row=total_row, column=total_column, value=counts.total)


def save_results(counts: DocumentCounts, staff_members: list[str], departments: list[str],
                 template_path: str, output_dir: str) -> str:
    """集計結果をテンプレートに書いて保存し、Excel で開いて保存先のパスを返す"""
    wb = load_workbook(template_path)
    sheet = wb.active
    assert sheet is not None

    sheet['A1'] = (f'医療文書作成件数 {counts.first_date:%Y年%m月%d日}'
                   f'-{counts.last_date:%Y年%m月%d日}')
    write_counts(sheet, counts, staff_members, departments)

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    output_file_path = output_path / (
        f'医療文書作成件数{counts.first_date:%Y%m%d}-{counts.last_date:%Y%m%d}.xlsx')
    wb.save(output_file_path)

    os.startfile(output_file_path)

    return str(output_file_path)
