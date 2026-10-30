import os
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from service.service_data_analyzer import DocumentCounts

TOTAL_LABEL = '合計'


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
