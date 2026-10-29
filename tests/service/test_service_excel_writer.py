from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import polars as pl
import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from service import service_excel_writer
from service.service_data_analyzer import DocumentCounts
from service.service_excel_writer import backup_file, save_results, to_database_value, write_database
from tests.service.conftest import HEADERS, MakeWorkbook


def active_sheet(path: Path | str) -> Worksheet:
    sheet = load_workbook(path).active
    assert isinstance(sheet, Worksheet)
    return sheet


def sheet_values(sheet: Worksheet) -> list[list[object]]:
    return [list(row) for row in sheet.iter_rows(values_only=True)]


@pytest.mark.parametrize(('header', 'text', 'expected'), [
    ('患者ID', '123', 123),
    ('患者ID', 'A-1', 'A-1'),
    ('預り日', '2025/04/01', '2025/04/01'),
    ('備考', '', None),
])
def test_to_database_value(header: str, text: str, expected: object) -> None:
    assert to_database_value(header, text) == expected


def test_write_database_creates_new_file(tmp_path: Path) -> None:
    df = pl.DataFrame([('2025/04/01', '123', '山田', '診断書', '内科', '佐藤', '', '2025/04/02', '花﨑')],
                      schema={header: pl.String for header in HEADERS}, orient='row')
    path = tmp_path / 'database.xlsx'

    write_database(df, str(path))

    sheet = active_sheet(path)
    assert sheet_values(sheet) == [
        HEADERS,
        ['2025/04/01', 123, '山田', '診断書', '内科', '佐藤', None, '2025/04/02', '花﨑'],
    ]
    assert sheet['A2'].alignment.horizontal == 'center'
    assert sheet['C2'].alignment.horizontal == 'left'
    assert sheet['C2'].alignment.shrink_to_fit


def test_write_database_replaces_values_and_keeps_formats(tmp_path: Path, make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'database.xlsx', [
        ['2025/03/01', 1, '旧', '旧', '内科', '佐藤', None, '2025/03/02', '植田'],
        ['2025/03/02', 2, '旧', '旧', '内科', '佐藤', None, '2025/03/03', '植田'],
    ])
    wb = load_workbook(path)
    sheet = wb.active
    assert sheet is not None
    sheet.column_dimensions['C'].width = 30
    wb.save(path)
    df = pl.DataFrame([('2025/04/01', '9', '新', '新', '外科', '田中', '', '2025/04/02', '沖野')],
                      schema={header: pl.String for header in HEADERS}, orient='row')

    write_database(df, str(path))

    sheet = active_sheet(path)
    values = [row for row in sheet_values(sheet) if any(value is not None for value in row)]
    assert values == [HEADERS, ['2025/04/01', 9, '新', '新', '外科', '田中', None, '2025/04/02', '沖野']]
    assert sheet.column_dimensions['C'].width == 30


def test_backup_file(tmp_path: Path) -> None:
    path = tmp_path / 'database.xlsx'
    path.write_bytes(b'data')

    backup_path = backup_file(str(path), str(tmp_path / 'backup' / 'nested'))

    assert Path(backup_path) == tmp_path / 'backup' / 'nested' / 'backup_database.xlsx'
    assert Path(backup_path).read_bytes() == b'data'


def test_save_results(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    startfile = MagicMock()
    monkeypatch.setattr(service_excel_writer.os, 'startfile', startfile, raising=False)
    template_path = tmp_path / 'template.xlsx'
    Workbook().save(template_path)
    counts = DocumentCounts(
        by_staff_department={('植田', '内科'): 2, ('花﨑', '外科'): 1, ('外部', '内科'): 1, ('植田', '眼科'): 5},
        by_department={'内科': 3, '外科': 1, '眼科': 5},
        total=9,
        first_date=date(2025, 4, 1),
        last_date=date(2025, 5, 20),
    )

    output_path = save_results(counts, ['植田', '花﨑'], ['合計', '内科', '外科'],
                               str(template_path), str(tmp_path / 'output'))

    assert Path(output_path) == tmp_path / 'output' / '医療文書作成件数20250401-20250520.xlsx'
    assert sheet_values(active_sheet(output_path)) == [
        ['医療文書作成件数 2025年04月01日-2025年05月20日', None, None, None],
        ['氏名', '合計', '内科', '外科'],
        ['植田', 2, 2, 0],
        ['花﨑', 1, 0, 1],
        ['合計', 9, 3, 1],
    ]
    startfile.assert_called_once_with(Path(output_path))
