from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from service.service_excel_reader import read_documents, to_cell_text
from tests.service.conftest import HEADERS, MakeWorkbook


@pytest.mark.parametrize(('header', 'value', 'expected'), [
    ('預り日', datetime(2025, 4, 1, 9, 30), '2025/04/01'),
    ('預り日', date(2025, 4, 1), '2025/04/01'),
    ('医師依頼日', '2025-04-01 00:00:00', '2025/04/01'),
    ('医師依頼日', '2025/04/01', '2025/04/01'),
    ('医師依頼日', '   ', ''),
    ('患者ID', '00123', '123'),
    ('患者ID', 456, '456'),
    ('患者ID', 'A-1', 'A-1'),
    ('担当者名', None, ''),
    ('担当者名', '', ''),
    ('備考', 12, '12'),
])
def test_to_cell_text(header: str, value: object, expected: str) -> None:
    assert to_cell_text(header, value) == expected


def test_read_documents(tmp_path: Path, make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'source.xlsx', [
        [datetime(2025, 4, 3), '00123', '山田', '診断書', '内科', '佐藤', None, datetime(2025, 4, 4), '花﨑'],
        [None] * 9,
    ])

    df = read_documents(str(path))

    assert df.columns == HEADERS
    assert df.row(0) == ('2025/04/03', '123', '山田', '診断書', '内科', '佐藤', '', '2025/04/04', '花﨑')
    assert df.row(1) == ('',) * 9


def test_read_documents_ignores_columns_after_i(tmp_path: Path, make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'source.xlsx', [
        ['2025/04/03', 1, '山田', '診断書', '内科', '佐藤', None, '2025/04/04', '植田', '余分'],
    ])

    df = read_documents(str(path))

    assert df.width == 9
    assert df.row(0)[-1] == '植田'


def test_read_documents_empty_sheet(tmp_path: Path) -> None:
    path = tmp_path / 'empty.xlsx'
    Workbook().save(path)

    df = read_documents(str(path))

    assert df.height == 0
