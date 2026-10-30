from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from service.service_excel_reader import Document, read_documents, to_date
from tests.service.conftest import MakeWorkbook


@pytest.mark.parametrize(('value', 'expected'), [
    (datetime(2025, 4, 1, 9, 30), date(2025, 4, 1)),
    (date(2025, 4, 1), date(2025, 4, 1)),
    (45628, date(2024, 12, 2)),
    ('2025-04-01 00:00:00', date(2025, 4, 1)),
    ('2025/4/1', date(2025, 4, 1)),
    ('   ', None),
    (None, None),
])
def test_to_date(value: object, expected: date | None) -> None:
    assert to_date(value) == expected


@pytest.mark.parametrize('value', ['025/3/28', '202/3/25', '不明', '2025/13/01'])
def test_to_date_rejects_invalid_text(value: str) -> None:
    with pytest.raises(ValueError):
        to_date(value)


def test_read_documents(tmp_path: Path, make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'source.xlsx', [
        [datetime(2025, 4, 3), 123, '山田', ' 診断書 ', '内科', '佐藤', None, datetime(2025, 4, 4), '花﨑'],
        [None] * 9,
    ])

    source = read_documents(str(path))

    assert source.documents == [Document(date(2025, 4, 3), date(2025, 4, 4), '診断書', '内科', '佐藤', '花﨑')]
    assert source.errors == []
    assert source.first_received_date == date(2025, 4, 3)


def test_read_documents_skips_drafts_but_uses_their_received_date(tmp_path: Path,
                                                                  make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'source.xlsx', [
        ['2025/04/01', 1, '山田', '診断書', '内科', '佐藤', None, None, '花﨑'],
        ['2025/04/02', 2, '鈴木', '診断書', '内科', '佐藤', None, '2025/04/03', None],
        ['2025/04/05', 3, '田中', '診断書', '内科', '佐藤', None, '2025/04/06', '植田'],
    ])

    source = read_documents(str(path))

    assert [document.staff_name for document in source.documents] == ['植田']
    assert source.errors == []
    assert source.first_received_date == date(2025, 4, 1)


def test_read_documents_reports_error_rows(tmp_path: Path, make_workbook: MakeWorkbook) -> None:
    path = make_workbook(tmp_path / 'source.xlsx', [
        ['2025/04/01', 1, '山田', '診断書', '内科', '佐藤', None, '025/3/28', '花﨑'],
        ['2025/04/05', 2, '鈴木', '診断書', '内科', '佐藤', None, '2025/04/04', '花﨑'],
        ['不明', 3, '田中', '診断書', '内科', '佐藤', None, '2025/04/04', '花﨑'],
        ['2025/04/05', 4, '高橋', '診断書', '内科', '佐藤', None, '2025/04/06', '未登録'],
        ['2025/04/06', 5, '伊藤', '診断書', '内科', '佐藤', None, '2025/04/07', '花﨑'],
    ])

    source = read_documents(str(path), known_staff={'花﨑'})

    assert [(error.row_number, error.reason) for error in source.errors] == [
        (2, '医師依頼日が日付ではありません'),
        (3, '医師依頼日が預り日より前です'),
        (4, '預り日が日付ではありません'),
        (5, '担当者名がマスタに登録されていません'),
    ]
    assert source.errors[0].file_name == 'source.xlsx'
    assert source.errors[0].row_text == '2025/04/01 | 025/3/28 | 診断書 | 内科 | 佐藤 | 花﨑'
    assert len(source.documents) == 1
    assert source.first_received_date == date(2025, 4, 1)


def test_read_documents_rejects_missing_columns(tmp_path: Path) -> None:
    path = tmp_path / 'empty.xlsx'
    Workbook().save(path)

    with pytest.raises(ValueError, match='列がありません'):
        read_documents(str(path))
