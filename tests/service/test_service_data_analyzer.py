import sqlite3
from collections.abc import Iterator
from datetime import date

import pytest

from service.service_data_analyzer import count_documents, count_staff_documents, months_between
from service.service_database import connect, replace_documents_between
from service.service_excel_reader import Document, NoDataError


def document(received: date, department: str, staff: str) -> Document:
    return Document(received, received, '診断書', department, '佐藤', staff)


@pytest.fixture
def connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    replace_documents_between(connection, date(2025, 1, 1), date(2099, 12, 31), [
        document(date(2025, 3, 31), '内科', '植田'),
        document(date(2025, 4, 5), '内科', '植田'),
        document(date(2025, 4, 6), '内科', '植田'),
        document(date(2025, 4, 1), '外科', '花﨑'),
        document(date(2025, 6, 30), '内科', '沖野'),
        document(date(2025, 7, 1), '内科', '沖野'),
    ])
    yield connection
    connection.close()


def test_count_documents_includes_both_ends(connection: sqlite3.Connection) -> None:
    counts = count_documents(connection, date(2025, 4, 1), date(2025, 6, 30))

    assert counts.by_staff_department == {('植田', '内科'): 2, ('花﨑', '外科'): 1, ('沖野', '内科'): 1}
    assert counts.by_department == {'内科': 3, '外科': 1}
    assert counts.total == 4
    assert counts.first_date == date(2025, 4, 1)
    assert counts.last_date == date(2025, 6, 30)


def test_months_between() -> None:
    assert months_between(date(2024, 11, 15), date(2025, 2, 1)) == ['2024/11', '2024/12', '2025/01', '2025/02']


def test_count_staff_documents(connection: sqlite3.Connection) -> None:
    replace_documents_between(connection, date(2025, 8, 1), date(2099, 12, 31), [
        Document(date(2025, 8, 1), date(2025, 8, 2), '紹介状', '内科', '佐藤', '植田'),
        Document(date(2025, 8, 3), date(2025, 8, 4), '紹介状', '外科', '佐藤', '植田'),
        Document(date(2025, 8, 5), date(2025, 8, 6), '診断書', '外科', '佐藤', '植田'),
        Document(date(2025, 10, 1), date(2025, 10, 2), '診断書', '外科', '佐藤', '植田'),
        Document(date(2025, 8, 1), date(2025, 8, 2), '診断書', '内科', '佐藤', '花﨑'),
    ])

    counts = count_staff_documents(connection, '植田', date(2025, 8, 1), date(2025, 10, 31), document_limit=1)

    assert counts.total == 4
    assert counts.average == 2.5
    assert counts.by_department == [('外科', 3), ('内科', 1)]
    assert counts.by_document == [('紹介状', 2)]
    assert counts.by_month == [('2025/08', 3), ('2025/09', 0), ('2025/10', 1)]


def test_count_staff_documents_without_data(connection: sqlite3.Connection) -> None:
    counts = count_staff_documents(connection, '植田', date(2026, 1, 1), date(2026, 1, 31), document_limit=10)

    assert (counts.total, counts.average, counts.by_month) == (0, 0.0, [('2026/01', 0)])


def test_count_documents_raises_when_no_data(connection: sqlite3.Connection) -> None:
    with pytest.raises(NoDataError):
        count_documents(connection, date(2026, 4, 1), date(2026, 6, 30))
