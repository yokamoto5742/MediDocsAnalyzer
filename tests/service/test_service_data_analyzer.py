import sqlite3
from collections.abc import Iterator
from datetime import date

import pytest

from service.service_data_analyzer import count_documents
from service.service_database import connect, replace_documents_from
from service.service_excel_reader import Document, NoDataError


def document(received: date, department: str, staff: str) -> Document:
    return Document(received, received, '診断書', department, '佐藤', staff)


@pytest.fixture
def connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    replace_documents_from(connection, date(2025, 1, 1), [
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


def test_count_documents_raises_when_no_data(connection: sqlite3.Connection) -> None:
    with pytest.raises(NoDataError):
        count_documents(connection, date(2026, 4, 1), date(2026, 6, 30))
