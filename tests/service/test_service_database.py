from datetime import date
from pathlib import Path

import pytest

from service.service_database import (count_all_documents, count_documents_between, open_database,
                                      replace_documents_between)
from service.service_excel_reader import Document


def document(received: date, staff: str = '植田') -> Document:
    return Document(received, received, '診断書', '内科', '佐藤', staff)


def test_open_database_creates_file_with_rollback_journal(tmp_path: Path) -> None:
    path = tmp_path / 'nested' / 'database.db'

    with open_database(str(path)) as connection:
        assert connection.execute('PRAGMA journal_mode').fetchone()[0] == 'delete'

    assert path.exists()


def test_replace_documents_between_keeps_documents_outside_range(tmp_path: Path) -> None:
    path = str(tmp_path / 'database.db')
    with open_database(path) as connection:
        replace_documents_between(connection, date(2025, 1, 1), date(2025, 12, 31), [
            document(date(2025, 3, 31)), document(date(2025, 4, 1)), document(date(2025, 4, 2)),
            document(date(2025, 4, 3))])

    with open_database(path) as connection:
        assert count_documents_between(connection, date(2025, 4, 1), date(2025, 4, 2)) == 2
        replace_documents_between(connection, date(2025, 4, 1), date(2025, 4, 2), [document(date(2025, 4, 1), '花﨑')])

    with open_database(path) as connection:
        assert connection.execute(
            'SELECT received_date, staff_name FROM documents ORDER BY received_date').fetchall() == [
            ('2025-03-31', '植田'), ('2025-04-01', '花﨑'), ('2025-04-03', '植田')]


def test_open_database_rolls_back_on_error(tmp_path: Path) -> None:
    path = str(tmp_path / 'database.db')

    with pytest.raises(RuntimeError):
        with open_database(path) as connection:
            replace_documents_between(connection, date(2025, 1, 1), date(2025, 12, 31), [document(date(2025, 4, 1))])
            raise RuntimeError

    with open_database(path) as connection:
        assert count_all_documents(connection) == 0
