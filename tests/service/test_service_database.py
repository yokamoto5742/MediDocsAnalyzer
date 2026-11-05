import sqlite3
from datetime import date
from pathlib import Path

import pytest

from service.service_database import SCHEMA_VERSION, open_database, open_documents
from tests.documents_database import Document, make_documents_database


def test_open_database_sets_schema_version(tmp_path: Path) -> None:
    with open_database(str(tmp_path / 'database.db')) as connection:
        assert connection.execute('PRAGMA user_version').fetchone()[0] == SCHEMA_VERSION


def test_open_database_rejects_newer_schema(tmp_path: Path) -> None:
    path = tmp_path / 'database.db'
    newer_connection = sqlite3.connect(path)
    newer_connection.execute(f'PRAGMA user_version = {SCHEMA_VERSION + 1}')
    newer_connection.close()

    with pytest.raises(RuntimeError, match='新しい'):
        with open_database(str(path)):
            pass
    # 新しい版の DB にはテーブルを足さない
    check_connection = sqlite3.connect(path)
    assert check_connection.execute("SELECT COUNT(*) FROM sqlite_master WHERE type = 'table'").fetchone()[0] == 0
    check_connection.close()


def test_open_database_creates_file_with_rollback_journal(tmp_path: Path) -> None:
    path = tmp_path / 'nested' / 'database.db'

    with open_database(str(path)) as connection:
        assert connection.execute('PRAGMA journal_mode').fetchone()[0] == 'delete'

    assert path.exists()


def test_open_database_drops_imported_documents_of_old_database(tmp_path: Path) -> None:
    path = tmp_path / 'database.db'
    old_connection = sqlite3.connect(path)
    old_connection.execute('CREATE TABLE documents (id INTEGER PRIMARY KEY)')
    old_connection.execute('CREATE TABLE staff (id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE,'
                           ' display_order INTEGER NOT NULL)')
    old_connection.execute("INSERT INTO staff (name, display_order) VALUES ('植田', 1)")
    old_connection.execute('PRAGMA user_version = 1')
    old_connection.commit()
    old_connection.close()

    with open_database(str(path)) as connection:
        tables = {name for (name,) in connection.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
        assert tables == {'staff', 'departments'}
        assert connection.execute('SELECT name FROM staff').fetchall() == [('植田',)]


def test_open_database_rolls_back_on_error(tmp_path: Path) -> None:
    path = str(tmp_path / 'database.db')

    with pytest.raises(RuntimeError):
        with open_database(path) as connection:
            connection.execute("INSERT INTO departments (name, display_order) VALUES ('内科', 1)")
            raise RuntimeError

    with open_database(path) as connection:
        assert connection.execute('SELECT COUNT(*) FROM departments').fetchone()[0] == 0


def test_open_documents_reads_only_valid_documents(tmp_path: Path) -> None:
    path = make_documents_database(tmp_path / '医療文書担当一覧.db', [
        Document(date(2025, 4, 1), date(2025, 4, 2), '診断書', '内科', '佐藤', '植田', patient_id=1),
        Document(date(2025, 4, 1), date(2025, 4, 2), '診断書', '内科', '佐藤', '植田', patient_id=2, deleted=True),
        Document(date(2025, 4, 1), None, '診断書', '内科', '佐藤', '植田', patient_id=3),
        Document(date(2025, 4, 1), date(2025, 4, 2), '診断書', '内科', '佐藤', None, patient_id=4),
    ])

    with open_documents(str(path)) as connection:
        assert connection.execute('SELECT patient_id FROM valid_documents').fetchall() == [(1,)]


def test_open_documents_does_not_write(tmp_path: Path) -> None:
    path = make_documents_database(tmp_path / 'documents.db')

    with open_documents(str(path)) as connection:
        with pytest.raises(sqlite3.OperationalError):
            connection.execute("UPDATE documents SET note = 'x'")


def test_open_documents_does_not_create_missing_file(tmp_path: Path) -> None:
    path = tmp_path / 'missing.db'

    with pytest.raises(sqlite3.OperationalError):
        with open_documents(str(path)):
            pass
    assert not path.exists()
