import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import date
from pathlib import Path

from service.service_excel_reader import Document

BUSY_TIMEOUT_SECONDS = 10

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY,
    received_date TEXT NOT NULL,
    requested_date TEXT NOT NULL,
    document_name TEXT NOT NULL,
    department TEXT NOT NULL,
    doctor_name TEXT NOT NULL,
    staff_name TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_documents_received_date ON documents (received_date);
CREATE TABLE IF NOT EXISTS staff (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    display_order INTEGER NOT NULL,
    active INTEGER NOT NULL DEFAULT 1,
    emr_id TEXT UNIQUE,
    role TEXT NOT NULL DEFAULT 'staff',
    password_hash TEXT,
    password_salt TEXT,
    must_change_password INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS departments (
    name TEXT PRIMARY KEY,
    display_order INTEGER NOT NULL
);
"""


def connect(database_path: str) -> sqlite3.Connection:
    Path(database_path).parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(database_path, timeout=BUSY_TIMEOUT_SECONDS)
    # 共有フォルダ上では WAL が使えないため、ロールバックジャーナルにする
    connection.execute('PRAGMA journal_mode=DELETE')
    connection.executescript(SCHEMA)
    return connection


@contextmanager
def open_database(database_path: str) -> Iterator[sqlite3.Connection]:
    """接続を開き、抜けるときに例外がなければコミット、あればロールバックして閉じる"""
    connection = connect(database_path)
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def count_all_documents(connection: sqlite3.Connection) -> int:
    return connection.execute('SELECT COUNT(*) FROM documents').fetchone()[0]


def count_documents_from(connection: sqlite3.Connection, start_date: date) -> int:
    return connection.execute('SELECT COUNT(*) FROM documents WHERE received_date >= ?',
                              (start_date.isoformat(),)).fetchone()[0]


def replace_documents_from(connection: sqlite3.Connection, start_date: date, documents: list[Document]) -> None:
    """預り日が start_date 以降の文書を消して、documents を入れ直す"""
    connection.execute('DELETE FROM documents WHERE received_date >= ?', (start_date.isoformat(),))
    connection.executemany(
        'INSERT INTO documents (received_date, requested_date, document_name, department, doctor_name, staff_name)'
        ' VALUES (?, ?, ?, ?, ?, ?)',
        [(document.received_date.isoformat(), document.requested_date.isoformat(), document.document_name,
          document.department, document.doctor_name, document.staff_name) for document in documents])
