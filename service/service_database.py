import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import date
from pathlib import Path

from service.models import Document

BUSY_TIMEOUT_SECONDS = 10
# スキーマを変えたら上げ、migrate に古い版からの変更を足す
SCHEMA_VERSION = 1
# 集計・取り込みの対象にする期間は預り日で決める
PERIOD_CONDITION = 'received_date BETWEEN ? AND ?'

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY,
    received_date TEXT NOT NULL,
    requested_date TEXT NOT NULL,
    document_name TEXT NOT NULL,
    department TEXT NOT NULL,
    doctor_name TEXT NOT NULL,
    staff_name TEXT NOT NULL,
    patient_id TEXT NOT NULL DEFAULT ''
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
    try:
        migrate(connection)
    except Exception:
        connection.close()
        raise
    return connection


def migrate(connection: sqlite3.Connection) -> None:
    """PRAGMA user_version で DB の版を見て、テーブルを作り、古い DB を今のスキーマにそろえる。
    ほかのアプリ(移植版)が新しい版にした DB は、壊さないよう開かない"""
    version = connection.execute('PRAGMA user_version').fetchone()[0]
    if version > SCHEMA_VERSION:
        raise RuntimeError(f'データベースの版({version})がこのアプリ(版 {SCHEMA_VERSION})より新しいため開けません')
    connection.executescript(SCHEMA)
    if version < 1:
        _add_patient_id_column(connection)
    connection.execute(f'PRAGMA user_version = {SCHEMA_VERSION}')


def _add_patient_id_column(connection: sqlite3.Connection) -> None:
    """患者IDの列がない古い DB に列を足す"""
    columns = [row[1] for row in connection.execute('PRAGMA table_info(documents)')]
    if 'patient_id' not in columns:
        connection.execute("ALTER TABLE documents ADD COLUMN patient_id TEXT NOT NULL DEFAULT ''")


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


def period_parameters(start_date: date, end_date: date) -> tuple[str, str]:
    """PERIOD_CONDITION に渡す値"""
    return start_date.isoformat(), end_date.isoformat()


def count_documents_between(connection: sqlite3.Connection, start_date: date, end_date: date) -> int:
    return connection.execute(f'SELECT COUNT(*) FROM documents WHERE {PERIOD_CONDITION}',
                              period_parameters(start_date, end_date)).fetchone()[0]


def replace_documents_between(connection: sqlite3.Connection, start_date: date, end_date: date,
                              documents: list[Document]) -> None:
    """預り日が start_date〜end_date の文書を消して、documents を入れ直す"""
    connection.execute(f'DELETE FROM documents WHERE {PERIOD_CONDITION}', period_parameters(start_date, end_date))
    connection.executemany(
        'INSERT INTO documents (received_date, requested_date, document_name, department, doctor_name, staff_name,'
        ' patient_id) VALUES (?, ?, ?, ?, ?, ?, ?)',
        [(document.received_date.isoformat(), document.requested_date.isoformat(), document.document_name,
          document.department, document.doctor_name, document.staff_name, document.patient_id)
         for document in documents])
