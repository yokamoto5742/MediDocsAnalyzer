import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import date
from pathlib import Path

BUSY_TIMEOUT_SECONDS = 10
# スキーマを変えたら上げ、migrate に古い版からの変更を足す
SCHEMA_VERSION = 2
# 集計の期間は預り日で決める
PERIOD_CONDITION = 'received_date BETWEEN ? AND ?'

SCHEMA = """
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

# CSV2XL の DB で集計する文書。論理削除した行と、担当者名か医師依頼日が空の行(下書き)を除く
VALID_DOCUMENTS_VIEW = """
CREATE TEMP VIEW IF NOT EXISTS valid_documents AS
SELECT * FROM documents
WHERE deleted = 0 AND IFNULL(staff_name, '') <> '' AND IFNULL(requested_date, '') <> ''
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
    if version < 2:
        # 文書は CSV2XL の DB から読むようになったため、取り込んでいた文書を消す
        connection.execute('DROP TABLE IF EXISTS documents')
    connection.execute(f'PRAGMA user_version = {SCHEMA_VERSION}')


@contextmanager
def open_database(database_path: str) -> Iterator[sqlite3.Connection]:
    """マスタの DB の接続を開き、抜けるときに例外がなければコミット、あればロールバックして閉じる"""
    connection = connect(database_path)
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def create_valid_documents_view(connection: sqlite3.Connection) -> None:
    connection.execute(VALID_DOCUMENTS_VIEW)


@contextmanager
def open_documents(database_path: str) -> Iterator[sqlite3.Connection]:
    """CSV2XL の文書の DB を読み取り専用で開く。無いファイルは作らずにエラーにする"""
    uri = Path(database_path).resolve().as_uri() + '?mode=ro'
    connection = sqlite3.connect(uri, uri=True, timeout=BUSY_TIMEOUT_SECONDS)
    try:
        create_valid_documents_view(connection)
        yield connection
    finally:
        connection.close()


def period_parameters(start_date: date, end_date: date) -> tuple[str, str]:
    """PERIOD_CONDITION に渡す値"""
    return start_date.isoformat(), end_date.isoformat()
