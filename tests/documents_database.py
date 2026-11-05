"""CSV2XL が作る文書の DB をテスト用に作る"""
import sqlite3
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from service.service_database import create_valid_documents_view

# CSV2XL の services/schema.sql の documents テーブル(版 2)の写し
DOCUMENTS_SCHEMA = """
CREATE TABLE documents (
    id             INTEGER PRIMARY KEY,
    received_date  TEXT    NOT NULL,
    patient_id     INTEGER NOT NULL,
    patient_name   TEXT    NOT NULL,
    document_name  TEXT    NOT NULL,
    department     TEXT    NOT NULL,
    doctor_name    TEXT    NOT NULL,
    staff_name     TEXT,
    requested_date TEXT,
    note           TEXT,
    version        INTEGER NOT NULL DEFAULT 1,
    deleted        INTEGER NOT NULL DEFAULT 0,
    created_at     TEXT    NOT NULL,
    updated_at     TEXT    NOT NULL
);
PRAGMA user_version = 2;
"""
CREATED_AT = '2026-11-05T00:00:00'


@dataclass(frozen=True)
class Document:
    """CSV2XL の DB の 1 行。担当者名か医師依頼日が None なら下書き"""
    received_date: date
    requested_date: date | None
    document_name: str
    department: str
    doctor_name: str
    staff_name: str | None
    patient_id: int = 1
    deleted: bool = False


def insert_documents(connection: sqlite3.Connection, documents: Sequence[Document]) -> None:
    connection.executemany(
        'INSERT INTO documents (received_date, patient_id, patient_name, document_name, department, doctor_name,'
        ' staff_name, requested_date, deleted, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
        [(document.received_date.isoformat(), document.patient_id, '患者', document.document_name,
          document.department, document.doctor_name, document.staff_name,
          None if document.requested_date is None else document.requested_date.isoformat(),
          int(document.deleted), CREATED_AT, CREATED_AT) for document in documents])


def make_documents_connection(documents: Sequence[Document] = ()) -> sqlite3.Connection:
    """メモリ上に文書の DB を作り、集計で使うビューも作る"""
    connection = sqlite3.connect(':memory:')
    connection.executescript(DOCUMENTS_SCHEMA)
    insert_documents(connection, documents)
    create_valid_documents_view(connection)
    return connection


def make_documents_database(path: Path, documents: Sequence[Document] = ()) -> Path:
    connection = sqlite3.connect(path)
    with connection:
        connection.executescript(DOCUMENTS_SCHEMA)
        insert_documents(connection, documents)
    connection.close()
    return path
