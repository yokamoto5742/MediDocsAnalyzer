import json
import sqlite3
from collections.abc import Iterator
from datetime import date
from pathlib import Path
from typing import Any

import pytest

from service.service_data_analyzer import DocumentType, count_documents
from service.service_overview import load_overview
from tests.documents_database import Document, make_documents_connection

GOLDEN_PATH = Path(__file__).parent.parent / 'fixtures' / 'source_golden.json'


def date_or_none(value: str | None) -> date | None:
    return None if value is None else date.fromisoformat(value)


@pytest.fixture
def golden() -> dict[str, Any]:
    return json.loads(GOLDEN_PATH.read_text(encoding='utf-8'))


@pytest.fixture
def connection(golden: dict[str, Any]) -> Iterator[sqlite3.Connection]:
    """見本の行を CSV2XL と同じスキーマの DB に入れる"""
    documents = []
    for row in golden['rows']:
        values = dict(zip(golden['columns'], row))
        documents.append(Document(
            date.fromisoformat(values['received_date']), date_or_none(values['requested_date']),
            values['document_name'], values['department'], values['doctor_name'], values['staff_name'],
            values['patient_id'], bool(values['deleted'])))
    connection = make_documents_connection(documents)
    yield connection
    connection.close()


def test_count_documents_matches_golden(golden: dict[str, Any], connection: sqlite3.Connection) -> None:
    expected = golden['expected']['overall']
    start_date = date.fromisoformat(expected['start_date'])
    end_date = date.fromisoformat(expected['end_date'])

    counts = count_documents(connection, start_date, end_date)
    by_document_type = {
        name: sorted([staff, department, count] for (staff, department), count
                     in count_documents(connection, start_date, end_date, DocumentType(name)).by_staff_department.items())
        for name in expected['by_document_type']}

    assert sorted([staff, department, count] for (staff, department), count
                  in counts.by_staff_department.items()) == expected['by_staff_department']
    assert counts.total == expected['total']
    assert by_document_type == expected['by_document_type']


def test_load_overview_matches_golden(golden: dict[str, Any], connection: sqlite3.Connection) -> None:
    expected = golden['expected']['overview']

    overview = load_overview(connection, date.fromisoformat(expected['today']), expected['days'])

    assert overview.start_date.isoformat() == expected['start_date']
    assert [[item.received_date.isoformat(), None if item.requested_date is None else item.requested_date.isoformat(),
             item.document_name, item.department, item.doctor_name, item.staff_name, item.patient_id]
            for item in overview.drafts] == expected['drafts']
