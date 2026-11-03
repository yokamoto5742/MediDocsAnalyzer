import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook

from service.service_data_analyzer import DocumentType, count_documents
from service.service_database import connect, replace_documents_between
from service.service_excel_reader import read_documents

GOLDEN_PATH = Path(__file__).parent.parent / 'fixtures' / 'source_golden.json'


def to_cell(value: Any) -> object:
    """見本のセルの値を openpyxl に書く値にする。{"datetime": ...} だけ日時に変える"""
    if isinstance(value, dict):
        return datetime.fromisoformat(value['datetime'])
    return value


def iso_or_none(value: date | None) -> str | None:
    return None if value is None else value.isoformat()


@pytest.fixture
def golden() -> dict[str, Any]:
    return json.loads(GOLDEN_PATH.read_text(encoding='utf-8'))


@pytest.fixture
def golden_path(tmp_path: Path, golden: dict[str, Any]) -> Path:
    wb = Workbook()
    sheet = wb.active
    assert sheet is not None
    sheet.append(golden['headers'])
    for row in golden['rows']:
        sheet.append([to_cell(value) for value in row])
    path = tmp_path / '医療文書担当一覧_202504300000.xlsm'
    wb.save(path)
    return path


def test_read_documents_matches_golden(golden: dict[str, Any], golden_path: Path) -> None:
    source = read_documents(str(golden_path), set(golden['known_staff']))
    expected = golden['expected']

    assert [[item.received_date.isoformat(), item.requested_date.isoformat(), item.document_name,
             item.department, item.doctor_name, item.staff_name, item.patient_id]
            for item in source.documents] == expected['documents']
    assert [[item.received_date.isoformat(), iso_or_none(item.requested_date), item.document_name,
             item.department, item.doctor_name, item.staff_name, item.patient_id]
            for item in source.drafts] == expected['drafts']
    assert [[error.row_number, error.patient_id, error.reason] for error in source.errors] == expected['errors']
    assert iso_or_none(source.first_received_date) == expected['first_received_date']
    assert iso_or_none(source.last_received_date) == expected['last_received_date']


def test_count_documents_matches_golden(golden: dict[str, Any], golden_path: Path) -> None:
    source = read_documents(str(golden_path), set(golden['known_staff']))
    expected = golden['expected']['overall']
    connection = connect(':memory:')
    assert source.first_received_date is not None and source.last_received_date is not None
    replace_documents_between(connection, source.first_received_date, source.last_received_date, source.documents)

    start_date = date.fromisoformat(expected['start_date'])
    end_date = date.fromisoformat(expected['end_date'])
    counts = count_documents(connection, start_date, end_date)
    by_document_type = {
        name: sorted([staff, department, count] for (staff, department), count
                     in count_documents(connection, start_date, end_date, DocumentType(name)).by_staff_department.items())
        for name in expected['by_document_type']}
    connection.close()

    assert sorted([staff, department, count] for (staff, department), count
                  in counts.by_staff_department.items()) == expected['by_staff_department']
    assert counts.total == expected['total']
    assert by_document_type == expected['by_document_type']
