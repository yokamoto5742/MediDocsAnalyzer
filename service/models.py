from dataclasses import dataclass, field
from datetime import date


class NoDataError(Exception):
    """対象となるデータが 1 件もない"""


@dataclass(frozen=True)
class Document:
    received_date: date
    requested_date: date
    document_name: str
    department: str
    doctor_name: str
    staff_name: str
    patient_id: str = ''


@dataclass(frozen=True)
class Draft:
    """医師依頼日か担当者名が空の行(下書き)"""
    received_date: date
    requested_date: date | None
    document_name: str
    department: str
    doctor_name: str
    staff_name: str
    patient_id: str


@dataclass(frozen=True)
class RowError:
    file_name: str
    row_number: int
    patient_id: str
    reason: str
    row_text: str


@dataclass
class SourceData:
    documents: list[Document] = field(default_factory=list)
    errors: list[RowError] = field(default_factory=list)
    drafts: list[Draft] = field(default_factory=list)
    first_received_date: date | None = None
    last_received_date: date | None = None
