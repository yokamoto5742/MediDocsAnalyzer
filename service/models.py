from dataclasses import dataclass
from datetime import date


class NoDataError(Exception):
    """対象となるデータが 1 件もない"""


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
