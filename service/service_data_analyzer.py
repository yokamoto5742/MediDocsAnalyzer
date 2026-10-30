import sqlite3
from collections import Counter
from dataclasses import dataclass
from datetime import date

from service.service_excel_reader import NoDataError

PERIOD_CONDITION = 'received_date BETWEEN ? AND ?'


@dataclass
class DocumentCounts:
    by_staff_department: dict[tuple[str, str], int]
    by_department: dict[str, int]
    total: int
    first_date: date
    last_date: date


def period_parameters(start_date: date, end_date: date) -> tuple[str, str]:
    return start_date.isoformat(), end_date.isoformat()


def count_documents(connection: sqlite3.Connection, start_date: date, end_date: date) -> DocumentCounts:
    """預り日が期間内の文書を、担当者×診療科、診療科ごと、全体で数える"""
    parameters = period_parameters(start_date, end_date)
    by_staff_department = {
        (staff, department): count for staff, department, count in connection.execute(
            f'SELECT staff_name, department, COUNT(*) FROM documents WHERE {PERIOD_CONDITION}'
            ' GROUP BY staff_name, department', parameters)
    }
    if not by_staff_department:
        raise NoDataError

    by_department: Counter[str] = Counter()
    for (_, department), count in by_staff_department.items():
        by_department[department] += count
    first_date, last_date = connection.execute(
        f'SELECT MIN(received_date), MAX(received_date) FROM documents WHERE {PERIOD_CONDITION}',
        parameters).fetchone()

    return DocumentCounts(
        by_staff_department=by_staff_department,
        by_department=dict(by_department),
        total=sum(by_staff_department.values()),
        first_date=date.fromisoformat(first_date),
        last_date=date.fromisoformat(last_date),
    )
