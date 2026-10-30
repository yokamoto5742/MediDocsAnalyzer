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


@dataclass
class StaffCounts:
    staff_name: str
    total: int
    average: float  # 期間内に件数がある担当者 1 人あたりの件数
    by_department: list[tuple[str, int]]
    by_document: list[tuple[str, int]]
    by_month: list[tuple[str, int]]


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


def months_between(start_date: date, end_date: date) -> list[str]:
    """期間に含まれる月を 'YYYY/MM' で返す"""
    months = []
    year, month = start_date.year, start_date.month
    while (year, month) <= (end_date.year, end_date.month):
        months.append(f'{year:04d}/{month:02d}')
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return months


def count_staff_documents(connection: sqlite3.Connection, staff_name: str, start_date: date, end_date: date,
                          document_limit: int) -> StaffCounts:
    """1 人の担当者について、診療科別・文書名別(多い順に document_limit 件)・月別の件数を数える"""
    parameters = (staff_name, *period_parameters(start_date, end_date))
    condition = f'staff_name = ? AND {PERIOD_CONDITION}'
    by_department = connection.execute(
        f'SELECT department, COUNT(*) AS count FROM documents WHERE {condition}'
        ' GROUP BY department ORDER BY count DESC, department', parameters).fetchall()
    by_document = connection.execute(
        f'SELECT document_name, COUNT(*) AS count FROM documents WHERE {condition}'
        ' GROUP BY document_name ORDER BY count DESC, document_name LIMIT ?',
        (*parameters, document_limit)).fetchall()
    month_counts = dict(connection.execute(
        f"SELECT strftime('%Y/%m', received_date), COUNT(*) FROM documents WHERE {condition}"
        ' GROUP BY 1', parameters).fetchall())
    staff_count, all_total = connection.execute(
        f'SELECT COUNT(DISTINCT staff_name), COUNT(*) FROM documents WHERE {PERIOD_CONDITION}',
        period_parameters(start_date, end_date)).fetchone()

    return StaffCounts(
        staff_name=staff_name,
        total=sum(count for _, count in by_department),
        average=all_total / staff_count if staff_count else 0.0,
        by_department=by_department,
        by_document=by_document,
        by_month=[(month, month_counts.get(month, 0)) for month in months_between(start_date, end_date)],
    )
