import sqlite3
from collections import Counter
from dataclasses import dataclass
from datetime import date
from enum import Enum

from service.models import NoDataError
from service.service_database import PERIOD_CONDITION, period_parameters
from utils.date_period import fiscal_year, get_period_range

TOTAL_LABEL = '合計'
# 文書名にこの語を含む文書を「主治医意見書」とする
OPINION_LABEL = '主治医意見書'


class DocumentType(Enum):
    """集計対象。値は画面と Excel に出す名前"""
    ALL = '全書類'
    MEDICAL = '医療文書'  # 主治医意見書を除いた全書類
    OPINION = OPINION_LABEL


def document_type_condition(document_type: DocumentType) -> str:
    """WHERE 句に足す集計対象の条件。全書類なら空文字"""
    if document_type == DocumentType.OPINION:
        return f" AND document_name LIKE '%{OPINION_LABEL}%'"
    if document_type == DocumentType.MEDICAL:
        return f" AND document_name NOT LIKE '%{OPINION_LABEL}%'"
    return ''


@dataclass(frozen=True)
class DocumentFilter:
    """推移で数える文書の絞り込み。既定は全診療科の全書類"""
    department: str | None = None
    document_type: DocumentType = DocumentType.ALL


@dataclass
class DocumentCounts:
    by_staff_department: dict[tuple[str, str], int]
    by_department: dict[str, int]
    total: int
    first_date: date
    last_date: date


@dataclass
class OverallSummary:
    counts: DocumentCounts
    staff_members: list[str]
    departments: list[str]
    document_type: DocumentType = DocumentType.ALL

    @property
    def rows(self) -> list[list[str | int]]:
        """全体の表の行。[氏名, 合計, 診療科ごとの件数...] を担当者ごとに並べ、最後に合計行を置く。
        担当者の合計は並べた診療科だけ、合計行の合計は一覧にない担当者・診療科も含めた全体の件数"""
        counts = self.counts
        rows: list[list[str | int]] = []
        for staff in self.staff_members:
            department_counts = [counts.by_staff_department.get((staff, department), 0)
                                 for department in self.departments]
            rows.append([staff, sum(department_counts), *department_counts])
        rows.append([TOTAL_LABEL, counts.total,
                     *(counts.by_department.get(department, 0) for department in self.departments)])
        return rows


@dataclass
class StaffCounts:
    staff_name: str
    total: int
    by_department: list[tuple[str, int]]
    by_document: list[tuple[str, int]]
    by_month: list[tuple[str, int]]
    document_type: DocumentType = DocumentType.ALL


@dataclass
class TrendCounts:
    """年(年度)ごとの月別件数。範囲外と本日より後の月は None"""
    year_start_month: int
    by_year: list[tuple[int, list[int | None]]]

    @property
    def year_header(self) -> str:
        return '年' if self.year_start_month == 1 else '年度'

    def year_label(self, year: int) -> str:
        return f'{year}{self.year_header}'

    @property
    def month_labels(self) -> list[str]:
        return [f'{(self.year_start_month - 1 + offset) % 12 + 1}月' for offset in range(12)]

    @property
    def table_rows(self) -> list[list[str | int | None]]:
        """1 行が 1 年で、年の名前・12 か月の件数・合計を並べる"""
        return [[self.year_label(year), *counts, sum(count for count in counts if count is not None)]
                for year, counts in self.by_year]


def count_documents(connection: sqlite3.Connection, start_date: date, end_date: date,
                    document_type: DocumentType = DocumentType.ALL) -> DocumentCounts:
    """預り日が期間内で集計対象の文書を、担当者×診療科、診療科ごと、全体で数える"""
    parameters = period_parameters(start_date, end_date)
    condition = PERIOD_CONDITION + document_type_condition(document_type)
    by_staff_department = {
        (staff, department): count for staff, department, count in connection.execute(
            f'SELECT staff_name, department, COUNT(*) FROM documents WHERE {condition}'
            ' GROUP BY staff_name, department', parameters)
    }
    if not by_staff_department:
        raise NoDataError

    by_department: Counter[str] = Counter()
    for (_, department), count in by_staff_department.items():
        by_department[department] += count
    first_date, last_date = connection.execute(
        f'SELECT MIN(received_date), MAX(received_date) FROM documents WHERE {condition}',
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
                          document_limit: int | None = None,
                          document_type: DocumentType = DocumentType.ALL) -> StaffCounts:
    """1 人の担当者について、集計対象の文書の診療科別・文書名別(多い順。document_limit が None なら全件)・
    月別の件数を数える"""
    parameters = (staff_name, *period_parameters(start_date, end_date))
    condition = f'staff_name = ? AND {PERIOD_CONDITION}' + document_type_condition(document_type)
    by_department = connection.execute(
        f'SELECT department, COUNT(*) AS count FROM documents WHERE {condition}'
        ' GROUP BY department ORDER BY count DESC, department', parameters).fetchall()
    # SQLite の LIMIT -1 は件数の上限なし
    by_document = connection.execute(
        f'SELECT document_name, COUNT(*) AS count FROM documents WHERE {condition}'
        ' GROUP BY document_name ORDER BY count DESC, document_name LIMIT ?',
        (*parameters, -1 if document_limit is None else document_limit)).fetchall()
    month_counts = dict(connection.execute(
        f"SELECT strftime('%Y/%m', received_date), COUNT(*) FROM documents WHERE {condition}"
        ' GROUP BY 1', parameters).fetchall())

    return StaffCounts(
        staff_name=staff_name,
        total=sum(count for _, count in by_department),
        by_department=by_department,
        by_document=by_document,
        by_month=[(month, month_counts.get(month, 0)) for month in months_between(start_date, end_date)],
        document_type=document_type,
    )


def trend_subject(document_filter: DocumentFilter, staff_name: str | None) -> str:
    """推移の表題に出す対象。「集計対象 診療科」で、診療科を絞っていなければ集計対象だけ。
    担当者を選んでいれば氏名を前に付ける"""
    subject = document_filter.document_type.value
    if document_filter.department is not None:
        subject += f' {document_filter.department}'
    return f'{staff_name} {subject}' if staff_name else subject


def count_trend(connection: sqlite3.Connection, document_filter: DocumentFilter, start_date: date, end_date: date,
                year_start_month: int, today: date, staff_name: str | None = None) -> TrendCounts:
    """医師依頼日が範囲内で document_filter に合う文書を、年(年度)ごとに月別に数える。
    staff_name が None なら全担当者"""
    condition = 'requested_date BETWEEN ? AND ?'
    parameters = [start_date.isoformat(), end_date.isoformat()]
    if document_filter.department is not None:
        condition += ' AND department = ?'
        parameters.append(document_filter.department)
    condition += document_type_condition(document_filter.document_type)
    if staff_name is not None:
        condition += ' AND staff_name = ?'
        parameters.append(staff_name)
    month_counts = dict(connection.execute(
        f"SELECT strftime('%Y/%m', requested_date), COUNT(*) FROM documents WHERE {condition} GROUP BY 1",
        parameters).fetchall())

    # 'YYYY/MM' は文字列のまま大小を比べられる。今月は途中までの件数を出す
    first_month = f'{start_date:%Y/%m}'
    last_month = min(f'{end_date:%Y/%m}', f'{today:%Y/%m}')
    by_year: list[tuple[int, list[int | None]]] = []
    for year in range(fiscal_year(start_date, year_start_month), fiscal_year(end_date, year_start_month) + 1):
        months = months_between(*get_period_range(year, '一年', year_start_month))
        by_year.append((year, [month_counts.get(month, 0) if first_month <= month <= last_month else None
                               for month in months]))
    return TrendCounts(year_start_month, by_year)
