import sqlite3
from collections import Counter
from dataclasses import dataclass, field
from datetime import date

from service.models import NoDataError
from service.service_database import PERIOD_CONDITION, period_parameters
from utils.date_period import fiscal_year, get_period_range

TOTAL_LABEL = '合計'
# 文書名にこの語を含む文書を、診療科を問わず「主治医意見書」として数える
OPINION_LABEL = '主治医意見書'
OPINION_CONDITION = f"document_name LIKE '%{OPINION_LABEL}%'"


@dataclass(frozen=True)
class DocumentFilter:
    """推移で数える文書の絞り込み。既定は全文書"""
    department: str | None = None
    opinion_only: bool = False


@dataclass
class DocumentCounts:
    by_staff_department: dict[tuple[str, str], int]
    by_department: dict[str, int]
    total: int
    first_date: date
    last_date: date
    # 担当者ごとの主治医意見書の件数。診療科の件数にも含まれている
    opinion_by_staff: dict[str, int] = field(default_factory=dict)


@dataclass
class OverallSummary:
    counts: DocumentCounts
    staff_members: list[str]
    departments: list[str]

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

    @property
    def opinion_counts(self) -> list[int]:
        """rows の各行に並べる主治医意見書の件数。最後は合計行で、一覧にない担当者の件数も含む"""
        opinion_by_staff = self.counts.opinion_by_staff
        return [*(opinion_by_staff.get(staff, 0) for staff in self.staff_members), sum(opinion_by_staff.values())]


@dataclass
class StaffCounts:
    staff_name: str
    total: int
    by_department: list[tuple[str, int]]
    by_document: list[tuple[str, int]]
    by_month: list[tuple[str, int]]


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
    opinion_by_staff = dict(connection.execute(
        f'SELECT staff_name, COUNT(*) FROM documents WHERE {PERIOD_CONDITION} AND {OPINION_CONDITION}'
        ' GROUP BY staff_name', parameters).fetchall())

    return DocumentCounts(
        by_staff_department=by_staff_department,
        by_department=dict(by_department),
        total=sum(by_staff_department.values()),
        first_date=date.fromisoformat(first_date),
        last_date=date.fromisoformat(last_date),
        opinion_by_staff=opinion_by_staff,
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
                          document_limit: int | None = None) -> StaffCounts:
    """1 人の担当者について、診療科別・文書名別(多い順。document_limit が None なら全件)・月別の件数を数える"""
    parameters = (staff_name, *period_parameters(start_date, end_date))
    condition = f'staff_name = ? AND {PERIOD_CONDITION}'
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
    opinion_count = connection.execute(
        f'SELECT COUNT(*) FROM documents WHERE {condition} AND {OPINION_CONDITION}', parameters).fetchone()[0]

    # 主治医意見書は診療科の件数にも含まれるため、合計に足さないよう後から並べる
    return StaffCounts(
        staff_name=staff_name,
        total=sum(count for _, count in by_department),
        by_department=by_department + ([(OPINION_LABEL, opinion_count)] if opinion_count else []),
        by_document=by_document,
        by_month=[(month, month_counts.get(month, 0)) for month in months_between(start_date, end_date)],
    )


def trend_subject(department: str, staff_name: str | None) -> str:
    """推移の表題に出す対象。担当者を選んでいれば氏名を前に付ける"""
    return f'{staff_name} {department}' if staff_name else department


def count_trend(connection: sqlite3.Connection, document_filter: DocumentFilter, start_date: date, end_date: date,
                year_start_month: int, today: date, staff_name: str | None = None) -> TrendCounts:
    """医師依頼日が範囲内で document_filter に合う文書を、年(年度)ごとに月別に数える。
    staff_name が None なら全担当者"""
    condition = 'requested_date BETWEEN ? AND ?'
    parameters = [start_date.isoformat(), end_date.isoformat()]
    if document_filter.department is not None:
        condition += ' AND department = ?'
        parameters.append(document_filter.department)
    if document_filter.opinion_only:
        condition += f' AND {OPINION_CONDITION}'
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
