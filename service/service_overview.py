import sqlite3
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, timedelta
from enum import StrEnum
from typing import Any

from service.models import Draft

DEFAULT_PENDING_DAYS = 7
# 論理削除していない行のうち、担当者名か医師依頼日が空の行(下書き)
DRAFTS_QUERY = (
    'SELECT received_date, requested_date, document_name, department, doctor_name, staff_name, patient_id'
    ' FROM documents WHERE deleted = 0 AND received_date BETWEEN ? AND ?'
    " AND (IFNULL(staff_name, '') = '' OR IFNULL(requested_date, '') = '')"
    ' ORDER BY received_date DESC, id')


class DraftStatus(StrEnum):
    NOT_STARTED = '未着手'
    INCOMPLETE = '未完了'
    UNASSIGNED = '担当不明'


def draft_status(draft: Draft) -> DraftStatus:
    """下書きは担当者名か医師依頼日の少なくとも一方が空"""
    if draft.staff_name:
        return DraftStatus.INCOMPLETE
    return DraftStatus.NOT_STARTED if draft.requested_date is None else DraftStatus.UNASSIGNED


@dataclass
class Overview:
    """預り日が start_date〜end_date の下書き(担当者名か医師依頼日が空の行)"""
    start_date: date
    end_date: date
    drafts: list[Draft]

    @property
    def status_counts(self) -> Counter[DraftStatus]:
        return Counter(draft_status(draft) for draft in self.drafts)

    @property
    def department_status_counts(self) -> list[tuple[str, Counter[DraftStatus]]]:
        """診療科ごとの状態別の件数を、件数の多い順に返す"""
        by_department: dict[str, Counter[DraftStatus]] = {}
        for draft in self.drafts:
            by_department.setdefault(draft.department, Counter())[draft_status(draft)] += 1
        return sorted(by_department.items(), key=lambda item: (-item[1].total(), item[0]))


def draft_from_row(row: Sequence[Any]) -> Draft:
    """DRAFTS_QUERY の順に並んだ行から下書きを作る。CSV2XL の DB は患者IDが整数で、空欄は NULL"""
    received, requested, document_name, department, doctor_name, staff_name, patient_id = row
    return Draft(date.fromisoformat(received), date.fromisoformat(requested) if requested else None,
                 document_name, department, doctor_name, staff_name or '', str(patient_id))


def load_overview(connection: sqlite3.Connection, today: date, days: int) -> Overview:
    """CSV2XL の DB から、本日を含めて days 日間に預かった下書きを、預り日の新しい順に読む"""
    start_date = today - timedelta(days=days - 1)
    rows = connection.execute(DRAFTS_QUERY, (start_date.isoformat(), today.isoformat()))
    return Overview(start_date, today, [draft_from_row(row) for row in rows])
