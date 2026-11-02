from collections import Counter
from dataclasses import dataclass
from datetime import date, timedelta
from enum import StrEnum

from service.models import Draft

DEFAULT_PENDING_DAYS = 7


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


def make_overview(drafts: list[Draft], today: date, days: int) -> Overview:
    """本日を含めて days 日間に預かった下書きを、預り日の新しい順に並べる"""
    start_date = today - timedelta(days=days - 1)
    targets = [draft for draft in drafts if start_date <= draft.received_date <= today]
    targets.sort(key=lambda draft: draft.received_date, reverse=True)
    return Overview(start_date, today, targets)
