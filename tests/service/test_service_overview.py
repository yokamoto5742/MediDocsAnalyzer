from datetime import date

import pytest

from service.models import Draft
from service.service_overview import DraftStatus, Overview, draft_status, load_overview
from tests.documents_database import Document, make_documents_connection


def make_draft(received: date, requested: date | None = None, staff: str = '') -> Draft:
    return Draft(received, requested, '診断書', '内科', '佐藤', staff, '1')


def test_load_overview_reads_drafts_within_days_newest_first() -> None:
    connection = make_documents_connection([
        Document(date(2026, 10, 24), None, '診断書', '内科', '佐藤', None, patient_id=1),
        Document(date(2026, 10, 25), None, '診断書', '内科', '佐藤', '植田', patient_id=2),
        Document(date(2026, 10, 31), date(2026, 10, 31), '診断書', '外科', '佐藤', None, patient_id=3),
        Document(date(2026, 11, 1), None, '診断書', '内科', '佐藤', None, patient_id=4),
        # 文書(担当者名と医師依頼日がある行)と論理削除した行は出さない
        Document(date(2026, 10, 30), date(2026, 10, 30), '診断書', '内科', '佐藤', '植田', patient_id=5),
        Document(date(2026, 10, 30), None, '診断書', '内科', '佐藤', None, patient_id=6, deleted=True),
    ])

    overview = load_overview(connection, date(2026, 10, 31), 7)

    assert (overview.start_date, overview.end_date) == (date(2026, 10, 25), date(2026, 10, 31))
    assert overview.drafts == [
        Draft(date(2026, 10, 31), date(2026, 10, 31), '診断書', '外科', '佐藤', '', '3'),
        Draft(date(2026, 10, 25), None, '診断書', '内科', '佐藤', '植田', '2'),
    ]
    connection.close()


@pytest.mark.parametrize(('requested', 'staff', 'expected'), [
    (None, '', DraftStatus.NOT_STARTED),
    (None, '植田', DraftStatus.INCOMPLETE),
    (date(2026, 10, 31), '', DraftStatus.UNASSIGNED),
])
def test_draft_status(requested: date | None, staff: str, expected: DraftStatus) -> None:
    assert draft_status(make_draft(date(2026, 10, 31), requested, staff)) == expected


def test_overview_counts_status_by_department() -> None:
    drafts = [make_draft(date(2026, 10, 31)), make_draft(date(2026, 10, 31), staff='植田'),
              make_draft(date(2026, 10, 31)),
              Draft(date(2026, 10, 31), date(2026, 10, 31), '診断書', '外科', '佐藤', '', '1')]

    overview = Overview(date(2026, 10, 31), date(2026, 10, 31), drafts)

    assert overview.status_counts == {DraftStatus.NOT_STARTED: 2, DraftStatus.INCOMPLETE: 1,
                                      DraftStatus.UNASSIGNED: 1}
    assert overview.department_status_counts == [
        ('内科', {DraftStatus.NOT_STARTED: 2, DraftStatus.INCOMPLETE: 1}),
        ('外科', {DraftStatus.UNASSIGNED: 1}),
    ]
