from datetime import date

import pytest

from service.models import Draft
from service.service_overview import DraftStatus, draft_status, make_overview


def make_draft(received: date, requested: date | None = None, staff: str = '') -> Draft:
    return Draft(received, requested, '診断書', '内科', '佐藤', staff, '1')


def test_make_overview_keeps_drafts_within_days_newest_first() -> None:
    drafts = [make_draft(date(2026, 10, 24)), make_draft(date(2026, 10, 25), staff='植田'),
              make_draft(date(2026, 10, 31), requested=date(2026, 10, 31)), make_draft(date(2026, 11, 1))]

    overview = make_overview(drafts, date(2026, 10, 31), 7)

    assert (overview.start_date, overview.end_date) == (date(2026, 10, 25), date(2026, 10, 31))
    assert [draft.received_date for draft in overview.drafts] == [date(2026, 10, 31), date(2026, 10, 25)]


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

    overview = make_overview(drafts, date(2026, 10, 31), 1)

    assert overview.status_counts == {DraftStatus.NOT_STARTED: 2, DraftStatus.INCOMPLETE: 1,
                                      DraftStatus.UNASSIGNED: 1}
    assert overview.department_status_counts == [
        ('内科', {DraftStatus.NOT_STARTED: 2, DraftStatus.INCOMPLETE: 1}),
        ('外科', {DraftStatus.UNASSIGNED: 1}),
    ]
