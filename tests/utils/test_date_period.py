from datetime import date

import pytest

from utils.date_period import PERIOD_NAMES, get_period_range, parse_date


@pytest.mark.parametrize(('year', 'period_name', 'year_start_month', 'expected'), [
    (2025, '1月', 1, (date(2025, 1, 1), date(2025, 1, 31))),
    (2024, '2月', 1, (date(2024, 2, 1), date(2024, 2, 29))),
    (2025, '第1四半期', 1, (date(2025, 1, 1), date(2025, 3, 31))),
    (2025, '一年', 1, (date(2025, 1, 1), date(2025, 12, 31))),
    (2025, '4月', 4, (date(2025, 4, 1), date(2025, 4, 30))),
    (2025, '3月', 4, (date(2026, 3, 1), date(2026, 3, 31))),
    (2025, '第4四半期', 4, (date(2026, 1, 1), date(2026, 3, 31))),
    (2025, '下半期', 4, (date(2025, 10, 1), date(2026, 3, 31))),
    (2025, '一年', 4, (date(2025, 4, 1), date(2026, 3, 31))),
])
def test_get_period_range(year: int, period_name: str, year_start_month: int,
                          expected: tuple[date, date]) -> None:
    assert get_period_range(year, period_name, year_start_month) == expected


def test_period_names_contains_months_and_spans() -> None:
    assert PERIOD_NAMES[:12] == [f'{month}月' for month in range(1, 13)]
    assert PERIOD_NAMES[12:] == ['第1四半期', '第2四半期', '第3四半期', '第4四半期', '上半期', '下半期', '一年']


def test_parse_date() -> None:
    assert parse_date(' 2025/04/01 ') == date(2025, 4, 1)


def test_parse_date_invalid() -> None:
    with pytest.raises(ValueError):
        parse_date('2025-04-01')
