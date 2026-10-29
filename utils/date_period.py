from datetime import date, datetime, timedelta

DATE_FORMAT = '%Y/%m/%d'

# 期間名 → (年の始まりの月から数えた開始位置(月数), 期間の長さ(月数))
PERIOD_SPANS: dict[str, tuple[int, int]] = {
    '第1四半期': (0, 3),
    '第2四半期': (3, 3),
    '第3四半期': (6, 3),
    '第4四半期': (9, 3),
    '上半期': (0, 6),
    '下半期': (6, 6),
    '一年': (0, 12),
}

PERIOD_NAMES: list[str] = [f'{month}月' for month in range(1, 13)] + list(PERIOD_SPANS)


def _first_day_after_months(year: int, month: int, months: int) -> date:
    total_months = year * 12 + (month - 1) + months
    return date(total_months // 12, total_months % 12 + 1, 1)


def get_period_range(year: int, period_name: str, year_start_month: int) -> tuple[date, date]:
    """year_start_month 月始まりの year 年(年度)における期間の開始日と終了日を返す"""
    if period_name in PERIOD_SPANS:
        offset, length = PERIOD_SPANS[period_name]
    else:
        month = int(period_name.removesuffix('月'))
        offset, length = (month - year_start_month) % 12, 1

    start = _first_day_after_months(year, year_start_month, offset)
    end = _first_day_after_months(year, year_start_month, offset + length) - timedelta(days=1)
    return start, end


def parse_date(text: str) -> date:
    return datetime.strptime(text.strip(), DATE_FORMAT).date()
