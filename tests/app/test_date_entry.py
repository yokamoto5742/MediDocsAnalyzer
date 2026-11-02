import re
import tkinter as tk
from datetime import date

import pytest
from tkcalendar import DateEntry

from app import constants
from app.date_entry import make_date_entry, read_date_range


def make_entries(root: tk.Tk, start: str, end: str) -> tuple[DateEntry, DateEntry]:
    entries = make_date_entry(root), make_date_entry(root)
    for entry, text in zip(entries, (start, end)):
        entry.delete(0, tk.END)
        entry.insert(0, text)
    return entries


def test_read_date_range(root: tk.Tk) -> None:
    assert read_date_range(*make_entries(root, '2025/04/01', '2025/04/01')) == (date(2025, 4, 1), date(2025, 4, 1))


@pytest.mark.parametrize(('start', 'end', 'expected_message'), [
    ('2025-04-01', '2025/04/30', constants.INVALID_DATE_MESSAGE),
    ('2025/05/01', '2025/04/30', constants.START_AFTER_END_MESSAGE),
])
def test_read_date_range_rejects_invalid_range(root: tk.Tk, start: str, end: str, expected_message: str) -> None:
    with pytest.raises(ValueError, match=re.escape(expected_message)):
        read_date_range(*make_entries(root, start, end))
