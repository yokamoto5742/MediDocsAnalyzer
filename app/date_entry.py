import tkinter as tk
from datetime import date

from tkcalendar import DateEntry

from app import constants
from utils.date_period import parse_date


def make_date_entry(parent: tk.Misc) -> DateEntry:
    return DateEntry(parent, width=12, background='darkblue', foreground='white', borderwidth=2,
                     locale='ja_JP', date_pattern='yyyy/mm/dd')


def read_date_range(start_entry: DateEntry, end_entry: DateEntry) -> tuple[date, date]:
    """開始日と終了日を返す。読めないか開始日が終了日より後なら、画面に出す理由を入れた ValueError"""
    # get_date() は手入力がフォーカスを外れるまで反映されないため、入力欄の文字列を読む
    try:
        start_date = parse_date(start_entry.get())
        end_date = parse_date(end_entry.get())
    except ValueError:
        raise ValueError(constants.INVALID_DATE_MESSAGE) from None
    if start_date > end_date:
        raise ValueError(constants.START_AFTER_END_MESSAGE)
    return start_date, end_date
