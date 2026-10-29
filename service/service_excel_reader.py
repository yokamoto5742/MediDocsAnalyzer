from datetime import date, datetime

import polars as pl
from openpyxl import load_workbook

DATE_FORMAT = '%Y/%m/%d'
COLUMN_COUNT = 9  # A〜I列
DATE_COLUMNS = ('預り日', '医師依頼日')
PATIENT_ID_COLUMN = '患者ID'


class NoDataError(Exception):
    """対象となるデータが 1 件もない"""


def to_date_text(value: object) -> str:
    """日付を YYYY/MM/DD の文字列にする(時刻があれば落とす)"""
    if isinstance(value, (datetime, date)):
        return value.strftime(DATE_FORMAT)
    text = str(value).strip()
    return text.split()[0].replace('-', '/') if text else ''


def to_patient_id_text(value: object) -> str:
    """数値として読める患者IDは先頭の 0 などをそろえる(例: '00123' → '123')"""
    try:
        return str(int(str(value).strip()))
    except ValueError:
        return str(value)


def to_cell_text(header: str, value: object) -> str:
    if value is None or value == '':
        return ''
    if header in DATE_COLUMNS:
        return to_date_text(value)
    if header == PATIENT_ID_COLUMN:
        return to_patient_id_text(value)
    return str(value)


def read_documents(file_path: str) -> pl.DataFrame:
    """1 行目を見出しとして A〜I列を読み、全セルを文字列(空は '')にして返す"""
    wb = load_workbook(file_path, read_only=True)
    sheet = wb.active
    assert sheet is not None
    rows = sheet.iter_rows(max_col=COLUMN_COUNT, values_only=True)
    header_row = next(rows, ())
    headers = [str(value) for value in header_row if value is not None]

    records = [
        [to_cell_text(header, value) for header, value in zip(headers, row)]
        for row in rows
    ]
    wb.close()
    return pl.DataFrame(records, schema={header: pl.String for header in headers}, orient='row')
