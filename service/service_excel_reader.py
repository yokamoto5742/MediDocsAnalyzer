import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from service.models import Document, Draft, RowError, SourceData

DATE_PATTERN = re.compile(r'\d{4}/\d{1,2}/\d{1,2}')
# 元データのファイル名の末尾の日時(…_YYYYMMDDhhmm.xlsm)
SOURCE_FILE_PATTERN = re.compile(r'_(\d{12})\.xlsm$', re.IGNORECASE)
# Excel が開いている間に作る一時ファイルの先頭
EXCEL_LOCK_FILE_PREFIX = '~$'
RECEIVED_DATE_COLUMN = '預り日'
REQUESTED_DATE_COLUMN = '医師依頼日'
DOCUMENT_NAME_COLUMN = '文書名'
DEPARTMENT_COLUMN = '診療科'
DOCTOR_COLUMN = '医師名'
STAFF_COLUMN = '担当者名'
PATIENT_ID_COLUMN = '患者ID'
# Document / Draft の文字列の項目と同じ順に並べる
TEXT_COLUMNS = (DOCUMENT_NAME_COLUMN, DEPARTMENT_COLUMN, DOCTOR_COLUMN, STAFF_COLUMN, PATIENT_ID_COLUMN)
REQUIRED_COLUMNS = (RECEIVED_DATE_COLUMN, REQUESTED_DATE_COLUMN, *TEXT_COLUMNS)


def to_date(value: object) -> date | None:
    """日付・Excel のシリアル値・'YYYY/MM/DD' 形式の文字列を日付にする。空なら None、読めなければ ValueError"""
    if value is None or str(value).strip() == '':
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return from_excel(value).date()
    text = str(value).strip().split()[0].replace('-', '/')
    if not DATE_PATTERN.fullmatch(text):
        raise ValueError(text)
    return datetime.strptime(text, '%Y/%m/%d').date()


def find_latest_source(directory: str) -> Path:
    """フォルダ内で、ファイル名の日時(…_YYYYMMDDhhmm.xlsm)が最も新しい元データを返す。無ければ FileNotFoundError"""
    candidates = [(match.group(1), path) for path in Path(directory).glob('*.xlsm')
                  if not path.name.startswith(EXCEL_LOCK_FILE_PREFIX)
                  and (match := SOURCE_FILE_PATTERN.search(path.name))]
    if not candidates:
        raise FileNotFoundError(f'{directory} に元データ(ファイル名が「_年月日時分.xlsm」で終わるもの)がありません')
    return max(candidates)[1]


def to_text(value: object) -> str:
    return '' if value is None else str(value).strip()


def read_received_date(cells: dict[str, object]) -> date:
    """預り日を返す。空か日付でなければ、エラーの理由を入れた ValueError"""
    try:
        received = to_date(cells[RECEIVED_DATE_COLUMN])
    except ValueError:
        raise ValueError('預り日が日付ではありません') from None
    if received is None:
        raise ValueError('預り日が空です')
    return received


def read_requested_date(cells: dict[str, object], received: date, known_staff: set[str] | None) -> date | None:
    """医師依頼日を返す(空なら None)。日付でない・預り日より前・担当者がマスタにない行は、理由を入れた ValueError"""
    try:
        requested = to_date(cells[REQUESTED_DATE_COLUMN])
    except ValueError:
        raise ValueError('医師依頼日が日付ではありません') from None
    if requested is not None and requested < received:
        raise ValueError('医師依頼日が預り日より前です')
    staff = to_text(cells[STAFF_COLUMN])
    if requested is not None and staff and known_staff is not None and staff not in known_staff:
        raise ValueError('担当者名がマスタに登録されていません')
    return requested


def read_documents(file_path: str, known_staff: set[str] | None = None) -> SourceData:
    """1 行目を見出しとして読み、医師依頼日と担当者名がある行を文書として返す。

    空欄の行は下書きとして drafts に入れ、日付が読めない行などはエラーとして返す。
    first_received_date / last_received_date には、読めた預り日(下書きを含む)の最も古い日・新しい日を入れる。
    """
    wb = load_workbook(file_path, read_only=True, data_only=True)
    sheet = wb.active
    assert sheet is not None
    rows = sheet.iter_rows(values_only=True)
    headers = [to_text(value) for value in next(rows, ())]
    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        wb.close()
        raise ValueError(f'{Path(file_path).name} に列がありません: {", ".join(missing)}')

    column_indexes = {column: headers.index(column) for column in REQUIRED_COLUMNS}
    source = SourceData()
    for row_number, row in enumerate(rows, start=2):
        cells: dict[str, object] = {column: row[index] if index < len(row) else None
                                    for column, index in column_indexes.items()}
        if all(to_text(value) == '' for value in cells.values()):
            continue
        _add_row(source, cells, row_number, Path(file_path).name, known_staff)
    wb.close()
    return source


def _add_row(source: SourceData, cells: dict[str, object], row_number: int, file_name: str,
             known_staff: set[str] | None) -> None:
    try:
        received = read_received_date(cells)
        # 預り日が読める行は、エラーでも取り込み範囲に含める
        source.first_received_date = min(received, source.first_received_date or received)
        source.last_received_date = max(received, source.last_received_date or received)
        requested = read_requested_date(cells, received, known_staff)
    except ValueError as e:
        row_text = ' | '.join(to_text(value) for column, value in cells.items() if column != PATIENT_ID_COLUMN)
        source.errors.append(RowError(file_name, row_number, to_text(cells[PATIENT_ID_COLUMN]), str(e), row_text))
        return

    texts = [to_text(cells[column]) for column in TEXT_COLUMNS]
    if requested is None or not to_text(cells[STAFF_COLUMN]):
        source.drafts.append(Draft(received, requested, *texts))
        return
    source.documents.append(Document(received, requested, *texts))
