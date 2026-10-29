import re
from configparser import ConfigParser
from dataclasses import dataclass
from datetime import date, datetime
from typing import TypedDict

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

NAME_PATTERN = re.compile(r'\((.*?)\)')
INDEX_SHEET_NAME = 'シート一覧'
SHEET_DATE_FORMAT = '%Y年%m月%d日'


class TaskRecord(TypedDict):
    date: date
    content: str
    minutes: float


class CommunicationRecord(TaskRecord):
    name: str


@dataclass
class WorkbookData:
    tasks: list[TaskRecord]
    daily_tasks: list[TaskRecord]
    communication_tasks: list[CommunicationRecord]
    all_items: list[TaskRecord]
    first_date: date
    last_date: date


class NoDataError(Exception):
    """指定された期間内のシートが 1 枚もない"""


def parse_sheet_date(value: object) -> date | None:
    """A1 セルの日付を返す。日付として読めなければ None"""
    if isinstance(value, datetime):
        return value.date()
    try:
        return datetime.strptime(str(value), SHEET_DATE_FORMAT).date()
    except ValueError:
        return None


class ExcelTaskReader:
    def __init__(self, config: ConfigParser) -> None:
        self.config = config

    @staticmethod
    def extract_cell_data(sheet: Worksheet, row: int, sheet_date: date,
                          remove_name: bool = False) -> TaskRecord | None:
        content = str(sheet[f'B{row}'].value or '')
        time = sheet[f'C{row}'].value

        if not content or not time:
            return None

        try:
            minutes = float(time)
        except (ValueError, TypeError):
            return None

        if remove_name:
            content = NAME_PATTERN.sub('', content)
        words = content.split()
        if not words:
            return None

        return {'date': sheet_date, 'content': words[0], 'minutes': minutes}

    def _rows(self, start_key: str, end_key: str) -> range:
        return range(self.config.getint('Analysis', start_key),
                     self.config.getint('Analysis', end_key) + 1)

    def _read_tasks(self, sheet: Worksheet, rows: range, sheet_date: date) -> list[TaskRecord]:
        tasks = (self.extract_cell_data(sheet, row, sheet_date) for row in rows)
        return [task for task in tasks if task]

    def _read_communications(self, sheet: Worksheet, sheet_date: date) -> list[CommunicationRecord]:
        communications: list[CommunicationRecord] = []
        for row in self._rows('communication_start_row', 'communication_end_row'):
            # 「内容(名前)」形式で名前があるものだけを対象にする
            name_match = NAME_PATTERN.search(str(sheet[f'B{row}'].value))
            task = self.extract_cell_data(sheet, row, sheet_date, remove_name=True)
            if name_match and task:
                communications.append({**task, 'name': name_match.group(1)})
        return communications

    def load_excel_task_data(self, sheet: Worksheet, sheet_date: date) -> tuple[
            list[TaskRecord], list[TaskRecord], list[CommunicationRecord]]:
        tasks = self._read_tasks(sheet, self._rows('start_row', 'end_row'), sheet_date)
        daily_tasks = self._read_tasks(sheet, self._rows('daily_task_start_row', 'daily_task_end_row'), sheet_date)
        return tasks, daily_tasks, self._read_communications(sheet, sheet_date)

    def load_excel_sheet_all_items(self, sheet: Worksheet, sheet_date: date) -> list[TaskRecord]:
        comm_rows = self._rows('communication_start_row', 'communication_end_row')
        items = (self.extract_cell_data(sheet, row, sheet_date, remove_name=row in comm_rows)
                 for row in self._rows('start_row', 'daily_task_end_row'))
        return [item for item in items if item]

    def read_workbook(self, file_path: str, start_date: date, end_date: date) -> WorkbookData:
        wb = load_workbook(filename=file_path)
        tasks: list[TaskRecord] = []
        daily_tasks: list[TaskRecord] = []
        communication_tasks: list[CommunicationRecord] = []
        all_items: list[TaskRecord] = []
        dates: list[date] = []

        for sheet in wb.worksheets:
            if sheet.title == INDEX_SHEET_NAME:
                continue

            sheet_date = parse_sheet_date(sheet['A1'].value)
            if sheet_date is None or not start_date <= sheet_date <= end_date:
                continue

            sheet_tasks, sheet_daily_tasks, sheet_communications = self.load_excel_task_data(sheet, sheet_date)
            tasks.extend(sheet_tasks)
            daily_tasks.extend(sheet_daily_tasks)
            communication_tasks.extend(sheet_communications)
            all_items.extend(self.load_excel_sheet_all_items(sheet, sheet_date))
            dates.append(sheet_date)

        if not dates:
            raise NoDataError

        return WorkbookData(tasks, daily_tasks, communication_tasks, all_items, min(dates), max(dates))
