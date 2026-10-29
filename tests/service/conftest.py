from collections.abc import Callable, Sequence
from pathlib import Path

import pytest
from openpyxl import Workbook

HEADERS = ['預り日', '患者ID', '患者氏名', '文書名', '診療科', '医師名', '備考', '医師依頼日', '担当者名']

MakeWorkbook = Callable[[Path, Sequence[Sequence[object]]], Path]


@pytest.fixture
def make_workbook() -> MakeWorkbook:
    """1 行目に見出しを入れた Excel ファイルを作る"""
    def _make(path: Path, rows: Sequence[Sequence[object]]) -> Path:
        wb = Workbook()
        sheet = wb.active
        assert sheet is not None
        sheet.append(HEADERS)
        for row in rows:
            sheet.append(list(row))
        wb.save(path)
        return path
    return _make
