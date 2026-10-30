import csv
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

import pytest

from app import constants
from app.import_errors_dialog import ImportErrorsDialog
from service.service_excel_reader import RowError

ERRORS = [RowError('元データ.xlsm', 5, '預り日が空です', ' | 花﨑')]


@pytest.fixture
def dialog(root: tk.Tk) -> ImportErrorsDialog:
    return ImportErrorsDialog(root, ERRORS)


def test_init_lists_errors(dialog: ImportErrorsDialog) -> None:
    items = dialog.tree.get_children()
    assert [dialog.tree.item(item, 'values') for item in items] == [('元データ.xlsm', '5', '預り日が空です', ' | 花﨑')]


def test_copy_puts_tab_separated_text(dialog: ImportErrorsDialog) -> None:
    dialog.copy()

    assert dialog.dialog.clipboard_get() == '\t'.join(constants.ERROR_COLUMNS) + '\n元データ.xlsm\t5\t預り日が空です\t | 花﨑'


def test_save_csv_writes_utf8_with_bom(dialog: ImportErrorsDialog, tmp_path: Path) -> None:
    path = tmp_path / 'errors.csv'
    with patch('app.import_errors_dialog.filedialog.asksaveasfilename', return_value=str(path)):
        dialog.save_csv()

    assert path.read_bytes().startswith(b'\xef\xbb\xbf')
    with open(path, encoding='utf-8-sig', newline='') as f:
        assert list(csv.reader(f)) == [list(constants.ERROR_COLUMNS), ['元データ.xlsm', '5', '預り日が空です', ' | 花﨑']]
