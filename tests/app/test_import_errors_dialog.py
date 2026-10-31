import csv
import tkinter as tk
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pytest

from app import constants
from app.import_errors_dialog import ImportErrorsDialog
from service.service_excel_reader import RowError

ERRORS = [RowError('元データ.xlsm', 5, '123', '預り日が空です', ' | 花﨑')]


@pytest.fixture
def dialog(root: tk.Tk) -> ImportErrorsDialog:
    return ImportErrorsDialog(root, ERRORS)


def test_init_lists_errors(dialog: ImportErrorsDialog) -> None:
    items = dialog.tree.get_children()
    assert [dialog.tree.item(item, 'values') for item in items] == [('元データ.xlsm', '5', '123', '預り日が空です', ' | 花﨑')]


def test_copy_puts_tab_separated_text(dialog: ImportErrorsDialog) -> None:
    dialog.copy()

    assert dialog.dialog.clipboard_get() == '\t'.join(constants.ERROR_COLUMNS) + '\n元データ.xlsm\t5\t123\t預り日が空です\t | 花﨑'


def test_save_csv_writes_to_output_dir_and_opens_it(dialog: ImportErrorsDialog, config_path: Path) -> None:
    output_dir = config_path.parent / 'output'
    with patch('app.import_errors_dialog.datetime') as mock_datetime, \
            patch('app.import_errors_dialog.os.startfile', create=True) as startfile:
        mock_datetime.now.return_value = datetime(2026, 10, 31, 14, 30, 5)
        dialog.save_csv()

    path = output_dir / '取込エラー行_20261031_143005.csv'
    assert path.read_bytes().startswith(b'\xef\xbb\xbf')
    with open(path, encoding='utf-8-sig', newline='') as f:
        assert list(csv.reader(f)) == [list(constants.ERROR_COLUMNS), ['元データ.xlsm', '5', '123', '預り日が空です', ' | 花﨑']]
    startfile.assert_called_once_with(output_dir)


def test_save_csv_shows_error(dialog: ImportErrorsDialog, config_path: Path) -> None:
    with patch('app.import_errors_dialog.open', side_effect=PermissionError('locked'), create=True), \
            patch('app.import_errors_dialog.messagebox') as messagebox:
        dialog.save_csv()

    messagebox.showerror.assert_called_once()
    assert 'locked' in messagebox.showerror.call_args.args[1]
