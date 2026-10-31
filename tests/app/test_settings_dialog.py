import tkinter as tk
from pathlib import Path
from unittest.mock import patch

import pytest

from app.settings_dialog import SettingsDialog
from utils.config_manager import load_config


@pytest.fixture
def dialog(config_path: Path, root: tk.Tk) -> SettingsDialog:
    return SettingsDialog(root)


def test_init_shows_saved_paths(dialog: SettingsDialog, config_path: Path) -> None:
    assert list(dialog.path_vars) == ['source_file_path', 'database_path', 'legacy_database_path', 'backup_dir',
                                      'template_path', 'output_dir']
    assert dialog.path_vars['database_path'].get() == str(config_path.parent / 'database.db')


def test_choose_database_allows_new_file(dialog: SettingsDialog) -> None:
    with patch('app.settings_dialog.filedialog.asksaveasfilename', return_value='D:/共有/new.db') as ask:
        dialog._choose_database('database_path')

    assert dialog.path_vars['database_path'].get() == 'D:\\共有\\new.db'
    assert ask.call_args.kwargs['confirmoverwrite'] is False


def test_choose_file_converts_slashes(dialog: SettingsDialog) -> None:
    with patch('app.settings_dialog.filedialog.askopenfilename', return_value='C:/data/元データ.xlsm'):
        dialog._choose_file('source_file_path')

    assert dialog.path_vars['source_file_path'].get() == 'C:\\data\\元データ.xlsm'


def test_choose_directory_keeps_value_when_cancelled(dialog: SettingsDialog) -> None:
    before = dialog.path_vars['output_dir'].get()

    with patch('app.settings_dialog.filedialog.askdirectory', return_value=''):
        dialog._choose_directory('output_dir')

    assert dialog.path_vars['output_dir'].get() == before


def test_save_writes_paths_and_keeps_other_settings(dialog: SettingsDialog) -> None:
    dialog.path_vars['backup_dir'].set('  D:\\backup  ')
    set_valid_numbers(dialog)

    dialog.save()

    config = load_config()
    assert config['PATHS']['backup_dir'] == 'D:\\backup'
    assert config['Analysis']['ordered_names'] == '植田,沖野,花﨑'


def set_valid_numbers(dialog: SettingsDialog) -> None:
    for key, value in (('window_width', ' 1250 '), ('window_height', '760'), ('keep_generations', '14'),
                       ('pending_days', '10')):
        dialog.number_vars[key].set(value)


def test_init_shows_saved_numbers(dialog: SettingsDialog) -> None:
    # pending_days は一時の設定ファイルに無いため、初期値の 7 を出す
    assert {key: var.get() for key, var in dialog.number_vars.items()} == {
        'window_width': '300', 'window_height': '450', 'keep_generations': '30', 'pending_days': '7'}


def test_save_writes_numbers(dialog: SettingsDialog) -> None:
    set_valid_numbers(dialog)

    dialog.save()

    config = load_config()
    assert config['Appearance']['window_width'] == '1250'
    assert config['Appearance']['window_height'] == '760'
    assert config['Backup']['keep_generations'] == '14'
    assert config['Overview']['pending_days'] == '10'


def test_save_rejects_invalid_pending_days(dialog: SettingsDialog) -> None:
    set_valid_numbers(dialog)
    dialog.number_vars['pending_days'].set('0')

    with patch('app.settings_dialog.messagebox') as messagebox:
        dialog.save()

    messagebox.showerror.assert_called_once()
    assert not load_config().has_section('Overview')


@pytest.mark.parametrize('value', ['abc', '0', '366', '-1', ''])
def test_save_rejects_invalid_number(dialog: SettingsDialog, value: str) -> None:
    set_valid_numbers(dialog)
    dialog.number_vars['keep_generations'].set(value)

    with patch('app.settings_dialog.messagebox') as messagebox:
        dialog.save()

    messagebox.showerror.assert_called_once()
    assert load_config()['Backup']['keep_generations'] == '30'
