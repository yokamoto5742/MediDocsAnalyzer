import tkinter as tk
from pathlib import Path
from unittest.mock import patch

import pytest

from app.path_settings_dialog import PathSettingsDialog
from utils.config_manager import load_config


@pytest.fixture
def dialog(config_path: Path, root: tk.Tk) -> PathSettingsDialog:
    return PathSettingsDialog(root)


def test_init_shows_saved_paths(dialog: PathSettingsDialog, config_path: Path) -> None:
    assert list(dialog.path_vars) == ['source_file_path', 'database_path', 'legacy_database_path', 'backup_dir',
                                      'template_path', 'output_dir']
    assert dialog.path_vars['database_path'].get() == str(config_path.parent / 'database.db')


def test_choose_database_allows_new_file(dialog: PathSettingsDialog) -> None:
    with patch('app.path_settings_dialog.filedialog.asksaveasfilename', return_value='D:/共有/new.db') as ask:
        dialog._choose_database('database_path')

    assert dialog.path_vars['database_path'].get() == 'D:\\共有\\new.db'
    assert ask.call_args.kwargs['confirmoverwrite'] is False


def test_choose_file_converts_slashes(dialog: PathSettingsDialog) -> None:
    with patch('app.path_settings_dialog.filedialog.askopenfilename', return_value='C:/data/元データ.xlsm'):
        dialog._choose_file('source_file_path')

    assert dialog.path_vars['source_file_path'].get() == 'C:\\data\\元データ.xlsm'


def test_choose_directory_keeps_value_when_cancelled(dialog: PathSettingsDialog) -> None:
    before = dialog.path_vars['output_dir'].get()

    with patch('app.path_settings_dialog.filedialog.askdirectory', return_value=''):
        dialog._choose_directory('output_dir')

    assert dialog.path_vars['output_dir'].get() == before


def test_save_writes_paths_and_keeps_other_settings(dialog: PathSettingsDialog) -> None:
    dialog.path_vars['backup_dir'].set('  D:\\backup  ')

    dialog.save()

    config = load_config()
    assert config['PATHS']['backup_dir'] == 'D:\\backup'
    assert config['Analysis']['ordered_names'] == '植田,沖野,花﨑'
