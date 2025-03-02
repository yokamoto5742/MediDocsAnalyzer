import pytest
from unittest.mock import Mock, patch, mock_open
from datetime import datetime
import configparser
import os
from app_window import TaskAnalyzerGUI
from version import VERSION


@pytest.fixture
def mock_tk():
    with patch('app_window.tk.Tk') as mock_tk_class:
        root = Mock()
        root.title = Mock()
        root.geometry = Mock()
        root.quit = Mock()
        mock_tk_class.return_value = root
        yield root


@pytest.fixture
def mock_date_entry():
    with patch('app_window.DateEntry') as mock_date_entry_class:
        date_entry = Mock()
        date_entry.get_date = Mock()
        mock_date_entry_class.return_value = date_entry
        yield date_entry


@pytest.fixture
def mock_config():
    config = configparser.ConfigParser()
    original_config = configparser.ConfigParser()

    config['Appearance'] = {
        'window_width': '800',
        'window_height': '600'
    }
    config['PATHS'] = {
        'config_path': 'test_config.ini'
    }
    config['Analysis'] = {}

    for section in config.sections():
        if not original_config.has_section(section):
            original_config.add_section(section)
        for key, value in config[section].items():
            original_config[section][key] = value

    with patch('app_window.load_config', return_value=config), \
            patch('app_window.save_config') as mock_save:
        yield config
        restore_config(config, original_config)


def restore_config(config, original_config):
    """configを元の状態に復元するヘルパーメソッド"""
    for section in config.sections():
        config.remove_section(section)
    for section in original_config.sections():
        if not config.has_section(section):
            config.add_section(section)
        for key, value in original_config[section].items():
            config[section][key] = value



@pytest.fixture
def mock_analyzer():
    with patch('app_window.TaskAnalyzer') as mock_task_analyzer:
        analyzer = Mock()
        mock_task_analyzer.return_value = analyzer
        yield analyzer


@pytest.fixture
def mock_messagebox():
    with patch('app_window.messagebox') as mock_msg:
        # showerrorのデフォルト実装を設定
        mock_msg.showerror = Mock()
        yield mock_msg


@pytest.fixture
def gui(mock_tk, mock_config, mock_analyzer, mock_date_entry):
    with patch('app_window.ttk.Frame'), \
            patch('app_window.ttk.LabelFrame'), \
            patch('app_window.ttk.Label'), \
            patch('app_window.ttk.Button'):
        gui = TaskAnalyzerGUI(mock_tk)
        yield gui


def test_init(gui, mock_tk):
    """初期化のテスト"""
    mock_tk.title.assert_called_with(f'業務分析 v{VERSION}')


def test_analysis_error(gui, mock_config, mock_analyzer, mock_date_entry, mock_messagebox):
    """分析エラー時のテスト"""
    error_message = "テストエラー"
    mock_analyzer.run_analysis.return_value = (False, error_message)

    start_date = datetime(2025, 2, 1)
    end_date = datetime(2025, 2, 28)

    gui.start_date.get_date.return_value = start_date
    gui.end_date.get_date.return_value = end_date

    with patch('app_window.save_config'):
        gui.start_analysis()

    mock_messagebox.showerror.assert_called_with("エラー", error_message)


def test_open_config(gui, mock_config):
    """設定ファイルを開く機能のテスト"""
    test_config_path = 'test_config.ini'

    with patch('app_window.subprocess.Popen') as mock_popen:
        gui.open_config()
        mock_popen.assert_called_once_with(['notepad.exe', test_config_path])


def test_open_config_error(gui, mock_config):
    """設定ファイルを開く際のエラーテスト"""
    with patch('app_window.subprocess.Popen', side_effect=Exception("テストエラー")):
        with patch('app_window.messagebox.showerror') as mock_error:
            gui.open_config()
            mock_error.assert_called_with("エラー", "設定ファイルを開けませんでした：\nテストエラー")
