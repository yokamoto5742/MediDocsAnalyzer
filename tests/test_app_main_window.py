import os
import sys
import pytest
import configparser
from pathlib import Path
from unittest.mock import patch, MagicMock

from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt6.QtTest import QTest
from PyQt6.QtCore import Qt

from app_main_window import MainWindow
from config_manager import ConfigManager, CONFIG_PATH
from version import VERSION


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
def app():
    """テスト用のQApplicationを提供するフィクスチャ"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


@pytest.fixture
def backup_config():
    """設定ファイルのバックアップと復元を行うフィクスチャ"""
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH, encoding='utf-8')
    # 現在の設定内容を保存
    original_config = configparser.ConfigParser()
    for section in config.sections():
        original_config.add_section(section)
        for key, value in config[section].items():
            original_config[section][key] = value

    yield original_config

    # テスト後に設定を元に戻す
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH, encoding='utf-8')
    restore_config(config, original_config)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        config.write(f)


class TestMainWindow:
    def test_init(self, app, backup_config):
        """MainWindowの初期化テスト"""
        window = MainWindow()

        # ウィンドウタイトルを確認
        assert window.windowTitle() == f"CSV取込アプリ v{VERSION}"

        # 設定に基づくフォントサイズとウィンドウサイズの確認
        config = ConfigManager()
        font_size = config.get_font_size()
        assert window.font().pointSize() == font_size

        window_size = config.get_window_size()
        assert window.width() == window_size[0]
        assert window.height() == window_size[1]

        # 必要なウィジェットが存在することを確認
        central_widget = window.centralWidget()
        assert central_widget is not None

        # レイアウト内のボタンの数を確認
        layout = central_widget.layout()
        assert layout is not None

        # 少なくとも7つのウィジェットがあることを確認
        # (タイトルラベル、CSVボタン、設定ラベル、除外文書ボタン、除外医師ボタン、外観ボタン、座標ボタン、フォルダボタン、閉じるボタン)
        assert layout.count() >= 9

    @patch('app_main_window.transfer_csv_to_excel')
    def test_import_csv_success(self, mock_transfer, app, backup_config):
        """CSVインポート成功のテスト"""
        window = MainWindow()

        window.import_csv()

        # transfer_csv_to_excel関数が呼ばれたことを確認
        mock_transfer.assert_called_once()

    @patch('app_main_window.transfer_csv_to_excel')
    @patch('app_main_window.QMessageBox.critical')
    def test_import_csv_error(self, mock_critical, mock_transfer, app, backup_config):
        """CSVインポートエラーのテスト"""
        # エラーをシミュレート
        mock_transfer.side_effect = Exception("テストエラー")

        window = MainWindow()
        window.import_csv()

        # エラーメッセージが表示されたことを確認
        mock_critical.assert_called_once()
        args = mock_critical.call_args[0]
        assert args[0] == window  # 親ウィンドウ
        assert args[1] == "エラー"  # タイトル
        assert "テストエラー" in args[2]  # エラーメッセージ

    @patch('app_main_window.ExcludeDocsDialog')
    def test_show_exclude_docs_dialog(self, mock_dialog, app, backup_config):
        """除外文書ダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_exclude_docs_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    @patch('app_main_window.ExcludeDoctorsDialog')
    def test_show_exclude_doctors_dialog(self, mock_dialog, app, backup_config):
        """除外医師ダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_exclude_doctors_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    @patch('app_main_window.AppearanceDialog')
    def test_show_appearance_dialog(self, mock_dialog, app, backup_config):
        """外観設定ダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_appearance_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    @patch('app_main_window.FolderPathDialog')
    def test_show_folder_path_dialog(self, mock_dialog, app, backup_config):
        """フォルダパスダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_folder_path_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    def test_show_coordinate_tracker(self, app, backup_config):
        """座標トラッカー表示テスト"""
        window = MainWindow()

        # trackerの初期状態を確認
        tracker = window.tracker
        assert tracker is not None

        # showメソッドをモック化
        tracker.show = MagicMock()

        window.show_coordinate_tracker()

        # showメソッドが呼ばれたことを確認
        tracker.show.assert_called_once()

    @patch('app_main_window.QMainWindow.close')
    def test_close_button(self, mock_close, app, backup_config):
        """閉じるボタンのテスト"""
        window = MainWindow()

        # 閉じるボタンを探して押下
        central_widget = window.centralWidget()
        layout = central_widget.layout()

        close_button = None
        for i in range(layout.count()):
            widget = layout.itemAt(i).widget()
            if widget is not None and hasattr(widget, 'text') and widget.text() == "閉じる":
                close_button = widget
                break

        assert close_button is not None
        close_button.click()

        # closeメソッドが呼ばれたことを確認
        mock_close.assert_called_once()
