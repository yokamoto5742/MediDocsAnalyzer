import os
import sys
import pytest
import configparser
from pathlib import Path
from unittest.mock import patch, MagicMock

from PyQt6.QtWidgets import QApplication, QDialogButtonBox, QMessageBox, QDialog
from PyQt6.QtTest import QTest
from PyQt6.QtCore import Qt

from app_dialogs import ExcludeItemDialog, ExcludeDocsDialog, ExcludeDoctorsDialog, FolderPathDialog, AppearanceDialog
from config_manager import ConfigManager, CONFIG_PATH


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


class TestExcludeItemDialog:
    def test_init(self, app, backup_config):
        """ExcludeItemDialogの初期化テスト"""
        dialog = ExcludeItemDialog("テストダイアログ", "テスト項目", "TestSection")
        assert dialog.windowTitle() == "テストダイアログ"
        assert dialog.config_section == "TestSection"
        assert dialog.item_list is not None

    def test_add_item(self, app, backup_config):
        """項目追加機能のテスト"""
        dialog = ExcludeItemDialog("テストダイアログ", "テスト項目", "TestSection")

        # 項目を追加
        dialog.input_field.setText("テスト項目1")
        dialog.add_item()

        assert dialog.item_list.count() == 1
        assert dialog.item_list.item(0).text() == "テスト項目1"

        # 空白の項目は追加されないことを確認
        dialog.input_field.setText("  ")
        dialog.add_item()
        assert dialog.item_list.count() == 1

    def test_delete_selected(self, app, backup_config):
        """項目削除機能のテスト"""
        dialog = ExcludeItemDialog("テストダイアログ", "テスト項目", "TestSection")

        # 項目を追加
        dialog.input_field.setText("テスト項目1")
        dialog.add_item()
        dialog.input_field.setText("テスト項目2")
        dialog.add_item()

        assert dialog.item_list.count() == 2

        # 項目を選択して削除
        dialog.item_list.setCurrentRow(0)
        dialog.delete_selected()

        assert dialog.item_list.count() == 1
        assert dialog.item_list.item(0).text() == "テスト項目2"

    def test_load_items(self, app, backup_config):
        """設定ファイルからの項目読み込みテスト"""
        # 設定ファイルにテスト用セクションと項目を追加
        config = ConfigManager()
        if 'TestSection' not in config.config:
            config.config['TestSection'] = {}
        config.config['TestSection']['list'] = 'アイテム1,アイテム2,アイテム3'
        config.save_config()

        # ダイアログを初期化して項目が読み込まれるか確認
        dialog = ExcludeItemDialog("テストダイアログ", "テスト項目", "TestSection")

        assert dialog.item_list.count() == 3
        assert dialog.item_list.item(0).text() == "アイテム1"
        assert dialog.item_list.item(1).text() == "アイテム2"
        assert dialog.item_list.item(2).text() == "アイテム3"

    def test_accept(self, app, backup_config):
        """ダイアログ確定時の設定保存テスト"""
        dialog = ExcludeItemDialog("テストダイアログ", "テスト項目", "TestSection")

        # 項目を追加
        dialog.input_field.setText("新しい項目1")
        dialog.add_item()
        dialog.input_field.setText("新しい項目2")
        dialog.add_item()

        # モック化したsuperクラスのacceptメソッドでテスト
        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()
            mock_accept.assert_called_once()

        # 設定が保存されたことを確認
        config = ConfigManager()
        saved_items = config.config['TestSection'].get('list', '').split(',')
        assert len(saved_items) == 2
        assert "新しい項目1" in saved_items
        assert "新しい項目2" in saved_items


class TestExcludeDocsDialog:
    def test_init(self, app, backup_config):
        """ExcludeDocsDialogの初期化テスト"""
        dialog = ExcludeDocsDialog()
        assert dialog.windowTitle() == "除外する文書名"
        assert dialog.config_section == "ExcludeDocs"


class TestExcludeDoctorsDialog:
    def test_init(self, app, backup_config):
        """ExcludeDoctorsDialogの初期化テスト"""
        dialog = ExcludeDoctorsDialog()
        assert dialog.windowTitle() == "除外する医師名"
        assert dialog.config_section == "ExcludeDoctors"


class TestFolderPathDialog:
    def test_init(self, app, backup_config):
        """FolderPathDialogの初期化テスト"""
        dialog = FolderPathDialog()
        assert dialog.windowTitle() == "フォルダの場所"
        assert dialog.downloads_path is not None
        assert dialog.excel_path is not None
        assert dialog.backup_path is not None

    def test_load_paths(self, app, backup_config):
        """パス読み込み機能のテスト"""
        # 設定ファイルにテスト用パスを設定
        config = ConfigManager()
        if 'Paths' not in config.config:
            config.config['Paths'] = {}
        config.config['Paths']['downloads_path'] = 'C:\\Test\\Downloads'
        config.config['Paths']['excel_path'] = 'C:\\Test\\Excel\\test.xlsm'
        config.config['Paths']['backup_path'] = 'C:\\Test\\Backup'
        config.save_config()

        # ダイアログを初期化してパスが読み込まれるか確認
        dialog = FolderPathDialog()

        assert dialog.downloads_path.text() == 'C:\\Test\\Downloads'
        assert dialog.excel_path.text() == 'C:\\Test\\Excel\\test.xlsm'
        assert dialog.backup_path.text() == 'C:\\Test\\Backup'

    @patch('app_dialogs.QFileDialog.getExistingDirectory')
    def test_browse_folder_downloads(self, mock_get_dir, app, backup_config):
        """ダウンロードフォルダブラウズ機能のテスト"""
        mock_get_dir.return_value = 'C:\\New\\Downloads'

        dialog = FolderPathDialog()
        dialog.browse_folder('downloads')

        mock_get_dir.assert_called_once()
        assert dialog.downloads_path.text() == 'C:\\New\\Downloads'

    @patch('app_dialogs.QFileDialog.getExistingDirectory')
    def test_browse_folder_backup(self, mock_get_dir, app, backup_config):
        """バックアップフォルダブラウズ機能のテスト"""
        mock_get_dir.return_value = 'C:\\New\\Backup'

        dialog = FolderPathDialog()
        dialog.browse_folder('backup')

        mock_get_dir.assert_called_once()
        assert dialog.backup_path.text() == 'C:\\New\\Backup'

    @patch('app_dialogs.QFileDialog.getOpenFileName')
    def test_browse_folder_excel(self, mock_get_file, app, backup_config):
        """Excelファイルブラウズ機能のテスト"""
        mock_get_file.return_value = ('C:\\New\\Excel\\new.xlsm', 'Excel Files (*.xlsx *.xlsm)')

        dialog = FolderPathDialog()
        dialog.browse_folder('excel')

        mock_get_file.assert_called_once()
        assert dialog.excel_path.text() == 'C:\\New\\Excel\\new.xlsm'

    @patch('app_dialogs.os.path.exists')
    @patch('app_dialogs.os.startfile')
    def test_open_folder_exists(self, mock_startfile, mock_exists, app, backup_config):
        """フォルダオープン機能のテスト（フォルダが存在する場合）"""
        mock_exists.return_value = True

        dialog = FolderPathDialog()
        dialog.open_folder('C:\\Test\\Folder')

        mock_exists.assert_called_once_with('C:\\Test\\Folder')
        mock_startfile.assert_called_once_with('C:\\Test\\Folder')

    @patch('app_dialogs.os.path.exists')
    @patch('app_dialogs.QMessageBox.warning')
    def test_open_folder_not_exists(self, mock_warning, mock_exists, app, backup_config):
        """フォルダオープン機能のテスト（フォルダが存在しない場合）"""
        mock_exists.return_value = False

        dialog = FolderPathDialog()
        dialog.open_folder('C:\\Test\\NonExistingFolder')

        mock_exists.assert_called_once_with('C:\\Test\\NonExistingFolder')
        mock_warning.assert_called_once()

    def test_accept(self, app, backup_config):
        """ダイアログ確定時の設定保存テスト"""
        dialog = FolderPathDialog()

        # パスを設定
        dialog.downloads_path.setText('C:\\New\\Downloads')
        dialog.excel_path.setText('C:\\New\\Excel\\new.xlsm')
        dialog.backup_path.setText('C:\\New\\Backup')

        # モック化したsuperクラスのacceptメソッドでテスト
        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()
            mock_accept.assert_called_once()

        # 設定が保存されたことを確認
        config = ConfigManager()
        assert config.get_downloads_path() == 'C:\\New\\Downloads'
        assert config.get_excel_path() == 'C:\\New\\Excel\\new.xlsm'
        assert config.get_backup_path() == 'C:\\New\\Backup'


class TestAppearanceDialog:
    def test_init(self, app, backup_config):
        """AppearanceDialogの初期化テスト"""
        dialog = AppearanceDialog()
        assert dialog.windowTitle() == "フォントとウインドウサイズ"
        assert dialog.font_size_input is not None
        assert dialog.window_width_input is not None
        assert dialog.window_height_input is not None

    def test_load_settings(self, app, backup_config):
        """設定読み込み機能のテスト"""
        # 設定ファイルにテスト用サイズを設定
        config = ConfigManager()
        if 'Appearance' not in config.config:
            config.config['Appearance'] = {}
        config.config['Appearance']['font_size'] = '14'
        config.config['Appearance']['window_width'] = '400'
        config.config['Appearance']['window_height'] = '300'
        config.save_config()

        # ダイアログを初期化して設定が読み込まれるか確認
        dialog = AppearanceDialog()

        assert dialog.font_size_input.text() == '14'
        assert dialog.window_width_input.text() == '400'
        assert dialog.window_height_input.text() == '300'

    @patch('app_dialogs.QMessageBox.information')
    def test_accept(self, mock_info, app, backup_config):
        """ダイアログ確定時の設定保存テスト"""
        dialog = AppearanceDialog()

        # サイズを設定
        dialog.font_size_input.setText('16')
        dialog.window_width_input.setText('450')
        dialog.window_height_input.setText('350')

        # モック化したsuperクラスのacceptメソッドでテスト
        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()
            mock_accept.assert_called_once()
            mock_info.assert_called_once()

        # 設定が保存されたことを確認
        config = ConfigManager()
        assert config.get_font_size() == 16
        window_size = config.get_window_size()
        assert window_size[0] == 450
        assert window_size[1] == 350
