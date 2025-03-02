import sys
import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMessageBox

from service_csv_excel_transfer import transfer_csv_to_excel
from config_manager import ConfigManager


@pytest.fixture
def app():
    """テスト用のQApplicationを提供するフィクスチャ"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


class TestCsvExcelTransfer:
    @patch('service_csv_excel_transfer.ConfigManager')
    @patch('service_csv_excel_transfer.ensure_directories_exist')
    @patch('service_csv_excel_transfer.cleanup_old_csv_files')
    @patch('service_csv_excel_transfer.find_latest_csv')
    @patch('service_csv_excel_transfer.read_csv_with_encoding')
    @patch('service_csv_excel_transfer.convert_date_format')
    @patch('service_csv_excel_transfer.process_csv_data')
    @patch('service_csv_excel_transfer.write_data_to_excel')
    @patch('service_csv_excel_transfer.backup_excel_file')
    @patch('service_csv_excel_transfer.process_completed_csv')
    @patch('service_csv_excel_transfer.open_and_sort_excel')
    def test_transfer_csv_to_excel_success(self, mock_open_sort, mock_process_csv,
                                           mock_backup, mock_write, mock_process_data,
                                           mock_convert_date, mock_read_csv, mock_find_csv,
                                           mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """CSVからExcelへの正常な転送処理のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 各関数のモック戻り値設定
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = "mock_dataframe"
        mock_convert_date.return_value = "mock_dataframe_with_date"
        mock_process_data.return_value = "mock_processed_dataframe"
        mock_write.return_value = True  # 書き込み成功

        # 関数実行
        transfer_csv_to_excel()

        # 各関数が正しく呼ばれたことを確認
        mock_ensure_dirs.assert_called_once()
        mock_cleanup.assert_called_once_with(Path("C:/Processed"))
        mock_find_csv.assert_called_once_with("C:/Downloads")
        mock_read_csv.assert_called_once_with("C:/Downloads/test.csv")
        mock_convert_date.assert_called_once_with("mock_dataframe")
        mock_process_data.assert_called_once_with("mock_dataframe_with_date")
        mock_write.assert_called_once_with("C:/Excel/test.xlsm", "mock_processed_dataframe")
        mock_backup.assert_called_once_with("C:/Excel/test.xlsm")
        mock_process_csv.assert_called_once_with("C:/Downloads/test.csv")
        mock_open_sort.assert_called_once_with("C:/Excel/test.xlsm")

    @patch('service_csv_excel_transfer.ConfigManager')
    @patch('service_csv_excel_transfer.ensure_directories_exist')
    @patch('service_csv_excel_transfer.cleanup_old_csv_files')
    @patch('service_csv_excel_transfer.find_latest_csv')
    @patch('service_csv_excel_transfer.QMessageBox.warning')
    def test_transfer_csv_to_excel_no_csv(self, mock_warning, mock_find_csv,
                                          mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """CSVファイルが見つからない場合のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # CSVファイルが見つからない場合
        mock_find_csv.return_value = None

        # 関数実行
        transfer_csv_to_excel()

        # 警告が表示されることを確認
        mock_warning.assert_called_once()
        args = mock_warning.call_args[0]
        assert args[1] == "警告"
        assert "CSVファイルが見つかりません" in args[2]

    @patch('service_csv_excel_transfer.ConfigManager')
    @patch('service_csv_excel_transfer.ensure_directories_exist')
    @patch('service_csv_excel_transfer.cleanup_old_csv_files')
    @patch('service_csv_excel_transfer.find_latest_csv')
    @patch('service_csv_excel_transfer.read_csv_with_encoding')
    @patch('service_csv_excel_transfer.convert_date_format')
    @patch('service_csv_excel_transfer.process_csv_data')
    @patch('service_csv_excel_transfer.write_data_to_excel')
    @patch('service_csv_excel_transfer.QMessageBox.critical')
    def test_transfer_csv_to_excel_write_error(self, mock_critical, mock_write, mock_process_data,
                                               mock_convert_date, mock_read_csv, mock_find_csv,
                                               mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """Excel書き込みエラーのテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 各関数のモック戻り値設定
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = "mock_dataframe"
        mock_convert_date.return_value = "mock_dataframe_with_date"
        mock_process_data.return_value = "mock_processed_dataframe"
        mock_write.return_value = False  # 書き込み失敗

        # 関数実行
        transfer_csv_to_excel()

        # バックアップや後続処理が呼ばれないことを確認
        assert not mock_critical.called

    @patch('service_csv_excel_transfer.ConfigManager')
    @patch('service_csv_excel_transfer.ensure_directories_exist')
    @patch('service_csv_excel_transfer.cleanup_old_csv_files')
    @patch('service_csv_excel_transfer.find_latest_csv')
    @patch('service_csv_excel_transfer.QMessageBox.critical')
    def test_transfer_csv_to_excel_exception(self, mock_critical, mock_find_csv,
                                             mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """例外発生時のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 例外を発生させる
        mock_find_csv.side_effect = Exception("テストエラー")

        # 関数実行
        transfer_csv_to_excel()

        # エラーメッセージが表示されることを確認
        mock_critical.assert_called_once()
        args = mock_critical.call_args[0]
        assert args[1] == "エラー"
        assert "テストエラー" in args[2]
