import sys
import pytest
import datetime
from pathlib import Path
from unittest.mock import patch, MagicMock

from service_file_manager import backup_excel_file, cleanup_old_csv_files, ensure_directories_exist


class TestFileManager:
    @patch('service_file_manager.shutil.copy2')
    @patch('service_file_manager.Path')
    @patch('service_file_manager.ConfigManager')
    def test_backup_excel_file(self, mock_config_manager, mock_path, mock_copy2):
        """バックアップファイル作成機能のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定
        mock_backup_dir = MagicMock()
        mock_backup_dir.exists.return_value = True
        mock_path.return_value = mock_backup_dir

        # ファイル名を取得するためのPathモック
        mock_excel_path = MagicMock()
        mock_excel_path.name = 'test.xlsm'
        mock_path.side_effect = [mock_backup_dir, mock_excel_path]

        # バックアップパスのモック設定
        backup_path = MagicMock()
        mock_backup_dir.__truediv__.return_value = backup_path

        # 関数実行
        backup_excel_file('C:/Excel/test.xlsm')

        # 結果を確認
        mock_config_manager.assert_called_once()
        mock_config.get_backup_path.assert_called_once()

        # バックアップディレクトリの存在確認
        mock_backup_dir.exists.assert_called_once()

        # バックアップ処理（コピー）が実行されたことを確認
        mock_copy2.assert_called_once()
        args = mock_copy2.call_args[0]
        assert args[0] == 'C:/Excel/test.xlsm'  # コピー元
        assert args[1] == backup_path  # コピー先

    @patch('service_file_manager.shutil.copy2')
    @patch('service_file_manager.Path')
    @patch('service_file_manager.ConfigManager')
    def test_backup_excel_file_create_dir(self, mock_config_manager, mock_path, mock_copy2):
        """バックアップディレクトリが存在しない場合のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定（ディレクトリが存在しないケース）
        mock_backup_dir = MagicMock()
        mock_backup_dir.exists.return_value = False
        mock_path.return_value = mock_backup_dir

        # ファイル名を取得するためのPathモック
        mock_excel_path = MagicMock()
        mock_excel_path.name = 'test.xlsm'
        mock_path.side_effect = [mock_backup_dir, mock_excel_path]

        # 関数実行
        backup_excel_file('C:/Excel/test.xlsm')

        # ディレクトリが作成されたことを確認
        mock_backup_dir.mkdir.assert_called_once_with(parents=True)

        # バックアップ処理が実行されたことを確認
        mock_copy2.assert_called_once()

    @patch('service_file_manager.datetime.datetime')
    def test_cleanup_old_csv_files(self, mock_dt):
        """古いCSVファイルの削除テスト"""
        # 現在時刻を固定
        current_time = MagicMock()
        current_time.now = MagicMock(return_value=current_time)
        mock_dt.now.return_value = current_time

        # モックディレクトリと3つのファイル
        processed_dir = MagicMock(spec=Path)

        file1 = MagicMock()  # 新しいファイル
        file2 = MagicMock()  # 古いファイル
        file3 = MagicMock()  # 非CSVファイル

        file1.suffix = '.csv'
        file2.suffix = '.csv'
        file3.suffix = '.txt'

        # glob()がCSVファイルだけを返すように設定（実際の動作に合わせる）
        processed_dir.glob.return_value = [file1, file2]

        # fromtimestampメソッドをモック
        file1_time = MagicMock()
        file2_time = MagicMock()

        # 日付比較の結果を設定
        diff1 = MagicMock()
        diff1.days = 2  # 2日前（削除しない）

        diff2 = MagicMock()
        diff2.days = 4  # 4日前（削除する）

        # 引数に応じて異なる日付オブジェクトを返す
        def from_timestamp_side_effect(timestamp):
            if timestamp is file1.stat().st_mtime:
                return file1_time
            elif timestamp is file2.stat().st_mtime:
                return file2_time
            return MagicMock()

        mock_dt.fromtimestamp.side_effect = from_timestamp_side_effect

        # 日付の差分計算結果を設定
        current_time.__sub__.side_effect = lambda other: diff1 if other is file1_time else diff2

        # 関数を実行
        cleanup_old_csv_files(processed_dir)

        # 古いCSVファイル（3日以上前）のみが削除されることを確認
        file1.unlink.assert_not_called()  # 新しいファイルは削除されない
        file2.unlink.assert_called_once()  # 古いファイルは削除される
        file3.unlink.assert_not_called()  # 非CSVファイルは削除されない

    @patch('service_file_manager.ConfigManager')
    def test_ensure_directories_exist(self, mock_config_manager):
        """必要なディレクトリの存在確認と作成テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = 'C:/Downloads'
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config.get_processed_path.return_value = 'C:/Processed'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定
        with patch('service_file_manager.Path') as mock_path:
            # 各ディレクトリの存在状態をモック
            mock_downloads = MagicMock()
            mock_downloads.exists.return_value = True  # 既に存在する

            mock_backup = MagicMock()
            mock_backup.exists.return_value = False  # 存在しない

            mock_processed = MagicMock()
            mock_processed.exists.return_value = False  # 存在しない

            # Pathのインスタンス化の戻り値を設定
            mock_path.side_effect = [mock_downloads, mock_backup, mock_processed]

            # 関数実行
            ensure_directories_exist()

            # 結果を確認
            mock_config_manager.assert_called_once()
            mock_config.get_downloads_path.assert_called_once()
            mock_config.get_backup_path.assert_called_once()
            mock_config.get_processed_path.assert_called_once()

            # 存在するディレクトリは作成されないことを確認
            mock_downloads.mkdir.assert_not_called()

            # 存在しないディレクトリは作成されることを確認
            mock_backup.mkdir.assert_called_once_with(parents=True, exist_ok=True)
            mock_processed.mkdir.assert_called_once_with(parents=True, exist_ok=True)

    @patch('service_file_manager.ConfigManager')
    def test_ensure_directories_exist_exception(self, mock_config_manager):
        """ディレクトリ作成時の例外処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = 'C:/Downloads'
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config.get_processed_path.return_value = 'C:/Processed'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定
        with patch('service_file_manager.Path') as mock_path:
            # ディレクトリ作成で例外を発生させる
            mock_dir = MagicMock()
            mock_dir.exists.return_value = False
            mock_dir.mkdir.side_effect = PermissionError("アクセス拒否")

            # Pathのインスタンス化の戻り値を設定（すべて同じモックを返す）
            mock_path.return_value = mock_dir

            # 関数実行（例外が発生するはず）
            with pytest.raises(Exception) as excinfo:
                ensure_directories_exist()

            # 例外内容を確認
            # 元のコードが例外をラップしている場合はこれに合わせる
            assert isinstance(excinfo.value, Exception)
            # エラーメッセージにはPermissionErrorの内容が含まれるはず
            assert "アクセス拒否" in str(excinfo.value)
