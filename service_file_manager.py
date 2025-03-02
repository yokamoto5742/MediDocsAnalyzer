import os
import shutil
import datetime
from pathlib import Path

from config_manager import ConfigManager


def backup_excel_file(excel_path):
    config = ConfigManager()
    backup_dir = Path(config.get_backup_path())

    if not backup_dir.exists():
        backup_dir.mkdir(parents=True)

    backup_path = backup_dir / Path(excel_path).name

    try:
        shutil.copy2(excel_path, backup_path)
    except Exception as e:
        print(f"バックアップ作成中にエラーが発生しました: {str(e)}")
        raise


def cleanup_old_csv_files(processed_dir: Path):
    current_time = datetime.datetime.now()
    for file in processed_dir.glob("*.csv"):
        file_time = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        if (current_time - file_time).days >= 3:
            try:
                file.unlink()
            except Exception as e:
                print(f"ファイル削除中にエラーが発生しました: {file} - {str(e)}")


def ensure_directories_exist():
    config = ConfigManager()
    directories = [
        Path(config.get_downloads_path()),
        Path(config.get_backup_path()),
        Path(config.get_processed_path())
    ]

    for directory in directories:
        if not directory.exists():
            try:
                directory.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"ディレクトリの作成中にエラーが発生しました: {directory} - {str(e)}")
                raise
