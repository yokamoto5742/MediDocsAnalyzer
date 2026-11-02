import configparser
import logging
import os
from datetime import datetime
from pathlib import Path


def cleanup_old_files(config: configparser.ConfigParser) -> None:
    """
    設定ファイルの内容に基づき古いファイルを削除する

    Args:
        config: 設定ファイルオブジェクト
    """
    if not config.getboolean('FileCleanup', 'enabled', fallback=True):
        logging.info("ファイルクリーンアップが無効になっています")
        return

    retention_days = config.getint('FileCleanup', 'retention_days', fallback=1)
    output_path = config.get('PATHS', 'output_dir', fallback='')

    if not output_path or not os.path.exists(output_path):
        logging.warning(f"出力パスが存在しません: {output_path}")
        return

    _delete_old_files(output_path, retention_days)


def _delete_old_files(directory: str, retention_days: int) -> None:
    """
    指定ディレクトリ内の古いファイルを削除する

    Args:
        directory: 削除対象のディレクトリ
        retention_days: ファイル保持日数
    """
    now = datetime.now()
    deleted_count = 0

    try:
        for file_path in Path(directory).iterdir():
            if not file_path.is_file():
                continue

            file_modification_time = datetime.fromtimestamp(file_path.stat().st_mtime)
            age_days = (now - file_modification_time).days

            if age_days >= retention_days:
                try:
                    file_path.unlink()
                    deleted_count += 1
                    logging.info(f"古いファイルを削除しました: {file_path.name} (経過日数: {age_days}日)")
                except OSError as e:
                    logging.error(f"ファイルの削除中にエラーが発生しました {file_path.name}: {str(e)}")

        if deleted_count > 0:
            logging.info(f"合計 {deleted_count} 個のファイルを削除しました")
    except Exception as e:
        logging.error(f"ファイルクリーンアップ中にエラーが発生しました: {str(e)}")
