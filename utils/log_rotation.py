import configparser
import logging
import os
from datetime import datetime, timedelta
from logging.handlers import TimedRotatingFileHandler

from utils import config_manager

LOG_FILE_NAME = 'MediDocsAnalyzer.log'


def setup_logging(config: configparser.ConfigParser) -> None:
    # 凍結時も残るよう、config.ini と同じフォルダ(exe のフォルダ)を基準にする
    log_directory = os.path.join(os.path.dirname(config_manager.CONFIG_PATH),
                                 config.get('LOGGING', 'log_directory', fallback='logs'))
    log_retention_days = config.getint('LOGGING', 'log_retention_days', fallback=7)

    os.makedirs(log_directory, exist_ok=True)

    log_file = os.path.join(log_directory, LOG_FILE_NAME)
    # 古いログは日数で cleanup_old_logs が消すため、ハンドラには世代数を持たせない
    file_handler = TimedRotatingFileHandler(filename=log_file, when='midnight', encoding='utf-8')
    file_handler.suffix = "%Y-%m-%d.log"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[file_handler]
    )

    cleanup_old_logs(log_directory, log_retention_days)


def cleanup_old_logs(log_directory: str, retention_days: int) -> None:
    now = datetime.now()

    for filename in os.listdir(log_directory):
        if filename.endswith('.log') and filename != LOG_FILE_NAME:
            file_path = os.path.join(log_directory, filename)
            file_modification_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            if now - file_modification_time > timedelta(days=retention_days):
                try:
                    os.remove(file_path)
                    logging.info(f"古いログファイルを削除しました: {filename}")
                except OSError as e:
                    logging.error(f"ログファイルの削除中にエラーが発生しました {filename}: {e}")
