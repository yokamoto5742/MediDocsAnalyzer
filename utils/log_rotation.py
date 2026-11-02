import logging
import os
from datetime import datetime, timedelta
from logging.handlers import TimedRotatingFileHandler

import configparser


def setup_logging(config: configparser.ConfigParser):
    log_directory = os.path.join(os.path.dirname(__file__), config.get('LOGGING', 'log_directory', fallback='logs'))
    log_retention_days = config.getint('LOGGING', 'log_retention_days', fallback=7)

    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

    parent_dir_name = os.path.basename(os.path.dirname(os.path.dirname(log_directory)))
    log_file = os.path.join(log_directory, f'{parent_dir_name}.log')

    file_handler = TimedRotatingFileHandler(filename=log_file, when='midnight', backupCount=log_retention_days)
    file_handler.suffix = "%Y-%m-%d.log"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[file_handler]
    )

    cleanup_old_logs(log_directory, log_retention_days)


def cleanup_old_logs(log_directory: str, retention_days: int):
    now = datetime.now()
    parent_dir_name = os.path.basename(os.path.dirname(os.path.dirname(log_directory)))
    main_log_file = f'{parent_dir_name}.log'

    for filename in os.listdir(log_directory):
        if filename.endswith('.log') and filename != main_log_file:
            file_path = os.path.join(log_directory, filename)
            file_modification_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            if now - file_modification_time > timedelta(days=retention_days):
                try:
                    os.remove(file_path)
                    logging.info(f"古いログファイルを削除しました: {filename}")
                except OSError as e:
                    logging.error(f"ログファイルの削除中にエラーが発生しました {filename}: {str(e)}")
