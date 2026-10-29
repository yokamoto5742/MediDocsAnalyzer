import configparser
import os
import sys


def get_config_path() -> str:
    if getattr(sys, 'frozen', False):
        # PyInstallerでビルドされた実行ファイルの場合
        base_path = getattr(sys, '_MEIPASS')
    else:
        # 通常のPythonスクリプトとして実行される場合
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, 'config.ini')


CONFIG_PATH = get_config_path()


def load_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    with open(CONFIG_PATH, encoding='utf-8') as f:
        config.read_file(f)
    return config


def save_config(config: configparser.ConfigParser) -> None:
    with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
        config.write(configfile)
