import configparser
import os
import shutil
import sys


def get_config_path() -> str:
    if getattr(sys, 'frozen', False):
        # PyInstaller の展開先(_MEIPASS)は終了時に消えるため、exe と同じフォルダへ写して使う
        config_path = os.path.join(os.path.dirname(sys.executable), 'config.ini')
        if not os.path.exists(config_path):
            shutil.copyfile(os.path.join(getattr(sys, '_MEIPASS'), 'config.ini'), config_path)
        return config_path

    return os.path.join(os.path.dirname(__file__), 'config.ini')


CONFIG_PATH = get_config_path()


def load_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    with open(CONFIG_PATH, encoding='utf-8') as f:
        config.read_file(f)
    return config


def save_config(config: configparser.ConfigParser) -> None:
    with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
        config.write(configfile)


def update_config(sections: dict[str, dict[str, str]], replace: bool = False) -> None:
    """ほかの画面の変更を消さないよう、保存直前に読み直してから sections の値を書き換える。
    replace なら、そのセクションにある古い値を消してから書く"""
    config = load_config()
    for section, values in sections.items():
        if replace or not config.has_section(section):
            config[section] = {}
        config[section].update(values)
    save_config(config)
