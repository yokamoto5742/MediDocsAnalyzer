import configparser
import os
import sys
from pathlib import Path
from typing import Final, List

def get_config_path() -> Path:
    # 実行ファイルのディレクトリを取得
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(os.path.dirname(os.path.abspath(__file__)))
    return base_path / 'config.ini'

CONFIG_PATH = get_config_path()

class ConfigManager:
    def __init__(self, config_file: Path | str = CONFIG_PATH) -> None:
        self.config_file: Path = Path(config_file)
        self.config: configparser.ConfigParser = configparser.ConfigParser()
        self.load_config()

    def load_config(self) -> None:
        if not self.config_file.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_file}")

        try:
            self.config.read(self.config_file, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                content: str = self.config_file.read_bytes().decode('cp932')
                self.config.read_string(content)
            except (UnicodeDecodeError, OSError) as e:
                raise OSError(f"Failed to load config: {e}") from e

    def get_exclude_docs(self) -> List[str]:
        if 'ExcludeDocs' not in self.config:
            return []
        docs = self.config.get('ExcludeDocs', 'list', fallback='')
        return [doc.strip() for doc in docs.split(',') if doc.strip()]

    def get_exclude_doctors(self) -> List[str]:
        if 'ExcludeDoctors' not in self.config:
            return []
        doctors = self.config.get('ExcludeDoctors', 'list', fallback='')
        return [doctor.strip() for doctor in doctors.split(',') if doctor.strip()]

    def get_downloads_path(self) -> str:
        if 'Paths' not in self.config:
            return str(Path.home() / "Downloads")
        return self.config.get('Paths', 'downloads_path', fallback=str(Path.home() / "Downloads"))

    def get_excel_path(self) -> str:
        if 'Paths' not in self.config:
            return r"C:\Shinseikai\CSV2XL\医療文書担当一覧.xlsm"
        return self.config.get('Paths', 'excel_path', fallback=r"C:\Shinseikai\CSV2XL\医療文書担当一覧.xlsm")

    def set_downloads_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['downloads_path'] = path
        self.save_config()

    def set_excel_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['excel_path'] = path
        self.save_config()

    def get_backup_path(self) -> str:
        if 'Paths' not in self.config:
            return r"C:\Shinseikai\CSV2XL\backup"
        return self.config.get('Paths', 'backup_path', fallback=r"C:\Shinseikai\CSV2XL\backup")

    def set_backup_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['backup_path'] = path
        self.save_config()

    def get_folder_dialog_size(self) -> tuple[int, int]:
        if 'DialogSize' not in self.config:
            return 400, 200  # デフォルトのダイアログサイズ
        width = self.config.getint('DialogSize', 'folder_dialog_width', fallback=400)
        height = self.config.getint('DialogSize', 'folder_dialog_height', fallback=200)
        return width, height

    def get_processed_path(self) -> str:
        if 'Paths' not in self.config:
            return r"C:\Shinseikai\CSV2XL\processed"
        return self.config.get('Paths', 'processed_path', fallback=r"C:\Shinseikai\CSV2XL\processed")

    def set_processed_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['processed_path'] = path
        self.save_config()

    def get_font_size(self) -> int:
        if 'Appearance' not in self.config:
            return 9  # デフォルトのフォントサイズ
        return self.config.getint('Appearance', 'font_size', fallback=9)

    def get_window_size(self) -> tuple[int, int]:
        if 'Appearance' not in self.config:
            return 300, 200  # デフォルトのウィンドウサイズ
        width = self.config.getint('Appearance', 'window_width', fallback=300)
        height = self.config.getint('Appearance', 'window_height', fallback=200)
        return width, height

    def set_font_size(self, size: int) -> None:
        if 'Appearance' not in self.config:
            self.config['Appearance'] = {}
        self.config['Appearance']['font_size'] = str(size)
        self.save_config()

    def set_window_size(self, width: int, height: int) -> None:
        if 'Appearance' not in self.config:
            self.config['Appearance'] = {}
        self.config['Appearance']['window_width'] = str(width)
        self.config['Appearance']['window_height'] = str(height)
        self.save_config()

    def get_share_button_position(self) -> tuple[int, int]:
        if 'ButtonPosition' not in self.config:
            return 1450, 180  # デフォルトの座標
        x = self.config.getint('ButtonPosition', 'share_button_x', fallback=1450)
        y = self.config.getint('ButtonPosition', 'share_button_y', fallback=180)
        return x, y

    def set_share_button_position(self, x: int, y: int) -> None:
        if 'ButtonPosition' not in self.config:
            self.config['ButtonPosition'] = {}
        self.config['ButtonPosition']['share_button_x'] = str(x)
        self.config['ButtonPosition']['share_button_y'] = str(y)
        self.save_config()

    def get_share_button_wait_time(self) -> int:
        if 'ButtonPosition' not in self.config:
            return 3  # デフォルトの待機時間（秒）
        return self.config.getint('ButtonPosition', 'share_button_wait_time', fallback=3)

    def set_share_button_wait_time(self, seconds: int) -> None:
        if 'ButtonPosition' not in self.config:
            self.config['ButtonPosition'] = {}
        self.config['ButtonPosition']['share_button_wait_time'] = str(seconds)
        self.save_config()

    def save_config(self) -> None:
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except (IOError, OSError) as e:
            raise OSError(f"Failed to load config: {e}") from e

    def _ensure_section(self, section: str) -> None:
        """設定セクションが存在することを確認し、必要に応じて作成する"""
        if section not in self.config:
            self.config[section] = {}
