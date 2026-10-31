import tkinter as tk
from collections.abc import Callable
from tkinter import filedialog, messagebox, ttk

from app import constants
from utils.config_manager import load_config, save_config

EXCEL_FILE_TYPES = [(constants.EXCEL_FILE_TYPE_LABEL, '*.xlsx *.xlsm')]
DATABASE_FILE_TYPES = [(constants.DATABASE_FILE_TYPE_LABEL, '*.db')]
# (セクション, キー, ラベル, 最小値, 最大値)
NUMBER_SETTINGS: list[tuple[str, str, str, int, int]] = [
    ('Appearance', 'window_width', constants.WINDOW_WIDTH_LABEL, 800, 3840),
    ('Appearance', 'window_height', constants.WINDOW_HEIGHT_LABEL, 600, 2160),
    ('Backup', 'keep_generations', constants.KEEP_GENERATIONS_LABEL, 1, 365),
]


class SettingsDialog:
    """ファイルパスと数値の設定を画面で変える。期間と担当者・診療科はメイン画面とマスタで変えるので出さない"""

    def __init__(self, parent: tk.Tk) -> None:
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(constants.SETTINGS_TITLE)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        config = load_config()
        path_rows: list[tuple[str, str, Callable[[str], None]]] = [
            ('source_file_path', constants.SOURCE_FILE_LABEL, self._choose_file),
            ('database_path', constants.DATABASE_FILE_LABEL, self._choose_database),
            ('legacy_database_path', constants.LEGACY_DATABASE_FILE_LABEL, self._choose_file),
            ('backup_dir', constants.BACKUP_DIR_LABEL, self._choose_directory),
            ('template_path', constants.TEMPLATE_FILE_LABEL, self._choose_file),
            ('output_dir', constants.OUTPUT_DIR_LABEL, self._choose_directory),
        ]
        self.path_vars = {key: tk.StringVar(value=config['PATHS'].get(key, '')) for key, _, _ in path_rows}
        self.number_vars = {key: tk.StringVar(value=config.get(section, key, fallback=''))
                            for section, key, _, _, _ in NUMBER_SETTINGS}

        frame = ttk.Frame(self.dialog, padding="10")
        frame.grid(row=0, column=0)
        path_frame = ttk.LabelFrame(frame, text=constants.PATH_SETTINGS_HEADING, padding="5")
        path_frame.grid(row=0, column=0, sticky=tk.EW)
        for row, (key, label, choose_command) in enumerate(path_rows):
            self._add_path_row(path_frame, row, label, key, choose_command)

        number_frame = ttk.LabelFrame(frame, text=constants.OTHER_SETTINGS_HEADING, padding="5")
        number_frame.grid(row=1, column=0, sticky=tk.EW, pady=(10, 0))
        for row, (_, key, label, minimum, maximum) in enumerate(NUMBER_SETTINGS):
            ttk.Label(number_frame, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
            ttk.Spinbox(number_frame, from_=minimum, to=maximum, width=8,
                        textvariable=self.number_vars[key]).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(number_frame, text=constants.WINDOW_SIZE_NOTE).grid(
            row=len(NUMBER_SETTINGS), column=0, columnspan=2, sticky=tk.W, padx=5)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, pady=10)
        ttk.Button(button_frame, text=constants.SAVE_BUTTON_LABEL, command=self.save).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text=constants.CANCEL_BUTTON_LABEL, command=self.dialog.destroy).grid(
            row=0, column=1, padx=5)

    def _add_path_row(self, parent: ttk.LabelFrame, row: int, label: str, key: str,
                      choose_command: Callable[[str], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent, width=50, textvariable=self.path_vars[key]).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(parent, text=constants.BROWSE_BUTTON_LABEL,
                   command=lambda: choose_command(key)).grid(row=row, column=2, padx=5, pady=5)

    def _choose_file(self, key: str) -> None:
        selected = filedialog.askopenfilename(parent=self.dialog, filetypes=EXCEL_FILE_TYPES)
        if selected:
            self.path_vars[key].set(selected.replace('/', '\\'))

    def _choose_database(self, key: str) -> None:
        # まだ無いファイルも選べるよう、保存ダイアログで選ぶ
        selected = filedialog.asksaveasfilename(parent=self.dialog, filetypes=DATABASE_FILE_TYPES,
                                                defaultextension='.db', confirmoverwrite=False)
        if selected:
            self.path_vars[key].set(selected.replace('/', '\\'))

    def _choose_directory(self, key: str) -> None:
        selected = filedialog.askdirectory(parent=self.dialog)
        if selected:
            self.path_vars[key].set(selected.replace('/', '\\'))

    def _find_invalid_number(self) -> tuple[str, int, int] | None:
        """範囲外か整数でない項目の (ラベル, 最小値, 最大値) を返す。すべて正しければ None"""
        for _, key, label, minimum, maximum in NUMBER_SETTINGS:
            text = self.number_vars[key].get().strip()
            if not text.isdigit() or not minimum <= int(text) <= maximum:
                return label, minimum, maximum
        return None

    def save(self) -> None:
        invalid = self._find_invalid_number()
        if invalid is not None:
            label, minimum, maximum = invalid
            messagebox.showerror(constants.ERROR_TITLE, constants.INVALID_NUMBER_SETTING_MESSAGE.format(
                label=label, minimum=minimum, maximum=maximum), parent=self.dialog)
            return

        # メイン画面が保存した期間を消さないよう、保存直前に読み直す
        config = load_config()
        config['PATHS'].update({key: var.get().strip() for key, var in self.path_vars.items()})
        for section, key, _, _, _ in NUMBER_SETTINGS:
            config[section][key] = str(int(self.number_vars[key].get().strip()))
        save_config(config)
        self.dialog.destroy()
