import tkinter as tk
from collections.abc import Callable
from tkinter import filedialog, ttk

from app import constants
from utils.config_manager import load_config, save_config

EXCEL_FILE_TYPES = [(constants.EXCEL_FILE_TYPE_LABEL, '*.xlsx *.xlsm')]
DATABASE_FILE_TYPES = [(constants.DATABASE_FILE_TYPE_LABEL, '*.db')]


class PathSettingsDialog:
    def __init__(self, parent: tk.Tk) -> None:
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(constants.PATH_SETTINGS_TITLE)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        paths_config = load_config()['PATHS']
        path_rows: list[tuple[str, str, Callable[[str], None]]] = [
            ('source_file_path', constants.SOURCE_FILE_LABEL, self._choose_file),
            ('database_path', constants.DATABASE_FILE_LABEL, self._choose_database),
            ('legacy_database_path', constants.LEGACY_DATABASE_FILE_LABEL, self._choose_file),
            ('backup_dir', constants.BACKUP_DIR_LABEL, self._choose_directory),
            ('template_path', constants.TEMPLATE_FILE_LABEL, self._choose_file),
            ('output_dir', constants.OUTPUT_DIR_LABEL, self._choose_directory),
        ]
        self.path_vars = {key: tk.StringVar(value=paths_config.get(key, '')) for key, _, _ in path_rows}

        frame = ttk.Frame(self.dialog, padding="10")
        frame.grid(row=0, column=0)
        for row, (key, label, choose_command) in enumerate(path_rows):
            self._add_path_row(frame, row, label, key, choose_command)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=len(path_rows), column=0, columnspan=3, pady=10)
        ttk.Button(button_frame, text=constants.SAVE_BUTTON_LABEL, command=self.save).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text=constants.CANCEL_BUTTON_LABEL, command=self.dialog.destroy).grid(
            row=0, column=1, padx=5)

    def _add_path_row(self, parent: ttk.Frame, row: int, label: str, key: str,
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

    def save(self) -> None:
        # メモ帳での編集を消さないよう、保存直前に読み直す
        config = load_config()
        config['PATHS'].update({key: var.get().strip() for key, var in self.path_vars.items()})
        save_config(config)
        self.dialog.destroy()
