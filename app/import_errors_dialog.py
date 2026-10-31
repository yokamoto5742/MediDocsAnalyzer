import csv
import tkinter as tk
from tkinter import filedialog, ttk

from app import constants
from service.service_excel_reader import RowError

CSV_FILE_TYPES = [(constants.CSV_FILE_TYPE_LABEL, '*.csv')]


def error_rows(errors: list[RowError]) -> list[tuple[str, int, str, str, str]]:
    return [(error.file_name, error.row_number, error.patient_id, error.reason, error.row_text) for error in errors]


class ImportErrorsDialog:
    """取り込めなかった行を一覧にし、コピーと CSV 保存をできるようにする"""

    def __init__(self, parent: tk.Tk, errors: list[RowError]) -> None:
        self.errors = errors
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(constants.IMPORT_ERRORS_TITLE)
        self.dialog.transient(parent)

        frame = ttk.Frame(self.dialog, padding="10")
        frame.grid(row=0, column=0, sticky=tk.NSEW)
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text=constants.IMPORT_ERRORS_HEADING).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tree = ttk.Treeview(frame, columns=constants.ERROR_COLUMNS, show='headings', height=15)
        for column, width in zip(constants.ERROR_COLUMNS, (180, 50, 100, 220, 500)):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width)
        for row in error_rows(errors):
            self.tree.insert('', tk.END, values=row)
        self.tree.grid(row=1, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky=tk.NS)
        self.tree.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, pady=10)
        for column, (label, command) in enumerate((
                (constants.COPY_BUTTON_LABEL, self.copy),
                (constants.SAVE_CSV_BUTTON_LABEL, self.save_csv),
                (constants.CLOSE_BUTTON_LABEL, self.dialog.destroy))):
            ttk.Button(button_frame, text=label, command=command).grid(row=0, column=column, padx=5)

    def copy(self) -> None:
        """Excel に貼り付けられるよう、タブ区切りでクリップボードに入れる"""
        lines = ['\t'.join(constants.ERROR_COLUMNS)]
        lines += ['\t'.join(str(value) for value in row) for row in error_rows(self.errors)]
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append('\n'.join(lines))

    def save_csv(self) -> None:
        path = filedialog.asksaveasfilename(parent=self.dialog, filetypes=CSV_FILE_TYPES, defaultextension='.csv')
        if not path:
            return
        # Excel で開いても文字化けしないよう BOM 付きにする
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(constants.ERROR_COLUMNS)
            writer.writerows(error_rows(self.errors))
