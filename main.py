import logging
import tkinter as tk
from tkinter import messagebox
from types import TracebackType

from app import constants
from app.app_window import MedicalDocsAnalyzerGUI
from app.login_dialog import login_if_required
from utils.config_manager import load_config
from utils.file_cleaner import cleanup_old_files
from utils.log_rotation import setup_logging


def report_callback_error(exc_type: type[BaseException], exc_value: BaseException,
                          exc_traceback: TracebackType | None) -> None:
    """ボタンなどの処理で拾えなかった例外をログに残して表示する。既定では stderr に出るだけで、exe では見えない"""
    logging.error('想定外のエラー', exc_info=(exc_type, exc_value, exc_traceback))
    messagebox.showerror(constants.ERROR_TITLE, constants.UNEXPECTED_ERROR_MESSAGE.format(error=exc_value))


def main() -> None:
    config = load_config()
    setup_logging(config)
    cleanup_old_files(config)

    root = tk.Tk()
    root.report_callback_exception = report_callback_error
    # ログインが済むまでメイン画面は隠しておく
    root.withdraw()
    can_start, user = login_if_required(root)
    if not can_start:
        root.destroy()
        return
    root.deiconify()
    MedicalDocsAnalyzerGUI(root, user)
    root.mainloop()


if __name__ == "__main__":
    main()
