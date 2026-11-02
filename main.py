import tkinter as tk

from app.app_window import MedicalDocsAnalyzerGUI
from app.login_dialog import login_if_required
from utils.config_manager import load_config
from utils.file_cleaner import cleanup_old_files
from utils.log_rotation import setup_logging


def main() -> None:
    config = load_config()
    setup_logging(config)
    cleanup_old_files(config)

    root = tk.Tk()
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
