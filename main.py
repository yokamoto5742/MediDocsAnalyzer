import tkinter as tk

from app.app_window import MedicalDocsAnalyzerGUI
from app.login_dialog import login_if_required


def main() -> None:
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
