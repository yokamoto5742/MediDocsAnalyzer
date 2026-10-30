import tkinter as tk
from tkinter import messagebox, ttk

from app import constants
from service.service_auth import MIN_PASSWORD_LENGTH
from service.service_master import Staff
from service.service_medical_docs_analyzer import MedicalDocsAnalyzer


def make_dialog(parent: tk.Tk, title: str) -> tuple[tk.Toplevel, ttk.Frame]:
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.resizable(False, False)
    frame = ttk.Frame(dialog, padding="15")
    frame.grid(row=0, column=0)
    return dialog, frame


def add_entry(frame: ttk.Frame, row: int, label: str, variable: tk.StringVar, secret: bool) -> ttk.Entry:
    ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
    entry = ttk.Entry(frame, width=24, textvariable=variable, show='*' if secret else '')
    entry.grid(row=row, column=1, padx=5, pady=5)
    return entry


class LoginDialog:
    """電子カルテIDとパスワードでログインする。成功すると user に担当者が入る"""

    def __init__(self, parent: tk.Tk, analyzer: MedicalDocsAnalyzer) -> None:
        self.analyzer = analyzer
        self.user: Staff | None = None
        self.must_change_password = False
        self.dialog, frame = make_dialog(parent, constants.LOGIN_TITLE)
        self.emr_id_var = tk.StringVar()
        self.password_var = tk.StringVar()
        add_entry(frame, 0, constants.EMR_ID_LABEL, self.emr_id_var, secret=False).focus_set()
        add_entry(frame, 1, constants.PASSWORD_LABEL, self.password_var, secret=True)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text=constants.LOGIN_BUTTON_LABEL, command=self.login).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text=constants.CANCEL_BUTTON_LABEL, command=self.dialog.destroy).grid(
            row=0, column=1, padx=5)
        self.dialog.bind('<Return>', lambda _event: self.login())

    def login(self) -> None:
        result = self.analyzer.authenticate(self.emr_id_var.get().strip(), self.password_var.get())
        if result is None:
            messagebox.showerror(constants.ERROR_TITLE, constants.LOGIN_FAILED_MESSAGE, parent=self.dialog)
            self.password_var.set('')
            return
        self.user = result.staff
        self.must_change_password = result.must_change_password
        self.dialog.destroy()


class ChangePasswordDialog:
    """新しいパスワードを 2 回入力して変更する。変更できると changed が True になる"""

    def __init__(self, parent: tk.Tk, analyzer: MedicalDocsAnalyzer, user: Staff) -> None:
        self.analyzer = analyzer
        self.user = user
        self.changed = False
        self.dialog, frame = make_dialog(parent, constants.CHANGE_PASSWORD_TITLE)
        self.password_var = tk.StringVar()
        self.confirm_var = tk.StringVar()
        add_entry(frame, 0, constants.NEW_PASSWORD_LABEL, self.password_var, secret=True).focus_set()
        add_entry(frame, 1, constants.CONFIRM_PASSWORD_LABEL, self.confirm_var, secret=True)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text=constants.SAVE_BUTTON_LABEL, command=self.save).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text=constants.CANCEL_BUTTON_LABEL, command=self.dialog.destroy).grid(
            row=0, column=1, padx=5)

    def save(self) -> None:
        password = self.password_var.get()
        if len(password) < MIN_PASSWORD_LENGTH:
            message = constants.PASSWORD_TOO_SHORT_MESSAGE.format(length=MIN_PASSWORD_LENGTH)
        elif password != self.confirm_var.get():
            message = constants.PASSWORD_MISMATCH_MESSAGE
        else:
            assert self.user.id is not None
            self.analyzer.change_password(self.user.id, password)
            self.changed = True
            self.dialog.destroy()
            return
        messagebox.showerror(constants.ERROR_TITLE, message, parent=self.dialog)


def login_if_required(root: tk.Tk) -> tuple[bool, Staff | None]:
    """ログインが必要ならログイン画面を出す。(起動してよいか, ログインした担当者) を返す。

    担当者が None ならログインなしで、すべての操作をできる。
    DB を開けないときは、ファイルパスを直せるようログインなしで起動する。
    """
    analyzer = MedicalDocsAnalyzer()
    try:
        required = analyzer.login_required()
    except Exception as e:
        messagebox.showerror(constants.ERROR_TITLE, constants.LOGIN_CHECK_ERROR_MESSAGE.format(error=e))
        return True, None
    if not required:
        return True, None
    user = ask_login(root, analyzer)
    return user is not None, user


def ask_login(root: tk.Tk, analyzer: MedicalDocsAnalyzer) -> Staff | None:
    """ログイン画面を出し、ログインした担当者を返す。やめたときや仮パスワードを変えなかったときは None"""
    login_dialog = LoginDialog(root, analyzer)
    root.wait_window(login_dialog.dialog)
    if login_dialog.user is None:
        return None
    if login_dialog.must_change_password:
        messagebox.showinfo(constants.CHANGE_PASSWORD_TITLE, constants.MUST_CHANGE_PASSWORD_MESSAGE)
        change_dialog = ChangePasswordDialog(root, analyzer, login_dialog.user)
        root.wait_window(change_dialog.dialog)
        if not change_dialog.changed:
            return None
    return login_dialog.user
