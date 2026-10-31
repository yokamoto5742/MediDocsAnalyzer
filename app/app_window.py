import tkinter as tk
from datetime import date
from tkinter import messagebox, ttk

from tkcalendar import DateEntry

from app import __version__, constants
from app.import_errors_dialog import ImportErrorsDialog
from app.login_dialog import ChangePasswordDialog
from app.master_dialog import MasterDialog
from app.result_view import ResultView
from app.settings_dialog import SettingsDialog
from service.service_excel_reader import Draft, NoDataError
from service.service_master import ADMIN_ROLE, Staff
from service.service_medical_docs_analyzer import ImportPlan, MedicalDocsAnalyzer
from utils.config_manager import load_config, save_config
from utils.date_period import PERIOD_NAMES, get_period_range, parse_date

DEFAULT_YEAR_START_MONTH = 4


class MedicalDocsAnalyzerGUI:
    def __init__(self, root: tk.Tk, user: Staff | None = None) -> None:
        """user はログインした担当者。None ならログインなしで、すべての操作をできる"""
        self.root = root
        self.user = user
        self.root.title(constants.LOGGED_IN_TITLE_FORMAT.format(version=__version__, name=user.name) if user
                        else constants.WINDOW_TITLE_FORMAT.format(version=__version__))
        self.config = load_config()
        self.analyzer = MedicalDocsAnalyzer()

        window_width = self.config.getint('Appearance', 'window_width')
        window_height = self.config.getint('Appearance', 'window_height')
        self.root.geometry(f"{window_width}x{window_height}")

        # 画面に出している集計の期間。未集計なら None
        self.period: tuple[date, date] | None = None
        self._setup_gui()
        self.backup_database()
        self.refresh_overview()

    def _setup_gui(self) -> None:
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self._setup_date_frame(main_frame)
        self._setup_buttons(main_frame)

        self.result_view = ResultView(self.root, self.export_overall, self.export_staff, self.show_staff_summary)
        self.result_view.notebook.grid(row=0, column=1, sticky=tk.NSEW, padx=(0, 10), pady=10)
        self._apply_permissions()

    def _apply_permissions(self) -> None:
        """職員は Excel読込・マスタ・設定を使えず、自分の個人集計だけを見られる"""
        if self.user is None or self.user.role == ADMIN_ROLE:
            return
        for button in self.admin_buttons:
            button.state(['disabled'])
        self.result_view.restrict_to_staff(self.user.name)

    def _setup_date_frame(self, parent: ttk.Frame) -> None:
        date_frame = ttk.LabelFrame(parent, text="分析期間", padding="5")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        saved_start_date, saved_end_date = self._load_saved_dates()
        year_start_month = self.config.getint('Analysis', 'year_start_month', fallback=DEFAULT_YEAR_START_MONTH)
        # 4月始まりなら 1〜3月は前年の年度として表示する
        initial_year = saved_start_date.year - (saved_start_date.month < year_start_month)
        self._setup_period_selector(date_frame, initial_year, year_start_month)

        self.start_date = self._create_date_entry(date_frame, 3, "開始日", saved_start_date)
        self.end_date = self._create_date_entry(date_frame, 4, "終了日", saved_end_date)

    @staticmethod
    def _create_date_entry(parent: ttk.LabelFrame, row: int, label: str, initial_date: date) -> DateEntry:
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=5)
        date_entry = DateEntry(parent, width=12, background='darkblue',
                               foreground='white', borderwidth=2,
                               locale='ja_JP', date_pattern='yyyy/mm/dd')
        date_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        date_entry.set_date(initial_date)
        return date_entry

    def _load_saved_dates(self) -> tuple[date, date]:
        today = date.today().isoformat()
        start = self.config.get('Analysis', 'start_date', fallback=today)
        end = self.config.get('Analysis', 'end_date', fallback=today)
        return date.fromisoformat(start), date.fromisoformat(end)

    def _setup_period_selector(self, date_frame: ttk.LabelFrame, initial_year: int,
                               year_start_month: int) -> None:
        ttk.Label(date_frame, text="年").grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(value=str(initial_year))
        ttk.Spinbox(date_frame, from_=2000, to=2100, width=6, textvariable=self.year_var,
                    command=self._apply_period).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        self.year_start_month_var = tk.IntVar(value=year_start_month)
        year_start_frame = ttk.Frame(date_frame)
        year_start_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W)
        for column, month in enumerate((1, 4)):
            ttk.Radiobutton(year_start_frame, text=f"{month}月始まり", value=month,
                            variable=self.year_start_month_var,
                            command=self._apply_period).grid(row=0, column=column, padx=5)

        ttk.Label(date_frame, text="期間").grid(row=2, column=0, padx=5, pady=5)
        self.period_var = tk.StringVar()
        period_combobox = ttk.Combobox(date_frame, width=10, state='readonly',
                                       values=PERIOD_NAMES, textvariable=self.period_var)
        period_combobox.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        period_combobox.bind('<<ComboboxSelected>>', self._apply_period)

    def _apply_period(self, _event: tk.Event | None = None) -> None:
        """選んだ年・年の始まり・期間から開始日と終了日を入れる"""
        period_name = self.period_var.get()
        year_text = self.year_var.get()
        if not period_name or not year_text.isdigit():
            return

        start_date, end_date = get_period_range(
            int(year_text), period_name, self.year_start_month_var.get())
        self.start_date.set_date(start_date)
        self.end_date.set_date(end_date)

    def _setup_buttons(self, parent: ttk.Frame) -> None:
        ttk.Button(parent, text="集計開始", command=self.start_analysis).grid(
            row=2, column=0, columnspan=2, pady=10)
        master_button = ttk.Button(parent, text=constants.MASTER_BUTTON_LABEL, command=self.open_master)
        master_button.grid(row=3, column=0, columnspan=2, pady=5)
        load_button = ttk.Button(parent, text=constants.EXCEL_IMPORT_BUTTON_LABEL, command=self.load_excel)
        load_button.grid(row=4, column=0, columnspan=2, pady=5)
        settings_button = ttk.Button(parent, text=constants.SETTINGS_BUTTON_LABEL, command=self.open_settings)
        settings_button.grid(row=5, column=0, columnspan=2, pady=5)
        if self.user is not None:
            ttk.Button(parent, text=constants.CHANGE_PASSWORD_BUTTON_LABEL, command=self.change_password).grid(
                row=6, column=0, columnspan=2, pady=5)
        ttk.Button(parent, text="閉じる", command=self.root.quit).grid(
            row=7, column=0, columnspan=2, pady=5)
        self.admin_buttons = [load_button, master_button, settings_button]

    def backup_database(self) -> None:
        try:
            self.analyzer.backup_database(date.today())
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.BACKUP_ERROR_MESSAGE.format(error=e))

    def refresh_overview(self, drafts: list[Draft] | None = None) -> None:
        """drafts が None なら元データを読み直す。起動のたびにダイアログを出さないよう、失敗はタブ内に出す"""
        try:
            overview = (self.analyzer.load_overview(date.today()) if drafts is None
                        else self.analyzer.overview(drafts, date.today()))
            self.result_view.show_overview(overview)
        except Exception as e:
            self.result_view.show_overview_error(e)

    def load_excel(self) -> None:
        """旧 DB(xlsx)を確認してから取り込む"""
        try:
            plan = self.analyzer.prepare_legacy_import()
            if not messagebox.askyesno(constants.CONFIRM_TITLE, self._import_confirm_message(plan)):
                return
            count = self.analyzer.apply_import(plan)
        except NoDataError:
            messagebox.showerror(constants.ERROR_TITLE, constants.NO_SOURCE_DATA_MESSAGE)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.LOAD_ERROR_MESSAGE.format(error=e))
        else:
            messagebox.showinfo(constants.COMPLETE_TITLE, constants.LOAD_COMPLETE_MESSAGE.format(count=count))
            if plan.errors:
                ImportErrorsDialog(self.root, plan.errors)

    @staticmethod
    def _import_confirm_message(plan: ImportPlan) -> str:
        return constants.IMPORT_CONFIRM_MESSAGE.format(
            start_date=plan.start_date, end_date=plan.end_date, deleted=plan.deleted_count,
            inserted=len(plan.source.documents), errors=len(plan.errors))

    def _import_source(self) -> ImportPlan | None:
        """元データ(xlsm)を確認なしで取り込む。取り込めたら計画を、失敗したら None を返す"""
        try:
            plan = self.analyzer.prepare_source_import()
            self.analyzer.apply_import(plan)
        except NoDataError:
            messagebox.showerror(constants.ERROR_TITLE, constants.NO_SOURCE_DATA_MESSAGE)
            return None
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.LOAD_ERROR_MESSAGE.format(error=e))
            return None
        if plan.errors:
            ImportErrorsDialog(self.root, plan.errors)
        return plan

    def start_analysis(self) -> None:
        # get_date() は手入力がフォーカスを外れるまで反映されないため、入力欄の文字列を読む
        try:
            start_date = parse_date(self.start_date.get())
            end_date = parse_date(self.end_date.get())
        except ValueError:
            messagebox.showerror(constants.ERROR_TITLE, constants.INVALID_DATE_MESSAGE)
            return

        if start_date > end_date:
            messagebox.showerror(constants.ERROR_TITLE, constants.START_AFTER_END_MESSAGE)
            return

        plan = self._import_source()
        if plan is None:
            return
        self.refresh_overview(plan.source.drafts)

        try:
            self._save_analysis_settings(start_date, end_date)
            summary = self.analyzer.overall_summary(start_date, end_date)
        except NoDataError:
            messagebox.showerror(constants.ERROR_TITLE, constants.NO_DATA_MESSAGE)
            return
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE,
                                 constants.ANALYSIS_ERROR_MESSAGE.format(error=e))
            return

        self.period = (start_date, end_date)
        self.result_view.show_overall(summary)
        self.result_view.set_staff_names(summary.staff_members)
        self.show_staff_summary()

    def show_staff_summary(self) -> None:
        staff_name = self.result_view.staff_var.get()
        if self.period is None or not staff_name:
            return
        try:
            self.result_view.show_staff(self.analyzer.staff_summary(staff_name, *self.period))
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.ANALYSIS_ERROR_MESSAGE.format(error=e))

    def export_overall(self) -> None:
        if self.period is None:
            messagebox.showerror(constants.ERROR_TITLE, constants.NOT_ANALYZED_ERROR_MESSAGE)
            return
        try:
            self.analyzer.run_analysis(*self.period)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error=e))

    def export_staff(self) -> None:
        staff_name = self.result_view.staff_var.get()
        if self.period is None or not staff_name:
            messagebox.showerror(constants.ERROR_TITLE, constants.NOT_ANALYZED_ERROR_MESSAGE)
            return
        try:
            self.analyzer.export_staff(staff_name, *self.period)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error=e))

    def _save_analysis_settings(self, start_date: date, end_date: date) -> None:
        # 設定画面での変更を消さないよう、保存直前に読み直す
        self.config = load_config()
        self.config['Analysis'].update({
            'start_date': start_date.isoformat(),
            'end_date': end_date.isoformat(),
            'year_start_month': str(self.year_start_month_var.get())
        })
        save_config(self.config)

    def change_password(self) -> None:
        assert self.user is not None
        try:
            dialog = ChangePasswordDialog(self.root, self.analyzer, self.user)
            self.root.wait_window(dialog.dialog)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.ANALYSIS_ERROR_MESSAGE.format(error=e))
            return
        if dialog.changed:
            messagebox.showinfo(constants.COMPLETE_TITLE, constants.PASSWORD_CHANGED_MESSAGE)

    def open_master(self) -> None:
        try:
            MasterDialog(self.root, self.analyzer)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.MASTER_OPEN_ERROR_MESSAGE.format(error=e))

    def open_settings(self) -> None:
        SettingsDialog(self.root)
