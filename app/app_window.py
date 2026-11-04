import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from tkinter import messagebox, ttk

from tkcalendar import DateEntry

from app import __version__, constants
from app.date_entry import make_date_entry, read_date_range
from app.folder import open_folder_of
from app.import_errors_dialog import ImportErrorsDialog
from app.login_dialog import ChangePasswordDialog
from app.master_dialog import MasterDialog
from app.result_view import ResultView, TrendPanel
from app.settings_dialog import SettingsDialog
from service.models import Draft, NoDataError
from service.service_data_analyzer import DocumentFilter, DocumentType, trend_subject
from service.service_master import ADMIN_ROLE, Staff
from service.service_medical_docs_analyzer import ImportPlan, MedicalDocsAnalyzer
from utils.config_manager import load_config, update_config
from utils.date_period import PERIOD_NAMES, fiscal_year, get_period_range

# 同梱の config.ini と同じ値にする
DEFAULT_YEAR_START_MONTH = 1
TREND_SECTION = 'Trend'
STAFF_TREND_SECTION = 'StaffTrend'


@dataclass
class TrendRequest:
    """推移タブで選んだ内容"""
    department: str  # 選択欄の表示。「合計」なら全診療科
    document_type: DocumentType  # 左側の集計対象
    staff_name: str | None  # None なら全担当者(全体推移)
    start_date: date
    end_date: date

    @property
    def document_filter(self) -> DocumentFilter:
        """選択欄の表示と集計対象を、サービスへ渡す絞り込みにする"""
        department = None if self.department == constants.TOTAL_ROW_LABEL else self.department
        return DocumentFilter(department, self.document_type)

    @property
    def subject(self) -> str:
        return trend_subject(self.document_filter, self.staff_name)


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
        # 最後に「集計開始」で確定した集計対象。推移タブもこれで絞る
        self.document_type = self._load_saved_document_type()
        self._setup_gui()

    def start(self) -> None:
        """起動時の処理。DB のバックアップ、概要と推移の表示、前回の期間の集計をする"""
        self.backup_database()
        self.refresh_overview()
        self.load_trend()
        self.analyze_saved_period()

    def _setup_gui(self) -> None:
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self._setup_date_frame(main_frame)
        self._setup_document_type_frame(main_frame)
        self._setup_buttons(main_frame)

        self.result_view = ResultView(
            self.root,
            on_export_overall=self.export_overall,
            on_export_staff=self.export_staff,
            on_staff_selected=self.show_staff_summary,
            on_export_trend=self.export_trend,
            on_trend_changed=self.refresh_trend,
            on_export_staff_trend=self.export_staff_trend,
            on_staff_trend_changed=self.refresh_staff_trend,
        )
        self.result_view.notebook.grid(row=0, column=1, sticky=tk.NSEW, padx=(0, 10), pady=10)
        self._apply_permissions()

    @property
    def restricted_staff(self) -> str | None:
        """職員としてログインしたときの本人の氏名。None なら全員分を見られる"""
        if self.user is None or self.user.role == ADMIN_ROLE:
            return None
        return self.user.name

    def _apply_permissions(self) -> None:
        """職員は Excel読込・マスタ・設定を使えず、自分の個人集計だけを見られる"""
        if self.restricted_staff is None:
            return
        for button in self.admin_buttons:
            button.state(['disabled'])
        self.result_view.restrict_to_staff(self.restricted_staff)

    def _setup_date_frame(self, parent: ttk.Frame) -> None:
        date_frame = ttk.LabelFrame(parent, text=constants.ANALYSIS_PERIOD_LABEL, padding="5")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        saved_start_date, saved_end_date = self._load_saved_dates()
        year_start_month = self.config.getint('Analysis', 'year_start_month', fallback=DEFAULT_YEAR_START_MONTH)
        self._setup_period_selector(date_frame, fiscal_year(saved_start_date, year_start_month), year_start_month)

        self.start_date = self._create_date_entry(date_frame, 3, constants.START_DATE_LABEL, saved_start_date)
        self.end_date = self._create_date_entry(date_frame, 4, constants.END_DATE_LABEL, saved_end_date)

    @staticmethod
    def _create_date_entry(parent: ttk.LabelFrame, row: int, label: str, initial_date: date) -> DateEntry:
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=5)
        date_entry = make_date_entry(parent)
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
        ttk.Label(date_frame, text=constants.YEAR_LABEL).grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(value=str(initial_year))
        year_entry = ttk.Entry(date_frame, width=6, textvariable=self.year_var)
        year_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        year_entry.bind('<Return>', self._apply_period)
        year_entry.bind('<FocusOut>', self._apply_period)

        self.year_start_month_var = tk.IntVar(value=year_start_month)
        year_start_frame = ttk.Frame(date_frame)
        year_start_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W)
        for column, month in enumerate((1, 4)):
            ttk.Radiobutton(year_start_frame, text=constants.YEAR_START_MONTH_FORMAT.format(month=month), value=month,
                            variable=self.year_start_month_var,
                            command=self._apply_period).grid(row=0, column=column, padx=5)

        ttk.Label(date_frame, text=constants.PERIOD_LABEL).grid(row=2, column=0, padx=5, pady=5)
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

    def _load_saved_document_type(self) -> DocumentType:
        try:
            return DocumentType(self.config.get('Analysis', 'document_type', fallback=DocumentType.ALL.value))
        except ValueError:
            return DocumentType.ALL

    def _setup_document_type_frame(self, parent: ttk.Frame) -> None:
        """集計対象の選択欄。選び直すとすぐに集計し直す"""
        document_type_frame = ttk.LabelFrame(parent, text=constants.DOCUMENT_TYPE_LABEL, padding="5")
        document_type_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.document_type_var = tk.StringVar(value=self.document_type.value)
        combobox = ttk.Combobox(document_type_frame, width=12, state='readonly', textvariable=self.document_type_var,
                                values=[document_type.value for document_type in DocumentType])
        combobox.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        combobox.bind('<<ComboboxSelected>>', self.change_document_type)

    def _setup_buttons(self, parent: ttk.Frame) -> None:
        ttk.Button(parent, text=constants.START_ANALYSIS_BUTTON_LABEL, command=self.start_analysis).grid(
            row=3, column=0, columnspan=2, pady=10)
        master_button = ttk.Button(parent, text=constants.MASTER_BUTTON_LABEL, command=self.open_master)
        master_button.grid(row=4, column=0, columnspan=2, pady=5)
        load_button = ttk.Button(parent, text=constants.EXCEL_IMPORT_BUTTON_LABEL, command=self.load_excel)
        load_button.grid(row=5, column=0, columnspan=2, pady=5)
        settings_button = ttk.Button(parent, text=constants.SETTINGS_BUTTON_LABEL, command=self.open_settings)
        settings_button.grid(row=6, column=0, columnspan=2, pady=5)
        if self.user is not None:
            ttk.Button(parent, text=constants.CHANGE_PASSWORD_BUTTON_LABEL, command=self.change_password).grid(
                row=7, column=0, columnspan=2, pady=5)
        ttk.Button(parent, text=constants.CLOSE_BUTTON_LABEL, command=self.root.quit).grid(
            row=8, column=0, columnspan=2, pady=5)
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
        self._import(self.analyzer.prepare_legacy_import, confirm=True)

    @staticmethod
    def _import_confirm_message(plan: ImportPlan) -> str:
        return constants.IMPORT_CONFIRM_MESSAGE.format(
            start_date=plan.start_date, end_date=plan.end_date, deleted=plan.deleted_count,
            inserted=len(plan.source.documents), errors=len(plan.source.errors))

    def _import(self, prepare: Callable[[], ImportPlan], confirm: bool,
                show_errors: bool = True) -> ImportPlan | None:
        """prepare で作った計画どおりに取り込み、計画を返す。やめたか失敗したら None。
        confirm なら取り込む前に件数を確かめ、取り込んだ後に DB の件数を知らせる。
        show_errors なら取り込めなかった行の一覧を出す"""
        try:
            plan = prepare()
            if confirm and not messagebox.askyesno(constants.CONFIRM_TITLE, self._import_confirm_message(plan)):
                return None
            count = self.analyzer.apply_import(plan)
        except NoDataError:
            messagebox.showerror(constants.ERROR_TITLE, constants.NO_SOURCE_DATA_MESSAGE)
            return None
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.LOAD_ERROR_MESSAGE.format(error=e))
            return None
        if confirm:
            messagebox.showinfo(constants.COMPLETE_TITLE, constants.LOAD_COMPLETE_MESSAGE.format(count=count))
        if show_errors and plan.source.errors:
            ImportErrorsDialog(self.root, plan.source.errors)
        return plan

    def start_analysis(self, show_errors: bool = True) -> None:
        try:
            start_date, end_date = read_date_range(self.start_date, self.end_date)
        except ValueError as e:
            messagebox.showerror(constants.ERROR_TITLE, str(e))
            return

        # 元データ(xlsm)は確認なしで取り込む
        plan = self._import(self.analyzer.prepare_source_import, confirm=False, show_errors=show_errors)
        if plan is None:
            return
        self.refresh_overview(plan.source.drafts)
        self._show_analysis(start_date, end_date, save_settings=True)

    def change_document_type(self, _event: tk.Event | None = None) -> None:
        """集計対象を選び直したら、元データは取り込まずに DB から推移と、集計済みなら全体・個人を集計し直す"""
        document_type = DocumentType(self.document_type_var.get())
        if document_type == self.document_type:
            return
        self.document_type = document_type
        update_config({'Analysis': {'document_type': document_type.value}})
        self.refresh_trends()
        if self.period is not None:
            self._show_analysis(*self.period)

    def _show_analysis(self, start_date: date, end_date: date, save_settings: bool = False) -> None:
        """期間内で集計対象の文書を DB から集計し、全体と個人を表示する。save_settings なら期間も保存する"""
        try:
            if save_settings:
                self._save_analysis_settings(start_date, end_date)
            summary = self.analyzer.overall_summary(start_date, end_date, self.document_type)
        except NoDataError:
            messagebox.showerror(constants.ERROR_TITLE, constants.NO_DATA_MESSAGE)
            return
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE,
                                 constants.ANALYSIS_ERROR_MESSAGE.format(error=e))
            return

        self.period = (start_date, end_date)
        self.result_view.show_overall(summary)
        self.result_view.set_staff_names(
            summary.staff_members if self.restricted_staff is None else [self.restricted_staff])
        self.show_staff_summary()

    def analyze_saved_period(self) -> None:
        """前回集計した期間が設定に残っていれば、起動時に「集計開始」と同じく取り込んでから集計する。
        起動のたびに出ると煩わしいため、取り込めなかった行の一覧は出さない"""
        if self.config.has_option('Analysis', 'start_date') and self.config.has_option('Analysis', 'end_date'):
            self.start_analysis(show_errors=False)

    def show_staff_summary(self) -> None:
        staff_name = self.result_view.staff_var.get()
        if self.period is None or not staff_name:
            return
        try:
            self.result_view.show_staff(self.analyzer.staff_summary(staff_name, *self.period,
                                                                    document_type=self.document_type))
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.ANALYSIS_ERROR_MESSAGE.format(error=e))

    def export_overall(self) -> None:
        if self.period is None:
            messagebox.showerror(constants.ERROR_TITLE, constants.NOT_ANALYZED_ERROR_MESSAGE)
            return
        try:
            open_folder_of(self.analyzer.run_analysis(*self.period, self.document_type))
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error=e))

    def export_staff(self) -> None:
        staff_name = self.result_view.staff_var.get()
        if self.period is None or not staff_name:
            messagebox.showerror(constants.ERROR_TITLE, constants.NOT_ANALYZED_ERROR_MESSAGE)
            return
        try:
            open_folder_of(self.analyzer.export_staff(staff_name, *self.period, self.document_type))
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error=e))

    def load_trend(self) -> None:
        """全体推移と個人推移を、前回選んだ範囲・診療科・氏名で表示する。職員は全体推移を見られない"""
        view = self.result_view
        if self.restricted_staff is None:
            self._load_trend_panel(view.trend, TREND_SECTION)
        self._load_trend_panel(view.staff_trend, STAFF_TREND_SECTION)

    def _load_trend_panel(self, panel: TrendPanel, section: str) -> None:
        """選択肢を作り、前回の選択で表示する。既定は前年(年度)の初日〜本日の合計と、先頭の担当者"""
        today = date.today()
        year_start_month = self.year_start_month_var.get()
        default_start = get_period_range(fiscal_year(today, year_start_month) - 1, '一年', year_start_month)[0]
        panel.start_date.set_date(date.fromisoformat(
            self.config.get(section, 'start_date', fallback=default_start.isoformat())))
        panel.end_date.set_date(date.fromisoformat(
            self.config.get(section, 'end_date', fallback=today.isoformat())))
        try:
            departments = [department.name for department in self.analyzer.list_departments()]
            if panel.has_staff:
                panel.set_staff_names(self._trend_staff_names(), self.config.get(section, 'staff_name', fallback=''))
        except Exception as e:
            panel.show_error(constants.TREND_ERROR_FORMAT.format(name=panel.name, error=e))
            return
        panel.set_departments([constants.TOTAL_ROW_LABEL, *departments],
                              self.config.get(section, 'department', fallback=constants.TOTAL_ROW_LABEL))
        self._refresh_trend_panel(panel, section)

    def _trend_staff_names(self) -> list[str]:
        """個人推移の氏名の選択肢。退職者の過去の推移も見られるよう、マスタの全担当者を並べる。職員は本人だけ"""
        if self.restricted_staff is not None:
            return [self.restricted_staff]
        return [staff.name for staff in self.analyzer.list_staff()]

    def _read_trend_request(self, panel: TrendPanel) -> TrendRequest | None:
        """推移タブの選択と範囲。個人推移で担当者がいなければ None。範囲が正しくなければ ValueError"""
        department = panel.selected_department()
        staff_name = panel.selected_staff()
        # selected_staff() は、全体推移なら None、個人推移で担当者がいなければ '' を返す
        if not department or staff_name == '':
            return None
        return TrendRequest(department, self.document_type, staff_name,
                            *read_date_range(panel.start_date, panel.end_date))

    def refresh_trends(self) -> None:
        """集計対象を変えたときに推移を集計し直す。職員は全体推移を見られない"""
        if self.restricted_staff is None:
            self.refresh_trend()
        self.refresh_staff_trend()

    def refresh_trend(self) -> None:
        self._refresh_trend_panel(self.result_view.trend, TREND_SECTION)

    def refresh_staff_trend(self) -> None:
        self._refresh_trend_panel(self.result_view.staff_trend, STAFF_TREND_SECTION)

    def _refresh_trend_panel(self, panel: TrendPanel, section: str) -> None:
        """推移を DB から集計し直し、選んだ範囲・診療科・氏名を保存する。選ぶたびに呼ぶため、失敗はタブ内に出す"""
        try:
            request = self._read_trend_request(panel)
        except ValueError:
            panel.show_error(constants.INVALID_TREND_RANGE_MESSAGE)
            return
        if request is None:
            return
        try:
            self._save_trend_settings(section, request)
            trend = self.analyzer.trend_summary(request.document_filter, request.start_date, request.end_date,
                                                self.year_start_month_var.get(), date.today(), request.staff_name)
            panel.show(trend, request.subject)
        except Exception as e:
            panel.show_error(constants.TREND_ERROR_FORMAT.format(name=panel.name, error=e))

    def export_trend(self) -> None:
        self._export_trend_panel(self.result_view.trend)

    def export_staff_trend(self) -> None:
        self._export_trend_panel(self.result_view.staff_trend)

    def _export_trend_panel(self, panel: TrendPanel) -> None:
        try:
            request = self._read_trend_request(panel)
        except ValueError:
            request = None
        if request is None:
            messagebox.showerror(constants.ERROR_TITLE, constants.INVALID_TREND_RANGE_MESSAGE)
            return
        try:
            open_folder_of(self.analyzer.export_trend(
                request.subject, request.document_filter, request.start_date, request.end_date,
                self.year_start_month_var.get(), date.today(), request.staff_name))
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error=e))

    @staticmethod
    def _save_trend_settings(section: str, request: TrendRequest) -> None:
        values = {'start_date': request.start_date.isoformat(), 'end_date': request.end_date.isoformat(),
                  'department': request.department}
        if request.staff_name is not None:
            values['staff_name'] = request.staff_name
        # 旧版の開始年・終了年は消す
        update_config({section: values}, replace=True)

    def _save_analysis_settings(self, start_date: date, end_date: date) -> None:
        update_config({'Analysis': {
            'start_date': start_date.isoformat(),
            'end_date': end_date.isoformat(),
            'year_start_month': str(self.year_start_month_var.get()),
            'document_type': self.document_type.value,
        }})

    def change_password(self) -> None:
        assert self.user is not None
        dialog = ChangePasswordDialog(self.root, self.analyzer, self.user)
        self.root.wait_window(dialog.dialog)
        if dialog.changed:
            messagebox.showinfo(constants.COMPLETE_TITLE, constants.PASSWORD_CHANGED_MESSAGE)

    def open_master(self) -> None:
        try:
            MasterDialog(self.root, self.analyzer)
        except Exception as e:
            messagebox.showerror(constants.ERROR_TITLE, constants.MASTER_OPEN_ERROR_MESSAGE.format(error=e))

    def open_settings(self) -> None:
        """保存したら、概要の日数や元データのパスの変更をすぐ反映する"""
        dialog = SettingsDialog(self.root)
        self.root.wait_window(dialog.dialog)
        if dialog.saved:
            self.refresh_overview()
