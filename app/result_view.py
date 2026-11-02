import tkinter as tk
from collections.abc import Callable, Sequence
from tkinter import ttk
from typing import cast

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from tkcalendar import DateEntry

from app import constants
from app.charts import overall_figure, staff_figure, trend_figure
from service.service_data_analyzer import OverallSummary, StaffCounts, TrendCounts
from service.service_excel_reader import Draft
from service.service_medical_docs_analyzer import DraftStatus, Overview, draft_status

OVERVIEW_COLUMN_WIDTHS = (90, 80, 240, 100, 90, 80, 90, 150)


def overview_row(draft: Draft) -> list[object]:
    requested = f'{draft.requested_date:%Y/%m/%d}' if draft.requested_date else ''
    return [f'{draft.received_date:%Y/%m/%d}', draft.patient_id, draft.document_name, draft.department,
            draft.doctor_name, draft.staff_name, requested, draft_status(draft)]


def overview_department_rows(overview: Overview) -> list[list[object]]:
    return [[department, counts[DraftStatus.NOT_STARTED], counts[DraftStatus.INCOMPLETE],
             counts[DraftStatus.UNASSIGNED], counts.total()]
            for department, counts in overview.department_status_counts]


def add_vertical_scrollbar(parent: ttk.Frame, tree: ttk.Treeview, row: int, column: int) -> None:
    scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.grid(row=row, column=column, sticky=tk.NS, pady=5)
    tree.configure(yscrollcommand=scrollbar.set)


def fill_tree(tree: ttk.Treeview, columns: list[str], rows: Sequence[Sequence[object]],
              first_width: int = 120) -> None:
    tree.delete(*tree.get_children())
    tree.configure(columns=columns)
    for index, column in enumerate(columns):
        tree.heading(column, text=column)
        # 見出しが欠けないよう、件数列は見出しの文字数に合わせた幅にする
        width = first_width if index == 0 else max(60, 16 * len(column) + 16)
        tree.column(column, width=width, minwidth=width, anchor=tk.W if index == 0 else tk.E)
    for row in rows:
        tree.insert('', tk.END, values=list(row))


class ChartArea:
    """matplotlib の図を差し替えて表示する枠"""

    def __init__(self, parent: ttk.Frame) -> None:
        self.frame = ttk.Frame(parent)
        self.canvas: FigureCanvasTkAgg | None = None

    def show(self, figure: Figure) -> None:
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
        self.canvas = FigureCanvasTkAgg(figure, master=self.frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def clear(self) -> None:
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None


class TrendPanel:
    """全体推移・個人推移タブの中身。with_staff なら先頭に氏名の選択欄を置く"""

    def __init__(self, tab: ttk.Frame, name: str, on_export: Callable[[], None], on_changed: Callable[[], None],
                 with_staff: bool) -> None:
        self.name = name
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)
        header = ttk.Frame(tab)
        header.grid(row=0, column=0, sticky=tk.EW)
        self.staff_var = tk.StringVar()
        self.staff_combobox: ttk.Combobox | None = None
        if with_staff:
            ttk.Label(header, text=constants.STAFF_NAME_HEADER).pack(side=tk.LEFT, padx=5)
            self.staff_combobox = self._create_combobox(header, 12, self.staff_var, on_changed)
        ttk.Label(header, text=constants.TREND_DEPARTMENT_LABEL).pack(side=tk.LEFT, padx=(20 if with_staff else 5, 5))
        self.department_var = tk.StringVar()
        self.department_combobox = self._create_combobox(header, 14, self.department_var, on_changed)

        ttk.Label(header, text=constants.TREND_RANGE_LABEL).pack(side=tk.LEFT, padx=(20, 5))
        self.start_date = self._create_date_entry(header, on_changed)
        ttk.Label(header, text=constants.TREND_RANGE_SEPARATOR).pack(side=tk.LEFT, padx=5)
        self.end_date = self._create_date_entry(header, on_changed)
        ttk.Button(header, text=constants.EXPORT_BUTTON_LABEL, command=on_export).pack(side=tk.RIGHT)

        self.title = ttk.Label(tab, text='')
        self.title.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        self.tree = ttk.Treeview(tab, show='headings', height=5)
        self.tree.grid(row=3, column=0, sticky=tk.EW, pady=5)
        self.chart = ChartArea(tab)
        self.chart.frame.grid(row=4, column=0, sticky=tk.NSEW)

    @staticmethod
    def _create_combobox(parent: ttk.Frame, width: int, variable: tk.StringVar,
                         on_changed: Callable[[], None]) -> ttk.Combobox:
        combobox = ttk.Combobox(parent, width=width, state='readonly', textvariable=variable)
        combobox.pack(side=tk.LEFT)
        combobox.bind('<<ComboboxSelected>>', lambda _event: on_changed())
        return combobox

    @staticmethod
    def _create_date_entry(parent: ttk.Frame, on_changed: Callable[[], None]) -> DateEntry:
        """カレンダーで選んだときと、手入力して Enter を押したときに表示し直す"""
        date_entry = DateEntry(parent, width=12, background='darkblue', foreground='white', borderwidth=2,
                               locale='ja_JP', date_pattern='yyyy/mm/dd')
        date_entry.pack(side=tk.LEFT)
        date_entry.bind('<<DateEntrySelected>>', lambda _event: on_changed())
        date_entry.bind('<Return>', lambda _event: on_changed())
        return date_entry

    @property
    def has_staff(self) -> bool:
        return self.staff_combobox is not None

    def set_departments(self, departments: list[str], selected: str) -> None:
        """診療科の選択肢を入れ、selected が無ければ先頭を選ぶ"""
        self.department_combobox.configure(values=departments)
        self.department_var.set(selected if selected in departments else departments[0])

    def selected_department(self) -> str:
        return self.department_var.get()

    def set_staff_names(self, names: list[str], selected: str) -> None:
        """氏名の選択肢を入れ、selected が無ければ先頭を選ぶ"""
        assert self.staff_combobox is not None
        self.staff_combobox.configure(values=names)
        self.staff_var.set(selected if selected in names else (names[0] if names else ''))

    def selected_staff(self) -> str | None:
        """選んだ氏名。氏名の選択欄が無い全体推移なら None"""
        return self.staff_var.get() if self.has_staff else None

    def show(self, trend: TrendCounts, subject: str) -> None:
        self.title.configure(text='')
        # 範囲外と本日より後の月は空欄にする
        fill_tree(self.tree, [constants.TREND_CORNER_HEADER, *trend.month_labels, constants.TOTAL_ROW_LABEL],
                  [['' if value is None else value for value in row] for row in trend.table_rows], first_width=80)
        self.chart.show(trend_figure(trend, subject))

    def show_error(self, message: str) -> None:
        self.title.configure(text=message)
        self.tree.delete(*self.tree.get_children())
        self.chart.clear()


class ResultView:
    """「概要」タブで下書きの一覧を、「全体」「個人」「全体推移」「個人推移」タブで集計結果を表と図にして見せる"""

    def __init__(self, parent: tk.Misc, on_export_overall: Callable[[], None],
                 on_export_staff: Callable[[], None], on_staff_selected: Callable[[], None],
                 on_export_trend: Callable[[], None], on_trend_changed: Callable[[], None],
                 on_export_staff_trend: Callable[[], None], on_staff_trend_changed: Callable[[], None]) -> None:
        # 職員としてログインしたときの本人の氏名。None なら全員分を見られる
        self.restricted_staff: str | None = None
        # 全体の表の担当者の行と合計行。並べ替えても合計行は最後に置く
        self.overall_staff_rows: list[list[str | int]] = []
        self.overall_total_row: list[str | int] = []
        # 並べ替えている列の番号と降順かどうか。None なら表示順のまま
        self.overall_sort: tuple[int, bool] | None = None
        self.notebook = ttk.Notebook(parent)
        self.overview_tab = ttk.Frame(self.notebook, padding="5")
        self.overall_tab = ttk.Frame(self.notebook, padding="5")
        self.staff_tab = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.overview_tab, text=constants.OVERVIEW_TAB_LABEL)
        self.notebook.add(self.overall_tab, text=constants.OVERALL_TAB_LABEL)
        self.trend_tab = ttk.Frame(self.notebook, padding="5")
        self.staff_trend_tab = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.trend_tab, text=constants.TREND_TAB_LABEL)
        self.notebook.add(self.staff_tab, text=constants.STAFF_SUMMARY_TAB_LABEL)
        self.notebook.add(self.staff_trend_tab, text=constants.STAFF_TREND_TAB_LABEL)
        self._setup_overview_tab()
        self._setup_overall_tab(on_export_overall)
        self._setup_staff_tab(on_export_staff, on_staff_selected)
        self.trend = TrendPanel(self.trend_tab, constants.TREND_TAB_LABEL, on_export_trend, on_trend_changed,
                                with_staff=False)
        self.staff_trend = TrendPanel(self.staff_trend_tab, constants.STAFF_TREND_TAB_LABEL, on_export_staff_trend,
                                      on_staff_trend_changed, with_staff=True)

    def _setup_overview_tab(self) -> None:
        tab = self.overview_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)
        self.overview_title = ttk.Label(tab, text='')
        self.overview_title.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        self.overview_department_tree = ttk.Treeview(tab, show='headings', height=6)
        fill_tree(self.overview_department_tree, list(constants.OVERVIEW_DEPARTMENT_COLUMNS), [])
        self.overview_department_tree.grid(row=1, column=0, sticky=tk.EW, pady=5)
        add_vertical_scrollbar(tab, self.overview_department_tree, row=1, column=1)
        self.overview_tree = ttk.Treeview(tab, show='headings', columns=constants.OVERVIEW_COLUMNS)
        for column, width in zip(constants.OVERVIEW_COLUMNS, OVERVIEW_COLUMN_WIDTHS):
            self.overview_tree.heading(column, text=column)
            self.overview_tree.column(column, width=width, minwidth=width, anchor=tk.W)
        self.overview_tree.grid(row=2, column=0, sticky=tk.NSEW, pady=5)
        add_vertical_scrollbar(tab, self.overview_tree, row=2, column=1)

    def _setup_overall_tab(self, on_export: Callable[[], None]) -> None:
        tab = self.overall_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)
        self.overall_title = ttk.Label(tab, text=constants.NOT_ANALYZED_MESSAGE)
        self.overall_title.grid(row=0, column=0, sticky=tk.W)
        ttk.Button(tab, text=constants.EXPORT_BUTTON_LABEL, command=on_export).grid(row=0, column=1, sticky=tk.E)
        self.overall_tree = ttk.Treeview(tab, show='headings', height=16)
        self.overall_tree.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=5)
        self.overall_chart = ChartArea(tab)
        self.overall_chart.frame.grid(row=2, column=0, columnspan=2, sticky=tk.NSEW)

    def _setup_staff_tab(self, on_export: Callable[[], None], on_staff_selected: Callable[[], None]) -> None:
        tab = self.staff_tab
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(2, weight=1)
        header = ttk.Frame(tab)
        header.grid(row=0, column=0, columnspan=3, sticky=tk.EW)
        ttk.Label(header, text=constants.STAFF_TAB_LABEL).pack(side=tk.LEFT, padx=5)
        self.staff_var = tk.StringVar()
        self.staff_combobox = ttk.Combobox(header, width=12, state='readonly', textvariable=self.staff_var)
        self.staff_combobox.pack(side=tk.LEFT)
        self.staff_combobox.bind('<<ComboboxSelected>>', lambda _event: on_staff_selected())
        self.staff_total_label = ttk.Label(header, text='')
        self.staff_total_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(header, text=constants.EXPORT_BUTTON_LABEL, command=on_export).pack(side=tk.RIGHT)

        self.department_tree = ttk.Treeview(tab, show='headings', height=10)
        self.department_tree.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, 5), pady=5)
        self.document_tree = ttk.Treeview(tab, show='headings', height=10)
        self.document_tree.grid(row=1, column=1, sticky=tk.NSEW, pady=5)
        # 10 行を表示し、それより多い文書名はスクロールで見る
        add_vertical_scrollbar(tab, self.document_tree, row=1, column=2)
        self.staff_chart = ChartArea(tab)
        self.staff_chart.frame.grid(row=2, column=0, columnspan=3, sticky=tk.NSEW)

    def show_overview(self, overview: Overview) -> None:
        status_counts = overview.status_counts
        self.overview_title.configure(text=constants.OVERVIEW_TITLE_FORMAT.format(
            start=overview.start_date, end=overview.end_date, total=len(overview.drafts),
            not_started=status_counts[DraftStatus.NOT_STARTED], incomplete=status_counts[DraftStatus.INCOMPLETE],
            unassigned=status_counts[DraftStatus.UNASSIGNED]))
        fill_tree(self.overview_department_tree, list(constants.OVERVIEW_DEPARTMENT_COLUMNS),
                  overview_department_rows(overview))
        self.overview_tree.delete(*self.overview_tree.get_children())
        for draft in overview.drafts:
            self.overview_tree.insert('', tk.END, values=overview_row(draft))

    def show_overview_error(self, error: Exception) -> None:
        self.overview_title.configure(text=constants.OVERVIEW_ERROR_FORMAT.format(error=error))
        self.overview_department_tree.delete(*self.overview_department_tree.get_children())
        self.overview_tree.delete(*self.overview_tree.get_children())

    def show_overall(self, summary: OverallSummary) -> None:
        counts = summary.counts
        self.overall_title.configure(text=constants.OVERALL_TITLE_FORMAT.format(
            first=counts.first_date, last=counts.last_date, total=counts.total))
        columns = [constants.STAFF_NAME_HEADER, constants.TOTAL_ROW_LABEL, *summary.departments]
        rows = summary.rows
        self.overall_staff_rows, self.overall_total_row = rows[:-1], rows[-1]
        self.overall_sort = None
        fill_tree(self.overall_tree, columns, rows)
        for index, column in enumerate(columns[1:], start=1):
            self.overall_tree.heading(column, command=lambda i=index: self.sort_overall(i))
        self.overall_chart.show(overall_figure(summary))

    def sort_overall(self, column_index: int) -> None:
        """件数の列で並べ替える。同じ列をもう一度押すと昇順と降順を切り替える。合計行はいつも最後"""
        descending = self.overall_sort != (column_index, True)
        self.overall_sort = (column_index, descending)
        rows = sorted(self.overall_staff_rows, key=lambda row: cast(int, row[column_index]), reverse=descending)
        self.overall_tree.delete(*self.overall_tree.get_children())
        for row in [*rows, self.overall_total_row]:
            self.overall_tree.insert('', tk.END, values=row)

    def restrict_to_staff(self, staff_name: str) -> None:
        """全体・全体推移タブを隠し、個人・個人推移タブを staff_name だけにする。概要タブは割り振りに使うため職員にも見せる"""
        self.restricted_staff = staff_name
        self.notebook.hide(self.overall_tab)
        self.notebook.hide(self.trend_tab)
        self.staff_combobox.state(['disabled'])
        self.set_staff_names([staff_name])
        assert self.staff_trend.staff_combobox is not None
        self.staff_trend.staff_combobox.state(['disabled'])

    def set_staff_names(self, names: list[str]) -> None:
        if self.restricted_staff is not None:
            names = [self.restricted_staff]
        self.staff_combobox.configure(values=names)
        if self.staff_var.get() not in names:
            self.staff_var.set(names[0] if names else '')

    def show_staff(self, staff_counts: StaffCounts) -> None:
        self.staff_total_label.configure(text=constants.STAFF_TOTAL_FORMAT.format(total=staff_counts.total))
        fill_tree(self.department_tree, list(constants.STAFF_DEPARTMENT_COLUMNS),
                  [list(row) for row in staff_counts.by_department])
        fill_tree(self.document_tree, list(constants.STAFF_DOCUMENT_COLUMNS),
                  [list(row) for row in staff_counts.by_document], first_width=360)
        self.staff_chart.show(staff_figure(staff_counts))
