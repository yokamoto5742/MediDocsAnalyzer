import tkinter as tk
from collections.abc import Callable
from tkinter import ttk
from typing import cast

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from app import constants
from app.charts import overall_figure, staff_figure
from service.service_data_analyzer import StaffCounts
from service.service_excel_reader import Draft
from service.service_medical_docs_analyzer import OverallSummary, Overview

OVERVIEW_COLUMN_WIDTHS = (90, 80, 240, 100, 90, 80, 90, 150)


def overview_row(draft: Draft) -> list[object]:
    blanks = [label for label, blank in ((constants.STAFF_BLANK_LABEL, not draft.staff_name),
                                         (constants.REQUESTED_BLANK_LABEL, draft.requested_date is None)) if blank]
    requested = f'{draft.requested_date:%Y/%m/%d}' if draft.requested_date else ''
    return [f'{draft.received_date:%Y/%m/%d}', draft.patient_id, draft.document_name, draft.department,
            draft.doctor_name, draft.staff_name, requested, '・'.join(blanks)]


def overall_rows(summary: OverallSummary) -> list[list[object]]:
    """Excel 出力と同じ行。合計列は並べた診療科の合計、最終行の合計は一覧にない診療科も含めた全体の件数"""
    counts = summary.counts
    departments = summary.departments[1:]
    rows: list[list[object]] = []
    for staff in summary.staff_members:
        department_counts = [counts.by_staff_department.get((staff, d), 0) for d in departments]
        rows.append([staff, sum(department_counts), *department_counts])
    rows.append([constants.TOTAL_ROW_LABEL, counts.total, *(counts.by_department.get(d, 0) for d in departments)])
    return rows


def fill_tree(tree: ttk.Treeview, columns: list[str], rows: list[list[object]], first_width: int = 120) -> None:
    tree.delete(*tree.get_children())
    tree.configure(columns=columns)
    for index, column in enumerate(columns):
        tree.heading(column, text=column)
        # 見出しが欠けないよう、件数列は見出しの文字数に合わせた幅にする
        width = first_width if index == 0 else max(60, 16 * len(column) + 16)
        tree.column(column, width=width, minwidth=width, anchor=tk.W if index == 0 else tk.E)
    for row in rows:
        tree.insert('', tk.END, values=row)


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


class ResultView:
    """「概要」タブで下書きの一覧を、「全体」「個人」タブで集計結果を表と図にして見せる"""

    def __init__(self, parent: tk.Misc, on_export_overall: Callable[[], None],
                 on_export_staff: Callable[[], None], on_staff_selected: Callable[[], None]) -> None:
        # 職員としてログインしたときの本人の氏名。None なら全員分を見られる
        self.restricted_staff: str | None = None
        # 全体の表の担当者の行と合計行。並べ替えても合計行は最後に置く
        self.overall_staff_rows: list[list[object]] = []
        self.overall_total_row: list[object] = []
        # 並べ替えている列の番号と降順かどうか。None なら表示順のまま
        self.overall_sort: tuple[int, bool] | None = None
        self.notebook = ttk.Notebook(parent)
        self.overview_tab = ttk.Frame(self.notebook, padding="5")
        self.overall_tab = ttk.Frame(self.notebook, padding="5")
        self.staff_tab = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.overview_tab, text=constants.OVERVIEW_TAB_LABEL)
        self.notebook.add(self.overall_tab, text=constants.OVERALL_TAB_LABEL)
        self.notebook.add(self.staff_tab, text=constants.STAFF_SUMMARY_TAB_LABEL)
        self._setup_overview_tab()
        self._setup_overall_tab(on_export_overall)
        self._setup_staff_tab(on_export_staff, on_staff_selected)

    def _setup_overview_tab(self) -> None:
        tab = self.overview_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        self.overview_title = ttk.Label(tab, text='')
        self.overview_title.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        self.overview_tree = ttk.Treeview(tab, show='headings', columns=constants.OVERVIEW_COLUMNS)
        for column, width in zip(constants.OVERVIEW_COLUMNS, OVERVIEW_COLUMN_WIDTHS):
            self.overview_tree.heading(column, text=column)
            self.overview_tree.column(column, width=width, minwidth=width, anchor=tk.W)
        self.overview_tree.grid(row=1, column=0, sticky=tk.NSEW, pady=5)
        scrollbar = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=self.overview_tree.yview)
        scrollbar.grid(row=1, column=1, sticky=tk.NS, pady=5)
        self.overview_tree.configure(yscrollcommand=scrollbar.set)

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
        header.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
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
        self.staff_chart = ChartArea(tab)
        self.staff_chart.frame.grid(row=2, column=0, columnspan=2, sticky=tk.NSEW)

    def show_overview(self, overview: Overview) -> None:
        self.overview_title.configure(text=constants.OVERVIEW_TITLE_FORMAT.format(
            start=overview.start_date, end=overview.end_date, total=len(overview.drafts),
            unassigned=overview.unassigned_count, unrequested=overview.unrequested_count))
        self.overview_tree.delete(*self.overview_tree.get_children())
        for draft in overview.drafts:
            self.overview_tree.insert('', tk.END, values=overview_row(draft))

    def show_overview_error(self, error: Exception) -> None:
        self.overview_title.configure(text=constants.OVERVIEW_ERROR_FORMAT.format(error=error))
        self.overview_tree.delete(*self.overview_tree.get_children())

    def show_overall(self, summary: OverallSummary) -> None:
        counts = summary.counts
        self.overall_title.configure(text=constants.OVERALL_TITLE_FORMAT.format(
            first=counts.first_date, last=counts.last_date, total=counts.total))
        columns = [constants.STAFF_NAME_HEADER, *summary.departments]
        rows = overall_rows(summary)
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
        """全体タブを隠し、個人タブを staff_name だけにする。概要タブは割り振りに使うため職員にも見せる"""
        self.restricted_staff = staff_name
        self.notebook.hide(self.overall_tab)
        self.staff_combobox.state(['disabled'])
        self.set_staff_names([staff_name])

    def set_staff_names(self, names: list[str]) -> None:
        if self.restricted_staff is not None:
            names = [self.restricted_staff]
        self.staff_combobox.configure(values=names)
        if self.staff_var.get() not in names:
            self.staff_var.set(names[0] if names else '')

    def show_staff(self, staff_counts: StaffCounts) -> None:
        self.staff_total_label.configure(text=constants.STAFF_TOTAL_FORMAT.format(
            total=staff_counts.total, average=staff_counts.average))
        fill_tree(self.department_tree, list(constants.STAFF_DEPARTMENT_COLUMNS),
                  [list(row) for row in staff_counts.by_department])
        fill_tree(self.document_tree, list(constants.STAFF_DOCUMENT_COLUMNS),
                  [list(row) for row in staff_counts.by_document], first_width=360)
        self.staff_chart.show(staff_figure(staff_counts))
