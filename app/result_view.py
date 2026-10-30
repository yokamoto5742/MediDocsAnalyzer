import tkinter as tk
from collections.abc import Callable
from tkinter import ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from app import constants
from app.charts import overall_figure, staff_figure
from service.service_data_analyzer import StaffCounts
from service.service_medical_docs_analyzer import OverallSummary


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
    """集計結果を「全体」「個人」のタブで表と図にして見せる"""

    def __init__(self, parent: tk.Misc, on_export_overall: Callable[[], None],
                 on_export_staff: Callable[[], None], on_staff_selected: Callable[[], None]) -> None:
        # 職員としてログインしたときの本人の氏名。None なら全員分を見られる
        self.restricted_staff: str | None = None
        self.notebook = ttk.Notebook(parent)
        self.overall_tab = ttk.Frame(self.notebook, padding="5")
        self.staff_tab = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.overall_tab, text=constants.OVERALL_TAB_LABEL)
        self.notebook.add(self.staff_tab, text=constants.STAFF_SUMMARY_TAB_LABEL)
        self._setup_overall_tab(on_export_overall)
        self._setup_staff_tab(on_export_staff, on_staff_selected)

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

    def show_overall(self, summary: OverallSummary) -> None:
        counts = summary.counts
        self.overall_title.configure(text=constants.OVERALL_TITLE_FORMAT.format(
            first=counts.first_date, last=counts.last_date, total=counts.total))
        fill_tree(self.overall_tree, [constants.STAFF_NAME_HEADER, *summary.departments], overall_rows(summary))
        self.overall_chart.show(overall_figure(summary))

    def restrict_to_staff(self, staff_name: str) -> None:
        """全体タブを隠し、個人タブを staff_name だけにする"""
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
