import tkinter as tk
from datetime import date
from tkinter import ttk
from unittest.mock import MagicMock

from app.charts import overall_figure, staff_figure
from app.result_view import ResultView, overall_rows
from service.service_data_analyzer import DocumentCounts, StaffCounts
from service.service_medical_docs_analyzer import OverallSummary

SUMMARY = OverallSummary(
    DocumentCounts({('植田', '内科'): 2, ('植田', '小児科'): 1, ('花﨑', '外科'): 1},
                   {'内科': 2, '小児科': 1, '外科': 1}, 4, date(2025, 4, 1), date(2025, 4, 30)),
    ['植田', '花﨑'], ['合計', '内科', '外科'])
STAFF_COUNTS = StaffCounts('植田', 3, 2.0, [('内科', 2), ('小児科', 1)], [('診断書', 3)], [('2025/04', 3)])


def tree_values(tree: ttk.Treeview) -> list[tuple[str, ...]]:
    return [tuple(tree.item(item, 'values')) for item in tree.get_children()]


def test_overall_rows_match_excel_layout() -> None:
    # 担当者の合計は並べた診療科だけ、最終行の合計は一覧にない診療科も含める
    assert overall_rows(SUMMARY) == [['植田', 2, 2, 0], ['花﨑', 1, 0, 1], ['合計', 4, 2, 1]]


def test_figures_draw_bars() -> None:
    assert len(overall_figure(SUMMARY).axes[0].patches) == 4
    assert len(staff_figure(STAFF_COUNTS).axes[0].patches) == 1


def test_show_overall_and_staff(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())

    view.show_overall(SUMMARY)
    view.set_staff_names(SUMMARY.staff_members)
    view.show_staff(STAFF_COUNTS)

    assert tree_values(view.overall_tree)[-1] == ('合計', '4', '2', '1')
    assert view.staff_var.get() == '植田'
    assert tree_values(view.department_tree) == [('内科', '2'), ('小児科', '1')]
    assert tree_values(view.document_tree) == [('診断書', '3')]
    assert '平均 2.0件' in view.staff_total_label.cget('text')


def test_set_staff_names_keeps_selection(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())
    view.staff_var.set('花﨑')

    view.set_staff_names(['植田', '花﨑'])

    assert view.staff_var.get() == '花﨑'
