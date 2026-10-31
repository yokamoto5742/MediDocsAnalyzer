import tkinter as tk
from datetime import date
from tkinter import ttk
from unittest.mock import MagicMock

from app.charts import overall_figure, staff_figure
from app import constants
from app.result_view import ResultView, overall_rows
from service.service_data_analyzer import DocumentCounts, StaffCounts
from service.service_excel_reader import Draft
from service.service_medical_docs_analyzer import OverallSummary, Overview

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


def test_tabs_are_overview_overall_staff(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())

    assert [view.notebook.tab(tab, 'text') for tab in view.notebook.tabs()] == ['概要', '全体', '個人']
    assert view.notebook.select() == str(view.overview_tab)


def test_show_overview(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())
    overview = Overview(date(2026, 10, 25), date(2026, 10, 31), [
        Draft(date(2026, 10, 31), None, '診断書', '内科', '佐藤', '', '123'),
        Draft(date(2026, 10, 30), date(2026, 10, 31), '紹介状', '外科', '田中', '', '456'),
    ])

    view.show_overview(overview)

    assert view.overview_title.cget('text') == (
        '預り日 2026/10/25〜2026/10/31　対象 2件(担当者未定 2件 / 医師依頼日未入力 1件)')
    assert tree_values(view.overview_tree) == [
        ('2026/10/31', '123', '診断書', '内科', '佐藤', '', '', '担当者名・医師依頼日'),
        ('2026/10/30', '456', '紹介状', '外科', '田中', '', '2026/10/31', '担当者名'),
    ]


def test_show_overview_error_clears_rows(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())
    view.show_overview(Overview(date(2026, 10, 25), date(2026, 10, 31),
                                [Draft(date(2026, 10, 31), None, '診断書', '内科', '佐藤', '植田', '1')]))

    view.show_overview_error(FileNotFoundError('missing'))

    assert view.overview_title.cget('text') == constants.OVERVIEW_ERROR_FORMAT.format(error='missing')
    assert view.overview_tree.get_children() == ()


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


def test_sort_overall_toggles_order_and_keeps_total_last(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())
    view.show_overall(SUMMARY)

    view.sort_overall(3)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['花﨑', '植田', '合計']

    view.sort_overall(3)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['植田', '花﨑', '合計']

    view.sort_overall(1)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['植田', '花﨑', '合計']


def test_set_staff_names_keeps_selection(root: tk.Tk) -> None:
    view = ResultView(root, MagicMock(), MagicMock(), MagicMock())
    view.staff_var.set('花﨑')

    view.set_staff_names(['植田', '花﨑'])

    assert view.staff_var.get() == '花﨑'
