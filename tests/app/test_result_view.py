import math
import tkinter as tk
from datetime import date
from tkinter import ttk
from typing import cast
from unittest.mock import MagicMock

from app.charts import overall_figure, staff_figure, trend_figure
from app import constants
from app.result_view import ResultView
from service.service_data_analyzer import DocumentCounts, OverallSummary, StaffCounts, TrendCounts
from service.service_excel_reader import Draft
from service.service_medical_docs_analyzer import Overview

SUMMARY = OverallSummary(
    DocumentCounts({('植田', '内科'): 2, ('植田', '小児科'): 1, ('花﨑', '外科'): 1},
                   {'内科': 2, '小児科': 1, '外科': 1}, 4, date(2025, 4, 1), date(2025, 4, 30)),
    ['植田', '花﨑'], ['内科', '外科'])
STAFF_COUNTS = StaffCounts('植田', 3, [('内科', 2), ('小児科', 1)], [('診断書', 3)], [('2025/04', 3)])
TREND = TrendCounts(4, [(2025, [1] * 12), (2026, [2, 3, *[None] * 10])])


def make_view(root: tk.Tk) -> ResultView:
    return ResultView(root, *(MagicMock() for _ in range(7)))


def tree_values(tree: ttk.Treeview) -> list[tuple[str, ...]]:
    return [tuple(tree.item(item, 'values')) for item in tree.get_children()]


def test_figures() -> None:
    assert len(overall_figure(SUMMARY).axes[0].patches) == 4
    assert len(staff_figure(STAFF_COUNTS).axes[0].lines) == 1
    trend_axes = trend_figure(TREND, '内科').axes[0]
    assert [line.get_label() for line in trend_axes.lines] == ['2025年度', '2026年度']


def test_trend_figure_does_not_draw_zero_months() -> None:
    trend = TrendCounts(1, [(2026, [0, 5, *[None] * 10])])

    line = trend_figure(trend, '内科').axes[0].lines[0]
    y_values = cast(list[float], line.get_ydata())

    assert math.isnan(y_values[0]) and y_values[1] == 5 and math.isnan(y_values[2])
    # 前後が 0 件でも見えるよう点の印を付ける
    assert line.get_marker() == 'o'


def test_trend_figure_always_shows_twelve_month_labels() -> None:
    # 後ろの月に値が無くても、4月〜3月のラベルをすべて出す
    trend = TrendCounts(4, [(2026, [5, 3, *[None] * 10])])

    axes = trend_figure(trend, '内科').axes[0]

    assert [label.get_text() for label in axes.get_xticklabels()] == trend.month_labels
    assert axes.get_xlim() == (-0.5, 11.5)


def test_tabs_are_overview_overall_staff_trend(root: tk.Tk) -> None:
    view = make_view(root)

    assert [view.notebook.tab(tab, 'text') for tab in view.notebook.tabs()] == [
        '概要', '全体', '全体推移', '個人', '個人推移']
    assert view.notebook.select() == str(view.overview_tab)


def test_show_overview(root: tk.Tk) -> None:
    view = make_view(root)
    overview = Overview(date(2026, 10, 25), date(2026, 10, 31), [
        Draft(date(2026, 10, 31), None, '診断書', '内科', '佐藤', '', '123'),
        Draft(date(2026, 10, 30), date(2026, 10, 31), '紹介状', '外科', '田中', '', '456'),
        Draft(date(2026, 10, 30), None, '意見書', '外科', '田中', '植田', '789'),
        Draft(date(2026, 10, 29), None, '紹介状', '外科', '田中', '', '012'),
    ])

    view.show_overview(overview)

    assert view.overview_title.cget('text') == (
        '預り日 2026/10/25〜2026/10/31　対象 4件(未着手 2件 / 未完了 1件 / 担当不明 1件)')
    assert tree_values(view.overview_department_tree) == [('外科', '1', '1', '1', '3'), ('内科', '1', '0', '0', '1')]
    assert tree_values(view.overview_tree) == [
        ('2026/10/31', '123', '診断書', '内科', '佐藤', '', '', '未着手'),
        ('2026/10/30', '456', '紹介状', '外科', '田中', '', '2026/10/31', '担当不明'),
        ('2026/10/30', '789', '意見書', '外科', '田中', '植田', '', '未完了'),
        ('2026/10/29', '012', '紹介状', '外科', '田中', '', '', '未着手'),
    ]


def test_show_overview_error_clears_rows(root: tk.Tk) -> None:
    view = make_view(root)
    view.show_overview(Overview(date(2026, 10, 25), date(2026, 10, 31),
                                [Draft(date(2026, 10, 31), None, '診断書', '内科', '佐藤', '植田', '1')]))

    view.show_overview_error(FileNotFoundError('missing'))

    assert view.overview_title.cget('text') == constants.OVERVIEW_ERROR_FORMAT.format(error='missing')
    assert view.overview_tree.get_children() == ()
    assert view.overview_department_tree.get_children() == ()


def test_show_overall_and_staff(root: tk.Tk) -> None:
    view = make_view(root)

    view.show_overall(SUMMARY)
    view.set_staff_names(SUMMARY.staff_members)
    view.show_staff(STAFF_COUNTS)

    assert view.overall_tree.cget('columns') == ('氏名', '合計', '内科', '外科')
    assert tree_values(view.overall_tree)[-1] == ('合計', '4', '2', '1')
    assert view.staff_var.get() == '植田'
    assert tree_values(view.department_tree) == [('内科', '2'), ('小児科', '1')]
    assert tree_values(view.document_tree) == [('診断書', '3')]
    assert view.staff_total_label.cget('text') == '合計 3件'


def test_sort_overall_toggles_order_and_keeps_total_last(root: tk.Tk) -> None:
    view = make_view(root)
    view.show_overall(SUMMARY)

    view.sort_overall(3)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['花﨑', '植田', '合計']

    view.sort_overall(3)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['植田', '花﨑', '合計']

    view.sort_overall(1)
    assert [row[0] for row in tree_values(view.overall_tree)] == ['植田', '花﨑', '合計']


def test_set_staff_names_keeps_selection(root: tk.Tk) -> None:
    view = make_view(root)
    view.staff_var.set('花﨑')

    view.set_staff_names(['植田', '花﨑'])

    assert view.staff_var.get() == '花﨑'


def test_set_trend_departments_selects_saved_department(root: tk.Tk) -> None:
    trend = make_view(root).trend

    trend.set_departments(['合計', '内科', '外科'], '外科')
    assert trend.selected_department() == '外科'

    trend.set_departments(['合計', '内科'], '外科')
    assert trend.department_combobox.cget('values') == ('合計', '内科')
    assert trend.selected_department() == '合計'


def test_staff_trend_selects_saved_staff(root: tk.Tk) -> None:
    view = make_view(root)

    # 全体推移には氏名の選択欄が無い
    assert (view.trend.has_staff, view.trend.selected_staff()) == (False, None)

    view.staff_trend.set_staff_names(['植田', '花﨑'], '花﨑')
    assert view.staff_trend.selected_staff() == '花﨑'

    view.staff_trend.set_staff_names(['植田'], '花﨑')
    assert view.staff_trend.selected_staff() == '植田'

    view.staff_trend.set_staff_names([], '花﨑')
    assert view.staff_trend.selected_staff() == ''


def test_restrict_to_staff_keeps_staff_trend_with_own_name(root: tk.Tk) -> None:
    view = make_view(root)

    view.restrict_to_staff('花﨑')

    assert view.notebook.tab(view.staff_trend_tab, 'state') == 'normal'
    assert view.staff_trend.staff_combobox is not None
    assert view.staff_trend.staff_combobox.instate(['disabled'])


def test_show_trend_rows_follow_years(root: tk.Tk) -> None:
    trend = make_view(root).trend

    trend.show(TREND, '内科')

    assert trend.title.cget('text') == ''
    assert (trend.tree.heading(0, 'text'), trend.tree.heading(1, 'text')) == ('件数', '4月')
    assert tree_values(trend.tree) == [
        ('2025年度', *['1'] * 12, '12'),
        ('2026年度', '2', '3', *[''] * 10, '5'),
    ]

    trend.show_error('エラー')

    assert trend.title.cget('text') == 'エラー'
    assert trend.tree.get_children() == ()
    assert trend.chart.canvas is None
