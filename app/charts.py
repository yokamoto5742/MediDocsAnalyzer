import math

from matplotlib import rcParams
from matplotlib.axes import Axes
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator

from app import constants
from service.service_data_analyzer import StaffCounts, TrendCounts
from service.service_medical_docs_analyzer import OverallSummary

# 日本語を表示できるフォントを、Windows に入っている順に試す
rcParams['font.family'] = ['Meiryo', 'Yu Gothic', 'MS Gothic', 'sans-serif']

FIGURE_SIZE = (8, 3.6)


def overall_figure(summary: OverallSummary) -> Figure:
    """担当者ごとの件数を、診療科で色分けした横棒グラフにする"""
    figure = Figure(figsize=FIGURE_SIZE, layout='constrained')
    axes = figure.add_subplot()
    # 上から表と同じ順に並べるため、下から積む
    staff_members = summary.staff_members[::-1]
    lefts = [0] * len(staff_members)
    for department in summary.departments[1:]:
        widths = [summary.counts.by_staff_department.get((staff, department), 0) for staff in staff_members]
        if any(widths):
            axes.barh(staff_members, widths, left=lefts, label=department)
            lefts = [left + width for left, width in zip(lefts, widths)]
    axes.set_xlabel(constants.COUNT_AXIS_LABEL)
    if axes.get_legend_handles_labels()[0]:
        # 棒に重ならないよう、凡例はグラフの右外に置く
        axes.legend(loc='upper left', bbox_to_anchor=(1.01, 1), fontsize='small')
    return figure


def set_count_axis(axes: Axes, counts: list[int]) -> None:
    axes.set_ylabel(constants.COUNT_AXIS_LABEL)
    # 件数が 0 だけでも軸が負にならないよう、0 から始めて目盛りを整数にする
    axes.set_ylim(0, max(counts, default=0) * 1.1 or 1)
    axes.yaxis.set_major_locator(MaxNLocator(integer=True))


def staff_figure(staff_counts: StaffCounts) -> Figure:
    """1 人の担当者の月別件数を折れ線グラフにする"""
    figure = Figure(figsize=FIGURE_SIZE, layout='constrained')
    axes = figure.add_subplot()
    months = [month for month, _ in staff_counts.by_month]
    counts = [count for _, count in staff_counts.by_month]
    axes.plot(months, counts, marker='o')
    axes.set_title(constants.MONTHLY_COUNT_TITLE_FORMAT.format(subject=staff_counts.staff_name))
    set_count_axis(axes, counts)
    axes.tick_params(axis='x', labelrotation=45)
    return figure


def trend_figure(trend: TrendCounts, subject: str) -> Figure:
    """年(年度)ごとの月別件数を、1 年 1 本の折れ線グラフにする。範囲外・本日より後・0 件の月は線を引かない。
    前後が 0 件の月も見えるよう、点の印を付ける。subject は診療科(個人推移なら「氏名 診療科」)"""
    figure = Figure(figsize=FIGURE_SIZE, layout='constrained')
    axes = figure.add_subplot()
    month_positions = range(len(trend.month_labels))
    for year, counts in trend.by_year:
        axes.plot(month_positions, [count or math.nan for count in counts], marker='o', markersize=4,
                  label=trend.year_label(year))
    # 値の無い月も軸から外れないよう、12 か月の目盛りと範囲を固定する
    axes.set_xticks(month_positions, trend.month_labels)
    axes.set_xlim(-0.5, len(trend.month_labels) - 0.5)
    axes.set_title(constants.MONTHLY_COUNT_TITLE_FORMAT.format(subject=subject))
    set_count_axis(axes, [count for _, counts in trend.by_year for count in counts if count is not None])
    # 凡例は年を横に並べてグラフの下に置く
    axes.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), ncols=max(len(trend.by_year), 1),
                fontsize='small', frameon=False)
    return figure
