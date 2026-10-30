from matplotlib import rcParams
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator

from service.service_data_analyzer import StaffCounts
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
    axes.set_xlabel('件数')
    if axes.get_legend_handles_labels()[0]:
        # 棒に重ならないよう、凡例はグラフの右外に置く
        axes.legend(loc='upper left', bbox_to_anchor=(1.01, 1), fontsize='small')
    return figure


def staff_figure(staff_counts: StaffCounts) -> Figure:
    """1 人の担当者の月別件数を棒グラフにする"""
    figure = Figure(figsize=FIGURE_SIZE, layout='constrained')
    axes = figure.add_subplot()
    months = [month for month, _ in staff_counts.by_month]
    counts = [count for _, count in staff_counts.by_month]
    axes.bar(months, counts)
    axes.set_title(f'{staff_counts.staff_name} 月別件数')
    axes.set_ylabel('件数')
    # 件数が 0 だけでも軸が負にならないよう、0 から始めて目盛りを整数にする
    axes.set_ylim(0, max(counts, default=0) * 1.1 or 1)
    axes.yaxis.set_major_locator(MaxNLocator(integer=True))
    axes.tick_params(axis='x', labelrotation=45)
    return figure
