import tkinter as tk
from collections.abc import Iterator
from datetime import date
from pathlib import Path
from typing import cast
from unittest.mock import MagicMock, patch

import pytest

from app import constants
from app.app_window import MedicalDocsAnalyzerGUI
from app.result_view import TrendPanel
from service.models import Document, Draft, NoDataError, RowError, SourceData
from service.service_data_analyzer import (DocumentCounts, DocumentFilter, DocumentType, OverallSummary, StaffCounts,
                                           TrendCounts)
from service.service_master import ADMIN_ROLE, Department, Staff
from service.service_medical_docs_analyzer import ImportPlan
from service.service_overview import Overview
from utils.config_manager import load_config, save_config

DRAFT = Draft(date(2025, 4, 1), None, '診断書', '内科', '佐藤', '', '1')
TREND = TrendCounts(4, [(2025, [1] * 12)])
STAFF_LIST = [Staff(1, '植田', 1), Staff(2, '花﨑', 2, active=False)]
ANALYZE_SAVED_PERIOD = MedicalDocsAnalyzerGUI.analyze_saved_period


@pytest.fixture(autouse=True)
def analyze_saved_period() -> Iterator[MagicMock]:
    """起動時の自動集計は専用のテストで確かめ、ほかのテストでは止める"""
    with patch.object(MedicalDocsAnalyzerGUI, 'analyze_saved_period') as mock:
        yield mock


def start_gui(root: tk.Tk, user: Staff | None = None) -> MedicalDocsAnalyzerGUI:
    """main.py と同じく、画面を作ってから起動時の処理をする"""
    gui = MedicalDocsAnalyzerGUI(root, user)
    gui.start()
    return gui


def make_overview(drafts: list[Draft]) -> Overview:
    return Overview(date(2025, 4, 1), date(2025, 4, 7), drafts)


@pytest.fixture
def messagebox() -> Iterator[MagicMock]:
    with patch('app.app_window.messagebox') as mock:
        yield mock


@pytest.fixture
def open_folder_of() -> Iterator[MagicMock]:
    with patch('app.app_window.open_folder_of') as mock:
        yield mock


@pytest.fixture
def gui(config_path: Path, root: tk.Tk, messagebox: MagicMock,
        open_folder_of: MagicMock) -> Iterator[MedicalDocsAnalyzerGUI]:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.load_overview.return_value = make_overview([])
        analyzer_class.return_value.list_departments.return_value = [Department('内科', 1), Department('外科', 2)]
        analyzer_class.return_value.list_staff.return_value = STAFF_LIST
        analyzer_class.return_value.trend_summary.return_value = TREND
        yield start_gui(root)


@pytest.fixture
def analyzer(gui: MedicalDocsAnalyzerGUI) -> MagicMock:
    return cast(MagicMock, gui.analyzer)


def set_dates(gui: MedicalDocsAnalyzerGUI, start: str, end: str) -> None:
    for entry, text in ((gui.start_date, start), (gui.end_date, end)):
        entry.delete(0, tk.END)
        entry.insert(0, text)


def test_init_restores_saved_settings(gui: MedicalDocsAnalyzerGUI) -> None:
    assert gui.start_date.get() == '2025/04/01'
    assert gui.end_date.get() == '2025/06/30'
    assert gui.year_start_month_var.get() == 4
    assert gui.year_var.get() == '2025'


def test_init_shows_fiscal_year_for_january_to_march(config_path: Path, root: tk.Tk) -> None:
    config = load_config()
    config['Analysis']['start_date'] = '2026-01-01'
    with open(config_path, 'w', encoding='utf-8') as f:
        config.write(f)

    with patch('app.app_window.MedicalDocsAnalyzer'):
        gui = start_gui(root)

    assert gui.year_var.get() == '2025'


@pytest.mark.parametrize(('year_start_month', 'expected'), [
    (4, ('2025/04/01', '2026/03/31')),
    (1, ('2025/01/01', '2025/12/31')),
])
def test_apply_period(gui: MedicalDocsAnalyzerGUI, year_start_month: int, expected: tuple[str, str]) -> None:
    gui.year_start_month_var.set(year_start_month)
    gui.period_var.set('一年')

    gui._apply_period()

    assert (gui.start_date.get(), gui.end_date.get()) == expected


def test_apply_period_ignores_invalid_year(gui: MedicalDocsAnalyzerGUI) -> None:
    gui.year_var.set('abc')
    gui.period_var.set('一年')

    gui._apply_period()

    assert gui.start_date.get() == '2025/04/01'


def make_plan(errors: list[RowError]) -> ImportPlan:
    source = SourceData(documents=[Document(date(2025, 4, 1), date(2025, 4, 2), '診断書', '内科', '佐藤', '植田')],
                        errors=errors, drafts=[DRAFT],
                        first_received_date=date(2025, 4, 1), last_received_date=date(2025, 4, 30))
    return ImportPlan(source, date(2025, 4, 1), date(2025, 4, 30), deleted_count=3)


def test_init_backs_up_database(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    analyzer.backup_database.assert_called_once_with(date.today())


def test_init_only_builds_window(config_path: Path, root: tk.Tk) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        MedicalDocsAnalyzerGUI(root)

    # 起動時の処理は start() で行う
    analyzer_class.return_value.backup_database.assert_not_called()
    analyzer_class.return_value.load_overview.assert_not_called()


def test_init_shows_backup_error(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.backup_database.side_effect = OSError('offline')
        start_gui(root)

    messagebox.showerror.assert_called_once_with(
        constants.ERROR_TITLE, constants.BACKUP_ERROR_MESSAGE.format(error='offline'))


def test_init_shows_overview(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    analyzer.load_overview.assert_called_once_with(date.today())
    assert gui.result_view.overview_title.cget('text').startswith('預り日 2025/04/01〜2025/04/07')


def test_init_shows_overview_error_without_dialog(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.load_overview.side_effect = FileNotFoundError('missing')
        gui = start_gui(root)

    assert gui.result_view.overview_title.cget('text') == constants.OVERVIEW_ERROR_FORMAT.format(error='missing')
    messagebox.showerror.assert_not_called()


def test_load_excel_shows_completion(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                     messagebox: MagicMock) -> None:
    analyzer.prepare_legacy_import.return_value = make_plan([])
    analyzer.apply_import.return_value = 12
    messagebox.askyesno.return_value = True

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.load_excel()

    assert messagebox.askyesno.call_args.args[1] == constants.IMPORT_CONFIRM_MESSAGE.format(
        start_date=date(2025, 4, 1), end_date=date(2025, 4, 30), deleted=3, inserted=1, errors=0)
    messagebox.showinfo.assert_called_once_with(
        constants.COMPLETE_TITLE, constants.LOAD_COMPLETE_MESSAGE.format(count=12))
    dialog.assert_not_called()


def test_load_excel_shows_error_rows(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                     messagebox: MagicMock) -> None:
    errors = [RowError('legacy.xlsx', 5, '123', '預り日が空です', '')]
    analyzer.prepare_legacy_import.return_value = make_plan(errors)
    messagebox.askyesno.return_value = True

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.load_excel()

    dialog.assert_called_once_with(gui.root, errors)


def test_start_analysis_without_error_rows_dialog(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                  messagebox: MagicMock) -> None:
    analyzer.prepare_source_import.return_value = make_plan([RowError('source.xlsm', 5, '123', '預り日が空です', '')])
    analyzer.overview.return_value = make_overview([])
    analyzer.overall_summary.return_value = make_summary()
    analyzer.staff_summary.return_value = make_staff_counts('植田')

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.start_analysis(show_errors=False)

    # 起動時はエラー行の一覧を出さずに集計まで進める
    dialog.assert_not_called()
    analyzer.overall_summary.assert_called_once()
    messagebox.showerror.assert_not_called()


def test_load_excel_cancelled(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    analyzer.prepare_legacy_import.return_value = make_plan([])
    messagebox.askyesno.return_value = False

    gui.load_excel()

    analyzer.apply_import.assert_not_called()
    messagebox.showinfo.assert_not_called()


@pytest.mark.parametrize(('error', 'expected_message'), [
    (NoDataError(), constants.NO_SOURCE_DATA_MESSAGE),
    (FileNotFoundError('missing'), constants.LOAD_ERROR_MESSAGE.format(error='missing')),
])
def test_load_excel_shows_error(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock,
                                error: Exception, expected_message: str) -> None:
    analyzer.prepare_legacy_import.side_effect = error

    gui.load_excel()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)


def make_summary() -> OverallSummary:
    counts = DocumentCounts({('植田', '内科'): 2, ('花﨑', '外科'): 1}, {'内科': 2, '外科': 1}, 3,
                            date(2025, 10, 1), date(2025, 12, 31))
    return OverallSummary(counts, ['植田', '花﨑'], ['内科', '外科'])


def make_staff_counts(name: str) -> StaffCounts:
    return StaffCounts(name, 2, [('内科', 2)], [('診断書', 2)], [('2025/10', 2), ('2025/11', 0)])


def test_start_analysis_shows_results_and_saves_settings(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                         messagebox: MagicMock) -> None:
    plan = make_plan([])
    analyzer.prepare_source_import.return_value = plan
    analyzer.overview.return_value = make_overview([DRAFT])
    analyzer.overall_summary.return_value = make_summary()
    analyzer.staff_summary.return_value = make_staff_counts('植田')
    set_dates(gui, '2025/10/01', '2025/12/31')
    gui.year_start_month_var.set(1)

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.start_analysis()

    analyzer.apply_import.assert_called_once_with(plan)
    analyzer.overview.assert_called_once_with([DRAFT], date.today())
    assert len(gui.result_view.overview_tree.get_children()) == 1
    messagebox.askyesno.assert_not_called()
    dialog.assert_not_called()
    analyzer.overall_summary.assert_called_once_with(date(2025, 10, 1), date(2025, 12, 31), DocumentType.ALL)
    analyzer.staff_summary.assert_called_once_with('植田', date(2025, 10, 1), date(2025, 12, 31),
                                                   document_type=DocumentType.ALL)
    assert gui.period == (date(2025, 10, 1), date(2025, 12, 31))
    assert len(gui.result_view.overall_tree.get_children()) == 3
    analysis_config = load_config()['Analysis']
    assert (analysis_config['start_date'], analysis_config['end_date']) == ('2025-10-01', '2025-12-31')
    assert analysis_config['year_start_month'] == '1'
    assert analysis_config['document_type'] == '全書類'
    messagebox.showerror.assert_not_called()


def test_init_document_type_defaults_to_all(gui: MedicalDocsAnalyzerGUI) -> None:
    assert gui.document_type == DocumentType.ALL
    assert gui.document_type_var.get() == '全書類'


@pytest.mark.parametrize(('saved', 'expected'), [('主治医意見書', DocumentType.OPINION), ('不明', DocumentType.ALL)])
def test_init_restores_saved_document_type(config_path: Path, root: tk.Tk, saved: str,
                                           expected: DocumentType) -> None:
    config = load_config()
    config['Analysis']['document_type'] = saved
    save_config(config)

    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.list_departments.return_value = [Department('内科', 1)]
        analyzer_class.return_value.list_staff.return_value = STAFF_LIST
        analyzer_class.return_value.trend_summary.return_value = TREND
        gui = start_gui(root)

    assert (gui.document_type, gui.document_type_var.get()) == (expected, expected.value)
    # 推移も保存した集計対象で絞る
    assert analyzer_class.return_value.trend_summary.call_args.args[0].document_type == expected


def test_start_analysis_applies_selected_document_type(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                       open_folder_of: MagicMock) -> None:
    analyzer.prepare_source_import.return_value = make_plan([])
    analyzer.overall_summary.return_value = make_summary()
    analyzer.staff_summary.return_value = make_staff_counts('植田')
    gui.document_type_var.set('医療文書')
    # 「集計開始」を押すまでは確定しない
    assert gui.document_type == DocumentType.ALL
    analyzer.trend_summary.reset_mock()

    gui.start_analysis()
    gui.export_overall()
    gui.export_staff()

    assert gui.document_type == DocumentType.MEDICAL
    analyzer.overall_summary.assert_called_once_with(date(2025, 4, 1), date(2025, 6, 30), DocumentType.MEDICAL)
    analyzer.run_analysis.assert_called_once_with(date(2025, 4, 1), date(2025, 6, 30), DocumentType.MEDICAL)
    analyzer.export_staff.assert_called_once_with('植田', date(2025, 4, 1), date(2025, 6, 30), DocumentType.MEDICAL)
    # 全体推移と個人推移も集計し直す
    assert [call.args[0] for call in analyzer.trend_summary.call_args_list] == [
        DocumentFilter(document_type=DocumentType.MEDICAL)] * 2
    assert load_config()['Analysis']['document_type'] == '医療文書'


@pytest.mark.parametrize(('start', 'end', 'expected_message'), [
    ('2025-10-01', '2025/12/31', constants.INVALID_DATE_MESSAGE),
    ('2025/12/31', '2025/10/01', constants.START_AFTER_END_MESSAGE),
])
def test_start_analysis_rejects_invalid_dates(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock,
                                              start: str, end: str, expected_message: str) -> None:
    set_dates(gui, start, end)

    gui.start_analysis()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)
    analyzer.overall_summary.assert_not_called()


@pytest.mark.parametrize(('error', 'expected_message'), [
    (NoDataError(), constants.NO_DATA_MESSAGE),
    (PermissionError('locked'), constants.ANALYSIS_ERROR_MESSAGE.format(error='locked')),
])
def test_start_analysis_shows_error(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock,
                                    error: Exception, expected_message: str) -> None:
    analyzer.prepare_source_import.return_value = make_plan([])
    analyzer.overall_summary.side_effect = error

    gui.start_analysis()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)
    assert gui.period is None


def test_start_analysis_shows_import_error_rows_then_analyzes(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    errors = [RowError('source.xlsm', 5, '123', '預り日が空です', '')]
    analyzer.prepare_source_import.return_value = make_plan(errors)
    analyzer.overall_summary.return_value = make_summary()
    analyzer.staff_summary.return_value = make_staff_counts('植田')

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.start_analysis()

    dialog.assert_called_once_with(gui.root, errors)
    analyzer.overall_summary.assert_called_once()


@pytest.mark.parametrize(('error', 'expected_message'), [
    (NoDataError(), constants.NO_SOURCE_DATA_MESSAGE),
    (FileNotFoundError('missing'), constants.LOAD_ERROR_MESSAGE.format(error='missing')),
])
def test_start_analysis_stops_when_import_fails(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                messagebox: MagicMock, error: Exception,
                                                expected_message: str) -> None:
    analyzer.prepare_source_import.side_effect = error

    gui.start_analysis()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)
    analyzer.overall_summary.assert_not_called()


def test_export_requires_analysis(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    gui.export_overall()
    gui.export_staff()

    assert messagebox.showerror.call_count == 2
    analyzer.run_analysis.assert_not_called()
    analyzer.export_staff.assert_not_called()


def test_export_uses_shown_period_and_staff(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                            messagebox: MagicMock, open_folder_of: MagicMock) -> None:
    gui.period = (date(2025, 4, 1), date(2025, 4, 30))
    gui.result_view.staff_var.set('花﨑')

    gui.export_overall()
    gui.export_staff()

    analyzer.run_analysis.assert_called_once_with(date(2025, 4, 1), date(2025, 4, 30), DocumentType.ALL)
    analyzer.export_staff.assert_called_once_with('花﨑', date(2025, 4, 1), date(2025, 4, 30), DocumentType.ALL)
    # 保存したファイルのフォルダを開く
    assert open_folder_of.call_args_list == [((analyzer.run_analysis.return_value,),),
                                             ((analyzer.export_staff.return_value,),)]
    messagebox.showerror.assert_not_called()


def test_export_shows_error(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    gui.period = (date(2025, 4, 1), date(2025, 4, 30))
    analyzer.run_analysis.side_effect = PermissionError('locked')

    gui.export_overall()

    messagebox.showerror.assert_called_once_with(
        constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error='locked'))


def test_staff_user_sees_only_own_summary(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        gui = start_gui(root, Staff(2, '花﨑', 2, emr_id='S01'))
        analyzer = analyzer_class.return_value
        analyzer.prepare_source_import.return_value = make_plan([])
        analyzer.overall_summary.return_value = make_summary()
        analyzer.staff_summary.return_value = make_staff_counts('花﨑')

        gui.start_analysis()

    assert gui.restricted_staff == '花﨑'
    assert gui.root.title().endswith('花﨑')
    assert all(button.instate(['disabled']) for button in gui.admin_buttons)
    assert gui.result_view.notebook.tab(gui.result_view.overall_tab, 'state') == 'hidden'
    assert gui.result_view.notebook.tab(gui.result_view.overview_tab, 'state') == 'normal'
    # 集計結果に他の担当者がいても、本人だけを選べる
    assert gui.result_view.staff_combobox.cget('values') == ('花﨑',)
    analyzer.staff_summary.assert_called_once_with('花﨑', date(2025, 4, 1), date(2025, 6, 30),
                                                   document_type=DocumentType.ALL)


def test_admin_user_can_use_everything(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer'):
        gui = start_gui(root, Staff(1, '植田', 1, emr_id='A01', role=ADMIN_ROLE))

    assert gui.restricted_staff is None
    assert not any(button.instate(['disabled']) for button in gui.admin_buttons)
    assert gui.result_view.notebook.tab(gui.result_view.overall_tab, 'state') == 'normal'


@pytest.mark.parametrize('saved', [True, False])
def test_open_settings_refreshes_overview_when_saved(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                     saved: bool) -> None:
    analyzer.load_overview.reset_mock()

    with patch('app.app_window.SettingsDialog') as dialog, patch.object(gui.root, 'wait_window') as wait_window:
        dialog.return_value.saved = saved
        gui.open_settings()

    dialog.assert_called_once_with(gui.root)
    wait_window.assert_called_once_with(dialog.return_value.dialog)
    assert analyzer.load_overview.call_count == (1 if saved else 0)


def test_init_analyzes_saved_period(gui: MedicalDocsAnalyzerGUI, analyze_saved_period: MagicMock) -> None:
    analyze_saved_period.assert_called_once_with()


def test_analyze_saved_period_runs_analysis(gui: MedicalDocsAnalyzerGUI) -> None:
    with patch.object(gui, 'start_analysis') as start_analysis:
        ANALYZE_SAVED_PERIOD(gui)

    start_analysis.assert_called_once_with(show_errors=False)


def test_analyze_saved_period_skips_without_saved_dates(gui: MedicalDocsAnalyzerGUI) -> None:
    gui.config.remove_option('Analysis', 'start_date')

    with patch.object(gui, 'start_analysis') as start_analysis:
        ANALYZE_SAVED_PERIOD(gui)

    start_analysis.assert_not_called()


def set_trend_range(gui: MedicalDocsAnalyzerGUI, start: str, end: str, panel: TrendPanel | None = None) -> None:
    panel = panel or gui.result_view.trend
    for date_entry, text in ((panel.start_date, start), (panel.end_date, end)):
        date_entry.delete(0, tk.END)
        date_entry.insert(0, text)


def test_init_shows_trend_from_last_year_to_today(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    today = date.today()
    last_year_start = date(today.year - (today.month < 4) - 1, 4, 1)
    trend = gui.result_view.trend

    assert (trend.start_date.get_date(), trend.end_date.get_date()) == (last_year_start, today)
    assert trend.department_combobox.cget('values') == ('合計', '内科', '外科')
    analyzer.trend_summary.assert_any_call(DocumentFilter(), last_year_start, today, 4, today, None)
    assert len(trend.tree.get_children()) == 1
    assert load_config()['Trend']['department'] == '合計'


def test_init_shows_staff_trend_for_first_staff(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    today = date.today()
    last_year_start = date(today.year - (today.month < 4) - 1, 4, 1)
    staff_trend = gui.result_view.staff_trend

    # 退職者も選べる
    assert staff_trend.staff_combobox is not None
    assert staff_trend.staff_combobox.cget('values') == ('植田', '花﨑')
    analyzer.trend_summary.assert_any_call(DocumentFilter(), last_year_start, today, 4, today, '植田')
    assert len(staff_trend.tree.get_children()) == 1
    assert dict(load_config()['StaffTrend']) == {
        'start_date': last_year_start.isoformat(), 'end_date': today.isoformat(), 'department': '合計',
        'staff_name': '植田'}


def test_refresh_staff_trend_restores_and_saves_selection(config_path: Path, root: tk.Tk, messagebox: MagicMock,
                                                        open_folder_of: MagicMock) -> None:
    config = load_config()
    config['StaffTrend'] = {'start_date': '2024-01-01', 'end_date': '2026-10-01', 'department': '外科',
                            'staff_name': '花﨑'}
    save_config(config)
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer = analyzer_class.return_value
        analyzer.list_departments.return_value = [Department('内科', 1), Department('外科', 2)]
        analyzer.list_staff.return_value = STAFF_LIST
        analyzer.trend_summary.return_value = TREND
        gui = start_gui(root)

    staff_trend = gui.result_view.staff_trend
    assert (staff_trend.selected_staff(), staff_trend.selected_department()) == ('花﨑', '外科')
    analyzer.trend_summary.assert_any_call(DocumentFilter('外科'), date(2024, 1, 1), date(2026, 10, 1), 4, date.today(), '花﨑')

    staff_trend.staff_var.set('植田')
    gui.refresh_staff_trend()
    gui.export_staff_trend()

    analyzer.trend_summary.assert_called_with(DocumentFilter('外科'), date(2024, 1, 1), date(2026, 10, 1), 4, date.today(), '植田')
    analyzer.export_trend.assert_called_once_with('植田 全書類 外科', DocumentFilter('外科'), date(2024, 1, 1), date(2026, 10, 1), 4,
                                                  date.today(), '植田')
    assert load_config()['StaffTrend']['staff_name'] == '植田'
    open_folder_of.assert_called_once_with(analyzer.export_trend.return_value)
    messagebox.showerror.assert_not_called()


def test_init_restores_saved_trend_settings(config_path: Path, root: tk.Tk) -> None:
    config = load_config()
    config['Trend'] = {'start_date': '2024-01-01', 'end_date': '2026-10-01', 'department': '外科'}
    save_config(config)

    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.list_departments.return_value = [Department('内科', 1), Department('外科', 2)]
        analyzer_class.return_value.trend_summary.return_value = TREND
        gui = start_gui(root)

    assert gui.result_view.trend.selected_department() == '外科'
    analyzer_class.return_value.trend_summary.assert_called_once_with(
        DocumentFilter('外科'), date(2024, 1, 1), date(2026, 10, 1), 4, date.today(), None)


def test_refresh_trend_saves_range_and_department(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    config = load_config()
    config['Trend'] = {'start_year': '2022', 'end_year': '2024', 'department': '外科'}
    save_config(config)
    set_trend_range(gui, '2021/01/01', '2023/06/30')
    gui.result_view.trend.department_var.set('内科')

    gui.refresh_trend()

    analyzer.trend_summary.assert_called_with(DocumentFilter('内科'), date(2021, 1, 1), date(2023, 6, 30), 4, date.today(), None)
    # 旧版の開始年・終了年は残さない
    assert dict(load_config()['Trend']) == {'start_date': '2021-01-01', 'end_date': '2023-06-30',
                                            'department': '内科'}


@pytest.mark.parametrize(('start', 'end'), [('2026/01/01', '2025/12/31'), ('abc', '2025/12/31')])
def test_refresh_trend_rejects_invalid_range(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                            messagebox: MagicMock, start: str, end: str) -> None:
    analyzer.trend_summary.reset_mock()
    set_trend_range(gui, start, end)

    gui.refresh_trend()
    gui.export_trend()

    analyzer.trend_summary.assert_not_called()
    assert gui.result_view.trend.title.cget('text') == constants.INVALID_TREND_RANGE_MESSAGE
    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, constants.INVALID_TREND_RANGE_MESSAGE)
    analyzer.export_trend.assert_not_called()


def test_refresh_trend_shows_error_in_tab(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                          messagebox: MagicMock) -> None:
    analyzer.trend_summary.side_effect = PermissionError('locked')

    gui.refresh_trend()

    assert gui.result_view.trend.title.cget('text') == constants.TREND_ERROR_FORMAT.format(name='全体推移', error='locked')
    messagebox.showerror.assert_not_called()


def test_export_trend_uses_selected_department(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    set_trend_range(gui, '2024/01/01', '2025/12/31')
    gui.result_view.trend.department_var.set('外科')

    gui.export_trend()

    analyzer.export_trend.assert_called_once_with('全書類 外科', DocumentFilter('外科'), date(2024, 1, 1),
                                                  date(2025, 12, 31), 4, date.today(), None)


def test_export_trend_passes_total_as_all_departments(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    set_trend_range(gui, '2024/01/01', '2025/12/31')
    gui.result_view.trend.department_var.set(constants.TOTAL_ROW_LABEL)

    gui.export_trend()

    analyzer.export_trend.assert_called_once_with('全書類', DocumentFilter(), date(2024, 1, 1),
                                                  date(2025, 12, 31), 4, date.today(), None)


def test_export_trend_uses_applied_document_type(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    set_trend_range(gui, '2024/01/01', '2025/12/31')
    gui.result_view.trend.department_var.set(constants.TOTAL_ROW_LABEL)
    gui.document_type = DocumentType.OPINION

    gui.export_trend()

    analyzer.export_trend.assert_called_once_with('主治医意見書', DocumentFilter(document_type=DocumentType.OPINION),
                                                  date(2024, 1, 1), date(2025, 12, 31), 4, date.today(), None)


def test_staff_user_sees_only_own_staff_trend(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.list_departments.return_value = [Department('内科', 1)]
        analyzer_class.return_value.trend_summary.return_value = TREND
        gui = start_gui(root, Staff(2, '花﨑', 2, emr_id='S01'))

    view = gui.result_view
    assert view.notebook.tab(view.trend_tab, 'state') == 'hidden'
    assert view.staff_trend.staff_combobox is not None
    assert view.staff_trend.staff_combobox.cget('values') == ('花﨑',)
    # 全体推移は集計せず、本人の個人推移だけを集計する
    analyzer_class.return_value.list_staff.assert_not_called()
    analyzer_class.return_value.trend_summary.assert_called_once()
    assert analyzer_class.return_value.trend_summary.call_args.args[-1] == '花﨑'
