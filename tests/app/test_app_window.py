import tkinter as tk
from collections.abc import Iterator
from datetime import date
from pathlib import Path
from typing import cast
from unittest.mock import MagicMock, patch

import pytest

from app import constants
from app.app_window import MedicalDocsAnalyzerGUI
from service.service_data_analyzer import DocumentCounts, StaffCounts
from service.service_excel_reader import Document, NoDataError, RowError, SourceData
from service.service_medical_docs_analyzer import ImportPlan, OverallSummary
from utils.config_manager import load_config


@pytest.fixture
def messagebox() -> Iterator[MagicMock]:
    with patch('app.app_window.messagebox') as mock:
        yield mock


@pytest.fixture
def gui(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> Iterator[MedicalDocsAnalyzerGUI]:
    with patch('app.app_window.MedicalDocsAnalyzer'):
        yield MedicalDocsAnalyzerGUI(root)


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
        gui = MedicalDocsAnalyzerGUI(root)

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


def make_plan(errors: list[RowError], legacy: SourceData | None = None) -> ImportPlan:
    source = SourceData(documents=[Document(date(2025, 4, 1), date(2025, 4, 2), '診断書', '内科', '佐藤', '植田')],
                        errors=errors, first_received_date=date(2025, 4, 1), last_received_date=date(2025, 4, 30))
    return ImportPlan(source, legacy, deleted_count=3)


def test_init_backs_up_database(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock) -> None:
    analyzer.backup_database.assert_called_once_with(date.today())


def test_init_shows_backup_error(config_path: Path, root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.app_window.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.backup_database.side_effect = OSError('offline')
        MedicalDocsAnalyzerGUI(root)

    messagebox.showerror.assert_called_once_with(
        constants.ERROR_TITLE, constants.BACKUP_ERROR_MESSAGE.format(error='offline'))


def test_load_data_shows_completion(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                    messagebox: MagicMock) -> None:
    analyzer.prepare_import.return_value = make_plan([])
    analyzer.apply_import.return_value = 12
    messagebox.askyesno.return_value = True

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.load_data()

    assert messagebox.askyesno.call_args.args[1] == constants.IMPORT_CONFIRM_MESSAGE.format(
        start_date=date(2025, 4, 1), end_date=date(2025, 4, 30), deleted=3, inserted=1, errors=0)
    messagebox.showinfo.assert_called_once_with(
        constants.COMPLETE_TITLE, constants.LOAD_COMPLETE_MESSAGE.format(count=12))
    dialog.assert_not_called()


def test_load_data_shows_error_rows(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                    messagebox: MagicMock) -> None:
    errors = [RowError('source.xlsx', 5, '預り日が空です', '')]
    analyzer.prepare_import.return_value = make_plan(errors, legacy=SourceData(documents=[]))
    messagebox.askyesno.return_value = True

    with patch('app.app_window.ImportErrorsDialog') as dialog:
        gui.load_data()

    assert messagebox.askyesno.call_args.args[1].startswith(constants.LEGACY_IMPORT_MESSAGE.format(count=0))
    dialog.assert_called_once_with(gui.root, errors)


def test_load_data_cancelled(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    analyzer.prepare_import.return_value = make_plan([])
    messagebox.askyesno.return_value = False

    gui.load_data()

    analyzer.apply_import.assert_not_called()
    messagebox.showinfo.assert_not_called()


@pytest.mark.parametrize(('error', 'expected_message'), [
    (NoDataError(), constants.NO_SOURCE_DATA_MESSAGE),
    (FileNotFoundError('missing'), constants.LOAD_ERROR_MESSAGE.format(error='missing')),
])
def test_load_data_shows_error(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock,
                               error: Exception, expected_message: str) -> None:
    analyzer.prepare_import.side_effect = error

    gui.load_data()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)


def make_summary() -> OverallSummary:
    counts = DocumentCounts({('植田', '内科'): 2, ('花﨑', '外科'): 1}, {'内科': 2, '外科': 1}, 3,
                            date(2025, 10, 1), date(2025, 12, 31))
    return OverallSummary(counts, ['植田', '花﨑'], ['合計', '内科', '外科'])


def make_staff_counts(name: str) -> StaffCounts:
    return StaffCounts(name, 2, 1.5, [('内科', 2)], [('診断書', 2)], [('2025/10', 2), ('2025/11', 0)])


def test_start_analysis_shows_results_and_saves_settings(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                                         messagebox: MagicMock) -> None:
    analyzer.overall_summary.return_value = make_summary()
    analyzer.staff_summary.return_value = make_staff_counts('植田')
    set_dates(gui, '2025/10/01', '2025/12/31')
    gui.year_start_month_var.set(1)

    gui.start_analysis()

    analyzer.overall_summary.assert_called_once_with(date(2025, 10, 1), date(2025, 12, 31))
    analyzer.staff_summary.assert_called_once_with('植田', date(2025, 10, 1), date(2025, 12, 31))
    assert gui.period == (date(2025, 10, 1), date(2025, 12, 31))
    assert len(gui.result_view.overall_tree.get_children()) == 3
    analysis_config = load_config()['Analysis']
    assert (analysis_config['start_date'], analysis_config['end_date']) == ('2025-10-01', '2025-12-31')
    assert analysis_config['year_start_month'] == '1'
    messagebox.showerror.assert_not_called()


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
    analyzer.overall_summary.side_effect = error

    gui.start_analysis()

    messagebox.showerror.assert_called_once_with(constants.ERROR_TITLE, expected_message)
    assert gui.period is None


def test_export_requires_analysis(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    gui.export_overall()
    gui.export_staff()

    assert messagebox.showerror.call_count == 2
    analyzer.run_analysis.assert_not_called()
    analyzer.export_staff.assert_not_called()


def test_export_uses_shown_period_and_staff(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock,
                                            messagebox: MagicMock) -> None:
    gui.period = (date(2025, 4, 1), date(2025, 4, 30))
    gui.result_view.staff_var.set('花﨑')

    gui.export_overall()
    gui.export_staff()

    analyzer.run_analysis.assert_called_once_with(date(2025, 4, 1), date(2025, 4, 30))
    analyzer.export_staff.assert_called_once_with('花﨑', date(2025, 4, 1), date(2025, 4, 30))


def test_export_shows_error(gui: MedicalDocsAnalyzerGUI, analyzer: MagicMock, messagebox: MagicMock) -> None:
    gui.period = (date(2025, 4, 1), date(2025, 4, 30))
    analyzer.run_analysis.side_effect = PermissionError('locked')

    gui.export_overall()

    messagebox.showerror.assert_called_once_with(
        constants.ERROR_TITLE, constants.EXPORT_ERROR_MESSAGE.format(error='locked'))


def test_open_config_opens_notepad(gui: MedicalDocsAnalyzerGUI) -> None:
    with patch('app.app_window.subprocess.Popen') as popen:
        gui.open_config()

    popen.assert_called_once()
    assert popen.call_args.args[0][0] == 'notepad.exe'


def test_open_config_shows_error(gui: MedicalDocsAnalyzerGUI, messagebox: MagicMock) -> None:
    with patch('app.app_window.subprocess.Popen', side_effect=OSError('denied')):
        gui.open_config()

    messagebox.showerror.assert_called_once_with(
        constants.ERROR_TITLE, constants.CONFIG_OPEN_ERROR_MESSAGE.format(error='denied'))
