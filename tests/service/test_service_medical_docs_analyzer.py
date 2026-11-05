import sqlite3
from collections.abc import Sequence
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from openpyxl import load_workbook

from service.models import NoDataError
from service.service_data_analyzer import DocumentFilter, DocumentType
from service.service_master import Department
from service.service_medical_docs_analyzer import TOP_DOCUMENT_COUNT, MedicalDocsAnalyzer, split_list
from tests.documents_database import Document, make_documents_database
from utils.config_manager import load_config, save_config


def make_documents(config_path: Path, documents: Sequence[Document]) -> None:
    """設定した文書の DB(CSV2XL の DB)を作る"""
    make_documents_database(config_path.parent / 'documents.db', documents)


def document(received: date, requested: date | None, department: str, staff: str | None,
             document_name: str = '診断書', patient_id: int = 1) -> Document:
    return Document(received, requested, document_name, department, '佐藤', staff, patient_id)


def test_split_list() -> None:
    assert split_list(' 植田, 花﨑 ,,') == ['植田', '花﨑']


def test_load_overview_reads_documents_database_with_configured_days(config_path: Path) -> None:
    make_documents(config_path, [
        document(date(2026, 10, 27), None, '内科', '植田', patient_id=1),
        document(date(2026, 10, 30), date(2026, 10, 31), '内科', None, patient_id=2),
        document(date(2026, 10, 30), date(2026, 10, 31), '内科', '植田', patient_id=3),
    ])
    config = load_config()
    config['Overview'] = {'pending_days': '3'}
    save_config(config)

    overview = MedicalDocsAnalyzer().load_overview(date(2026, 10, 31))

    assert overview.start_date == date(2026, 10, 29)
    assert [draft.patient_id for draft in overview.drafts] == ['2']


def test_load_overview_uses_seven_days_by_default(config_path: Path) -> None:
    make_documents(config_path, [document(date(2026, 10, 25), None, '内科', '植田')])

    overview = MedicalDocsAnalyzer().load_overview(date(2026, 10, 31))

    assert [draft.patient_id for draft in overview.drafts] == ['1']


def test_load_overview_raises_when_documents_database_is_missing(config_path: Path) -> None:
    with pytest.raises(sqlite3.OperationalError):
        MedicalDocsAnalyzer().load_overview(date(2026, 10, 31))


def test_run_analysis(config_path: Path) -> None:
    make_documents(config_path, [
        document(date(2025, 3, 31), date(2025, 4, 1), '内科', '植田'),
        document(date(2025, 4, 1), date(2025, 4, 2), '内科', '植田'),
        document(date(2025, 5, 10), date(2025, 5, 11), '外科', '花﨑', '紹介状'),
        # 下書きは数えない
        document(date(2025, 5, 20), None, '外科', '花﨑'),
    ])

    output_path = MedicalDocsAnalyzer().run_analysis(date(2025, 4, 1), date(2025, 6, 30))

    assert Path(output_path).name == '医療文書作成件数_全書類_20250401-20250510.xlsx'
    sheet = load_workbook(output_path).active
    assert sheet is not None
    assert [list(row) for row in sheet.iter_rows(min_row=2, values_only=True)] == [
        ['氏名', '合計', '内科', '外科'],
        ['植田', 1, 1, 0],
        ['沖野', 0, 0, 0],
        ['花﨑', 1, 0, 1],
        ['合計', 2, 1, 1],
    ]


def test_run_analysis_uses_masters(config_path: Path) -> None:
    make_documents(config_path, [document(date(2025, 4, 1), date(2025, 4, 2), '眼科', '植田')])
    analyzer = MedicalDocsAnalyzer()
    ueda, okino, _ = analyzer.list_staff()
    ueda.active = False
    okino.active = False
    analyzer.save_staff(ueda)
    analyzer.save_staff(okino)
    analyzer.save_department(None, Department('眼科', 0))

    output_path = analyzer.run_analysis(date(2025, 4, 1), date(2025, 4, 30))

    sheet = load_workbook(output_path).active
    assert sheet is not None
    assert [list(row) for row in sheet.iter_rows(min_row=2, values_only=True)] == [
        ['氏名', '合計', '眼科', '内科', '外科'],
        ['植田', 1, 1, 0, 0],
        ['花﨑', 0, 0, 0, 0],
        ['合計', 1, 1, 0, 0],
    ]


def test_export_staff(config_path: Path) -> None:
    make_documents(config_path, [document(date(2025, 4, 1), date(2025, 4, 2), '内科', '花﨑')])

    output_path = MedicalDocsAnalyzer().export_staff('花﨑', date(2025, 4, 1), date(2025, 4, 30))

    assert Path(output_path) == config_path.parent / 'output' / '医療文書作成件数_花﨑_全書類_20250401-20250430.xlsx'
    assert load_workbook(output_path)['個人']['B2'].value == 1


def test_export_staff_limits_documents_to_top(config_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    make_documents(config_path, [])
    analyzer = MedicalDocsAnalyzer()
    staff_summary = MagicMock(wraps=analyzer.staff_summary)
    monkeypatch.setattr(analyzer, 'staff_summary', staff_summary)
    monkeypatch.setattr('service.service_medical_docs_analyzer.save_staff_results', MagicMock(return_value='x'))

    analyzer.export_staff('花﨑', date(2025, 4, 1), date(2025, 4, 30), DocumentType.OPINION)

    staff_summary.assert_called_once_with('花﨑', date(2025, 4, 1), date(2025, 4, 30), TOP_DOCUMENT_COUNT,
                                          DocumentType.OPINION)


def test_trend_summary_and_export(config_path: Path) -> None:
    make_documents(config_path, [
        document(date(2025, 4, 1), date(2025, 4, 2), '内科', '花﨑'),
        document(date(2025, 4, 1), date(2025, 5, 2), '外科', '花﨑'),
    ])
    analyzer = MedicalDocsAnalyzer()

    start_date, end_date = date(2025, 4, 1), date(2026, 3, 31)
    total = analyzer.trend_summary(DocumentFilter(), start_date, end_date, 4, date(2026, 11, 1))
    surgery = analyzer.trend_summary(DocumentFilter('外科'), start_date, end_date, 4, date(2026, 11, 1))
    output_path = analyzer.export_trend('外科', DocumentFilter('外科'), start_date, end_date, 4, date(2026, 11, 1))

    assert total.by_year == [(2025, [1, 1, *[0] * 10])]
    assert surgery.by_year == [(2025, [0, 1, *[0] * 10])]
    assert Path(output_path) == config_path.parent / 'output' / '医療文書作成件数推移_外科_20250401-20260331.xlsx'


def test_staff_trend_summary_and_export(config_path: Path) -> None:
    make_documents(config_path, [
        document(date(2025, 4, 1), date(2025, 4, 2), '内科', '花﨑'),
        document(date(2025, 4, 1), date(2025, 5, 2), '内科', '植田'),
    ])
    analyzer = MedicalDocsAnalyzer()

    start_date, end_date = date(2025, 4, 1), date(2026, 3, 31)
    trend = analyzer.trend_summary(DocumentFilter(), start_date, end_date, 4, date(2026, 11, 1), '花﨑')
    output_path = analyzer.export_trend('花﨑 内科', DocumentFilter('内科'), start_date, end_date, 4, date(2026, 11, 1), '花﨑')

    assert trend.by_year == [(2025, [1, *[0] * 11])]
    assert Path(output_path) == config_path.parent / 'output' / '医療文書作成件数推移_花﨑_内科_20250401-20260331.xlsx'
    assert load_workbook(output_path)['推移']['A1'].value == (
        '医療文書作成件数推移(医師依頼日) 花﨑 内科 2025/04/01～2026/03/31')


def test_run_analysis_raises_when_no_data_in_period(config_path: Path) -> None:
    make_documents(config_path, [])

    with pytest.raises(NoDataError):
        MedicalDocsAnalyzer().run_analysis(date(2025, 4, 1), date(2025, 6, 30))
