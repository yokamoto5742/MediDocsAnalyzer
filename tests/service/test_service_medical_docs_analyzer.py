from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from openpyxl import Workbook, load_workbook

from service import service_excel_writer
from service.service_excel_reader import NoDataError
from service.service_medical_docs_analyzer import MedicalDocsAnalyzer, split_list
from tests.service.conftest import MakeWorkbook


@pytest.fixture
def startfile(monkeypatch: pytest.MonkeyPatch) -> MagicMock:
    mock = MagicMock()
    monkeypatch.setattr(service_excel_writer.os, 'startfile', mock, raising=False)
    return mock


def test_split_list() -> None:
    assert split_list(' 植田, 花﨑 ,,') == ['植田', '花﨑']


def test_load_data_merges_source_into_database(config_path: Path, make_workbook: MakeWorkbook) -> None:
    tmp_path = config_path.parent
    make_workbook(tmp_path / 'source.xlsx', [
        [datetime(2025, 4, 3), '00123', '山田', '診断書', '内科', '佐藤', None, datetime(2025, 4, 4), '植田'],
        [datetime(2025, 4, 3), 123, '山田', '診断書', '内科', '佐藤', '重複', datetime(2025, 4, 4), '植田'],
        [datetime(2025, 4, 5), 456, '鈴木', '紹介状', '外科', '田中', None, None, '花﨑'],
    ])
    make_workbook(tmp_path / 'database.xlsx', [
        ['2025/04/01', 111, '既存', '証明書', '内科', '佐藤', None, '2025/04/02', '沖野'],
    ])

    count = MedicalDocsAnalyzer().load_data()

    assert count == 2
    sheet = load_workbook(tmp_path / 'database.xlsx').active
    assert sheet is not None
    assert [row[:2] for row in sheet.iter_rows(min_row=2, values_only=True)] == [
        ('2025/04/01', 111), ('2025/04/03', 123)]
    assert (tmp_path / 'backup' / 'backup_database.xlsx').exists()


def test_load_data_creates_database(config_path: Path, make_workbook: MakeWorkbook) -> None:
    tmp_path = config_path.parent
    make_workbook(tmp_path / 'source.xlsx', [
        ['2025/04/03', 1, '山田', '診断書', '内科', '佐藤', None, '2025/04/04', '植田'],
    ])

    assert MedicalDocsAnalyzer().load_data() == 1
    assert (tmp_path / 'database.xlsx').exists()


def test_load_data_raises_when_source_is_empty(config_path: Path) -> None:
    Workbook().save(config_path.parent / 'source.xlsx')

    with pytest.raises(NoDataError):
        MedicalDocsAnalyzer().load_data()


def test_run_analysis(config_path: Path, make_workbook: MakeWorkbook, startfile: MagicMock) -> None:
    tmp_path = config_path.parent
    Workbook().save(tmp_path / 'template.xlsx')
    make_workbook(tmp_path / 'database.xlsx', [
        ['2025/03/31', 1, '対象外', '診断書', '内科', '佐藤', None, '2025/04/01', '植田'],
        ['2025/04/01', 2, '山田', '診断書', '内科', '佐藤', None, '2025/04/02', '植田'],
        ['2025/05/10', 3, '鈴木', '紹介状', '外科', '田中', None, '2025/05/11', '花﨑'],
    ])

    output_path = MedicalDocsAnalyzer().run_analysis(date(2025, 4, 1), date(2025, 6, 30))

    assert Path(output_path).name == '医療文書作成件数20250401-20250510.xlsx'
    sheet = load_workbook(output_path).active
    assert sheet is not None
    assert [list(row) for row in sheet.iter_rows(min_row=2, values_only=True)] == [
        ['氏名', '合計', '内科', '外科'],
        ['植田', 1, 1, 0],
        ['沖野', 0, 0, 0],
        ['花﨑', 1, 0, 1],
        ['合計', 2, 1, 1],
    ]
    startfile.assert_called_once()


def test_run_analysis_raises_when_no_data_in_period(config_path: Path, make_workbook: MakeWorkbook) -> None:
    make_workbook(config_path.parent / 'database.xlsx', [
        ['2025/03/31', 1, '対象外', '診断書', '内科', '佐藤', None, '2025/04/01', '植田'],
    ])

    with pytest.raises(NoDataError):
        MedicalDocsAnalyzer().run_analysis(date(2025, 4, 1), date(2025, 6, 30))
