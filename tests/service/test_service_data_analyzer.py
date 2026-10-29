from datetime import date

import polars as pl
import pytest

from service.service_data_analyzer import count_documents, filter_by_period, merge_documents
from tests.service.conftest import HEADERS


def make_df(rows: list[tuple[str, ...]]) -> pl.DataFrame:
    return pl.DataFrame(rows, schema={header: pl.String for header in HEADERS}, orient='row')


def document(received: str, patient_id: str, department: str, staff: str,
             requested: str = '2025/04/30', document_name: str = '診断書') -> tuple[str, ...]:
    return (received, patient_id, '患者', document_name, department, '佐藤', '', requested, staff)


def test_merge_documents_drops_rows_without_request_date_or_staff() -> None:
    source = make_df([
        document('2025/04/01', '1', '内科', '植田'),
        document('2025/04/02', '2', '内科', '植田', requested=''),
        document('2025/04/03', '3', '内科', ''),
    ])

    merged = merge_documents(source, None)

    assert merged['患者ID'].to_list() == ['1']


def test_merge_documents_removes_duplicates_with_database() -> None:
    source = make_df([
        document('2025/04/01', '1', '内科', '植田'),
        document('2025/04/01', '1', '内科', '沖野'),  # 担当者が違っても重複とみなす
        document('2025/04/01', '1', '内科', '植田', document_name='紹介状'),
    ])
    database = make_df([document('2025/04/01', '1', '内科', '花﨑')])

    merged = merge_documents(source, database)

    assert merged.select('文書名', '担当者名').rows() == [('診断書', '植田'), ('紹介状', '植田')]


def test_merge_documents_sorts_by_date_department_and_patient_id() -> None:
    source = make_df([
        document('2025/04/02', '1', '内科', '植田'),
        document('2025/04/01', '10', '内科', '植田'),
        document('2025/04/01', 'A', '内科', '植田'),
        document('2025/04/01', '9', '内科', '植田'),
        document('2025/04/01', '1', '外科', '植田'),
    ])

    merged = merge_documents(source, None)

    assert merged.select('預り日', '診療科', '患者ID').rows() == [
        ('2025/04/01', '内科', '9'),
        ('2025/04/01', '内科', '10'),
        ('2025/04/01', '内科', 'A'),
        ('2025/04/01', '外科', '1'),
        ('2025/04/02', '内科', '1'),
    ]


def test_merge_documents_rejects_different_columns() -> None:
    source = make_df([document('2025/04/01', '1', '内科', '植田')])
    database = source.rename({'備考': 'メモ'})

    with pytest.raises(ValueError, match='列が一致しません'):
        merge_documents(source, database)


def test_filter_by_period_includes_both_ends() -> None:
    df = make_df([
        document('2025/03/31', '1', '内科', '植田'),
        document('2025/04/01', '2', '内科', '植田'),
        document('2025/06/30', '3', '内科', '植田'),
        document('2025/07/01', '4', '内科', '植田'),
        document('不明', '5', '内科', '植田'),
    ])

    filtered = filter_by_period(df, date(2025, 4, 1), date(2025, 6, 30))

    assert filtered['患者ID'].to_list() == ['2', '3']


def test_count_documents() -> None:
    df = make_df([
        document('2025/04/05', '1', '内科', '植田'),
        document('2025/04/06', '2', '内科', '植田'),
        document('2025/04/01', '3', '外科', '花﨑'),
        document('2025/05/20', '4', '内科', '沖野'),
    ])

    counts = count_documents(df)

    assert counts.by_staff_department == {('植田', '内科'): 2, ('花﨑', '外科'): 1, ('沖野', '内科'): 1}
    assert counts.by_department == {'内科': 3, '外科': 1}
    assert counts.total == 4
    assert counts.first_date == date(2025, 4, 1)
    assert counts.last_date == date(2025, 5, 20)
