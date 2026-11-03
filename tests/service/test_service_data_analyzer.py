import sqlite3
from collections.abc import Iterator
from datetime import date

import pytest

from service.models import Document, NoDataError
from service.service_data_analyzer import (DocumentCounts, DocumentFilter, DocumentType, OverallSummary,
                                           count_documents, count_staff_documents, count_trend, months_between,
                                           trend_subject)
from service.service_database import connect, replace_documents_between


def document(received: date, department: str, staff: str) -> Document:
    return Document(received, received, '診断書', department, '佐藤', staff)


def test_overall_summary_rows() -> None:
    counts = DocumentCounts({('植田', '内科'): 2, ('植田', '小児科'): 1, ('花﨑', '外科'): 1, ('外部', '内科'): 1},
                            {'内科': 3, '小児科': 1, '外科': 1}, 5, date(2025, 4, 1), date(2025, 4, 30))

    rows = OverallSummary(counts, ['植田', '花﨑'], ['内科', '外科']).rows

    # 担当者の合計は並べた診療科だけ、合計行の合計は一覧にない担当者・診療科も含める
    assert rows == [['植田', 2, 2, 0], ['花﨑', 1, 0, 1], ['合計', 5, 3, 1]]


@pytest.fixture
def connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    replace_documents_between(connection, date(2025, 1, 1), date(2099, 12, 31), [
        document(date(2025, 3, 31), '内科', '植田'),
        document(date(2025, 4, 5), '内科', '植田'),
        document(date(2025, 4, 6), '内科', '植田'),
        document(date(2025, 4, 1), '外科', '花﨑'),
        document(date(2025, 6, 30), '内科', '沖野'),
        document(date(2025, 7, 1), '内科', '沖野'),
    ])
    yield connection
    connection.close()


def test_count_documents_includes_both_ends(connection: sqlite3.Connection) -> None:
    counts = count_documents(connection, date(2025, 4, 1), date(2025, 6, 30))

    assert counts.by_staff_department == {('植田', '内科'): 2, ('花﨑', '外科'): 1, ('沖野', '内科'): 1}
    assert counts.by_department == {'内科': 3, '外科': 1}
    assert counts.total == 4
    assert counts.first_date == date(2025, 4, 1)
    assert counts.last_date == date(2025, 6, 30)


@pytest.fixture
def opinion_connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    replace_documents_between(connection, date(2025, 1, 1), date(2099, 12, 31), [
        Document(date(2025, 4, 1), date(2025, 4, 2), '主治医意見書(新規)', '内科', '佐藤', '植田'),
        Document(date(2025, 4, 3), date(2025, 4, 4), '主治医意見書', '外科', '鈴木', '植田'),
        Document(date(2025, 4, 5), date(2025, 5, 6), '主治医意見書 更新', '外科', '鈴木', '花﨑'),
        Document(date(2025, 4, 7), date(2025, 4, 8), '診断書', '内科', '佐藤', '植田'),
    ])
    yield connection
    connection.close()


@pytest.mark.parametrize(('document_type', 'expected_rows'), [
    (DocumentType.ALL, [['植田', 3, 2, 1], ['花﨑', 1, 0, 1], ['合計', 4, 2, 2]]),
    (DocumentType.MEDICAL, [['植田', 1, 1, 0], ['花﨑', 0, 0, 0], ['合計', 1, 1, 0]]),
    (DocumentType.OPINION, [['植田', 2, 1, 1], ['花﨑', 1, 0, 1], ['合計', 3, 1, 2]]),
])
def test_count_documents_by_document_type(opinion_connection: sqlite3.Connection, document_type: DocumentType,
                                          expected_rows: list[list[str | int]]) -> None:
    counts = count_documents(opinion_connection, date(2025, 4, 1), date(2025, 4, 30), document_type)

    assert OverallSummary(counts, ['植田', '花﨑'], ['内科', '外科']).rows == expected_rows


def test_count_documents_medical_and_opinion_add_up_to_all(opinion_connection: sqlite3.Connection) -> None:
    def total(document_type: DocumentType) -> int:
        return count_documents(opinion_connection, date(2025, 4, 1), date(2025, 4, 30), document_type).total

    assert total(DocumentType.MEDICAL) + total(DocumentType.OPINION) == total(DocumentType.ALL)


def test_count_staff_documents_by_document_type(opinion_connection: sqlite3.Connection) -> None:
    all_counts = count_staff_documents(opinion_connection, '植田', date(2025, 4, 1), date(2025, 4, 30))
    opinion = count_staff_documents(opinion_connection, '植田', date(2025, 4, 1), date(2025, 4, 30),
                                    document_type=DocumentType.OPINION)
    medical = count_staff_documents(opinion_connection, '植田', date(2025, 4, 1), date(2025, 4, 30),
                                    document_type=DocumentType.MEDICAL)

    assert (all_counts.total, all_counts.by_department) == (3, [('内科', 2), ('外科', 1)])
    assert (opinion.total, opinion.by_department) == (2, [('内科', 1), ('外科', 1)])
    assert (medical.total, medical.by_department, medical.by_document) == (1, [('内科', 1)], [('診断書', 1)])
    assert medical.document_type == DocumentType.MEDICAL


def test_count_trend_by_document_type(opinion_connection: sqlite3.Connection) -> None:
    opinion = DocumentFilter(document_type=DocumentType.OPINION)
    trend = count_trend(opinion_connection, opinion, date(2025, 4, 1), date(2025, 5, 31), 4, date(2025, 5, 31))

    assert trend.by_year == [(2025, [2, 1, *[None] * 10])]
    staff_trend = count_trend(opinion_connection, opinion, date(2025, 4, 1),
                              date(2025, 5, 31), 4, date(2025, 5, 31), '花﨑')
    assert staff_trend.by_year == [(2025, [0, 1, *[None] * 10])]
    medical_surgery = count_trend(opinion_connection, DocumentFilter('内科', DocumentType.MEDICAL),
                                  date(2025, 4, 1), date(2025, 5, 31), 4, date(2025, 5, 31))
    assert medical_surgery.by_year == [(2025, [1, 0, *[None] * 10])]


def test_trend_subject() -> None:
    assert trend_subject(DocumentFilter(), None) == '全書類'
    assert trend_subject(DocumentFilter('外科', DocumentType.MEDICAL), '植田') == '植田 医療文書 外科'


def test_months_between() -> None:
    assert months_between(date(2024, 11, 15), date(2025, 2, 1)) == ['2024/11', '2024/12', '2025/01', '2025/02']


def test_count_staff_documents(connection: sqlite3.Connection) -> None:
    replace_documents_between(connection, date(2025, 8, 1), date(2099, 12, 31), [
        Document(date(2025, 8, 1), date(2025, 8, 2), '紹介状', '内科', '佐藤', '植田'),
        Document(date(2025, 8, 3), date(2025, 8, 4), '紹介状', '外科', '佐藤', '植田'),
        Document(date(2025, 8, 5), date(2025, 8, 6), '診断書', '外科', '佐藤', '植田'),
        Document(date(2025, 10, 1), date(2025, 10, 2), '診断書', '外科', '佐藤', '植田'),
        Document(date(2025, 8, 1), date(2025, 8, 2), '診断書', '内科', '佐藤', '花﨑'),
    ])

    counts = count_staff_documents(connection, '植田', date(2025, 8, 1), date(2025, 10, 31), document_limit=1)

    assert counts.total == 4
    assert counts.by_department == [('外科', 3), ('内科', 1)]
    assert counts.by_document == [('紹介状', 2)]
    assert counts.by_month == [('2025/08', 3), ('2025/09', 0), ('2025/10', 1)]

    all_documents = count_staff_documents(connection, '植田', date(2025, 8, 1), date(2025, 10, 31))
    assert all_documents.by_document == [('紹介状', 2), ('診断書', 2)]


def test_count_staff_documents_without_data(connection: sqlite3.Connection) -> None:
    counts = count_staff_documents(connection, '植田', date(2026, 1, 1), date(2026, 1, 31), document_limit=10)

    assert (counts.total, counts.by_month) == (0, [('2026/01', 0)])


def test_count_trend_uses_requested_date_by_fiscal_year() -> None:
    connection = connect(':memory:')
    replace_documents_between(connection, date(2024, 1, 1), date(2099, 12, 31), [
        # 預り日ではなく医師依頼日で数える
        Document(date(2025, 3, 20), date(2025, 4, 2), '診断書', '内科', '佐藤', '植田'),
        Document(date(2025, 4, 3), date(2025, 4, 3), '診断書', '外科', '佐藤', '植田'),
        Document(date(2026, 3, 1), date(2026, 3, 31), '診断書', '内科', '佐藤', '植田'),
        Document(date(2026, 4, 1), date(2026, 4, 1), '診断書', '内科', '佐藤', '植田'),
        Document(date(2024, 3, 1), date(2024, 3, 31), '診断書', '内科', '佐藤', '植田'),
    ])

    # 今月(5月)は途中までの件数を出し、本日より後の月は None
    trend = count_trend(connection, DocumentFilter('内科'), date(2025, 4, 1), date(2027, 3, 31), 4, date(2026, 5, 15))

    assert trend.month_labels[:3] == ['4月', '5月', '6月']
    assert trend.by_year == [
        (2025, [1, *[0] * 10, 1]),
        (2026, [1, 0, *[None] * 10]),
    ]
    assert trend.table_rows[0][0] == '2025年度'
    assert trend.table_rows[1][-1] == 1
    total = count_trend(connection, DocumentFilter(), date(2025, 4, 1), date(2026, 3, 31), 4, date(2026, 5, 15))
    assert total.by_year == [(2025, [2, *[0] * 10, 1])]
    connection.close()


def test_count_trend_limits_to_date_range() -> None:
    connection = connect(':memory:')
    replace_documents_between(connection, date(2024, 1, 1), date(2099, 12, 31), [
        Document(date(2025, 2, 1), date(2025, 2, 10), '診断書', '内科', '佐藤', '植田'),
        Document(date(2025, 2, 1), date(2025, 2, 20), '診断書', '内科', '佐藤', '植田'),
        Document(date(2025, 11, 1), date(2025, 11, 2), '診断書', '内科', '佐藤', '植田'),
    ])

    # 範囲の前後の月は None、範囲の端の月は範囲内の日だけ数える
    trend = count_trend(connection, DocumentFilter(), date(2025, 2, 15), date(2025, 11, 1), 1, date(2026, 11, 1))

    assert trend.by_year == [(2025, [None, 1, *[0] * 9, None])]
    connection.close()


def test_count_trend_calendar_year() -> None:
    trend = count_trend(connect(':memory:'), DocumentFilter(), date(2025, 1, 1), date(2025, 12, 31), 1, date(2026, 1, 1))

    assert (trend.year_header, trend.month_labels[0]) == ('年', '1月')
    assert trend.by_year == [(2025, [0] * 12)]


def test_count_documents_raises_when_no_data(connection: sqlite3.Connection) -> None:
    with pytest.raises(NoDataError):
        count_documents(connection, date(2026, 4, 1), date(2026, 6, 30))
