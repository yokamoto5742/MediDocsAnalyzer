from dataclasses import dataclass
from datetime import date

import polars as pl

from service.service_excel_reader import DATE_FORMAT, PATIENT_ID_COLUMN

REQUIRED_COLUMNS = ('医師依頼日', '担当者名')
DUPLICATE_KEY_COLUMNS = ['預り日', PATIENT_ID_COLUMN, '文書名', '診療科', '医師名']
SORT_COLUMNS = ['預り日', '診療科']


@dataclass
class DocumentCounts:
    by_staff_department: dict[tuple[str, str], int]
    by_department: dict[str, int]
    total: int
    first_date: date
    last_date: date


def received_date() -> pl.Expr:
    return pl.col('預り日').str.to_date(DATE_FORMAT, strict=False)


def merge_documents(source_df: pl.DataFrame, database_df: pl.DataFrame | None) -> pl.DataFrame:
    """元データと DB を結合し、集計対象外の行と重複行を除いて並べ替える"""
    df = source_df
    if database_df is not None and database_df.height > 0:
        if set(database_df.columns) != set(source_df.columns):
            raise ValueError(
                f'元データとデータベースの列が一致しません。\n'
                f'元データ: {source_df.columns}\nデータベース: {database_df.columns}')
        df = pl.concat([source_df, database_df.select(source_df.columns)])

    for column in REQUIRED_COLUMNS:
        df = df.filter(pl.col(column) != '')

    # 患者IDは数値として並べ、数値でないものは後ろに回す
    patient_id_number = pl.col(PATIENT_ID_COLUMN).cast(pl.Int64, strict=False)
    return (
        df.unique(subset=DUPLICATE_KEY_COLUMNS, keep='first', maintain_order=True)
        .sort([*SORT_COLUMNS, patient_id_number, PATIENT_ID_COLUMN], nulls_last=True)
    )


def filter_by_period(df: pl.DataFrame, start_date: date, end_date: date) -> pl.DataFrame:
    return df.filter(received_date().is_between(start_date, end_date))


def count_documents(df: pl.DataFrame) -> DocumentCounts:
    """担当者×診療科、診療科ごと、全体の作成件数を数える"""
    by_staff_department = df.group_by(['担当者名', '診療科']).len()
    by_department = df.group_by('診療科').len()
    dates = df.select(received_date())

    return DocumentCounts(
        by_staff_department={(staff, department): count
                             for staff, department, count in by_staff_department.iter_rows()},
        by_department=dict(by_department.iter_rows()),
        total=df.height,
        first_date=dates.min().item(),
        last_date=dates.max().item(),
    )
