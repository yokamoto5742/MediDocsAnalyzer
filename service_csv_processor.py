import datetime
import shutil
from pathlib import Path

import polars as pl

from config_manager import ConfigManager


def read_csv_with_encoding(file_path):
    encodings = ['shift-jis', 'utf-8']

    for encoding in encodings:
        try:
            schema = {
                "患者ID": pl.Int64,
            }
            df = pl.read_csv(
                file_path,
                encoding=encoding,
                separator=',',
                skip_rows=3,  # 最初の3行をスキップ
                has_header=True,  # 4行目をヘッダーとして使用
                infer_schema_length=0,
                schema_overrides=schema
            )

            if len(df.columns) > 1:
                print(f"エンコーディング {encoding} で正常に読み込みました")
                print(f"列数: {len(df.columns)}")
                print(f"行数: {len(df)}")
                print(f"列名: {df.columns}")
                return df
        except Exception as e:
            print(f"{encoding}での読み込み試行中にエラー: {str(e)}")
            continue

    raise Exception("CSVファイルの読み込みに失敗しました")


def process_csv_data(df):
    try:
        # 列名を一意にする
        original_columns = df.columns
        unique_columns = []
        for i, col in enumerate(original_columns):
            unique_columns.append(f"col_{i}_{col}")
        df = df.select([
            pl.col(old_name).alias(new_name)
            for old_name, new_name in zip(original_columns, unique_columns)
        ])

        # K列とI列を削除 (インデックスベースで削除)
        columns_to_keep = [i for i in range(len(df.columns)) if i not in [8, 10]]
        df = df.select([df.columns[i] for i in columns_to_keep])

        # A列からC列を削除
        df = df.select(df.columns[3:])

        config = ConfigManager()
        exclude_docs = config.get_exclude_docs()
        exclude_doctors = config.get_exclude_doctors()

        if exclude_docs:
            filter_conditions = [~pl.col(df.columns[3]).str.contains(doc) for doc in exclude_docs]
            combined_filter = filter_conditions[0]
            for condition in filter_conditions[1:]:
                combined_filter = combined_filter & condition
            df = df.filter(combined_filter)

        if exclude_doctors:
            doctor_filter_conditions = [~pl.col(df.columns[5]).str.contains(doc) for doc in exclude_doctors]
            doctor_combined_filter = doctor_filter_conditions[0]
            for condition in doctor_filter_conditions[1:]:
                doctor_combined_filter = doctor_combined_filter & condition
            df = df.filter(doctor_combined_filter)

            # D列とF列のスペースと*を除去（4列目と6列目）
            df = df.with_columns([
                pl.col(df.columns[3]).str.replace_all(r'[\s*]', ''),  # D列
                pl.col(df.columns[5]).str.replace_all(r'[\s*]', '')  # F列
            ])

            return df

        return df

    except Exception as e:
        print(f"データ処理中にエラーが発生しました: {str(e)}")
        raise


def convert_date_format(df):
    try:
        date_col = df.columns[3]
        df = df.with_columns([
            pl.col(date_col).str.strptime(pl.Date, format="%Y%m%d")
            .alias(date_col)
        ])
        return df
    except Exception as e:
        print(f"日付変換中にエラーが発生しましたが、処理を継続します: {str(e)}")
        return df


def process_completed_csv(csv_path: str):
    try:
        csv_file = Path(csv_path)
        if not csv_file.exists():
            return

        config = ConfigManager()
        processed_dir = Path(config.get_processed_path())
        processed_dir.mkdir(exist_ok=True, parents=True)

        new_path = processed_dir / csv_file.name
        shutil.move(str(csv_file), str(new_path))

    except Exception as e:
        print(f"CSVファイルの処理中にエラーが発生しました: {str(e)}")
        raise


def find_latest_csv(downloads_path):
    # ファイル名の形式がYYYYMMDD_HHmmss.csvのファイルを検索
    csv_files = [f for f in Path(downloads_path).glob('*.csv')
                if len(f.name.split('_')) == 2 and
                (3 <= len(f.name.split('_')[0]) <= 4) and
                len(f.name.split('_')[1].split('.')[0]) == 14]
    
    if not csv_files:
        return None

    return str(max(csv_files, key=lambda f: f.stat().st_mtime))
