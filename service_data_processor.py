import datetime

import polars as pl


def process_cell_value(cell):
    """
    セルの値を適切に処理する

    Args:
        cell: 処理するセル

    Returns:
        処理後の値
    """
    # セルがNoneまたはEmptyCellの場合は処理しない
    if cell is None:
        return None

    # hasattrを使用してcolumn属性があるか確認
    if not hasattr(cell, 'column'):
        return None

    if cell.column == 1 and cell.value:  # A列（預り日）
        if isinstance(cell.value, (datetime.datetime, datetime.date)):
            # 日付オブジェクトを文字列に変換
            return cell.value.strftime('%Y/%m/%d')
        return cell.value
    elif cell.column == 2 and cell.value:  # B列（患者ID）
        # 患者IDを数値として処理
        try:
            if isinstance(cell.value, str):
                # 文字列の場合は数値に変換
                return int(cell.value)
            return cell.value
        except (ValueError, TypeError):
            # 数値変換できない場合は元の値を使用
            return cell.value
    elif cell.column == 8 and cell.value:  # H列（医師依頼日）
        if isinstance(cell.value, (datetime.datetime, datetime.date)):
            # 日付オブジェクトを文字列に変換
            return cell.value.strftime('%Y/%m/%d')
        return cell.value
    else:
        return cell.value


def format_date_string(value):
    """
    日付文字列を標準形式に変換する

    Args:
        value: 日付文字列

    Returns:
        標準化された日付文字列
    """
    if isinstance(value, str) and value:
        # 日付部分のみを抽出（YYYY-MM-DD または YYYY/MM/DD 形式）
        date_parts = value.split()[0] if ' ' in value else value

        # 年月日の区切りを"/"に統一
        if '-' in date_parts:
            parts = date_parts.split('-')
            if len(parts) == 3:
                year, month, day = parts
                return f"{year}/{month}/{day}"
            else:
                return date_parts
        else:
            return date_parts
    return value


def format_output_cell_value(col_idx, value):
    """
    出力用のセル値を書式設定する

    Args:
        col_idx: 列インデックス
        value: セルの値

    Returns:
        書式設定後の値
    """
    # 1列目（預り日）または8列目（医師依頼日）の場合、日付形式を適用
    if (col_idx == 1 or col_idx == 8) and value:
        return format_date_string(value)
    # 2列目（患者ID）の場合、数値形式を適用
    elif col_idx == 2 and value is not None:
        # 数値として設定
        try:
            if isinstance(value, str) and value.strip():
                return int(value)
            else:
                return value
        except (ValueError, TypeError):
            return value
    else:
        return value


def parse_date_to_formats(date_value):
    """
    日付を各種フォーマットに変換する共通関数

    Args:
        date_value: 変換する日付値

    Returns:
        各種フォーマットの辞書
    """
    if date_value is None:
        return {
            'raw': None,
            'file_format': '',
            'display_format': ''
        }

    if isinstance(date_value, str):
        if not date_value.strip():
            return {
                'raw': '',
                'file_format': '',
                'display_format': ''
            }

        try:
            # 日付文字列を標準形式に変換してから処理
            date_str = format_date_string(date_value)
            date_obj = datetime.datetime.strptime(date_str, '%Y/%m/%d')
        except ValueError:
            # フォーマットが異なる場合
            return {
                'raw': date_value,
                'file_format': date_value.replace('/', ''),
                'display_format': date_value
            }
    else:
        date_obj = date_value

    return {
        'raw': date_obj,
        'file_format': date_obj.strftime('%Y%m%d'),
        'display_format': date_obj.strftime('%Y年%m月%d日')
    }


def filter_dataframe_by_date_range(df, start_date_str=None, end_date_str=None):
    """
    データフレームを日付範囲でフィルタリングする

    Args:
        df: フィルタリングするDataFrame
        start_date_str: 開始日（YYYY-MM-DD形式）
        end_date_str: 終了日（YYYY-MM-DD形式）

    Returns:
        フィルタリング結果と日付情報を含む辞書
    """
    # データフレームが空の場合は早期リターン
    if df is None or len(df.columns) == 0 or df.height == 0:
        return {
            'df': df if df is not None else pl.DataFrame(),
            'start_date_display': '該当なし',
            'end_date_display': '該当なし',
            'file_date_range': 'no_data'
        }

    # '預り日'カラムが存在するか確認
    if '預り日' not in df.columns:
        return {
            'df': df,
            'start_date_display': '該当なし',
            'end_date_display': '該当なし',
            'file_date_range': 'no_data'
        }

    # 日付列の前処理
    df = df.with_columns(
        pl.when(pl.col('預り日').is_not_null())
        .then(pl.col('預り日'))
        .otherwise(None)
        .alias('預り日')
    )

    if start_date_str and end_date_str:
        try:
            # 日付形式を統一
            start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d')

            df = df.with_columns([
                pl.when(pl.col('預り日').cast(pl.Utf8).str.contains('-'))
                .then(pl.col('預り日').cast(pl.Utf8).str.replace('-', '/'))
                .otherwise(pl.col('預り日'))
                .alias('預り日')
            ])

            # 日付範囲でフィルタリング
            df = df.filter(
                (pl.col('預り日') >= start_date.strftime('%Y/%m/%d')) &
                (pl.col('預り日') <= end_date.strftime('%Y/%m/%d'))
            )
        except (ValueError, TypeError) as e:
            print(f"日付フィルタリング中にエラーが発生: {e}")
            # エラーが発生した場合はフィルタリングせずに続行

    # 有効な日付を抽出
    try:
        valid_dates = df.filter(pl.col('預り日').is_not_null()).select('預り日')

        if valid_dates.height > 0:
            min_date = valid_dates.select(pl.col('預り日').min()).item()
            max_date = valid_dates.select(pl.col('預り日').max()).item()

            # 日付フォーマットを処理
            min_date_formats = parse_date_to_formats(min_date)
            max_date_formats = parse_date_to_formats(max_date)

            return {
                'df': df,
                'start_date_display': min_date_formats['display_format'],
                'end_date_display': max_date_formats['display_format'],
                'file_date_range': f"{min_date_formats['file_format']}-{max_date_formats['file_format']}"
            }
    except Exception as e:
        print(f"日付データ処理中にエラーが発生: {e}")

    return {
        'df': df,
        'start_date_display': '該当なし',
        'end_date_display': '該当なし',
        'file_date_range': 'no_data'
    }


def clean_and_standardize_dataframe(df):
    """
    データフレームをクリーニングして標準化する

    Args:
        df: 処理するDataFrame

    Returns:
        処理後のDataFrame
    """
    # データフレームが空の場合は早期リターン
    if df is None or len(df.columns) == 0:
        return pl.DataFrame()

    # すべての列を文字列型に、空のセルを空文字に変換して処理
    for col in df.columns:
        df = df.with_columns([
            pl.col(col).cast(pl.Utf8).fill_null("").alias(col)
        ])

    return df