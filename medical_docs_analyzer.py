import openpyxl
import polars as pl
import warnings
from datetime import datetime

warnings.simplefilter('ignore')


def analyze_medical_documents(file_path):
    """
    医療文書の作成件数を担当者・診療科別に集計して表示する関数

    Parameters:
        file_path (str): 分析対象のExcelファイルパス
    """
    try:
        # Excelファイルを読み込む
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook.active

        # ヘッダー行を取得して列名を特定
        headers = []
        for cell in sheet[1]:
            headers.append(cell.value)

        # データを格納するためのリスト
        data = []

        # 2行目からデータを読み込む（1行目はヘッダー）
        for row in list(sheet.rows)[1:]:
            row_data = {}
            for i, cell in enumerate(row):
                if i < len(headers) and headers[i] is not None:
                    row_data[headers[i]] = cell.value

            # 空でない行のみ追加
            if row_data.get('預り日') is not None:
                data.append(row_data)

        # データをpolarsのDataFrameに変換
        df = pl.DataFrame(data)

        # 空のDataFrameチェック
        if df.height == 0 or '担当者名' not in df.columns or '診療科' not in df.columns or '預り日' not in df.columns:
            print("分析対象となるデータがありません。")
            return

        # 預り日を日付型に変換
        df = df.with_columns(
            pl.when(pl.col('預り日').is_not_null())
            .then(pl.col('預り日'))
            .otherwise(None)
            .alias('預り日')
        )

        # 期間の取得（None値を除外）
        valid_dates = df.filter(pl.col('預り日').is_not_null()).select('預り日')
        if valid_dates.height == 0:
            print("有効な日付データがありません。")
            return

        # 日付データが文字列の場合は、datetime型に変換してから処理
        min_date = valid_dates.select(pl.col('預り日').min()).item()
        max_date = valid_dates.select(pl.col('預り日').max()).item()

        # 日付フォーマットの処理（文字列か日付型かを判断）
        if isinstance(min_date, str):
            # 文字列の場合はそのまま表示するか、必要に応じて日付型に変換
            start_date = min_date
            end_date = max_date
        else:
            # datetimeオブジェクトの場合はstrftimeを使用
            start_date = min_date.strftime('%Y年%m月%d日')
            end_date = max_date.strftime('%Y年%m月%d日')

        # 担当者と診療科で集計
        grouped = df.filter(
            pl.col('担当者名').is_not_null() & pl.col('診療科').is_not_null()
        ).group_by(['担当者名', '診療科']).agg(
            pl.count().alias('作成件数')
        )

        # 担当者ごとの合計件数を計算
        staff_totals = df.filter(
            pl.col('担当者名').is_not_null()
        ).group_by('担当者名').agg(
            pl.count().alias('作成件数')
        ).sort('担当者名')

        # 担当者のリストを取得（None値を除外）
        staff_members = df.select('担当者名').filter(
            pl.col('担当者名').is_not_null()
        ).unique().to_series().to_list()

        # 結果を表示
        print(f"\n=== 担当者別医療文書作成件数 ===")
        print(f"{start_date}-{end_date}\n")

        for staff in staff_members:
            # 担当者名を表示
            print(f"{staff}")

            # 合計件数を表示（安全に取得）
            total_row = staff_totals.filter(pl.col('担当者名') == staff).select('作成件数')
            if total_row.height == 0:
                total_docs = 0
            else:
                total_docs = total_row.item()
            print(f"合計: {total_docs}件")

            # 内訳のヘッダーを表示
            print("(内訳)")

            # この担当者の診療科別データを取得し、件数で降順ソート
            dept_data = grouped.filter(pl.col('担当者名') == staff).sort('作成件数', descending=True)

            # 診療科別の件数を表示
            for row in dept_data.iter_rows(named=True):
                print(f"{row['診療科']}: {row['作成件数']}件")
            print()

    except FileNotFoundError:
        print(f"エラー: ファイル '{file_path}' が見つかりません。")
    except Exception as e:
        print(f"エラー: データの読み込み中に問題が発生しました - {str(e)}")


# 使用例
if __name__ == "__main__":
    file_path = "医療文書担当一覧データベース.xlsx"
    analyze_medical_documents(file_path)
