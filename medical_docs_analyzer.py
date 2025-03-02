import polars as pl
import warnings
from datetime import datetime

warnings.simplefilter('ignore')


def analyze_medical_documents(file_path):
    """
    医療文書の作成件数を担当者・診療科別に集計して表示する関数

    Parameters:
        file_path (str): 分析対象のCSVファイルパス
    """
    try:
        # CSVファイルを読み込む
        df = pl.read_csv(file_path, encoding='utf-8')

        # 預り日を日付型に変換
        df = df.with_columns(pl.col('預り日').str.to_datetime())

        # 期間の取得
        start_date = df.select(pl.col('預り日').min()).item().strftime('%Y年%m月%d日')
        end_date = df.select(pl.col('預り日').max()).item().strftime('%Y年%m月%d日')

        # 担当者と診療科で集計
        grouped = df.group_by(['担当者名', '診療科']).agg(
            pl.count().alias('作成件数')
        )

        # 担当者ごとの合計件数を計算
        staff_totals = df.group_by('担当者名').agg(
            pl.count().alias('作成件数')
        ).sort('担当者名')

        # 担当者のリストを取得（None値を除外）
        staff_members = df.select('担当者名').filter(pl.col('担当者名').is_not_null()).unique().to_series().to_list()

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
    file_path = "医療文書作成件数.csv"  # CSVファイルのパス
    analyze_medical_documents(file_path)
