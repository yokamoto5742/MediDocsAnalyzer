import openpyxl
import polars as pl
import warnings
from datetime import datetime
import os
from shutil import copyfile

warnings.simplefilter('ignore')


def analyze_medical_documents(file_path, template_path="医療文書作成件数.xlsx"):
    """
    医療文書の作成件数を担当者・診療科別に集計して表示し、Excelに出力する関数

    Parameters:
        file_path (str): 分析対象のExcelファイルパス
        template_path (str): 出力用テンプレートのExcelファイルパス
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
            # 文字列の場合はdatetimeに変換
            try:
                min_date_obj = datetime.strptime(min_date, '%Y/%m/%d')
                max_date_obj = datetime.strptime(max_date, '%Y/%m/%d')
                start_date = min_date_obj.strftime('%Y年%m月%d日')
                end_date = max_date_obj.strftime('%Y年%m月%d日')
                file_date_range = f"{min_date_obj.strftime('%Y%m%d')}-{max_date_obj.strftime('%Y%m%d')}"
            except ValueError:
                # 日付形式が違う場合はそのまま使用
                start_date = min_date
                end_date = max_date
                file_date_range = f"{start_date}-{end_date}".replace('/', '')
        else:
            # datetimeオブジェクトの場合はstrftimeを使用
            start_date = min_date.strftime('%Y年%m月%d日')
            end_date = max_date.strftime('%Y年%m月%d日')
            file_date_range = f"{min_date.strftime('%Y%m%d')}-{max_date.strftime('%Y%m%d')}"

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

        # 診療科ごとの合計件数を計算
        dept_totals = df.filter(
            pl.col('診療科').is_not_null()
        ).group_by('診療科').agg(
            pl.count().alias('作成件数')
        ).sort('作成件数', descending=True)

        # 担当者のリストを取得（None値を除外）
        staff_members = df.select('担当者名').filter(
            pl.col('担当者名').is_not_null()
        ).unique().sort('担当者名').to_series().to_list()

        # 診療科のリストを取得（None値を除外）
        departments = df.select('診療科').filter(
            pl.col('診療科').is_not_null()
        ).unique().sort('診療科').to_series().to_list()

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

        # Excelに出力
        output_excel(template_path, staff_members, departments, grouped, staff_totals,
                     dept_totals, start_date, end_date, file_date_range)

    except FileNotFoundError:
        print(f"エラー: ファイル '{file_path}' が見つかりません。")
    except Exception as e:
        print(f"エラー: データの読み込み中に問題が発生しました - {str(e)}")


def output_excel(template_path, staff_members, departments, grouped_data, staff_totals,
                 dept_totals, start_date, end_date, file_date_range):
    """
    集計結果をExcelファイルに出力する関数

    Parameters:
        template_path (str): テンプレートExcelファイルのパス
        staff_members (list): 担当者リスト
        departments (list): 診療科リスト
        grouped_data (pl.DataFrame): 担当者×診療科の集計データ
        staff_totals (pl.DataFrame): 担当者ごとの合計件数
        dept_totals (pl.DataFrame): 診療科ごとの合計件数
        start_date (str): 期間の開始日
        end_date (str): 期間の終了日
        file_date_range (str): ファイル名用の日付範囲
    """
    try:
        # 出力ファイル名の設定
        output_file = f"医療文書作成件数{file_date_range}.xlsx"

        # テンプレートが存在する場合はコピー、存在しない場合は新規作成
        if os.path.exists(template_path):
            copyfile(template_path, output_file)
            workbook = openpyxl.load_workbook(output_file)
            sheet = workbook.active
        else:
            print(f"警告: テンプレートファイル '{template_path}' が見つかりません。新規ファイルを作成します。")
            workbook = openpyxl.Workbook()
            sheet = workbook.active

        # タイトル行の設定
        sheet['A1'] = f"医療文書作成件数 {start_date}-{end_date}"

        # ヘッダー行の設定（すでにテンプレートにある場合は上書き）
        if len(departments) > 0:
            sheet['A2'] = "氏名"
            for col_idx, dept in enumerate(departments, 2):
                sheet.cell(row=2, column=col_idx).value = dept
            sheet.cell(row=2, column=len(departments) + 2).value = "合計"

        # データの書き込み
        for row_idx, staff in enumerate(staff_members, 3):
            # 担当者名
            sheet.cell(row=row_idx, column=1).value = staff

            # 各診療科の件数
            staff_total = 0
            for col_idx, dept in enumerate(departments, 2):
                # この担当者×診療科の件数を取得
                filtered_data = grouped_data.filter(
                    (pl.col('担当者名') == staff) & (pl.col('診療科') == dept)
                )

                if filtered_data.height > 0:
                    count = filtered_data.select('作成件数').item()
                    sheet.cell(row=row_idx, column=col_idx).value = count
                    staff_total += count
                else:
                    sheet.cell(row=row_idx, column=col_idx).value = 0

            # 合計欄
            sheet.cell(row=row_idx, column=len(departments) + 2).value = staff_total

        # 合計行
        total_row = len(staff_members) + 3
        sheet.cell(row=total_row, column=1).value = "合計"

        # 各診療科の合計
        for col_idx, dept in enumerate(departments, 2):
            dept_data = dept_totals.filter(pl.col('診療科') == dept)
            if dept_data.height > 0:
                sheet.cell(row=total_row, column=col_idx).value = dept_data.select('作成件数').item()
            else:
                sheet.cell(row=total_row, column=col_idx).value = 0

        # 全体の合計
        total_docs = staff_totals.select('作成件数').sum().item()
        sheet.cell(row=total_row, column=len(departments) + 2).value = total_docs

        # ファイルを保存
        workbook.save(output_file)
        print(f"集計結果を '{output_file}' に保存しました。")

    except Exception as e:
        print(f"エラー: Excelファイルの出力中に問題が発生しました - {str(e)}")


# 使用例
if __name__ == "__main__":
    file_path = "医療文書担当一覧データベース.xlsx"
    template_path = "医療文書作成件数.xlsx"
    analyze_medical_documents(file_path, template_path)
