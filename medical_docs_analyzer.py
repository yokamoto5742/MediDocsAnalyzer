import openpyxl
import polars as pl
import warnings
from datetime import datetime
import os
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

warnings.simplefilter('ignore')


def analyze_medical_documents(file_path, template_path=None):
    """
    医療文書の作成件数を担当者・診療科別に集計してExcelファイルに出力する関数

    Parameters:
        file_path (str): 分析対象のExcelファイルパス
        template_path (str, optional): テンプレートファイルのパス。指定がない場合は新規作成
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
            # 文字列の場合はdatetimeに変換を試みる
            try:
                min_date_obj = datetime.strptime(min_date, '%Y-%m-%d')
                max_date_obj = datetime.strptime(max_date, '%Y-%m-%d')
                start_date = min_date_obj.strftime('%Y年%m月%d日')
                end_date = max_date_obj.strftime('%Y年%m月%d日')
                filename_start = min_date_obj.strftime('%Y%m%d')
                filename_end = max_date_obj.strftime('%Y%m%d')
            except:
                # 変換できない場合はそのまま使用
                start_date = min_date
                end_date = max_date
                filename_start = start_date
                filename_end = end_date
        else:
            # datetimeオブジェクトの場合はstrftimeを使用
            start_date = min_date.strftime('%Y年%m月%d日')
            end_date = max_date.strftime('%Y年%m月%d日')
            filename_start = min_date.strftime('%Y%m%d')
            filename_end = max_date.strftime('%Y%m%d')

        # 担当者と診療科で集計
        grouped = df.filter(
            pl.col('担当者名').is_not_null() & pl.col('診療科').is_not_null()
        ).group_by(['担当者名', '診療科']).agg(
            pl.count().alias('作成件数')
        )

        # 担当者のリストを取得（None値を除外）
        staff_members = df.select('担当者名').filter(
            pl.col('担当者名').is_not_null()
        ).unique().sort('担当者名').to_series().to_list()

        # 診療科のリストを取得（None値を除外）
        departments = df.select('診療科').filter(
            pl.col('診療科').is_not_null()
        ).unique().sort('診療科').to_series().to_list()

        # 結果をExcelファイルに書き込む
        if template_path and os.path.exists(template_path):
            # テンプレートファイルをベースに作成
            result_wb = openpyxl.load_workbook(template_path)
            result_ws = result_wb.active
        else:
            # 新規にワークブックを作成
            result_wb = openpyxl.Workbook()
            result_ws = result_wb.active

        # シートをクリア（テンプレート使用時も内容をクリア）
        for row in result_ws.iter_rows():
            for cell in row:
                cell.value = None

        # タイトル行の設定
        title = f"医療文書作成件数 {start_date}-{end_date}"
        result_ws['A1'] = title
        result_ws.merge_cells('A1:L1')

        # タイトルのスタイル設定
        title_cell = result_ws['A1']
        title_cell.font = Font(size=12, bold=True)
        title_cell.alignment = Alignment(horizontal='left', vertical='center')

        # ヘッダー行の設定
        result_ws['A2'] = '氏名'
        result_ws['B2'] = '内科'

        # 診療科をヘッダーに設定
        col_idx = 3  # C列から開始
        department_columns = {}  # 診療科と列の対応

        # テンプレートファイルの診療科順を維持
        if template_path and os.path.exists(template_path):
            template_wb = openpyxl.load_workbook(template_path, read_only=True)
            template_ws = template_wb.active

            # 2行目からヘッダーを読み取り
            template_depts = []
            for cell in template_ws[2][1:]:  # B列以降
                if cell.value and cell.value != '氏名' and cell.value != '合計':
                    template_depts.append(cell.value)
                    col_letter = openpyxl.utils.get_column_letter(cell.column)
                    result_ws[f'{col_letter}2'] = cell.value
                    department_columns[cell.value] = col_letter

            # テンプレートになかった診療科を追加
            for dept in departments:
                if dept not in template_depts and dept != '内科':  # 内科はB列で固定
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    result_ws[f'{col_letter}2'] = dept
                    department_columns[dept] = col_letter
                    col_idx += 1
        else:
            # 新規作成の場合は診療科をソートして配置
            for dept in departments:
                if dept != '内科':  # 内科はB列で固定
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    result_ws[f'{col_letter}2'] = dept
                    department_columns[dept] = col_letter
                    col_idx += 1

        # 内科の列を設定
        department_columns['内科'] = 'B'

        # 合計列を最後に配置
        total_col_letter = openpyxl.utils.get_column_letter(col_idx)
        result_ws[f'{total_col_letter}2'] = '合計'

        # ヘッダー行のスタイル設定
        header_style = Font(bold=True)
        for cell in result_ws[2]:
            cell.font = header_style
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

        # データの書き込み
        row_idx = 3  # 3行目からデータ開始

        for staff in staff_members:
            # 担当者名を設定
            result_ws[f'A{row_idx}'] = staff

            # 各診療科の件数を取得
            staff_total = 0

            for dept, col_letter in department_columns.items():
                # この担当者の特定診療科のデータを取得
                dept_data = grouped.filter(
                    (pl.col('担当者名') == staff) & (pl.col('診療科') == dept)
                )

                if dept_data.height > 0:
                    count = dept_data.select('作成件数').item()
                    result_ws[f'{col_letter}{row_idx}'] = count
                    staff_total += count

            # 合計を設定
            result_ws[f'{total_col_letter}{row_idx}'] = staff_total

            row_idx += 1

        # 合計行の追加
        result_ws[f'A{row_idx}'] = '合計'

        # 各診療科の合計を計算
        for dept, col_letter in department_columns.items():
            # 特定診療科の全担当者合計を計算
            dept_data = grouped.filter(pl.col('診療科') == dept)
            if dept_data.height > 0:
                total_count = dept_data.select('作成件数').sum().item()
                result_ws[f'{col_letter}{row_idx}'] = total_count

        # 総合計を計算（全ての担当者、全ての診療科）
        grand_total = grouped.select('作成件数').sum().item()
        result_ws[f'{total_col_letter}{row_idx}'] = grand_total

        # データセルのスタイル設定
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for row in range(3, row_idx + 1):
            for col in range(1, col_idx + 1):
                cell = result_ws.cell(row=row, column=col)
                cell.border = thin_border
                # 数値のセルは右寄せ
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left')

        # 合計行のスタイル強調
        for col in range(1, col_idx + 1):
            cell = result_ws.cell(row=row_idx, column=col)
            cell.font = Font(bold=True)

        # 列幅の調整
        for col in range(1, col_idx + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            result_ws.column_dimensions[col_letter].width = 15

        # A列（氏名）の幅を調整
        result_ws.column_dimensions['A'].width = 10

        # 出力ファイル名の設定
        output_filename = f"医療文書作成件数{filename_start}-{filename_end}.xlsx"
        result_wb.save(output_filename)

        print(f"集計結果を '{output_filename}' に保存しました。")
        return output_filename

    except FileNotFoundError:
        print(f"エラー: ファイル '{file_path}' が見つかりません。")
    except Exception as e:
        print(f"エラー: データの処理中に問題が発生しました - {str(e)}")


# 使用例
if __name__ == "__main__":
    input_file = "医療文書担当一覧データベース.xlsx"
    template_file = "医療文書作成件数.xlsx"

    if os.path.exists(template_file):
        print(f"テンプレート '{template_file}' を使用して集計します。")
        analyze_medical_documents(input_file, template_file)
    else:
        print(f"テンプレート '{template_file}' が見つからないため、新規に作成します。")
        analyze_medical_documents(input_file)