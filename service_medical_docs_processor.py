import os

import polars as pl

from config_manager import load_config
from service_excel_handler import (
    backup_excel_file, read_excel_to_dataframe, write_dataframe_to_excel
)
from service_data_processor import (
    process_cell_value, format_output_cell_value, clean_and_standardize_dataframe
)


def process_medical_documents(source_file, target_file):
    try:
        config = load_config()
        backup_dir = config['PATHS']['backup_dir']

        source_df, headers = read_excel_to_dataframe(source_file, process_cell_value)

        if source_df.height == 0:
            print("エラー: ソースシートにデータがありません")
            return False

        if os.path.exists(target_file):
            target_df, target_headers = read_excel_to_dataframe(target_file, process_cell_value)
            
            if target_df.height > 0:
                # 文字列型に統一
                source_df = source_df.select([pl.col(col).cast(pl.Utf8) for col in source_df.columns])
                target_df = target_df.select([pl.col(col).cast(pl.Utf8) for col in target_df.columns])

                # 列名が同じことを確保
                if set(source_df.columns) == set(target_df.columns):
                    target_df = target_df.select(source_df.columns)
                    df = pl.concat([source_df, target_df])
                else:
                    print(f"警告: ソースとターゲットのカラム構造が異なります。")
                    print(f"ソース: {source_df.columns}")
                    print(f"ターゲット: {target_df.columns}")
                    print("ソースファイルのデータのみを使用します。")
                    df = source_df
            else:
                df = source_df
        else:
            df = source_df

        # データの前処理
        df = clean_and_standardize_dataframe(df)
        
        # 医師依頼日が空欄の行を削除
        if "医師依頼日" in df.columns:
            df = df.filter(pl.col("医師依頼日") != "")
            print(f"空の医師依頼日を持つ行を削除した後: {len(df)} 行")
        else:
            print("警告: '医師依頼日'の列が見つかりません。この条件でのフィルタリングをスキップします。")

        # 担当者名が空欄の行を削除
        if "担当者名" in df.columns:
            df = df.filter(pl.col("担当者名") != "")
            print(f"空の担当者名を持つ行を削除した後: {len(df)} 行")
        else:
            print("警告: '担当者名'の列が見つかりません。この条件でのフィルタリングをスキップします。")

        # 重複行を削除（預り日、患者ID、文書名、診療科、医師名の組み合わせが同じ行）
        required_columns = ["預り日", "患者ID", "文書名", "診療科", "医師名"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if not missing_columns:
            df = df.unique(subset=required_columns)
        else:
            print(f"警告: 次の列が見つからないため、重複削除をスキップします: {missing_columns}")
            # 存在する列のみで重複削除を試みる
            existing_columns = [col for col in required_columns if col in df.columns]
            if existing_columns:
                print(f"代わりに次の列で重複削除を行います: {existing_columns}")
                df = df.unique(subset=existing_columns)

        print(f"重複削除後: {len(df)} 行")

        success = write_dataframe_to_excel(
            df, target_file, headers, 
            create_new=not os.path.exists(target_file),
            format_func=format_output_cell_value
        )

        if success:
            # 処理後のファイルをバックアップ
            backup_result = backup_excel_file(target_file, backup_dir)
            if backup_result:
                print(f"ファイルをバックアップしました: {backup_result}")
            else:
                print("警告: ファイルのバックアップに失敗しました。")

            print(f"処理完了: {len(df)} 行のデータを {target_file} に保存しました")
            return True
        else:
            print("エラー: ファイルの書き込みに失敗しました")
            return False

    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        return False

