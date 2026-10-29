from datetime import date
from pathlib import Path

from service.service_data_analyzer import count_documents, filter_by_period, merge_documents
from service.service_excel_reader import NoDataError, read_documents
from service.service_excel_writer import backup_file, save_results, write_database
from utils.config_manager import load_config


def split_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(',') if item.strip()]


class MedicalDocsAnalyzer:
    def load_data(self) -> int:
        """元データを DB に取り込んでバックアップし、DB の件数を返す"""
        # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
        paths_config = load_config()['PATHS']
        database_path = paths_config['database_path']

        source_df = read_documents(paths_config['source_file_path'])
        if source_df.height == 0:
            raise NoDataError

        database_df = read_documents(database_path) if Path(database_path).exists() else None
        documents = merge_documents(source_df, database_df)

        write_database(documents, database_path)
        backup_file(database_path, paths_config['backup_dir'])
        return documents.height

    def run_analysis(self, start_date: date, end_date: date) -> str:
        """期間内の文書を集計して Excel に保存し、保存先のパスを返す"""
        config = load_config()
        paths_config = config['PATHS']

        documents = filter_by_period(read_documents(paths_config['database_path']), start_date, end_date)
        if documents.height == 0:
            raise NoDataError

        return save_results(
            count_documents(documents),
            split_list(config['Analysis']['ordered_names']),
            split_list(config['Analysis']['clinical_departments']),
            paths_config['template_path'],
            paths_config['output_dir'],
        )
