from dataclasses import dataclass
from datetime import date
from pathlib import Path

from service.service_backup import backup_database_daily
from service.service_data_analyzer import count_documents
from service.service_database import (count_all_documents, count_documents_from, open_database,
                                      replace_documents_from)
from service.service_excel_reader import NoDataError, RowError, SourceData, read_documents
from service.service_excel_writer import save_results
from utils.config_manager import load_config

DEFAULT_KEEP_GENERATIONS = 30


def split_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(',') if item.strip()]


@dataclass
class ImportPlan:
    """取り込みの内容。確認のため、DB に書く前に画面へ見せる"""
    source: SourceData
    legacy: SourceData | None
    deleted_count: int

    @property
    def start_date(self) -> date:
        assert self.source.first_received_date is not None
        return self.source.first_received_date

    @property
    def errors(self) -> list[RowError]:
        legacy_errors = self.legacy.errors if self.legacy else []
        return legacy_errors + self.source.errors


class MedicalDocsAnalyzer:
    def prepare_import(self) -> ImportPlan:
        """元データを読み、預り日が最も古い日以降の DB の文書を入れ替える計画を返す。

        DB が空で旧 DB(xlsx)があれば、先に旧 DB の文書を取り込む計画にする。
        """
        # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
        paths_config = load_config()['PATHS']
        source = read_documents(paths_config['source_file_path'])
        if source.first_received_date is None:
            raise NoDataError
        start_date = source.first_received_date

        with open_database(paths_config['database_path']) as connection:
            is_empty = count_all_documents(connection) == 0
            deleted_count = count_documents_from(connection, start_date)

        legacy = None
        legacy_path = paths_config.get('legacy_database_path', '')
        if is_empty and legacy_path and Path(legacy_path).exists():
            legacy = read_documents(legacy_path)
            deleted_count = sum(document.received_date >= start_date for document in legacy.documents)
        return ImportPlan(source, legacy, deleted_count)

    def apply_import(self, plan: ImportPlan) -> int:
        """計画どおりに DB を書き換え、DB の件数を返す"""
        with open_database(load_config()['PATHS']['database_path']) as connection:
            if plan.legacy is not None and plan.legacy.first_received_date is not None:
                replace_documents_from(connection, plan.legacy.first_received_date, plan.legacy.documents)
            replace_documents_from(connection, plan.start_date, plan.source.documents)
            return count_all_documents(connection)

    def backup_database(self, today: date) -> Path | None:
        config = load_config()
        return backup_database_daily(
            config['PATHS']['database_path'],
            config['PATHS']['backup_dir'],
            config.getint('Backup', 'keep_generations', fallback=DEFAULT_KEEP_GENERATIONS),
            today,
        )

    def run_analysis(self, start_date: date, end_date: date) -> str:
        """期間内の文書を集計して Excel に保存し、保存先のパスを返す"""
        config = load_config()
        paths_config = config['PATHS']

        with open_database(paths_config['database_path']) as connection:
            counts = count_documents(connection, start_date, end_date)

        return save_results(
            counts,
            split_list(config['Analysis']['ordered_names']),
            split_list(config['Analysis']['clinical_departments']),
            paths_config['template_path'],
            paths_config['output_dir'],
        )
