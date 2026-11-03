import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from service import service_auth, service_master
from service.models import Draft, NoDataError, SourceData
from service.service_auth import LoginResult
from service.service_backup import backup_database_daily
from service.service_data_analyzer import (TOTAL_LABEL, DocumentFilter, DocumentType, OverallSummary, StaffCounts, TrendCounts,
                                           count_documents, count_staff_documents, count_trend)
from service.service_database import (count_all_documents, count_documents_between, open_database,
                                      replace_documents_between)
from service.service_excel_reader import find_latest_source, read_documents
from service.service_excel_writer import save_results, save_staff_results, save_trend_results
from service.service_master import Department, Staff
from service.service_overview import DEFAULT_PENDING_DAYS, Overview, make_overview
from utils.config_manager import load_config

DEFAULT_KEEP_GENERATIONS = 30
TOP_DOCUMENT_COUNT = 10
DEFAULT_SOURCE_DIR = r'C:\Shinseikai\MediDocsAnalyzer'


def split_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(',') if item.strip()]


def source_file_path() -> str:
    """設定した元データのフォルダで、ファイル名の日時が最も新しい xlsm"""
    return str(find_latest_source(load_config().get('PATHS', 'source_dir', fallback=DEFAULT_SOURCE_DIR)))


@dataclass
class ImportPlan:
    """取り込みの内容。確認のため、DB に書く前に画面へ見せる"""
    source: SourceData
    start_date: date
    end_date: date
    deleted_count: int


class MedicalDocsAnalyzer:
    @contextmanager
    def _open_database(self) -> Iterator[sqlite3.Connection]:
        """DB を開き、マスタが空なら設定ファイルの担当者・診療科を入れる"""
        # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
        config = load_config()
        with open_database(config['PATHS']['database_path']) as connection:
            # 旧版の設定ファイルは clinical_departments の先頭に「合計」があるため除く
            departments =[name for name in split_list(config.get('Analysis', 'clinical_departments', fallback=''))
                           if name != TOTAL_LABEL]
            service_master.seed_masters(
                connection, split_list(config.get('Analysis', 'ordered_names', fallback='')), departments)
            yield connection

    def prepare_source_import(self) -> ImportPlan:
        """元データ(xlsm)を取り込む計画を返す"""
        return self._prepare_import(source_file_path())

    def prepare_legacy_import(self) -> ImportPlan:
        """旧 DB(xlsx)を取り込む計画を返す"""
        return self._prepare_import(load_config()['PATHS']['legacy_database_path'])

    def _prepare_import(self, file_path: str) -> ImportPlan:
        """file_path を読み、その預り日の範囲(最も古い日〜新しい日)にある DB の文書を入れ替える計画を返す。

        読んだファイルが古いコピーでも、それより新しい DB の文書は消さない。
        """
        with self._open_database() as connection:
            known_staff = {staff.name for staff in service_master.list_staff(connection)}

        source = read_documents(file_path, known_staff)
        start_date, end_date = source.first_received_date, source.last_received_date
        if start_date is None or end_date is None:
            raise NoDataError

        with self._open_database() as connection:
            deleted_count = count_documents_between(connection, start_date, end_date)
        return ImportPlan(source, start_date, end_date, deleted_count)

    def apply_import(self, plan: ImportPlan) -> int:
        """計画どおりに DB を書き換え、DB の件数を返す"""
        with self._open_database() as connection:
            replace_documents_between(connection, plan.start_date, plan.end_date, plan.source.documents)
            return count_all_documents(connection)

    def overview(self, drafts: list[Draft], today: date) -> Overview:
        days = load_config().getint('Overview', 'pending_days', fallback=DEFAULT_PENDING_DAYS)
        return make_overview(drafts, today, days)

    def load_overview(self, today: date) -> Overview:
        """元データを読み、概要を返す。下書きの判定にマスタは使わないため DB は開かない"""
        source = read_documents(source_file_path())
        return self.overview(source.drafts, today)

    def backup_database(self, today: date) -> Path | None:
        config = load_config()
        return backup_database_daily(
            config['PATHS']['database_path'],
            config['PATHS']['backup_dir'],
            config.getint('Backup', 'keep_generations', fallback=DEFAULT_KEEP_GENERATIONS),
            today,
        )

    def list_staff(self) -> list[Staff]:
        with self._open_database() as connection:
            return service_master.list_staff(connection)

    def save_staff(self, staff: Staff) -> None:
        with self._open_database() as connection:
            service_master.save_staff(connection, staff)

    def login_required(self) -> bool:
        with self._open_database() as connection:
            return service_auth.login_required(connection)

    def authenticate(self, emr_id: str, password: str) -> LoginResult | None:
        with self._open_database() as connection:
            return service_auth.authenticate(connection, emr_id, password)

    def change_password(self, staff_id: int, password: str) -> None:
        with self._open_database() as connection:
            service_auth.set_password(connection, staff_id, password, must_change=False)

    def issue_temporary_password(self, staff_id: int) -> str:
        """仮パスワードを設定して返す。次のログインで変更を求める"""
        password = service_auth.make_temporary_password()
        with self._open_database() as connection:
            service_auth.set_password(connection, staff_id, password, must_change=True)
        return password

    def list_departments(self) -> list[Department]:
        with self._open_database() as connection:
            return service_master.list_departments(connection)

    def save_department(self, old_name: str | None, department: Department) -> None:
        with self._open_database() as connection:
            service_master.save_department(connection, old_name, department)

    def delete_department(self, name: str) -> None:
        with self._open_database() as connection:
            service_master.delete_department(connection, name)

    def overall_summary(self, start_date: date, end_date: date,
                        document_type: DocumentType = DocumentType.ALL) -> OverallSummary:
        """期間内で集計対象の文書を担当者×診療科で集計する。期間内に文書がなければ NoDataError"""
        with self._open_database() as connection:
            counts = count_documents(connection, start_date, end_date, document_type)
            staff_list = service_master.list_staff(connection)
            departments = service_master.list_departments(connection)

        # 退職した担当者も、期間内に件数があれば並べる
        counted_staff = {staff for staff, _ in counts.by_staff_department}
        staff_members = [staff.name for staff in staff_list if staff.active or staff.name in counted_staff]
        return OverallSummary(counts, staff_members, [department.name for department in departments],
                              document_type)

    def staff_summary(self, staff_name: str, start_date: date, end_date: date,
                      document_limit: int | None = None,
                      document_type: DocumentType = DocumentType.ALL) -> StaffCounts:
        """画面ではスクロールで全件を見るため、文書名は既定で全件を返す"""
        with self._open_database() as connection:
            return count_staff_documents(connection, staff_name, start_date, end_date, document_limit, document_type)

    def export_staff(self, staff_name: str, start_date: date, end_date: date,
                     document_type: DocumentType = DocumentType.ALL) -> str:
        """1 人分の集計を Excel に保存し、保存先のパスを返す。A4 1 枚に収めるため文書名は上位だけ出す"""
        output_dir = load_config()['PATHS']['output_dir']
        staff_counts = self.staff_summary(staff_name, start_date, end_date, TOP_DOCUMENT_COUNT, document_type)
        return save_staff_results(staff_counts, start_date, end_date, output_dir)

    def trend_summary(self, document_filter: DocumentFilter, start_date: date, end_date: date, year_start_month: int,
                      today: date, staff_name: str | None = None) -> TrendCounts:
        """医師依頼日が範囲内で document_filter に合う DB の文書を、年(年度)ごとの月別に数える。
        staff_name が None なら全担当者"""
        with self._open_database() as connection:
            return count_trend(connection, document_filter, start_date, end_date, year_start_month, today,
                               staff_name)

    def export_trend(self, subject: str, document_filter: DocumentFilter, start_date: date, end_date: date,
                     year_start_month: int, today: date, staff_name: str | None = None) -> str:
        """全体推移(staff_name を渡せば個人推移)を Excel に保存し、保存先のパスを返す。subject は表題に出す対象"""
        trend = self.trend_summary(document_filter, start_date, end_date, year_start_month, today, staff_name)
        return save_trend_results(trend, subject, start_date, end_date, load_config()['PATHS']['output_dir'])

    def run_analysis(self, start_date: date, end_date: date,
                     document_type: DocumentType = DocumentType.ALL) -> str:
        """期間内で集計対象の文書を集計して Excel に保存し、保存先のパスを返す"""
        paths_config = load_config()['PATHS']
        summary = self.overall_summary(start_date, end_date, document_type)
        return save_results(summary, paths_config['template_path'], paths_config['output_dir'])
