import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

from service import service_auth, service_master
from service.service_auth import LoginResult
from service.service_backup import backup_database_daily
from service.service_data_analyzer import DocumentCounts, StaffCounts, count_documents, count_staff_documents
from service.service_database import (count_all_documents, count_documents_between, open_database,
                                      replace_documents_between)
from service.service_excel_reader import Draft, NoDataError, RowError, SourceData, read_documents
from service.service_excel_writer import TOTAL_LABEL, save_results, save_staff_results
from service.service_master import Department, Staff
from utils.config_manager import load_config

DEFAULT_KEEP_GENERATIONS = 30
TOP_DOCUMENT_COUNT = 10
DEFAULT_PENDING_DAYS = 7


def split_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(',') if item.strip()]


@dataclass
class Overview:
    """預り日が start_date〜end_date の下書き(担当者名か医師依頼日が空の行)"""
    start_date: date
    end_date: date
    drafts: list[Draft]

    @property
    def unassigned_count(self) -> int:
        return sum(1 for draft in self.drafts if not draft.staff_name)

    @property
    def unrequested_count(self) -> int:
        return sum(1 for draft in self.drafts if draft.requested_date is None)


def make_overview(drafts: list[Draft], today: date, days: int) -> Overview:
    """本日を含めて days 日間に預かった下書きを、預り日の新しい順に並べる"""
    start_date = today - timedelta(days=days - 1)
    targets = [draft for draft in drafts if start_date <= draft.received_date <= today]
    targets.sort(key=lambda draft: draft.received_date, reverse=True)
    return Overview(start_date, today, targets)


@dataclass
class ImportPlan:
    """取り込みの内容。確認のため、DB に書く前に画面へ見せる"""
    source: SourceData
    deleted_count: int

    @property
    def start_date(self) -> date:
        assert self.source.first_received_date is not None
        return self.source.first_received_date

    @property
    def end_date(self) -> date:
        assert self.source.last_received_date is not None
        return self.source.last_received_date

    @property
    def errors(self) -> list[RowError]:
        return self.source.errors


@dataclass
class OverallSummary:
    counts: DocumentCounts
    staff_members: list[str]
    departments: list[str]  # 先頭は合計列


class MedicalDocsAnalyzer:
    @contextmanager
    def _open_database(self) -> Iterator[sqlite3.Connection]:
        """DB を開き、マスタが空なら設定ファイルの担当者・診療科を入れる"""
        # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
        config = load_config()
        with open_database(config['PATHS']['database_path']) as connection:
            departments = [name for name in split_list(config.get('Analysis', 'clinical_departments', fallback=''))
                           if name != TOTAL_LABEL]
            service_master.seed_masters(
                connection, split_list(config.get('Analysis', 'ordered_names', fallback='')), departments)
            yield connection

    def prepare_source_import(self) -> ImportPlan:
        """元データ(xlsm)を取り込む計画を返す"""
        return self._prepare_import(load_config()['PATHS']['source_file_path'])

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
        if source.first_received_date is None or source.last_received_date is None:
            raise NoDataError

        with self._open_database() as connection:
            deleted_count = count_documents_between(connection, source.first_received_date,
                                                    source.last_received_date)
        return ImportPlan(source, deleted_count)

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
        source = read_documents(load_config()['PATHS']['source_file_path'])
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

    def overall_summary(self, start_date: date, end_date: date) -> OverallSummary:
        """期間内の文書を担当者×診療科で集計する。期間内に文書がなければ NoDataError"""
        with self._open_database() as connection:
            counts = count_documents(connection, start_date, end_date)
            staff_list = service_master.list_staff(connection)
            departments = service_master.list_departments(connection)

        # 退職した担当者も、期間内に件数があれば並べる
        counted_staff = {staff for staff, _ in counts.by_staff_department}
        staff_members = [staff.name for staff in staff_list if staff.active or staff.name in counted_staff]
        return OverallSummary(counts, staff_members, [TOTAL_LABEL, *(department.name for department in departments)])

    def staff_summary(self, staff_name: str, start_date: date, end_date: date) -> StaffCounts:
        with self._open_database() as connection:
            return count_staff_documents(connection, staff_name, start_date, end_date, TOP_DOCUMENT_COUNT)

    def export_staff(self, staff_name: str, start_date: date, end_date: date) -> str:
        """1 人分の集計を Excel に保存し、保存先のパスを返す"""
        output_dir = load_config()['PATHS']['output_dir']
        return save_staff_results(self.staff_summary(staff_name, start_date, end_date), start_date, end_date,
                                  output_dir)

    def run_analysis(self, start_date: date, end_date: date) -> str:
        """期間内の文書を集計して Excel に保存し、保存先のパスを返す"""
        paths_config = load_config()['PATHS']
        summary = self.overall_summary(start_date, end_date)
        return save_results(
            summary.counts,
            summary.staff_members,
            summary.departments,
            paths_config['template_path'],
            paths_config['output_dir'],
        )
