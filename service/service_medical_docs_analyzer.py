import sqlite3
from collections.abc import Iterator
from contextlib import AbstractContextManager, contextmanager
from datetime import date

from service import service_auth, service_master
from service.service_auth import LoginResult
from service.service_data_analyzer import (TOTAL_LABEL, DocumentFilter, DocumentType, OverallSummary, StaffCounts, TrendCounts,
                                           count_documents, count_staff_documents, count_trend)
from service.service_database import open_database, open_documents
from service.service_excel_writer import save_results, save_staff_results, save_trend_results
from service.service_master import Department, Staff
from service.service_overview import DEFAULT_PENDING_DAYS, Overview, load_overview
from utils.config_manager import load_config

TOP_DOCUMENT_COUNT = 10
# CSV2XL が作る文書の DB
DEFAULT_DOCUMENTS_DATABASE_PATH = r'C:\Shinseikai\CSV2XL\医療文書担当一覧.db'


def split_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(',') if item.strip()]


def open_documents_database() -> AbstractContextManager[sqlite3.Connection]:
    """設定した CSV2XL の文書の DB を読み取り専用で開く"""
    return open_documents(load_config().get('PATHS', 'documents_database_path',
                                            fallback=DEFAULT_DOCUMENTS_DATABASE_PATH))


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

    def load_overview(self, today: date) -> Overview:
        """CSV2XL の DB から下書きを読み、概要を返す。下書きの判定にマスタは使わないためマスタの DB は開かない"""
        days = load_config().getint('Overview', 'pending_days', fallback=DEFAULT_PENDING_DAYS)
        with open_documents_database() as connection:
            return load_overview(connection, today, days)

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
        with open_documents_database() as connection:
            counts = count_documents(connection, start_date, end_date, document_type)
        with self._open_database() as connection:
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
        with open_documents_database() as connection:
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
        with open_documents_database() as connection:
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
        summary = self.overall_summary(start_date, end_date, document_type)
        return save_results(summary, load_config()['PATHS']['output_dir'])
