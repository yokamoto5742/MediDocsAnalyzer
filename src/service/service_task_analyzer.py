from datetime import date

from service.service_data_analyzer import TaskDataAnalyzer
from service.service_excel_reader import ExcelTaskReader
from service.service_excel_writer import ExcelResultWriter
from utils.config_manager import load_config


class TaskAnalyzer:
    def run_analysis(self, start_date: date, end_date: date) -> str:
        """期間内のデータを集計して Excel に保存し、保存先のパスを返す"""
        # 起動後に設定ファイルが編集されていても反映されるよう、実行のたびに読み込む
        config = load_config()
        paths_config = config['PATHS']

        workbook_data = ExcelTaskReader(config).read_workbook(
            paths_config['input_file_path'], start_date, end_date)

        analysis_results = TaskDataAnalyzer().analyze_task_data(
            workbook_data.tasks,
            workbook_data.daily_tasks,
            workbook_data.communication_tasks,
            workbook_data.all_items
        )

        return ExcelResultWriter.save_results(
            analysis_results,
            paths_config['template_path'],
            paths_config['output_dir'],
            workbook_data.first_date,
            workbook_data.last_date
        )
