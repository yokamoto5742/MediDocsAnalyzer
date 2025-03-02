from datetime import datetime
from config_manager import load_config
from service_excel_reader import ExcelTaskReader
from service_data_analyzer import TaskDataAnalyzer
from service_excel_writer import ExcelResultWriter


class TaskAnalyzer:
    def __init__(self):
        self.config = load_config()
        self.paths_config = self.config['PATHS']
        self.reader = ExcelTaskReader(self.config)
        self.analyzer = TaskDataAnalyzer()
        self.writer = ExcelResultWriter()

    def run_analysis(self, start_date_str, end_date_str):
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

            tasks, daily_tasks, comm_tasks, all_items, actual_start_date_str, actual_end_date_str = self.reader.read_workbook(
                self.paths_config['input_file_path'],
                start_date,
                end_date
            )

            actual_start_date = datetime.strptime(actual_start_date_str, '%Y%m%d')
            actual_end_date = datetime.strptime(actual_end_date_str, '%Y%m%d')

            analysis_results = self.analyzer.analyze_task_data(
                tasks, daily_tasks, comm_tasks, all_items
            )

            output_file = self.writer.save_results(
                analysis_results,
                self.paths_config['template_path'],
                self.paths_config['output_dir'],
                actual_start_date,
                actual_end_date
            )

            return True, f"分析が完了しました。結果は {output_file} に保存されました。"

        except ValueError as ve:
            return False, f"日付の形式が正しくありません: {str(ve)}"
        except Exception as e:
            return False, f"分析中にエラーが発生しました: {str(e)}"
