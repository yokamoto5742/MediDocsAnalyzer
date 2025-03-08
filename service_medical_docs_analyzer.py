from datetime import datetime

from config_manager import load_config
from service_analyze_medical_documents import analyze_medical_documents


class MedicalDocsAnalyzer:
    def __init__(self):
        self.config = load_config()
        self.paths_config = self.config['PATHS']

    def run_analysis(self, start_date_str, end_date_str):
        try:
            database_path = self.paths_config['database_path']
            template_path = self.paths_config['template_path']

            analyze_medical_documents(database_path, template_path, start_date_str, end_date_str)

            return True, "集計が完了しました。"

        except ValueError as ve:
            return False, f"日付の形式が正しくありません: {str(ve)}"
        except Exception as e:
            return False, f"分析中にエラーが発生しました: {str(e)}"
