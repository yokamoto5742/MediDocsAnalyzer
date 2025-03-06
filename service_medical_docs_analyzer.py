from datetime import datetime
from config_manager import load_config
from analyze_medical_documents import analyze_medical_documents


class MedicalDocsAnalyzer:
    def __init__(self):
        self.config = load_config()
        self.paths_config = self.config['PATHS']

    def run_analysis(self, start_date_str, end_date_str):
        try:
            # 設定ファイルからパスを取得
            database_path = self.paths_config['database_path']
            template_path = self.paths_config['template_path']

            # 医療文書の分析を実行
            # 開始日と終了日のパラメータを追加
            analyze_medical_documents(database_path, template_path, start_date_str, end_date_str)

            return True, "集計が完了しました。"

        except ValueError as ve:
            return False, f"日付の形式が正しくありません: {str(ve)}"
        except Exception as e:
            return False, f"分析中にエラーが発生しました: {str(e)}"
