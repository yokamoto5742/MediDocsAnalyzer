import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime
from service_task_analyzer import TaskAnalyzer


class TestTaskAnalyzer:
    @patch('service_task_analyzer.load_config')
    @patch('service_task_analyzer.ExcelTaskReader')
    @patch('service_task_analyzer.TaskDataAnalyzer')
    @patch('service_task_analyzer.ExcelResultWriter')
    def test_initialization(self, mock_writer, mock_analyzer, mock_reader, mock_load_config):
        # モックの設定
        mock_config = {
            'PATHS': {
                'input_file_path': 'test_input.xlsx',
                'template_path': 'test_template.xlsx',
                'output_dir': 'test_output'
            }
        }
        mock_load_config.return_value = mock_config
        
        # テスト実行
        analyzer = TaskAnalyzer()
        
        # 検証
        assert analyzer.config == mock_config
        assert analyzer.paths_config == mock_config['PATHS']
        mock_reader.assert_called_once_with(mock_config)
        mock_analyzer.assert_called_once()
        mock_writer.assert_called_once()

    @patch('service_task_analyzer.load_config')
    @patch('service_task_analyzer.ExcelTaskReader')
    @patch('service_task_analyzer.TaskDataAnalyzer')
    @patch('service_task_analyzer.ExcelResultWriter')
    def test_run_analysis_success(self, mock_writer_class, mock_analyzer_class, mock_reader_class, mock_load_config):
        # モックの設定
        mock_config = {
            'PATHS': {
                'input_file_path': 'test_input.xlsx',
                'template_path': 'test_template.xlsx',
                'output_dir': 'test_output'
            }
        }
        mock_load_config.return_value = mock_config
        
        # リーダーのモック
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        mock_reader.read_workbook.return_value = (
            [{'content': 'タスク1', 'minutes': 30}],  # tasks
            [{'content': '日次タスク1', 'minutes': 10}],  # daily_tasks
            [{'name': '田中', 'content': '打合せ', 'minutes': 30}],  # comm_tasks
            [{'content': 'すべて1', 'minutes': 30}],  # all_items
            '20240101',  # actual_start_date_str
            '20240105'   # actual_end_date_str
        )
        
        # アナライザーのモック
        mock_analyzer = MagicMock()
        mock_analyzer_class.return_value = mock_analyzer
        mock_analyzer.analyze_task_data.return_value = (
            'clerk_tasks',
            'non_clerk_tasks',
            'daily_tasks_agg',
            'communication_by_name',
            'communication_by_content',
            'all_items_summary'
        )
        
        # ライターのモック
        mock_writer = MagicMock()
        mock_writer_class.return_value = mock_writer
        mock_writer.save_results.return_value = 'output_file_path.xlsx'
        
        # テスト実行
        analyzer = TaskAnalyzer()
        success, message = analyzer.run_analysis('2024-01-01', '2024-01-05')
        
        # 検証
        assert success is True
        assert '分析が完了しました' in message
        assert 'output_file_path.xlsx' in message
        
        # 各メソッドが正しく呼び出されたことを検証
        mock_reader.read_workbook.assert_called_once_with(
            'test_input.xlsx',
            datetime(2024, 1, 1),
            datetime(2024, 1, 5)
        )
        
        mock_analyzer.analyze_task_data.assert_called_once_with(
            [{'content': 'タスク1', 'minutes': 30}],
            [{'content': '日次タスク1', 'minutes': 10}],
            [{'name': '田中', 'content': '打合せ', 'minutes': 30}],
            [{'content': 'すべて1', 'minutes': 30}]
        )
        
        mock_writer.save_results.assert_called_once_with(
            (
                'clerk_tasks',
                'non_clerk_tasks',
                'daily_tasks_agg',
                'communication_by_name',
                'communication_by_content',
                'all_items_summary'
            ),
            'test_template.xlsx',
            'test_output',
            datetime(2024, 1, 1),
            datetime(2024, 1, 5)
        )

    @patch('service_task_analyzer.load_config')
    def test_run_analysis_date_format_error(self, mock_load_config):
        # モックの設定
        mock_config = {'PATHS': {}}
        mock_load_config.return_value = mock_config
        
        # テスト実行 - 不正な日付形式
        analyzer = TaskAnalyzer()
        success, message = analyzer.run_analysis('2024/01/01', '2024-01-05')
        
        # 検証
        assert success is False
        assert '日付の形式が正しくありません' in message

    @patch('service_task_analyzer.load_config')
    @patch('service_task_analyzer.ExcelTaskReader')
    def test_run_analysis_general_exception(self, mock_reader_class, mock_load_config):
        # モックの設定
        mock_config = {
            'PATHS': {
                'input_file_path': 'test_input.xlsx',
                'template_path': 'test_template.xlsx',
                'output_dir': 'test_output'
            }
        }
        mock_load_config.return_value = mock_config
        
        # リーダーが例外を発生させるようにモックする
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        mock_reader.read_workbook.side_effect = Exception('テストエラー')
        
        # テスト実行
        analyzer = TaskAnalyzer()
        success, message = analyzer.run_analysis('2024-01-01', '2024-01-05')
        
        # 検証
        assert success is False
        assert '分析中にエラーが発生しました' in message
        assert 'テストエラー' in message
