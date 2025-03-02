import pytest
import polars as pl
from service_data_analyzer import TaskDataAnalyzer


@pytest.fixture
def sample_tasks():
    return [
        {'date': '2024-01-01', 'content': 'クラーク業務A', 'minutes': 30},
        {'date': '2024-01-01', 'content': 'クラーク業務B', 'minutes': 45},
        {'date': '2024-01-02', 'content': 'クラーク業務A', 'minutes': 25},
        {'date': '2024-01-02', 'content': '会議', 'minutes': 60},
        {'date': '2024-01-03', 'content': '資料作成', 'minutes': 90},
    ]


@pytest.fixture
def sample_daily_tasks():
    return [
        {'date': '2024-01-01', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-01', 'content': '毎日タスクB', 'minutes': 15},
        {'date': '2024-01-02', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-03', 'content': '毎日タスクB', 'minutes': 15},
    ]


@pytest.fixture
def sample_communication_tasks():
    return [
        {'date': '2024-01-01', 'name': '田中', 'content': '打合せ', 'minutes': 30},
        {'date': '2024-01-01', 'name': '佐藤', 'content': 'レビュー', 'minutes': 45},
        {'date': '2024-01-02', 'name': '田中', 'content': '打合せ', 'minutes': 25},
        {'date': '2024-01-02', 'name': '鈴木', 'content': '相談', 'minutes': 15},
        {'date': '2024-01-03', 'name': '佐藤', 'content': 'レビュー', 'minutes': 30},
    ]


@pytest.fixture
def sample_all_items():
    return [
        {'date': '2024-01-01', 'content': 'クラーク業務A', 'minutes': 30},
        {'date': '2024-01-01', 'content': '毎日タスクA', 'minutes': 10},
        {'date': '2024-01-02', 'content': '会議', 'minutes': 60},
        {'date': '2024-01-03', 'content': '資料作成', 'minutes': 90},
    ]


class TestTaskDataAnalyzer:
    def test_create_dataframes(self, sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items):
        # テスト実行
        df, daily_df, comm_df, all_items_df = TaskDataAnalyzer.create_dataframes(
            sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items
        )
        
        # 検証
        assert isinstance(df, pl.DataFrame)
        assert isinstance(daily_df, pl.DataFrame)
        assert isinstance(comm_df, pl.DataFrame)
        assert isinstance(all_items_df, pl.DataFrame)
        
        assert len(df) == len(sample_tasks)
        assert len(daily_df) == len(sample_daily_tasks)
        assert len(comm_df) == len(sample_communication_tasks)
        assert len(all_items_df) == len(sample_all_items)

    def test_aggregate_dataframe_without_filter(self, sample_tasks):
        # テスト準備
        df = pl.DataFrame(sample_tasks)
        
        # テスト実行
        result = TaskDataAnalyzer.aggregate_dataframe(df)
        
        # 検証
        assert isinstance(result, pl.DataFrame)
        assert 'content' in result.columns
        assert 'total_minutes' in result.columns
        assert 'total_hours' in result.columns
        assert 'frequency' in result.columns
        
        # 合計時間の確認
        assert result.filter(pl.col('content') == 'クラーク業務A').select('total_minutes')[0, 0] == 55
        
        # 頻度の確認
        assert result.filter(pl.col('content') == 'クラーク業務A').select('frequency')[0, 0] == 2

    def test_aggregate_dataframe_with_filter(self, sample_tasks):
        # テスト準備
        df = pl.DataFrame(sample_tasks)
        
        # テスト実行：クラーク業務のみフィルタリング
        result = TaskDataAnalyzer.aggregate_dataframe(
            df, filter_condition=pl.col('content').str.contains('クラーク業務')
        )
        
        # 検証
        assert len(result) == 2  # クラーク業務AとB
        assert set(result['content'].to_list()) == {'クラーク業務A', 'クラーク業務B'}

    def test_aggregate_dataframe_with_name_column(self, sample_communication_tasks):
        # テスト準備
        df = pl.DataFrame(sample_communication_tasks)
        
        # テスト実行
        result = TaskDataAnalyzer.aggregate_dataframe(df)
        
        # 検証
        assert 'name' in result.columns
        assert len(result) == 3  # 3つの組み合わせ
        
        # 田中+打合せの合計時間
        tanaka_meeting = result.filter((pl.col('name') == '田中') & (pl.col('content') == '打合せ'))
        assert tanaka_meeting.select('total_minutes')[0, 0] == 55

    def test_analyze_task_data(self, sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items):
        # テスト準備
        analyzer = TaskDataAnalyzer()
        
        # テスト実行
        results = analyzer.analyze_task_data(
            sample_tasks, sample_daily_tasks, sample_communication_tasks, sample_all_items
        )
        
        # 検証
        clerk_tasks, non_clerk_tasks, daily_tasks_agg, communication_by_name, communication_by_content, all_items_summary = results
        
        # クラーク業務の確認
        assert len(clerk_tasks) == 2
        
        # クラーク業務以外の確認
        assert len(non_clerk_tasks) == 2
        
        # デイリータスクの確認
        assert len(daily_tasks_agg) == 2
        
        # コミュニケーション（名前別）の確認
        assert len(communication_by_name) == 3
        
        # コミュニケーション（内容別）の確認
        assert len(communication_by_content) == 3
        
        # 全項目の確認
        assert len(all_items_summary) == 4
