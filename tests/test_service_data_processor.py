import pytest
import datetime
import polars as pl
from unittest.mock import MagicMock

from service_data_processor import (
    process_cell_value,
    format_date_string,
    format_output_cell_value,
    parse_date_to_formats,
    filter_dataframe_by_date_range,
    clean_and_standardize_dataframe
)


class TestProcessCellValue:
    def test_none_value(self):
        assert process_cell_value(None) is None

    def test_no_column_attribute(self):
        cell = MagicMock()
        del cell.column
        assert process_cell_value(cell) is None

    def test_column_1_datetime(self):
        cell = MagicMock()
        cell.column = 1
        cell.value = datetime.datetime(2023, 5, 15)
        assert process_cell_value(cell) == "2023/05/15"

    def test_column_1_string(self):
        cell = MagicMock()
        cell.column = 1
        cell.value = "2023-05-15"
        assert process_cell_value(cell) == "2023-05-15"

    def test_column_2_numeric_string(self):
        cell = MagicMock()
        cell.column = 2
        cell.value = "12345"
        assert process_cell_value(cell) == 12345

    def test_column_2_non_numeric_string(self):
        cell = MagicMock()
        cell.column = 2
        cell.value = "ABC123"
        assert process_cell_value(cell) == "ABC123"

    def test_column_8_datetime(self):
        cell = MagicMock()
        cell.column = 8
        cell.value = datetime.datetime(2023, 6, 20)
        assert process_cell_value(cell) == "2023/06/20"

    def test_other_column(self):
        cell = MagicMock()
        cell.column = 5
        cell.value = "テストデータ"
        assert process_cell_value(cell) == "テストデータ"


class TestFormatDateString:
    def test_empty_string(self):
        assert format_date_string("") == ""

    def test_none_value(self):
        assert format_date_string(None) is None

    def test_hyphen_date(self):
        assert format_date_string("2023-05-15") == "2023/05/15"

    def test_slash_date(self):
        assert format_date_string("2023/05/15") == "2023/05/15"

    def test_datetime_with_time(self):
        assert format_date_string("2023-05-15 10:30:00") == "2023/05/15"

    def test_invalid_date_format(self):
        assert format_date_string("2023.05.15") == "2023.05.15"

    def test_numeric_input(self):
        assert format_date_string(20230515) == 20230515


class TestFormatOutputCellValue:
    def test_column_1_date(self):
        assert format_output_cell_value(1, "2023-05-15") == "2023/05/15"

    def test_column_8_date(self):
        assert format_output_cell_value(8, "2023-05-15") == "2023/05/15"

    def test_column_2_numeric_string(self):
        assert format_output_cell_value(2, "12345") == 12345

    def test_column_2_none(self):
        assert format_output_cell_value(2, None) is None

    def test_column_2_empty_string(self):
        assert format_output_cell_value(2, "") == ""

    def test_column_2_non_numeric(self):
        assert format_output_cell_value(2, "ABC") == "ABC"

    def test_other_column(self):
        assert format_output_cell_value(5, "テストデータ") == "テストデータ"


class TestParseDateToFormats:
    def test_none_value(self):
        result = parse_date_to_formats(None)
        assert result == {'raw': None, 'file_format': '', 'display_format': ''}

    def test_empty_string(self):
        result = parse_date_to_formats("  ")
        assert result == {'raw': '', 'file_format': '', 'display_format': ''}

    def test_valid_date_string(self):
        result = parse_date_to_formats("2023/05/15")
        assert result['raw'] == datetime.datetime(2023, 5, 15, 0, 0)
        assert result['file_format'] == "20230515"
        assert result['display_format'] == "2023年05月15日"

    def test_invalid_date_string(self):
        result = parse_date_to_formats("不正な日付")
        assert result['raw'] == "不正な日付"
        assert result['file_format'] == "不正な日付"
        assert result['display_format'] == "不正な日付"

    def test_datetime_object(self):
        date_obj = datetime.datetime(2023, 5, 15)
        result = parse_date_to_formats(date_obj)
        assert result['raw'] == date_obj
        assert result['file_format'] == "20230515"
        assert result['display_format'] == "2023年05月15日"


class TestFilterDataframeByDateRange:
    def test_empty_dataframe(self):
        empty_df = pl.DataFrame()
        result = filter_dataframe_by_date_range(empty_df)
        assert result['df'].height == 0
        assert result['start_date_display'] == '該当なし'
        assert result['end_date_display'] == '該当なし'
        assert result['file_date_range'] == 'no_data'

    def test_none_dataframe(self):
        result = filter_dataframe_by_date_range(None)
        assert result['df'].height == 0
        assert result['start_date_display'] == '該当なし'
        assert result['end_date_display'] == '該当なし'
        assert result['file_date_range'] == 'no_data'

    def test_no_date_column(self):
        df = pl.DataFrame({
            'ID': [1, 2, 3],
            'Name': ['A', 'B', 'C']
        })
        result = filter_dataframe_by_date_range(df)
        assert result['df'].equals(df)
        assert result['start_date_display'] == '該当なし'
        assert result['end_date_display'] == '該当なし'
        assert result['file_date_range'] == 'no_data'

    def test_with_date_range(self):
        df = pl.DataFrame({
            '預り日': ['2023/05/01', '2023/05/15', '2023/05/30'],
            'ID': [1, 2, 3],
            'Name': ['A', 'B', 'C']
        })
        result = filter_dataframe_by_date_range(df, '2023-05-10', '2023-05-20')
        assert result['df'].height == 1
        assert result['df']['預り日'][0] == '2023/05/15'

    def test_date_formats(self):
        df = pl.DataFrame({
            '預り日': ['2023/05/01', '2023/05/15', '2023/05/30'],
            'ID': [1, 2, 3]
        })
        result = filter_dataframe_by_date_range(df)
        assert result['start_date_display'] == '2023年05月01日'
        assert result['end_date_display'] == '2023年05月30日'
        assert result['file_date_range'] == '20230501-20230530'


class TestCleanAndStandardizeDataframe:
    def test_empty_dataframe(self):
        empty_df = pl.DataFrame()
        result = clean_and_standardize_dataframe(empty_df)
        assert result.height == 0

    def test_none_dataframe(self):
        result = clean_and_standardize_dataframe(None)
        assert result.height == 0

    def test_mixed_types(self):
        df = pl.DataFrame({
            'A': [1, None, 3],
            'B': ['X', None, 'Z'],
            'C': [True, False, None]
        })
        result = clean_and_standardize_dataframe(df)
        # すべての値が文字列型に変換され、Nullが空文字に変換されていることを確認
        assert result.dtypes == [pl.Utf8, pl.Utf8, pl.Utf8]
        assert result['A'].to_list() == ['1', '', '3']
        assert result['B'].to_list() == ['X', '', 'Z']
        assert result['C'].to_list() == ['true', 'false', '']
