from collections.abc import Mapping, Sequence

import polars as pl

# 0 件でも列を持たせるためにスキーマを指定する(日付は集計に使わないので読み込まない)
TASK_SCHEMA = {'content': pl.String, 'minutes': pl.Float64}
COMMUNICATION_SCHEMA = {**TASK_SCHEMA, 'name': pl.String}
CLERK_TASK_KEYWORD = 'クラーク業務'

Records = Sequence[Mapping[str, object]]


def summarize(df: pl.DataFrame, keys: list[str]) -> pl.DataFrame:
    """keys ごとに合計時間と回数を集計し、合計時間の降順に並べる"""
    return (
        df.group_by(keys)
        .agg([
            pl.col('minutes').sum().alias('total_minutes'),
            (pl.col('minutes').sum() / 60).round(1).alias('total_hours'),
            pl.col('minutes').count().alias('frequency'),
        ])
        .sort('total_minutes', descending=True)
    )


class TaskDataAnalyzer:
    def analyze_task_data(self, tasks: Records, daily_tasks: Records, comm_tasks: Records,
                          all_items: Records) -> tuple[pl.DataFrame, ...]:
        """service_excel_writer.SHEET_NAMES と同じ順に集計結果を返す"""
        df = pl.DataFrame(tasks, schema=TASK_SCHEMA)
        comm_df = pl.DataFrame(comm_tasks, schema=COMMUNICATION_SCHEMA)
        is_clerk = pl.col('content').str.contains(CLERK_TASK_KEYWORD)

        return (
            summarize(df.filter(is_clerk), ['content']),
            summarize(df.filter(~is_clerk), ['content']),
            summarize(pl.DataFrame(daily_tasks, schema=TASK_SCHEMA), ['content']),
            summarize(comm_df, ['name']),
            summarize(comm_df, ['content', 'name']).sort(['name', 'total_minutes'], descending=[False, True]),
            summarize(pl.DataFrame(all_items, schema=TASK_SCHEMA), ['content']),
        )
