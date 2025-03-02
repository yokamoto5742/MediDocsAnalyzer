import polars as pl


class TaskDataAnalyzer:
    @staticmethod
    def create_dataframes(tasks, daily_tasks, communication_tasks, all_items):
        df = pl.DataFrame(tasks)
        daily_df = pl.DataFrame(daily_tasks)
        comm_df = pl.DataFrame(communication_tasks)
        all_items_df = pl.DataFrame(all_items)
        
        return df, daily_df, comm_df, all_items_df

    @staticmethod
    def aggregate_dataframe(data_frame, group_by_col='content', filter_condition=None):

        if filter_condition is not None:
            data_frame = data_frame.filter(filter_condition)

        if group_by_col == 'content' and 'name' in data_frame.columns:
            return (
                data_frame.group_by(['name', group_by_col])
                .agg([
                    pl.col('minutes').sum().alias('total_minutes'),
                    (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                    pl.col('minutes').count().alias('frequency')
                ])
                .sort(['name', 'total_minutes'], descending=[False, True])
            )

        return (
            data_frame.group_by(group_by_col)
            .agg([
                pl.col('minutes').sum().alias('total_minutes'),
                (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                pl.col('minutes').count().alias('frequency')
            ])
            .sort('total_minutes', descending=True)
        )

    def analyze_task_data(self, tasks, daily_tasks, comm_tasks, all_items):

        df, daily_df, comm_df, all_items_df = self.create_dataframes(
            tasks, daily_tasks, comm_tasks, all_items
        )
        
        # クラーク業務の集計
        clerk_tasks = self.aggregate_dataframe(
            df,
            filter_condition=pl.col('content').str.contains('クラーク業務')
        )
        
        # クラーク業務以外の集計
        non_clerk_tasks = self.aggregate_dataframe(
            df,
            filter_condition=~pl.col('content').str.contains('クラーク業務')
        )
        
        # デイリータスクの集計
        daily_tasks_agg = self.aggregate_dataframe(daily_df)
        
        # コミュニケーションの集計
        communication_by_name = self.aggregate_dataframe(comm_df, group_by_col='name')
        
        # コミュニケーション内容別の集計
        communication_by_content = (
            comm_df.group_by(['content', 'name'])
            .agg([
                pl.col('minutes').sum().alias('total_minutes'),
                (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                pl.col('minutes').count().alias('frequency')
            ])
            .sort(['name', 'total_minutes'], descending=[False, True])
            .select(['content', 'name', 'total_minutes', 'total_hours', 'frequency'])
        )
        
        # 全項目の集計
        all_items_summary = self.aggregate_dataframe(all_items_df)
        
        return (
            clerk_tasks,
            non_clerk_tasks,
            daily_tasks_agg,
            communication_by_name,
            communication_by_content,
            all_items_summary
        )
