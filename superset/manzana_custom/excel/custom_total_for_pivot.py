import pandas as pd
import numpy as np

class CustomTotalForPivot:
    def __init__(self, df: pd.DataFrame):
        self.df = df


    def generate(self, aggfunc: str, show_rows_total: bool, show_columns_total: bool):
        # "Sample Variance", "First", "Last", "Sample Standard Deviation", need to clarify
        aggfunc_type={
            "Count": self.generate_count,
            "Count Unique Values": self.generate_count_unique_values,
            "List Unique Values": self.generate_list_unique_values,
            "Count as Fraction of Total": self.generate_count_as_fraction_of_total,
            "Count as Fraction of Rows": self.generate_count_as_fraction_of_rows,
            "Count as Fraction of Columns": self.generate_count_as_fraction_of_columns,
            "Sample Variance": self.generate_sample_variance,
            "Sample Standard Deviation": self.generate_sample_standart_deviation,
        }

        generate_method = aggfunc_type.get(aggfunc)
        if generate_method:
            return generate_method(show_rows_total, show_columns_total)
        else:
            return self.df

    def generate_count(self, show_rows_total: bool, show_columns_total: bool):
        if show_rows_total:
            self.df['Подытог'] = self.df.count(axis=1)
            self.df['Total (Count)'] = self.df['Подытог']
            
        if show_columns_total:
            column_counts = self.df.count()
                
            if show_rows_total:
                column_counts['Подытог'] = column_counts.drop(['Подытог', 'Total (Count)']).sum()  
                column_counts['Total (Count)'] = self.df['Подытог'].sum()
            else:
                column_counts['Подытог'] = column_counts.sum()
            self.df.loc['Total (Count)'] = column_counts

        return self.df

        
    def generate_count_unique_values(self, show_rows_total: bool, show_columns_total: bool):
        if show_rows_total:
            self.df['Подытог'] = self.df.count(axis=1)
            self.df['Total (Count unique values)'] = self.df['Подытог']
            
        if show_columns_total:
            column_counts = self.df.count()
            
            if show_rows_total:
                column_counts['Подытог'] = column_counts.drop(['Подытог', 'Total (Count unique values)']).sum()  
                column_counts['Total (Count unique values)'] = self.df['Подытог'].sum()
            else:
                column_counts['Подытог'] = column_counts.sum()
            self.df.loc['Total (Count unique values)'] = column_counts
        return self.df

    def generate_list_unique_values(self, show_rows_total: bool, show_columns_total: bool):
        def unique_values(series):
            values = series.dropna().astype(str).unique()
            return ', '.join(values)

        if show_rows_total:
            self.df['Подытог'] = self.df.apply(unique_values, axis=1)
            self.df['Total (List unique values)'] = self.df['Подытог']


        if show_columns_total:
            column_unique = self.df.apply(unique_values)
            if show_rows_total:
                column_unique['Подытог'] = unique_values(self.df['Подытог'])
                column_unique['Total (List unique values)'] = unique_values(self.df['Total (List unique values)'])
        
            self.df.loc['Total (List unique values)'] = column_unique
        return self.df

    def generate_count_as_fraction_of_total(self, show_rows_total: bool, show_columns_total: bool):
        total_sum = self.df.sum().sum()

        if show_rows_total:
            self.df['Подытог'] = self.df.sum(axis=1) / total_sum

        if show_columns_total:
            column_totals = self.df.sum() / total_sum

            if show_rows_total:
                column_totals['Подытог'] = self.df['Подытог'].sum()

            self.df.loc['Total (Count as fraction of total)'] = column_totals

        if show_rows_total:
            self.df['Total (Count as fraction of total)'] = self.df['Подытог']

        return self.df

    def generate_count_as_fraction_of_rows(self, show_rows_total: bool, show_columns_total: bool):
        total_sum = self.df.sum().sum()

        if show_rows_total:
            self.df['Подытог'] = self.df.sum(axis=1)

        if show_columns_total:
            column_totals = self.df.sum()

            if show_rows_total:
                column_totals['Подытог'] = self.df['Подытог'].sum()

            self.df.loc['Total (Count as fraction of total)'] = column_totals / total_sum

        if show_rows_total:
            self.df['Total (Count as fraction of rows)'] = self.df['Подытог']

        return self.df
    
    def generate_count_as_fraction_of_columns(self, show_rows_total: bool, show_columns_total: bool):
        print(self.df, flush=True)
        column_sums = self.df.sum()

        df_normalized = self.df.divide(column_sums, axis=1)

        if show_rows_total:
            df_normalized['Подытог'] = df_normalized.mean(axis=1)

        if show_columns_total:
            df_normalized.loc['Total (Count as fraction of columns)'] = 1

        if show_rows_total:
            df_normalized['Total (Count as fraction of columns)'] = df_normalized['Подытог']

            avg_of_total_row = df_normalized.loc['Total (Count as fraction of columns)'].mean()
            df_normalized.loc['Total (Count as fraction of columns)', 'Подытог'] = avg_of_total_row

        self.df = df_normalized
        return df_normalized
    
    def generate_sample_variance(self, show_rows_total: bool, show_columns_total: bool):
        # Computing sample variance along rows
        def row_sample_variance(row):
            n = row.count()  # number of non-NaN data points
            if n <= 1:  # to avoid division by zero
                return np.nan
            mean = row.mean()
            return ((row - mean) ** 2).sum() / (n - 1)

        # Computing sample variance along columns
        def col_sample_variance(col):
            n = col.count()
            if n <= 1:
                return np.nan
            mean = col.mean()
            return ((col - mean) ** 2).sum() / (n - 1)

        if show_rows_total:
            self.df['Подытог'] = self.df.apply(row_sample_variance, axis=1)
            self.df['Total (Sample Variance)'] = self.df['Подытог']

        if show_columns_total:
            column_variances = self.df.apply(col_sample_variance)

            if show_rows_total:
                column_variances['Подытог'] = column_variances.drop(['Подытог', 'Total (Sample Variance)']).mean()
                column_variances['Total (Sample Variance)'] = self.df['Подытог'].mean()

            self.df.loc['Total (Sample Variance)'] = column_variances

        return self.df
        
    def generate_sample_standart_deviation(self, show_rows_total: bool, show_columns_total: bool):
    # Computing sample standard deviation along rows
        def row_sample_std(row):
            n = row.count()  # number of non-NaN data points
            if n <= 1:  # to avoid division by zero
                return np.nan
            mean = row.mean()
            variance = ((row - mean) ** 2).sum() / (n - 1)
            return np.sqrt(variance)

        # Computing sample standard deviation along columns
        def col_sample_std(col):
            n = col.count()
            if n <= 1:
                return np.nan
            mean = col.mean()
            variance = ((col - mean) ** 2).sum() / (n - 1)
            return np.sqrt(variance)

        if show_rows_total:
            self.df['Подытог'] = self.df.apply(row_sample_std, axis=1)
            self.df['Total (Sample Standard Deviation)'] = self.df['Подытог']

        if show_columns_total:
            column_std_devs = self.df.apply(col_sample_std)

            if show_rows_total:
                column_std_devs['Подытог'] = column_std_devs.drop(['Подытог', 'Total (Sample Standard Deviation)']).mean()
                column_std_devs['Total (Sample Standard Deviation)'] = self.df['Подытог'].mean()

            self.df.loc['Total (Sample Standard Deviation)'] = column_std_devs

        return self.df
        