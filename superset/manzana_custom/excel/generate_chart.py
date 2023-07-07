from xlsxwriter import Workbook
from pandas import DataFrame
from typing import Any
import logging

logger = logging.getLogger(__name__)


def col_num_to_letter(n):
    try:
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    except BaseException as err:
        logger.error("ERROR WHEN TRYING CENERATE PIE CHART")
        logger.error(err)


class GenerateChart:

    cell_for_rendering = "B4"

    def __init__(self, df: DataFrame, workbook: Workbook, worksheet: Any, slice_name: str):
        self.df = df
        self.workbook = workbook
        self.worksheet = worksheet
        self.slice_name = slice_name

    def generate(self, chart_type: str):
        if (chart_type == 'pie'):
            self.__generate_pie_chart()
        if (chart_type == 'echarts_timeseries_bar'):
            self.__generate_column_chart()
        if (chart_type == 'echarts_timeseries_line'):
            self.__generate_line_chart()

    def __generate_pie_chart(self):
        try:
            chart = self.workbook.add_chart({'type': 'doughnut'})
            num_rows = self.df.shape[0]
            num_cols = self.df.shape[1]
            sheet_name = 'Sheet1'
        # As Excel's index start from 1 and pandas from 0
            categories = f'={sheet_name}!A2:A{num_rows+1}'
            values = f'={sheet_name}!B2:B{num_rows+1}'
            chart.add_series({
                "name": self.slice_name,
                'categories': categories,
                'values': values,
            })
            self.worksheet.insert_chart(self.cell_for_rendering, chart)
        except BaseException as err:
            logger.error("ERROR WHEN TRYING CENERATE PIE CHART")
            logger.error(err)

    def __generate_column_chart(self):
        try:
            chart = self.workbook.add_chart({'type': 'column'})
            num_rows = self.df.shape[0]
            sheet_name = 'Sheet1'

    # As Excel's index start from 1 and pandas from 0
            categories = f'={sheet_name}!A2:A{num_rows+1}'

            for i in range(1, self.df.shape[1]):
                # +1 because Excel's index starts from 1
                col_letter = col_num_to_letter(i+1)
                values = f'={sheet_name}!{col_letter}2:{col_letter}{num_rows+1}'
                chart.add_series({
                    "name": str(self.df.columns[i]),
                    'categories': categories,
                    'values': values,
                })
    # Add a chart title and some axis labels.
            chart.set_title({"name": self.slice_name})

    # Set an Excel chart style.
            chart.set_style(11)

            self.worksheet.insert_chart(self.cell_for_rendering, chart, {
                "x_offset": 25, "y_offset": 10})
        except BaseException as err:
            logger.error("ERROR WHEN TRYING CENERATE COLUMN(BASE) CHART")
            logger.error(err)

    def __generate_line_chart(self):
        try:
            chart = self.workbook.add_chart({'type': 'line'})
            num_rows = self.df.shape[0]
            sheet_name = 'Sheet1'

    # As Excel's index start from 1 and pandas from 0
            categories = f'={sheet_name}!A2:A{num_rows+1}'

            for i in range(1, self.df.shape[1]):
                # +1 because Excel's index starts from 1
                col_letter = col_num_to_letter(i+1)
                values = f'={sheet_name}!{col_letter}2:{col_letter}{num_rows+1}'
                chart.add_series({
                    "name": str(self.df.columns[i]),
                    'categories': categories,
                    'values': values,
                })
    # Add a chart title and some axis labels.
            chart.set_title({"name": self.slice_name})

    # Set an Excel chart style.
            chart.set_style(10)

            self.worksheet.insert_chart(self.cell_for_rendering, chart, {
                                        "x_offset": 25, "y_offset": 10})
        except BaseException as err:
            logger.error("ERROR WHEN TRYING CENERATE LINE CHART")
            logger.error(err)
