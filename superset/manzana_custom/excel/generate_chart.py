from xlsxwriter import Workbook
from pandas import DataFrame
from typing import Any
import logging

logger = logging.getLogger(__name__)


def col_num_to_letter(n):
    """Converts a column number to an Excel column letter."""
    try:
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    except Exception as err:
        logger.error("Error when trying to convert column number to letter.")
        logger.error(err)


class GenerateChart:
    """Class to generate Excel charts using xlsxwriter."""

    cell_for_rendering = "B4"

    def __init__(self, df: DataFrame, workbook: Workbook, worksheet: Any, slice_name: str):
        self.df = df
        self.workbook = workbook
        self.worksheet = worksheet
        self.slice_name = slice_name
        self.sheet_name = 'Sheet1'
        self.num_rows = df.shape[0]

    def generate(self, chart_type: str):
        """Generates an Excel chart based on the provided chart type."""
        chart_types = {
            "pie": self.__generate_pie_chart,
            "donut": self.__generate_doughnut_chart,
            "echarts_timeseries_bar": self.__generate_column_chart,
            "bar": self.__generate_column_chart,  # dont use it old type
            "dist_bar": self.__generate_column_chart,  # dont use it old type
            "echarts_timeseries_line": self.__generate_line_chart,
            "line": self.__generate_line_chart,
            "echarts_area": self.__generate_area_chart,
            "echarts_timeseries_scatter": self.__generate_scatter_chart
        }
        chart_method = chart_types.get(chart_type)
        if chart_method:
            chart_method()

    def _generate_categories_values(self, col_letter='B'):
        """Generates categories and values for Excel charts."""
        categories = f'={self.sheet_name}!A2:A{self.num_rows+1}'
        values = f'={self.sheet_name}!{col_letter}2:{col_letter}{self.num_rows+1}'
        return categories, values

    def __generate_pie_chart(self):
        """Generates a pie chart."""
        try:
            chart = self.workbook.add_chart({'type': 'pie'})
            categories, values = self._generate_categories_values()
            chart.add_series({
                "name": self.slice_name,
                'categories': categories,
                'values': values,
            })
            self.worksheet.insert_chart(self.cell_for_rendering, chart)
        except Exception as err:
            logger.error("Error when trying to generate pie chart.")
            logger.error(err)

    def __generate_doughnut_chart(self):
        """Generates a doughnut chart."""
        try:
            chart = self.workbook.add_chart({'type': 'doughnut'})
            categories, values = self._generate_categories_values()
            chart.add_series({
                "name": self.slice_name,
                'categories': categories,
                'values': values,
            })
            self.worksheet.insert_chart(self.cell_for_rendering, chart)
        except Exception as err:
            logger.error("Error when trying to generate pie chart.")
            logger.error(err)

    def __generate_column_chart(self):
        """Generates a column chart."""
        self.__generate_chart('column', 11)

    def __generate_line_chart(self):
        """Generates a line chart."""
        self.__generate_chart('line', 10)

    def __generate_area_chart(self):
        """Generates an area chart."""
        self.__generate_chart('area', 11)

    def __generate_scatter_chart(self):
        """Generates a scatter chart."""
        self.__generate_chart('scatter', 11)

    def __generate_chart(self, chart_type: str, style: int):
        """Generates a chart of the specified type and style."""
        try:
            chart = self.workbook.add_chart({'type': chart_type})
            categories = self._generate_categories_values()[0]

            for i in range(1, self.df.shape[1]):
                col_letter = col_num_to_letter(i+1)
                _, values = self._generate_categories_values(col_letter)
                chart.add_series({
                    "name": str(self.df.columns[i]),
                    'categories': categories,
                    'values': values,
                })
            chart.set_title({"name": self.slice_name})
            chart.set_style(style)
            self.worksheet.insert_chart(self.cell_for_rendering, chart, {
                "x_offset": 25, "y_offset": 10})
        except Exception as err:
            logger.error(f"Error when trying to generate {chart_type} chart.")
            logger.error(err)
