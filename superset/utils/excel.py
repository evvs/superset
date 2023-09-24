import io
from typing import Any

import pandas as pd
from typing import Union
from superset.models.slice import Slice
from superset.manzana_custom.excel.generate_chart import GenerateChart
from superset.manzana_custom.excel.pivot import is_pivot
from superset.connectors.base.models import BaseDatasource

from superset.charts.post_processing import pivot_table_v2
from superset.utils.core import get_column_names

import logging
import datetime
import numpy as np

def is_valid_date(date_string: str, date_format: str) -> bool:
    try:
        datetime.datetime.strptime(date_string, date_format)
        return True
    except ValueError:
        return False


def translate_aggfunc_name(col_name):
    if not isinstance(col_name, str):
        return col_name
    translations = {
    "Count unique values": "Количество уникальных значений",
    "List unique values": "Список уникальных значений",
    "Average": "Среднее",
    "Median": "Медиана",
    "Sample variance": "Дисперсия",
    "Sample standard deviation": "Стандартное отклонение",
    "Minimum": "Минимум",
    "Maximum": "Максимум",
    "First": "Первый",
    "Last": "Последний",
    "Sum as fraction of columns": "Сумма как доля столбцов",
    "Sum as fraction of rows": "Сумма как доля строк",
    "Sum as fraction of total": "Сумма как доля целого",
    "Count as fraction of columns": "Количество, как доля от столбцов",
    "Count as fraction of rows": "Количество, как доля от строк",
    "Count as fraction of total": "Количество, как доля от целого",
    "Subtotal": "Подытог",
    "Sum": "Сумма",
    "Count": "Количество",
    "NaT": ""
    }
    for key, value in translations.items():
        col_name = col_name.replace(key, value)
    return col_name.strip()


def format_number(number, d3NumberFormat):
    if np.isnan(number):
        return ''
    if d3NumberFormat == ',d':
        return '{:,.0f}'.format(number)
    elif d3NumberFormat == '.1s':
        return '{:.1e}'.format(number).replace('e', ' ').split()[0]
    elif d3NumberFormat == '.3s':
        return '{:.3e}'.format(number).replace('e', ' ').split()[0]
    elif d3NumberFormat == ',.1%':
        return '{:,.1%}'.format(number)
    elif d3NumberFormat == '.2%':
        return '{:.2%}'.format(number)
    elif d3NumberFormat == '.3%':
        return '{:.3%}'.format(number)
    elif d3NumberFormat == '.4r':
        return '{:.4f}'.format(number).rstrip('0').rstrip('.')
    elif d3NumberFormat == ',.1f':
        return '{:,.1f}'.format(number)
    elif d3NumberFormat == ',.2f':
        return '{:,.2f}'.format(number)
    elif d3NumberFormat == ',.3f':
        return '{:,.3f}'.format(number)
    elif d3NumberFormat == '+,':
        return '{:+,.0f}'.format(number)
    elif d3NumberFormat == '$,.2f':
        return '${:,.2f}'.format(number)
    else:
        raise ValueError(f"Unsupported format: {d3NumberFormat}")

logger = logging.getLogger(__name__)

request_to_excel_format = {
    'smart_date': 'd mmmm yyyy',
    '%d.%m.%Y': 'dd/mm/yyyy',
    '%d/%m/%Y': 'dd/mm/yyyy',
    '%m/%d/%Y': 'mm/dd/yyyy',
    '%Y-%m-%d': 'yyyy-mm-dd',
    '%Y-%m-%d %H:%M:%S': 'yyyy-mm-dd hh:mm:ss',
    '%d-%m-%Y %H:%M:%S': 'dd-mm-yyyy hh:mm:ss',
    '%H:%M:%S': 'hh:mm:ss',
}

def df_to_excel(df: pd.DataFrame, sheet_name='Sheet1', from_report=False, slice: Union[Slice, None] = None,
                datasource: Union[BaseDatasource, None] = None,
                **kwargs: Any) -> bytes:
    output = io.BytesIO()
    date_format_by_column_name = {}

    if is_pivot(slice):

        if (not from_report):
            df = pivot_table_v2(df, slice.form_data, datasource)

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(map(str, col)).strip()
                          for col in df.columns.values]

        groupbyRows = slice.form_data.get("groupbyRows")
        if groupbyRows and len(groupbyRows) == 1:
            verbose_map = datasource.data["verbose_map"] if datasource else None
            df = df.reset_index().rename(
                columns={"level_0": get_column_names(groupbyRows, verbose_map)[0]}).rename(
                    columns={"index": get_column_names(
                        groupbyRows, verbose_map)[0]}
            )
        if bool(slice.form_data.get("rowTotals")):
            last_col_name = df.columns[-1]
            translated_col_name = translate_aggfunc_name(last_col_name)
            df.rename(columns={last_col_name: translated_col_name}, inplace=True)
        if bool(slice.form_data.get("colTotals")):
            last_row_values = df.iloc[-1, :]
    
            translated_values = last_row_values.map(translate_aggfunc_name)
    
            df.iloc[-1, :] = translated_values

    if slice:
        try:
            column_config = slice.form_data.get("column_config")
            if column_config:
                for idx, column in enumerate(df.columns):
                    col_name = column if not isinstance(
                        column, tuple) else column[0]
                    in_column_config = column_config.get(
                        col_name)
                    


                    if not in_column_config:
                        verbose_map = datasource.data["verbose_map"]
                        if verbose_map:
                            verbosed_name = next((key for key, value in verbose_map.items() if value == col_name), None)
                            in_column_config = column_config.get(verbosed_name)
                    if in_column_config:
                        date_format_from_request = in_column_config.get(
                            "d3TimeFormat")
                        number_format_from_request = in_column_config.get("d3NumberFormat")

                        if date_format_from_request and date_format_from_request != "smart_date":
                            date_format_by_column_name[df[column]
                                                       .name] = date_format_from_request
                            df[column] = pd.to_datetime(df[column])
                            df[column] = df[column].dt.strftime(
                                date_format_from_request)
                        elif number_format_from_request and number_format_from_request != "my_format":
                            df[column] = df[column].apply(lambda x: format_number(x, number_format_from_request))
        except Exception as err:
            logger.error(f"ERROR WHEN TRYING FORMAT DATE for column: {column}")
            logger.error(
                # f"Data type of the column before conversion: {df[column].dtype}") maybe check later
                f"Data type of the column before conversion: {df[column]}")
            logger.error(err)

    for column in df.columns:
        dtype = getattr(df[column], 'dtype', None)
        if dtype and str(dtype).startswith('datetime64[ns'):
            df[column] = df[column].dt.tz_localize(None)

    # remove NULL from dataframe before writing to Excel
    df.replace("NULL", np.nan, inplace=True)

    df = df.applymap(lambda x: x[0] if isinstance(x, tuple) else x)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, **kwargs)

        workbook = writer.book

        header_format = workbook.add_format(
            {'bg_color': '#96bfff', 'bold': True})  # manzana custom
        worksheet = writer.sheets[sheet_name]  # manzana custom

        for col_num, value in enumerate(df.columns.values):  # manzana custom
            if isinstance(value, tuple):  # manzana custom
                value = value[0]  # manzana custom

            worksheet.write(0, col_num, value, header_format)  # manzana custom

            if value in df.columns:
                column_data = df[value]

                if isinstance(column_data, pd.DataFrame):  # If multi-index column
                    column_data = column_data.iloc[:, 0]
                if isinstance(column_data, pd.Series) and column_data.dtype == 'object':
                    date_format_from_dict = date_format_by_column_name.get(
                        column_data.name, '')
                    valid_dates_mask = column_data.apply(
                        lambda x: is_valid_date(str(x), date_format_from_dict))
                    try:
                        # Convert only valid dates to datetime64 format first
                        column_data[valid_dates_mask] = pd.to_datetime(
                            column_data[valid_dates_mask], format=date_format_from_dict)
                        # Convert datetime64 format to Excel's datetime format
                        excel_format = request_to_excel_format.get(
                            date_format_from_dict, 'd mmmm yyyy')
                        date_format = workbook.add_format(
                            {'num_format': excel_format})
                        # Apply date formatting in Excel
                        for row_num, date_val in enumerate(column_data):
                            if valid_dates_mask.iloc[row_num]:
                                worksheet.write_datetime(
                                    row_num + 1, col_num, date_val, date_format)  # +1 to skip header
                    except BaseException as err:
                        logger.error("ERROR WITH CONVERTING pd.to_datetime")
                        logger.error(err)

        if from_report:  # manzana custom
            worksheet.write(0, len(df.columns.values), None,
                            workbook.add_format())  # manzana custom

        worksheet.autofit()  # manzana custom

        if slice and slice.datasource:
            if slice.datasource.type:
                try:
                    chart_name = str(slice.slice_name)
                    chart = GenerateChart(
                        df, workbook, worksheet, chart_name)  # type: ignore
                    chart_type = slice.form_data.get("viz_type")
                    if chart_type:
                        if chart_type == 'pie' and slice.form_data.get('donut'):
                            chart_type = 'donut'
                        chart.generate(chart_type)
                except BaseException as err:
                    logger.error("ERROR WITH GENERATING CHART")
                    logger.error(err)
    return output.getvalue()
