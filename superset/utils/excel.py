import io
from typing import Any

import pandas as pd
from typing import Union
from superset.models.slice import Slice
from superset.manzana_custom.excel.generate_chart import GenerateChart
from superset.manzana_custom.excel.pivot import is_pivot
from superset.manzana_custom.excel.conditional_formatting import ConditionalFormatting
from superset.connectors.base.models import BaseDatasource

from superset.charts.post_processing import pivot_table_v2
from superset.utils.core import get_column_names

import logging
import datetime
import numpy as np
import re

from xlsxwriter import Workbook

def contains_formatted_percentage(s):
    try:
        float(s.strip('%'))
        return True
    except ValueError:
        return False

def is_string_representation_of_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

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
    "Sample Variance": "Дисперсия",
    "Sample Standard Deviation": "Стандартное отклонение",
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

def percent_to_float(val):
    if isinstance(val, str) and 'nan%' in val:
        return np.nan
    elif isinstance(val, str) and '%' in val:
        return float(val.strip('%'))
    return val

def format_number(number, d3NumberFormat):
    if pd.isnull(number):
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
    'smart_date': '%d.%m.%Y',
    '%d.%m.%Y': 'dd/mm/yyyy',
    '%d/%m/%Y': 'dd/mm/yyyy',
    '%m/%d/%Y': 'mm/dd/yyyy',
    '%Y-%m-%d': 'yyyy-mm-dd',
    '%Y-%m-%d %H:%M:%S': 'yyyy-mm-dd hh:mm:ss',
    '%d-%m-%Y %H:%M:%S': 'dd-mm-yyyy hh:mm:ss',
    '%H:%M:%S': 'hh:mm:ss',
}

def format_unixtimestamp_in_colname(df, date_format, workbook: Workbook, worksheet, header_format):    
    def is_unix_timestamp(s):
        pattern = re.compile(r"\d{10,13}(\.\d+)?")
        return bool(pattern.fullmatch(s))
    
    new_col_names = []
    
    for col in df.columns:
        parts = col.split(" ")
        new_parts = [
            datetime.datetime.utcfromtimestamp(float(part)/1000).strftime(date_format)
            if is_unix_timestamp(part) else part
            for part in parts
        ]
        new_col_names.append(" ".join(new_parts))
    
    df.columns = new_col_names
    

    for col_num, value in enumerate(new_col_names):
        if isinstance(value, tuple):
            value = value[0]
        worksheet.write(0, col_num, value, header_format)


def df_to_excel(df: pd.DataFrame, sheet_name='Sheet1', from_report=False, slice: Union[Slice, None] = None,
                datasource: Union[BaseDatasource, None] = None,
                **kwargs: Any) -> bytes:
    output = io.BytesIO()
    date_format_by_column_name = {}
    index_already_reset = False
    if is_pivot(slice):
        if (not from_report):
            df = pivot_table_v2(df, slice.form_data, datasource)

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(map(str, col)).strip()
                          for col in df.columns.values]

        if isinstance(df.index, pd.MultiIndex):
            index_already_reset = True
            df.reset_index(inplace=True)
        else:
            index_already_reset = False

        groupbyRows = slice.form_data.get("groupbyRows")
        if groupbyRows:
            verbose_map = datasource.data["verbose_map"] if datasource else None

            # Only reset the index if it hasn't been reset earlier
            if not index_already_reset:
                df.reset_index(inplace=True)

            rename_dict = {}
            for i, groupby in enumerate(groupbyRows):
                rename_dict[f'level_{i}'] = get_column_names([groupby], verbose_map)[0]

            # Check if the "index" column exists and rename it to the first groupby column
            # This may be unnecessary, but added just in case.
            if 'index' in df.columns:
                rename_dict['index'] = get_column_names([groupbyRows[0]], verbose_map)[0]

            df.rename(columns=rename_dict, inplace=True)

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

                    # Check if column contains formatted percentages
                    if df[column].apply(lambda x: isinstance(x, str) and contains_formatted_percentage(x)).any():
                        continue
                    
                    # Check if column has string representations of numbers and try converting them to float
                    if df[column].apply(lambda x: isinstance(x, str) and is_string_representation_of_number(x)).any():
                        df[column] = df[column].apply(lambda x: float(x) if is_string_representation_of_number(x) else x)

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
                        small_number_format_from_request = in_column_config.get("d3SmallNumberFormat")

                        if date_format_from_request and date_format_from_request != "smart_date":
                            date_format_by_column_name[df[column]
                                                       .name] = date_format_from_request
                            df[column] = pd.to_datetime(df[column])
                            df[column] = df[column].dt.strftime(
                                date_format_from_request)
                        elif small_number_format_from_request and small_number_format_from_request not in ["my_format", "SMART_NUMBER"]:
                            df[column] = df[column].apply(lambda x: format_number(x, small_number_format_from_request))
                        elif not small_number_format_from_request and number_format_from_request and number_format_from_request not in ["my_format", "SMART_NUMBER"]:
                            df[column] = df[column].apply(lambda x: format_number(x, number_format_from_request))
        except Exception as err:
            logger.error(f"ERROR WHEN TRYING FORMAT DATE for column: {column}")
            logger.error(
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
                        valid_dates = column_data[valid_dates_mask]
                        valid_dates_converted = pd.to_datetime(valid_dates, format=date_format_from_dict, errors='coerce')

                        # Handle or replace NaT values after conversion if necessary
                        valid_dates_converted.fillna("YourPlaceholderOrAction", inplace=True)  # Consider what you want to do with NaT values

                        column_data[valid_dates_mask] = valid_dates_converted

                        # Convert datetime64 format to Excel's datetime format
                        excel_format = request_to_excel_format.get(date_format_from_dict, 'd mmmm yyyy')
                        date_format = workbook.add_format({'num_format': excel_format})

                        # Apply date formatting in Excel
                        for row_num, date_val in enumerate(column_data):
                            if valid_dates_mask.iloc[row_num]:
                                worksheet.write_datetime(row_num + 1, col_num, date_val, date_format)  # +1 to skip header

                    except Exception as err:  # Using BaseException might be too broad. Using Exception is generally recommended for catching unexpected errors
                        logger.error("ERROR WITH CONVERTING pd.to_datetime")
                        logger.error(f"Column: {value}, Error: {err}")

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

            
            if slice.form_data.get("conditional_formatting") and datasource.data["verbose_map"]: # type: ignore
                instance = ConditionalFormatting(df, workbook, worksheet, datasource.data["verbose_map"], is_pivot(slice)) # type: ignore
                instance.generate(slice.form_data.get("conditional_formatting")) # type: ignore
        
        if(from_report):
            try:
                date_format_key = slice.form_data.get("date_format")
                if (date_format_key):
                    date_format = request_to_excel_format.get(date_format_key, 'dd/mm/yyyy')
                    format_unixtimestamp_in_colname(df, date_format, workbook, worksheet, header_format) # type: ignore
            except Exception as err:
                logger.error(f"ERROR WHEN TRYING FORMAT DATE from report")

        bold_format = workbook.add_format({'bold': True})
        if is_pivot(slice):
            if bool(slice.form_data.get("rowTotals")):
                col_idx = len(df.columns) - 1  # index of the last column
    
                # For the header of the last column, use the header format.
                worksheet.write(0, col_idx, df.columns[col_idx], header_format)
    
                # For the rest of the rows in the last column, use the bold format.
                for row_num in range(1, len(df) + 1):
                    cell_value = df.iloc[row_num - 1, col_idx]
                    worksheet.write(row_num, col_idx, cell_value, bold_format)

        if bool(slice.form_data.get("colTotals")):
            row_idx = len(df)  # index of the last row
            bold_format = workbook.add_format({'bold': True})
        
            for col_num in range(len(df.columns)):
                cell_value = df.iloc[row_idx - 1, col_num]
                worksheet.write(row_idx, col_num, cell_value, bold_format)

    return output.getvalue()
