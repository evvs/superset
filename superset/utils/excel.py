# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
import io
from typing import Any

import pandas as pd
from typing import Union
from superset.models.slice import Slice
from superset.manzana_custom.excel.generate_chart import GenerateChart

import logging
import datetime

def is_valid_date(date_string: str, date_format: str) -> bool:
    try:
        datetime.datetime.strptime(date_string, date_format)
        return True
    except ValueError:
        return False

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

# Sample usage:
# request_format = '%d.%m.%Y'
# excel_format = request_to_excel_format.get(request_format, 'd mmmm yyyy')  # default to 'd mmmm yyyy' for 'smart_date'

def df_to_excel(df: pd.DataFrame, sheet_name='Sheet1', from_report=False, slice: Union[Slice, None] = None, **kwargs: Any) -> bytes:
    output = io.BytesIO()
    date_format_by_column_name = {}
    # Normalize and format datetime columns
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
                        mapped_field = ''
                        source_for_mapping = []
                        if slice.form_data.get("query_mode") == 'aggregate':
                            source_for_mapping = slice.form_data.get("groupby", [])
                        elif slice.form_data.get("query_mode") == 'raw':
                            source_for_mapping = slice.form_data.get(
                                "all_columns", [])
                        if len(source_for_mapping):
                            mapped_field = source_for_mapping[idx]
                        in_column_config = column_config.get(mapped_field)

                    if in_column_config:
                        date_format_from_request = in_column_config.get("d3TimeFormat")

                        if date_format_from_request and date_format_from_request != "smart_date":
                            date_format_by_column_name[df[column].name] = date_format_from_request
                            df[column] = pd.to_datetime(df[column])
                            df[column] = df[column].dt.strftime(date_format_from_request)
        except Exception as err:
            logger.error(f"ERROR WHEN TRYING FORMAT DATE for column: {column}")
            logger.error(
                f"Data type of the column before conversion: {df[column].dtype}")
            logger.error(err)

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
                if isinstance(column_data, pd.Series):
                    if column_data.dtype == 'object':
                        date_format_from_dict = date_format_by_column_name.get(column_data.name, '')
                        # Filter the data that can potentially be converted
                        valid_dates_mask = column_data.apply(lambda x: is_valid_date(str(x), date_format_from_dict))
                        
                        # Convert only valid dates
                        try:
                            excel_format = request_to_excel_format.get(date_format_from_dict, 'd mmmm yyyy')
                            date_format = workbook.add_format({'num_format': excel_format})
                            
                            column_data[valid_dates_mask] = pd.to_datetime(column_data[valid_dates_mask], format=date_format_from_dict)
                            
                            df[value] = column_data  # Update the main DataFrame
                            worksheet.set_column(col_num, col_num, None, date_format)
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
