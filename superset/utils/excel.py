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

import numpy as np
import logging
logger = logging.getLogger(__name__)


def df_to_excel(df: pd.DataFrame, sheet_name='Sheet1', from_report=False, slice: Union[Slice, None] = None, **kwargs: Any) -> bytes:
    output = io.BytesIO()

    # Normalize and format datetime columns

    if (slice):
        try:
            for idx, column in enumerate(df.columns):
                if df[column].dtype in ['datetime64[ns]', 'datetime']:
                    col_name = column if not isinstance(column, tuple) else column[0]
                    col_config = slice.form_data.get("column_config", {}).get(col_name) #.get("d3TimeFormat")

                    if (not col_config):
                        mapped_field = ''
                        all_columns = slice.form_data.get("all_columns")
                        if (all_columns):
                            mapped_field = all_columns[idx]
                        col_config =  slice.form_data.get("column_config", {}).get(mapped_field)

                    date_format = col_config.get("d3TimeFormat")

                    if date_format and date_format != "smart_date":
                        df[column] = df[column].dt.strftime(date_format)
        except Exception as err:
            logger.error(f"ERROR WHEN TRYING FORMAT DATE for column: {column}")
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

        if from_report:  # manzana custom
            worksheet.write(0, len(df.columns.values), None,
                            workbook.add_format())  # manzana custom

        worksheet.autofit()  # manzana custom

        if (slice and slice.datasource):
            if (slice.datasource.type):
                try:
                    chart_name = str(slice.slice_name)
                    chart = GenerateChart(
                        df, workbook, worksheet, chart_name)  # type: ignore
                    chart_type = slice.form_data.get("viz_type")
                    if (chart_type):
                        chart.generate(chart_type)
                except BaseException as err:
                    logger.error("ERROR WITH GENERATING CHART")
                    logger.error(err)

    return output.getvalue()
