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

def df_to_excel(df: pd.DataFrame, sheet_name='Sheet1', from_report=False, **kwargs: Any) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, **kwargs)

        workbook  = writer.book

        # worksheet = writer.sheets[sheet_name]
        # chart = workbook.add_chart({'type': 'doughnut'})
        # chart.add_series({
        #     "name": sheet_name,                           implemented next
        #     'categories': '=Sheet1!A2:A8',
        #     'values':     '=Sheet1!B2:B8',
        # })
        # worksheet.insert_chart('B4', chart)

        header_format = workbook.add_format({'bg_color': '#96bfff', 'bold': True}) # manzana custom
        worksheet = writer.sheets[sheet_name] # manzana custom

        for col_num, value in enumerate(df.columns.values): # manzana custom
            if isinstance(value, tuple): # manzana custom
                value = value[0] # manzana custom
            
            worksheet.write(0, col_num, value, header_format) # manzana custom
        
        if from_report: # manzana custom
            worksheet.write(0, len(df.columns.values), None, workbook.add_format()) # manzana custom

        worksheet.autofit()  # manzana custom

    return output.getvalue()