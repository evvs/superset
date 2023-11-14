import pandas as pd
import math
import decimal
import numpy as np

class ExcelColumnFormatter:
    def __init__(self, workbook):
        self.workbook = workbook
        self.formats = {}
        self.format_properties = {}
        self.type_mapping = {}

    def add_format(self, format_type):
        if format_type not in self.formats:
            if format_type == "~g":
                self.formats[format_type] = self.workbook.add_format()
                self.type_mapping[format_type] = '0'

            if format_type == ',d':
                self.formats[format_type] = self.workbook.add_format({'num_format': '#,##0'})
                self.type_mapping[format_type] = '#,##0'

            elif format_type == 'my_format':
                self.formats[format_type] = self.workbook.add_format({'num_format': '#,##0'})
                self.type_mapping[format_type] = '#,##0'

            elif format_type == ',.1%':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.0%'})
                self.type_mapping[format_type] = '0.0%'

            elif format_type == '.2%':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.00%'})
                self.type_mapping[format_type] = '0.00%'

            elif format_type == '.3%':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.000%'})
                self.type_mapping[format_type] = '0.000%'

            elif format_type == ',.1f':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.0'})
                self.type_mapping[format_type] = '0.0'

            elif format_type == ',.2f':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.00'})
                self.type_mapping[format_type] = '0.00'

            elif format_type == ',.3f':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.000'})
                self.type_mapping[format_type] = '0.000'

            elif format_type == '+,':  
                self.formats[format_type] = self.workbook.add_format({'num_format': '+#,##0'})
                self.type_mapping[format_type] = '+#,##0'

            elif format_type == '$,.2f':
                self.formats[format_type] = self.workbook.add_format({'num_format': '[$$-409]#,##0.00'})
                self.type_mapping[format_type] = '[$$-409]#,##0.00'

            else:
                self.formats[format_type] = self.workbook.add_format({'num_format': '0'})
                self.type_mapping[format_type] = '0'

        return self.formats.get(format_type)

    def apply_format_to_column(self, worksheet, column_index, df, format_type):
        formats = self.add_format(format_type)
        for row_num, value in enumerate(df.iloc[:, column_index], start=1):
            if pd.isna(value) or value in [np.inf, -np.inf]:
                worksheet.write_string(row_num, column_index, '')
            else:
                if isinstance(value, str):
                    value = value.replace(',', '')  # Remove commas
                    try:
                        value = float(value) if '.' in value else int(value)
                    except ValueError:
                        pass  

                elif isinstance(value, decimal.Decimal):
                    value = float(value)

                if isinstance(value, (int, float)):
                    worksheet.write_number(row_num, column_index, value, formats)
                else:
                    worksheet.write_string(row_num, column_index, str(value))
        self.format_properties[column_index] = {"num_format": self.type_mapping.get(format_type, {})}
