import pandas as pd
import math
import decimal
import numpy as np

class ExcelColumnFormatter:
    def __init__(self, workbook):
        self.workbook = workbook
        self.formats = {}

    # def si_format(self, num, sig_fig=3):
    #     num = Decimal(num)
    #     magnitude = int(math.floor(num.log10() / 3) * 3) if num != 0 else 0
    #     scaled_num = num / (10 ** magnitude)
    #     unit = {0: '', 3: 'k', 6: 'M', 9: 'G', 12: 'T', 15: 'P', 18: 'E', 21: 'Z'}.get(magnitude, 'Y')

        # # Format the number to the specified significant figures
        # return f"{scaled_num:.{sig_fig-1}f}{unit}"

    def add_format(self, format_type):
        if format_type not in self.formats:
            if format_type == ',d':  # Comma as thousands separator
                self.formats[format_type] = self.workbook.add_format({'num_format': '#,##0'})
            elif format_type == 'my_format':  # Custom space-separated format
                # Define custom formats for different number lengths as previously defined
                self.formats[format_type] = self._define_custom_formats()
            elif format_type == ',.1%':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.0%'})
            elif format_type == '.2%':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.00%'})
            elif format_type == '.3%':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.000%'})
            elif format_type == ',.1f':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.0'})
            elif format_type == ',.2f':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.00'})
            elif format_type == ',.3f':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '0.000'})
            elif format_type == '+,':  # Percentage format with one decimal place
                self.formats[format_type] = self.workbook.add_format({'num_format': '+#,##0'})
            elif format_type == '$,.2f':  # Dollar sign, comma as thousands separator, two decimal places
                self.formats[format_type] = self.workbook.add_format({'num_format': '[$$-409]#,##0.00'})
            # else:
            #     self.formats[format_type] = self.workbook.add_format({'num_format': '#,##0'})
        return self.formats.get(format_type)

    def _define_custom_formats(self):
        return {
            'small': self.workbook.add_format({'num_format': '0'}),
            'thousands': self.workbook.add_format({'num_format': '# ##0'}),
            'millions': self.workbook.add_format({'num_format': '# ##0, " "000'}),
            # Add more formats as needed
        }
    def round_to_significant_digits(self, num, sig_digits):
        if num != 0:
            return round(num, sig_digits - int(math.floor(math.log10(abs(num)))) - 1)
        return 0

    def apply_format_to_column(self, worksheet, column_index, df, format_type):
        formats = self.add_format(format_type)

        for row_num, value in enumerate(df.iloc[:, column_index], start=1):  # start=1 to skip header
            # Check for NaN or INF and write an empty cell
            if pd.isna(value) or value in [np.inf, -np.inf]:
                worksheet.write_string(row_num, column_index, '')
            else:
                # Remove commas from string values and attempt to convert to float or int
                if isinstance(value, str):
                    value = value.replace(',', '')  # Remove commas
                    try:
                        value = float(value) if '.' in value else int(value)
                    except ValueError:
                        pass  # If conversion fails, use the original string value

                # Convert decimal.Decimal to float
                elif isinstance(value, decimal.Decimal):
                    value = float(value)

                # Write the value as a number or string based on its type after conversion
                if isinstance(value, (int, float)):
                    if format_type == 'my_format':
                        # Apply different formats based on the length of the number
                        if value < 1000:
                            excel_format = formats['small']
                        elif value < 1000000:
                            excel_format = formats['thousands']
                        else:
                            excel_format = formats['millions']
                        worksheet.write_number(row_num, column_index, value, excel_format)
                    else:
                        worksheet.write_number(row_num, column_index, value, formats)
                else:
                    # If the value is not a number, write it as a string
                    worksheet.write_string(row_num, column_index, str(value))
