from xlsxwriter import Workbook
from pandas import DataFrame
from typing import Any, List, Dict, Optional
import pandas as pd
import logging

bands = [
    (0, 2000000, 50), 
    (2000000, 4000000, 100), 
    (4000000, 6000000, 150), 
    (6000000, 8000000, 200),
    (8000000, float('inf'), 255)
]

class ConditionalFormatting:
    def __init__(self, df: DataFrame, workbook: Workbook, worksheet: Any, verbose_map: Any, is_pivot: bool = False):
        self.df = df
        self.workbook = workbook
        self.worksheet = worksheet
        self.verbose_map = verbose_map
        self.is_pivot = is_pivot

    def _generate_dynamic_mappings(self, conditional_formatting_array: List[Dict[str, Any]]):
        original_verbose_map = self.verbose_map
        pivot_columns = self.df.columns  

        new_verbose_map = {}
        new_conditional_formatting_array = []

        # Extract desired columns and their settings from original conditional_formatting_array
        original_columns_settings = {item['column']: item for item in conditional_formatting_array}

        for original_column, verbose_name in original_verbose_map.items():
            # If the original_column is not in the settings, we don't process it
            if original_column not in original_columns_settings:
                continue
            
            # Check if the original verbose name is a substring in any pivot column
            matching_pivot_columns = [col for col in pivot_columns if verbose_name in col]

            # Only consider columns that are not exact matches
            matching_pivot_columns = [col for col in matching_pivot_columns if col != verbose_name]

            for pivot_column in matching_pivot_columns:
                new_verbose_map[f"{original_column} {pivot_column.split(' ')[-1]}"] = pivot_column

                # Extract original settings for the column
                orig_settings = original_columns_settings[original_column]

                # Add to new_conditional_formatting_array
                new_conditional_formatting_entry = {
                    'colorScheme': orig_settings['colorScheme'], 
                    'column': pivot_column, 
                    'operator': orig_settings['operator']
                }
                # Optionally add targetValue, targetValueLeft, and targetValueRight if they are present
                for optional_key in ['targetValue', 'targetValueLeft', 'targetValueRight']:
                    if optional_key in orig_settings:
                        new_conditional_formatting_entry[optional_key] = orig_settings[optional_key]

                new_conditional_formatting_array.append(new_conditional_formatting_entry)

        return new_verbose_map, new_conditional_formatting_array
        
    def generate(self, conditional_formatting_array: List[Dict[str, Any]]):
        if self.is_pivot:
            self.verbose_map, conditional_formatting_array = self._generate_dynamic_mappings(conditional_formatting_array)

        for condition in conditional_formatting_array:
            column = condition['column']
            color = condition['colorScheme']
            operator = condition['operator']
            mapped_column_name = self.verbose_map.get(column, column)
            col_idx = self.df.columns.get_loc(mapped_column_name)

            if operator == "None":
                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}")    

                min_val = float(self.df[mapped_column_name].min())
                max_val = float(self.df[mapped_column_name].max())
            
                scale = max_val - min_val if max_val != min_val else 1# make sure max_val is defined and appropriate

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value) or pd.isna(min_val):
                        continue
                    value = float(value)  # ensuring value is float
                    intensity = int(255 * ((value - min_val) / scale))

                    cell_format = self.workbook.add_format()
                    cell_format.set_pattern(1)
                    cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == ">":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator ">". Skipping...')
                    continue
                
                mapped_column_name = self.verbose_map.get(column, column)

                try:
                    for idx, item in self.df[mapped_column_name].iteritems():
                        try:
                            # Attempt to remove commas and convert to a number
                            self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                        except AttributeError:
                            # Handle non-string data - log, pass, or handle as appropriate
                            logging.debug(f"Non-string data at index {idx}: {item}")                    
                    
                    applicable_values = self.df[self.df[mapped_column_name] > target_value][mapped_column_name]
                except TypeError as e:
                    logging.error(f"TypeError encountered: {str(e)}")
                    logging.error(f"Skipping conditional formatting for {mapped_column_name}. Ensure that data types are consistent.")
                    continue
                if applicable_values.empty:
                    logging.warning(f"No values found larger than {target_value} in column {mapped_column_name}. Skipping...")
                    continue

                
                max_val = float(applicable_values.max())
                min_val = float(applicable_values.min())
                scale = max_val - min_val if max_val != min_val else 1

                # Lighter limit - the smaller the value, the lighter the minimum color will be.
                lightest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()
                    if value > target_value:
                        # Your existing code for setting color for values > target_value
                        intensity = lightest_color_limit + int((255 - lightest_color_limit) * ((value - min_val) / scale))
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))

                    # Writing the (numerical) value to Excel, using the (possibly modified) format
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "<":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator "<". Skipping...')
                    continue
                mapped_column_name = self.verbose_map.get(column, column)
                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}")    

                applicable_values = self.df[self.df[mapped_column_name] < target_value][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found smaller than {target_value} in column {mapped_column_name}. Skipping...")
                    continue
                
                max_val = float(applicable_values.max())
                min_val = float(applicable_values.min())
                scale = max_val - min_val if max_val != min_val else 1

                # Lighter limit - the smaller the value, the lighter the maximum color will be.
                lightest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value) or pd.isna(min_val):
                        continue
                    value = float(value) 
                    cell_format = self.workbook.add_format()
                    if value < target_value:
                        intensity = lightest_color_limit + int((255 - lightest_color_limit) * (1 - (value - min_val) / scale))
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "≥":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator "≥". Skipping...')
                    continue
                
                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}")    

                applicable_values = self.df[self.df[mapped_column_name] >= target_value][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found larger or equal to {target_value} in column {mapped_column_name}. Skipping...")
                    continue
                
                max_val = float(applicable_values.max())
                min_val = float(applicable_values.min())
                scale = max_val - min_val if max_val != min_val else 1

                # Lighter limit - the smaller the value, the lighter the minimum color will be.
                lightest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value) or pd.isna(min_val):
                        continue
                    value = float(value) 
                    cell_format = self.workbook.add_format()

                    if value >= target_value:
                        # Ensure higher values are darker and adjust intensity to not get too light.
                        intensity = lightest_color_limit + int((255 - lightest_color_limit) * ((value - min_val) / scale))
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "≤":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator "≤". Skipping...')
                    continue
                
                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                applicable_values = self.df[self.df[mapped_column_name] <= target_value][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found smaller or equal to {target_value} in column {mapped_column_name}. Skipping...")
                    continue
                
                max_val = float(applicable_values.max())
                min_val = float(applicable_values.min())
                scale = max_val - min_val if max_val != min_val else 1

                # Lighter limit - the smaller the value, the lighter the maximum color will be.
                lightest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value) or pd.isna(min_val):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()
                    if value <= target_value:
                        # Ensure lower values are darker and adjust intensity to not get too light.
                        intensity = lightest_color_limit + int((255 - lightest_color_limit) * (1 - (value - min_val) / scale))
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "=":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator "=". Skipping...')
                    continue
                
                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()
                    if value == target_value:
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(color)  # Using provided colorScheme directly.
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "≠":
                target_value = condition.get('targetValue')
                if not target_value:
                    logging.warning('Target value is not provided for operator "≠". Skipping...')
                    continue

                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                applicable_values = self.df[self.df[mapped_column_name] != target_value][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found not equal to {target_value} in column {mapped_column_name}. Skipping...")
                    continue
                
                max_val = float(applicable_values.max())
                min_val = float(applicable_values.min())
                scale = max(max_val - target_value, target_value - min_val)

                # Darker limit - the further the value from the target, the darker the maximum color will be.
                darkest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()
                    if value != target_value:
                        # Ensure values furthest from the target are darker.
                        intensity = darkest_color_limit + int((255 - darkest_color_limit) * abs(value - target_value) / scale)
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "< x <":
                target_value_left = condition.get('targetValueLeft')
                target_value_right = condition.get('targetValueRight')
                if not target_value_left or not target_value_right:
                    logging.warning('Target values are not provided for operator "< x <". Skipping...')
                    continue
                
                target_value_left, target_value_right = float(target_value_left), float(target_value_right)

                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                applicable_values = self.df[(self.df[mapped_column_name] > target_value_left) & (self.df[mapped_column_name] < target_value_right)][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found between {target_value_left} and {target_value_right} in column {mapped_column_name}. Skipping...")
                    continue
                
                scale = target_value_right - target_value_left if target_value_right != target_value_left else 1

                # Darker limit - the closer the value to the targetValueRight, the darker the maximum color will be.
                darkest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()
                    if target_value_left < value < target_value_right:
                        # Ensure values closer to the targetValueLeft are lighter.
                        intensity = darkest_color_limit + int((255 - darkest_color_limit) * (value - target_value_left) / scale)
                        cell_format = self.workbook.add_format()
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "≤ x ≤":
                target_value_left = condition.get('targetValueLeft')
                target_value_right = condition.get('targetValueRight')
                if not target_value_left or not target_value_right:
                    logging.warning('Target values are not provided for operator "≤ x ≤". Skipping...')
                    continue
                
                target_value_left, target_value_right = float(target_value_left), float(target_value_right)

                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                applicable_values = self.df[(self.df[mapped_column_name] >= target_value_left) & (self.df[mapped_column_name] <= target_value_right)][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found between {target_value_left} and {target_value_right} in column {mapped_column_name}. Skipping...")
                    continue
                
                scale = target_value_right - target_value_left if target_value_right != target_value_left else 1

                # Darker limit - the closer the value to the targetValueRight, the darker the maximum color will be.
                darkest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    cell_format = self.workbook.add_format()
                    value = float(value)
                    if target_value_left <= value <= target_value_right:
                        # Ensure values closer to the targetValueLeft are lighter.
                        intensity = darkest_color_limit + int((255 - darkest_color_limit) * (value - target_value_left) / scale)
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "≤ x <":
                target_value_left = condition.get('targetValueLeft')
                target_value_right = condition.get('targetValueRight')
                if not target_value_left or not target_value_right:
                    logging.warning('Target values are not provided for operator "≤ x <". Skipping...')
                    continue
                
                target_value_left, target_value_right = float(target_value_left), float(target_value_right)

                mapped_column_name = self.verbose_map.get(column, column)


                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                
                applicable_values = self.df[(self.df[mapped_column_name] >= target_value_left) & (self.df[mapped_column_name] < target_value_right)][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found between {target_value_left} and {target_value_right} in column {mapped_column_name}. Skipping...")
                    continue
                
                scale = target_value_right - target_value_left if target_value_right != target_value_left else 1
                
                # Darker limit - the closer the value to the targetValueRight, the darker the maximum color will be.
                darkest_color_limit = 100
                
                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue

                    cell_format = self.workbook.add_format()
                    value = float(value)
                    if target_value_left <= value < target_value_right:
                        # Ensure values closer to the targetValueLeft are lighter.
                        intensity = darkest_color_limit + int((255 - darkest_color_limit) * (value - target_value_left) / scale)
                        cell_format = self.workbook.add_format()
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            elif operator == "< x ≤":
                target_value_left = condition.get('targetValueLeft')
                target_value_right = condition.get('targetValueRight')
                if not target_value_left or not target_value_right:
                    logging.warning('Target values are not provided for operator "< x ≤". Skipping...')
                    continue
                
                target_value_left, target_value_right = float(target_value_left), float(target_value_right)

                mapped_column_name = self.verbose_map.get(column, column)

                for idx, item in self.df[mapped_column_name].iteritems():
                    try:
                        # Attempt to remove commas and convert to a number
                        self.df.at[idx, mapped_column_name] = pd.to_numeric(item.replace(',', ''), errors='coerce')
                    except AttributeError:
                        # Handle non-string data - log, pass, or handle as appropriate
                        logging.debug(f"Non-string data at index {idx}: {item}") 

                applicable_values = self.df[(self.df[mapped_column_name] > target_value_left) & (self.df[mapped_column_name] <= target_value_right)][mapped_column_name]
                if applicable_values.empty:
                    logging.warning(f"No values found between {target_value_left} and {target_value_right} in column {mapped_column_name}. Skipping...")
                    continue
                
                scale = target_value_right - target_value_left if target_value_right != target_value_left else 1

                # Darker limit - the closer the value to the targetValueRight, the darker the maximum color will be.
                darkest_color_limit = 100

                for idx, value in enumerate(self.df[mapped_column_name]):
                    if pd.isna(value):
                        continue
                    value = float(value)
                    cell_format = self.workbook.add_format()

                    if target_value_left < value <= target_value_right:
                        # Ensure values closer to the targetValueLeft are lighter.
                        intensity = darkest_color_limit + int((255 - darkest_color_limit) * (value - target_value_left) / scale)
                        cell_format.set_pattern(1)
                        cell_format.set_bg_color(self._get_color_intensity(color, intensity))
                    self.worksheet.write(idx + 1, col_idx, value, cell_format)

            else:
                # Additional logic for other operators can be added here.
                pass
    
    def _get_color_intensity(self, hex_color: str, intensity: int, is_pivot: bool = False) -> str:
        # Ensure intensity is between 0 and 255
        intensity = max(0, min(255, intensity))

        # Extract RGB components from the hex color code
        r_orig = int(hex_color[1:3], 16)
        g_orig = int(hex_color[3:5], 16)
        b_orig = int(hex_color[5:7], 16)

        # Adjust intensity if it's a pivot
        if is_pivot:
            intensity_factor = 0.5  # You can adjust this factor as needed
            intensity = int(intensity * intensity_factor)

        # Adjust each RGB component
        r = min(255, r_orig + (255 - r_orig) * (1 - intensity / 255))
        g = min(255, g_orig + (255 - g_orig) * (1 - intensity / 255))
        b = min(255, b_orig + (255 - b_orig) * (1 - intensity / 255))

        # Construct new hex color code
        new_hex_color = '#{:02X}{:02X}{:02X}'.format(int(r), int(g), int(b))

        return new_hex_color
