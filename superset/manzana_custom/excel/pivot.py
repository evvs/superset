from typing import Union
from superset.models.slice import Slice

def is_pivot(slice: Union[Slice, None]) -> bool:
    if slice and slice.form_data.get("viz_type") == 'pivot_table_v2':
        return True
    return False