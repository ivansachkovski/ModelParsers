from utils_transform import *


def get_array_vertical(sheet, col_name: str, start_row_index: int, transform_func=transform_none):
    """
    Get as array column `col_name` from `sheet` starting from `start_index_row`.
    Apply `transform_func` to each element.
    """
    out_array = []
    column = sheet[col_name]

    index = 0
    for cell in column:
        index += 1

        if index < start_row_index:
            continue

        if cell.value is None:
            break

        value = transform_func(cell.value)
        out_array.append(value)

    return out_array


def get_complex_array_vertical(sheet, col_names: list, start_row_index: int, separator='|'):
    """
    Get as array joined (via `separator`) columns belongs to `col_names` from `sheet` starting from `start_index_row`.
    """
    out_array = []

    is_first_column = True

    for col_name in col_names:
        column = sheet[col_name]

        index = 0
        for cell in column:
            index += 1

            if index < start_row_index:
                continue

            output_index = index - start_row_index

            if output_index >= len(out_array):
                out_array.append('')

            if cell.value is None:
                continue

            if not is_first_column:
                out_array[output_index] += separator

            out_array[output_index] += str(cell.value)

        is_first_column = False

    return out_array


def get_array_horizontal(sheet, start_col_name: str, row_index: int):
    pass
