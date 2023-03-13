import utils_transform


def get_array_vertical(sheet, col_name: str, start_row_index: int, transform_func=utils_transform.transform_none):
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


def get_complex_array_vertical(sheet, col_names: list, start_row_index: int, end_row_index=None, separator='|'):
    """
    Get as array joined (via `separator`) columns belongs to `col_names` from `sheet` starting from `start_index_row`.
    """
    out_array = []

    for col_name in col_names:

        index = 0
        column = sheet[col_name]

        for cell in column:
            index += 1

            # Do not process rows before the start one
            if index < start_row_index:
                continue

            # Do not process rows after the last one (if it configured)
            if end_row_index is not None and index > end_row_index:
                break

            output_index = index - start_row_index

            if output_index >= len(out_array):
                out_array.append('')

            # Do not process cells without values
            if cell.value is None:
                continue

            if out_array[output_index] != '':
                out_array[output_index] += separator

            out_array[output_index] += str(cell.value)

    return out_array


def get_array_horizontal(sheet, row: int, col_name_begin: str, col_name_end: str):

    range_begin = col_name_begin + str(row)
    range_end = col_name_end + str(row)

    return [item.value for item in sheet[range_begin:range_end][0]]
