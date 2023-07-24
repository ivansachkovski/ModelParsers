import utils_transform


def get_array_vertical(sheet,
                       col_name: str,
                       start_row_index: int,
                       transform_func=None):
    """
    Get as array column `col_name` from `sheet` starting from `start_index_row`.
    Apply `transform_func` to each element.
    """
    assert start_row_index > 0
    start_row_index -= 1

    transform_func = transform_func or utils_transform.transform_none
    column = sheet[col_name][start_row_index:]

    return [transform_func(cell.value) for cell in column if cell.value is not None]


def get_complex_array_vertical(sheet,
                               col_names: list,
                               start_row_index: int,
                               end_row_index: int = None,
                               separator: str = '|',
                               transform_func=None):
    """
    Get as array joined (via `separator`) columns belongs to `col_names` from `sheet` starting from `start_index_row`.
    """
    assert start_row_index > 0
    start_row_index -= 1

    transform_func = transform_func or utils_transform.transform_none

    assert len(col_names) > 0
    max_num_rows = len(sheet[col_names[0]][start_row_index:end_row_index])
    out_array = ['' for _ in range(max_num_rows)]

    for col_name in col_names:

        index = -1
        column = sheet[col_name][start_row_index:end_row_index]

        for cell in column:
            index += 1
            # Process only cells with values
            if cell.value is not None:
                if out_array[index] != '':
                    out_array[index] += separator
                out_array[index] += str(transform_func(cell.value))

    return [value for value in out_array if value != '']


def get_array_horizontal(sheet,
                         row: int,
                         col_name_begin: str,
                         col_name_end: str):

    range_begin = col_name_begin + str(row)
    range_end = col_name_end + str(row)

    return [item.value for item in sheet[range_begin:range_end][0]]


def get_table(sheet,
              start_row: int,
              num_rows: int,
              col_name_begin: str,
              col_name_end: str):
    return [get_array_horizontal(sheet, row, col_name_begin, col_name_end) for row in range(start_row, start_row + num_rows)]
