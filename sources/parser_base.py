import json
import os

import utils_array

# Folder that will contain generated files
output_dir = ''


def init(game_name: str, skin_id: str):
    global output_dir
    if os.name == 'posix':
        # save files to the game folder
        output_dir = '/home/ivan/sources/{}/resources/config/{}'.format(game_name, skin_id)
    else:
        # save files to tmp folder
        output_dir = skin_id


def prepare_out_directory_for_file(file_name):
    """
    Create directory for output files.
    """
    if not os.path.exists(output_dir) and output_dir != '':
        os.makedirs(output_dir, exist_ok=True)

    path = os.path.join(output_dir, file_name)

    # Debug log to find saved files
    print('Saving data to', path)

    return path


def create_game_settings_file(model: dict):
    """
    Generate `game_settings.json` according to `model`.
    """
    path = prepare_out_directory_for_file('game_settings.json')
    with open(path, 'w') as f:
        print(json.dumps(model, indent=4), file=f)


def create_strip_file(file_name: str, sheet, col_names: list, start_row_index: int, end_row_index=None):
    """
    Generate strip `file_name` using input parameters.
    """
    path = prepare_out_directory_for_file(file_name)
    with open(path, 'w') as f:
        strip = utils_array.get_complex_array_vertical(sheet, col_names, start_row_index, end_row_index, '\t')
        for row in strip:
            if row != '':
                print(row, file=f)
