import openpyxl

import sys

import parser_base
import utils_array

game_name = sys.argv[1]
skin_id = sys.argv[2]
model_path = sys.argv[3]

wb = openpyxl.load_workbook(model_path, data_only=True)
game_configs = ['LXF', 'LYF', 'HXF', 'HYF']


def create_all_strips_file(file_name):
    # consts
    columns_with_strips = ['H', 'J', 'L', 'N', 'P']
    sheets_with_strips = [
        'Reelstrip_$_BG',
        'Reelstrip_$_FG_1', 'Reelstrip_$_FG_2', 'Reelstrip_$_FG_3',
        'Reelstrip_$_FG_12', 'Reelstrip_$_FG_13', 'Reelstrip_$_FG_23',
        'Reelstrip_$_FG_123'
    ]
    num_sections = 9
    num_rows = 58
    rows_offset = 7

    sheets = [wb[mask.replace('$', config)] for config in game_configs for mask in sheets_with_strips]

    # create all_strips
    all_strips = []
    for column in columns_with_strips:
        for sheet in sheets:
            row = rows_offset
            for section in range(num_sections):
                values = utils_array.get_array_vertical(sheet, column, row, row + num_rows - 1)
                all_strips.append(values)
                row += num_rows

    # print to file
    path = parser_base.prepare_out_directory_for_file(file_name)
    with open(path, 'w') as f:
        for i in range(len(all_strips[0])):
            for j in range(len(all_strips)):
                print(str(all_strips[j][i]) + '\t', file=f, end='')
            print(file=f)


game_settings_model = {
    "game": {
        "type": 2,
        "id": 31289,
        "machine_id": 31289
    },
    "bonus_buy_enabled": True,
    "path_settings_math": "/b2b-real-games/ngs/weights/201155/settings_math.json",
    "normal_bet": {
        "config_type": {
            "values": [
                0,
                1,
                2,
                3
            ],
            "weights": utils_array.get_array_horizontal(wb['First draw weights'], 6, 'B', 'E')
        }
    }
}

parser_base.init(game_name, skin_id)
parser_base.create_game_settings_file(game_settings_model)
create_all_strips_file('all_strips.strip_set')
