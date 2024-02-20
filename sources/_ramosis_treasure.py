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


def get_array_row_by_row(sheet, cols: list, start_row: int, end_row):
    return [int(sheet[col][row].value) for row in range(start_row - 1, end_row) for col in cols]


def create_configuration(config):
    bg_strip_sheet = wb['Reelstrip_' + config + '_BG']
    bg_trigger_sheet = wb['Trigger_' + config + '_BG']
    fg_settings_sheet = wb['Feature_' + config + '_Settings']

    return {
        "paid": {
            "reel_strip": {
                "reward": utils_array.get_array_vertical(bg_strip_sheet, 'B', 7)
            },
            "trigger": {
                "events": utils_array.get_array_vertical(bg_trigger_sheet, 'B', 5),
                "odds": get_array_row_by_row(bg_trigger_sheet, ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'], 5, 60)
            }
        },
        "free": {
            "feature1": {
                "pos": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'F', 14),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'E', 14),
                },
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 18, 21),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 18, 21)
                }
            },
            "feature2": {

            },
            "feature3": {
                "reward": {
                    "values": [7, 8, 9, 4, 5, 6],
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'AB', 15, 20)
                }
            },
            "feature12": {
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 25, 28),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 25, 28)
                }
            },
            "feature13": {
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 32, 35),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 32, 35)
                },
                "reward": {
                    "values": [7, 8, 9, 4, 5, 6],
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'AB', 24, 29)
                }
            },
            "feature23": {

            },
            "feature123": {
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 39, 42),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 39, 42)
                }
            },
        }
    }


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
        },
        "LXF": create_configuration('LXF'),
        "LYF": create_configuration('LYF'),
        "HXF": create_configuration('HXF'),
        "HYF": create_configuration('HYF'),
    }
}

parser_base.init(game_name, skin_id)
parser_base.create_game_settings_file(game_settings_model)
create_all_strips_file('all_strips.strip_set')
