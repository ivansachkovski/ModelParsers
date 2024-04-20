import openpyxl

import sys

import parser_base
import utils_array

game_name = sys.argv[1]
skin_id = sys.argv[2]
model_path = sys.argv[3]

wb = openpyxl.load_workbook(model_path, data_only=True)


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

    sheets = [wb[mask.replace('$', config)] for config in ['LXF', 'LYF', 'HXF', 'HYF'] for mask in sheets_with_strips]

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


def get_array_row_by_row(sheet, cols: list, start_row: int, end_row: int):
    if start_row <= end_row:
        start_row -= 1
        step = 1
    else:
        start_row -= 1
        end_row -= 2
        step = -1

    return [int(sheet[col][row].value) for row in range(start_row, end_row, step) for col in cols]


def get_array_col_by_col(sheet, cols: list, start_row: int, end_row: int):
    if start_row <= end_row:
        start_row -= 1
        step = 1
    else:
        start_row -= 1
        end_row -= 2
        step = -1

    return [int(sheet[col][row].value) for col in cols for row in range(start_row, end_row, step)]


def create_configuration(config):
    bg_strip_sheet = wb['Reelstrip_' + config + '_BG']
    bg_trigger_sheet = wb['Trigger_' + config + '_BG']
    fg_trigger_sheet = wb['Trigger_' + config + '_FG']
    fg_settings_sheet = wb['Feature_' + config + '_Settings']

    pos_odds_list = ['T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
                     'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                     'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP',
                     'BQ']
    assert (len(pos_odds_list) == 5 * 10)

    return {
        "paid": {
            "reel_strip": {
                "reward": utils_array.get_array_vertical(bg_strip_sheet, 'B', 7, 96),
                "odds": get_array_col_by_col(bg_strip_sheet, pos_odds_list, 7, 528)
            },
            "trigger": {
                "odds": get_array_row_by_row(bg_trigger_sheet, ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'], 5, 60)
            }
        },
        "free": {
            "retrigger": {
                "odds_1": get_array_row_by_row(fg_trigger_sheet, ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'], 5, 60),
                "odds_3": get_array_row_by_row(fg_trigger_sheet, ['P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W'], 5, 60),
                "odds_13": get_array_row_by_row(fg_trigger_sheet, ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'], 5,
                                                60),
            },
            "feature1": {
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_1'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_1'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_1'], pos_odds_list, 123, 180)
                },
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
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_2'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_2'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_2'], pos_odds_list, 123, 180)
                },
                "set": utils_array.get_array_vertical(fg_settings_sheet, 'H', 14, 16),
                "rewards": [0, 1, 2, 3, 4, 5, 6],
                "reward_odds": [
                    get_array_row_by_row(fg_settings_sheet, ['J', 'K', 'L', 'M', 'N', 'O', 'P'], 85, 66),
                    get_array_row_by_row(fg_settings_sheet, ['J', 'K', 'L', 'M', 'N', 'O', 'P'], 62, 43),
                    get_array_row_by_row(fg_settings_sheet, ['J', 'K', 'L', 'M', 'N', 'O', 'P'], 39, 20),
                ]
            },
            "feature3": {
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_3'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_3'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_3'], pos_odds_list, 123, 180)
                },
                "reward": {
                    "values": [7, 8, 9, 4, 5, 6],
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'AB', 15, 20)
                }
            },
            "feature12": {
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_12'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_12'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_12'], pos_odds_list, 123, 180)
                },
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 25, 28),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 25, 28)
                },
                "set": utils_array.get_array_vertical(fg_settings_sheet, 'R', 14, 16),
                "rewards": [0, 1, 2, 3, 4, 5, 6],
                "reward_odds": [
                    get_array_row_by_row(fg_settings_sheet, ['T', 'U', 'V', 'W', 'X', 'Y', 'Z'], 85, 66),
                    get_array_row_by_row(fg_settings_sheet, ['T', 'U', 'V', 'W', 'X', 'Y', 'Z'], 62, 43),
                    get_array_row_by_row(fg_settings_sheet, ['T', 'U', 'V', 'W', 'X', 'Y', 'Z'], 39, 20),
                ]
            },
            "feature13": {
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_13'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_13'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_13'], pos_odds_list, 123, 180)
                },
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
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_23'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_23'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_23'], pos_odds_list, 123, 180)
                },
                "set": utils_array.get_array_vertical(fg_settings_sheet, 'AH', 14, 16),
                "rewards": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
                "reward_odds": [
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS'], 85, 66),
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS'], 62, 43),
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS'], 39, 20),
                ]
            },
            "feature123": {
                "reel_strip": {
                    "reward": utils_array.get_array_vertical(wb['Reelstrip_' + config + '_FG_123'], 'B', 7, 96),
                    "odds_normal": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_123'], pos_odds_list, 7, 64),
                    "odds_super": get_array_col_by_col(wb['Reelstrip_' + config + '_FG_123'], pos_odds_list, 123, 180)
                },
                "mult": {
                    "values": utils_array.get_array_vertical(fg_settings_sheet, 'C', 39, 42),
                    "weights": utils_array.get_array_vertical(fg_settings_sheet, 'B', 39, 42)
                },
                "set": utils_array.get_array_vertical(fg_settings_sheet, 'AU', 14, 16),
                "rewards": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
                "reward_odds": [
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF'], 85, 66),
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF'], 62, 43),
                    get_array_row_by_row(fg_settings_sheet,
                                         ['AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF'], 39, 20),
                ]
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
    "common": {
        "trigger_events": utils_array.get_array_vertical(wb['Trigger_LXF_BG'], 'B', 5),
    },
    "normal_bet": {
        "config_type": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_horizontal(wb['First draw weights'], 6, 'B', 'E')
        },
        "LXF": create_configuration('LXF'),
        "LYF": create_configuration('LYF'),
        "HXF": create_configuration('HXF'),
        "HYF": create_configuration('HYF'),
    },
    "bonus_buy": {
        "game_mode_odds_1": [0, 200, 180, 230, 120, 115, 115, 40],
        "game_mode_odds_2": [0, 0, 0, 0, 0, 0, 0, 1]
    }
}

parser_base.init(game_name, skin_id)
parser_base.create_game_settings_file(game_settings_model)
create_all_strips_file('all_strips.strip_set')
