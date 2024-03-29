import openpyxl

import sys

import utils_array
import parser_base
import utils_transform

game_name = sys.argv[1]
skin_id = sys.argv[2]
model_path = sys.argv[3]

wb = openpyxl.load_workbook(model_path, data_only=True)

game_settings_sheet = wb['Parameters']
strips_sheet = wb['Reels']


def get_pos_ranges(sheet, col_1, col_2, col_3, col_4, start_row, num_rows):
    r1 = utils_array.get_table(sheet, start_row, num_rows, col_1, col_2)
    r2 = utils_array.get_table(sheet, start_row, num_rows, col_3, col_4)
    return sum(r2, sum(r1, []))


def get_pos_ranges_table(sheet, col_1, col_2, col_3, col_4, start_row, num_rows=5, offset=7):
    return [get_pos_ranges(sheet, col_1, col_2, col_3, col_4, start_row + offset * i, num_rows) for i in range(5)]


game_settings_model = {
    "game": {
        "type": 2,
        "id": 27509,
        "machine_id": 27509
    },

    "fs_config": {
        "scatters_to_fs": [0, 0, 0, 8, 10, 12],
        "ranges": utils_array.get_table(game_settings_sheet, 357, 5, 'C', 'D'),
        "levels": utils_array.get_array_horizontal(game_settings_sheet, 138, 'C', 'F'),
    },

    "header": {
        "paid": {
            "values": [18, 19, 20],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 120, 'C', 'E'),
        },
        "free": {
            "values": [18, 19, 20],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 215, 'C', 'E'),
        }
    },

    "reels_option": {
        "paid": utils_array.get_array_horizontal(game_settings_sheet, 45, 'C', 'G'),
        "free": utils_array.get_table(game_settings_sheet, 129, 5, 'C', 'G'),
    },

    "pos_reels": {
        "type": {
            "paid": utils_array.get_table(game_settings_sheet, 86, 5, 'C', 'G'),
            "free": utils_array.get_table(game_settings_sheet, 181, 5, 'C', 'G'),
        },

        "ranges": {
            "paid": get_pos_ranges_table(game_settings_sheet, 'C', 'D', 'F', 'G', start_row=49),
            "free": get_pos_ranges_table(game_settings_sheet, 'C', 'D', 'F', 'G', start_row=144),
        }
    },

    "modifiers": {
        "mega_symbol": {
            "type": {
                "paid": utils_array.get_table(game_settings_sheet, 232, 5, 'C', 'F'),
                "free": utils_array.get_table(game_settings_sheet, 239, 5, 'C', 'F')
            },
            "max_symbol_type": [11, 11, 11],
            "symbol_weights": {
                "paid": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'M'),
                "free": utils_array.get_array_horizontal(game_settings_sheet, 255, 'C', 'M'),
            },
            "weight_s11_4x4": game_settings_sheet['D'][257].value,
        },
        "max_ways": {
            "type": {
                "paid": utils_array.get_table(game_settings_sheet, 266, 5, 'C', 'D'),
                "free": utils_array.get_table(game_settings_sheet, 273, 5, 'C', 'D')
            }
        },
        "power_block": {
            "type": {
                "paid": utils_array.get_table(game_settings_sheet, 285, 5, 'C', 'F'),
            }
        },
    },

    "power_block_multiplier": {
        "silver": {
            "paid": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 325, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 332, 'C', 'E'),
            },
            "free": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 325, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 337, 'C', 'E'),
            }
        },
        "gold": {
            "paid": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 326, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 333, 'C', 'E'),
            },
            "free": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 326, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 338, 'C', 'E'),
            }
        },
        "platinum": {
            "paid": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 327, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 334, 'C', 'E'),
            },
            "free": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 327, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 339, 'C', 'E'),
            }
        }
    },

    "place_power_blocks": {
        "reel_lose": {
            "paid": utils_array.get_table(game_settings_sheet, 95, 5, 'C', 'D'),
            "free": utils_array.get_table(game_settings_sheet, 190, 5, 'C', 'D'),
        },
        "reel_win": {
            "paid": utils_array.get_table(game_settings_sheet, 104, 5, 'C', 'D'),
            "free": utils_array.get_table(game_settings_sheet, 199, 5, 'C', 'D'),
        },
        "row": {
            "paid": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 112, 'C', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 113, 'C', 'L'),
            },
            "free": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 207, 'C', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 208, 'C', 'L'),
            }
        }
    },
}

parser_base.init(game_name, skin_id)

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_paid.strip_set', strips_sheet,
                              ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                               'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA'],
                              3, 102,
                              utils_transform.transform_inc)

parser_base.create_strip_file('1_free.strip_set', strips_sheet,
                              ['AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP',
                               'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC'],
                              3, 142,
                              utils_transform.transform_inc)
