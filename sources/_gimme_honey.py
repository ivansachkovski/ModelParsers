import openpyxl

import utils_array
import parser_base

model_path = '../../../math_models/GimmeHoney_11_06.03.2023.xlsx'

wb = openpyxl.load_workbook(model_path, data_only=True)

game_settings_sheet = wb['Parameters']
strips_sheet = wb['Reels']

game_settings_model = {
    "game": {
        "type": 2,
        "id": 26433,
        "machine_id": 26433
    },
    "common": {
    },
    "normal_bet": {
        "game_modes_odds": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 29, 'C', 'G'),
        },
        "paid_spins": {
            "reels_odds": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 47, 'C', 'F'),
            },
            "tracker_odds": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 47, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 48, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 49, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 50, 'I', 'L'),
                }
            ],
            "reels_heights_odds": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 47, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 48, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 49, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 50, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 51, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 52, 'Q', 'V'),
                }
            ]
        },
        "free_spins": {
            "reels_odds": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'C', 'F'),
            }
        }
    },
    "bonus_buy_x75": {
    },
    "bonus_buy_x150": {
    },
    "bonus_buy_enabled": True,
    "win_cap_total_bet_mult": 5000
}

parser_base.init('ngs_gimme_honey', '200993')

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_paid_spins.strip_set', strips_sheet,
                              ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                               'U', 'V', 'W', 'X', 'Y', 'Z'], 3, 142)
parser_base.create_strip_file('1_free_spins.strip_set', strips_sheet,
                              ['AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR',
                               'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA'], 3, 142)
