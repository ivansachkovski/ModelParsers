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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 'C29', 'G29')
        },
        "paid_spins": {
            "reels_odds": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 'C46', 'F46'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 'C47', 'F47'),
            }
        },
        "free_spins": {
            "reels_odds": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 'C57', 'F57'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 'C58', 'F58'),
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

parser_base.create_strip_file('0_paid_spins.strip_set', strips_sheet, ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
                                                                       'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V',
                                                                       'W', 'X', 'Y', 'Z'], 3, 142)
parser_base.create_strip_file('1_free_spins.strip_set', strips_sheet, ['AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK',
                                                                       'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS',
                                                                       'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA'],
                              3, 142)
