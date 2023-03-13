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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 'C', 29)
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

parser_base.create_strip_file('0_normal_paid.strip_set', strips_sheet, ['C', 'D', 'E', 'F', 'G'], 18)
parser_base.create_strip_file('1_normal_wild_mult.strip_set', strips_sheet, ['K', 'L', 'M', 'N', 'O'], 18)
parser_base.create_strip_file('2_normal_wild_reels.strip_set', strips_sheet, ['S', 'T', 'U', 'V', 'W'], 18)
parser_base.create_strip_file('3_side_bet_paid.strip_set', strips_sheet, ['AA', 'AB', 'AC', 'AD', 'AE'], 18)
