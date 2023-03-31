import openpyxl

import utils_array
import utils_transform
import parser_base

model_path = '../../../math_models/Server_HenHeist_v1.4.xlsx'

wb = openpyxl.load_workbook(model_path, data_only=True)

server_sheet = wb['Server']
server_bonus_buy_sheet = wb['Server Bonus Buy']
server_strips_sheet = wb['Server Strips']

game_settings_model = {
    "game": {
        "type": 2,
        "id": 21552,
        "machine_id": 21552
    },
    "common": {
        "num_non_triggering_scatters_odds": {
            "values": [0, 1, 2, 3, 4],
            "weights": [489, 250, 166, 55, 40]
        },
        "golden_egg_paid_odds": {
            "values": [0, 1],
            "weights": [99, 1]
        },
        "golden_egg_respins_odds": {
            "values": [0, 1],
            "weights": [99, 1]
        },
        "respins_losing_rows_odds": {
            "values": ["1:0:0", "3:0:0", "1:1:0", "3:3:0", "1:3:0"],
            "weights": [1, 1, 1, 1, 1]
        }
    },
    "normal_bet": {
        "game_modes_odds": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_vertical(server_sheet, 'C', 5)
        },
        "mod_wild_multiplier": {
            "multipliers_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'I', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'J', 6)
            },
            "placements_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'L', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'M', 6)
            }
        },
        "mod_wild_reels": {
            "placements_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'P', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'Q', 6)
            }
        },
        "respins": {
            "num_triggering_scatters_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'V', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'W', 6)
            },
            "bronze_multipliers_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'Y', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'Z', 6)
            },
            "scenarios_odds": {
                "values": utils_array.get_complex_array_vertical(server_sheet,
                                                                 ['AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
                                                                  'AK'], 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'AL', 6)
            }
        }
    },
    "side_bet": {
        "game_modes_odds": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_vertical(server_sheet, 'AO', 6)
        },
        "mod_wild_multiplier": {
            "multipliers_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'AU', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'AV', 6)
            },
            "placements_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'AX', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'AY', 6)
            }
        },
        "mod_wild_reels": {
            "placements_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'BB', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'BC', 6)
            }
        },
        "respins": {
            "num_triggering_scatters_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'BH', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'BI', 6)
            },
            "bronze_multipliers_odds": {
                "values": utils_array.get_array_vertical(server_sheet, 'BK', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'BL', 6)
            },
            "scenarios_odds": {
                "values": utils_array.get_complex_array_vertical(server_sheet,
                                                                 ['BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV',
                                                                  'BW'], 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'BX', 6)
            }
        }
    },
    "bonus_buy": {
        "respins": {
            "num_triggering_scatters_odds": {
                "values": utils_array.get_array_vertical(server_bonus_buy_sheet, 'D', 6),
                "weights": utils_array.get_array_vertical(server_bonus_buy_sheet, 'E', 6)
            },
            "bronze_multipliers_odds": {
                "values": utils_array.get_array_vertical(server_bonus_buy_sheet, 'G', 6),
                "weights": utils_array.get_array_vertical(server_bonus_buy_sheet, 'H', 6)
            },
            "scenarios_odds": {
                "values": utils_array.get_complex_array_vertical(server_bonus_buy_sheet,
                                                                 ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'], 6),
                "weights": utils_array.get_array_vertical(server_bonus_buy_sheet, 'T', 6)
            }
        }
    },
    "side_bet_enabled": True,
    "bonus_buy_enabled": True,
    "win_cap_total_bet_mult": 20000
}

parser_base.init('ngs_barnyard_bucks', '200967')

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_normal_paid.strip_set', server_strips_sheet, ['C', 'D', 'E', 'F', 'G'], 18)
parser_base.create_strip_file('1_normal_wild_mult.strip_set', server_strips_sheet, ['K', 'L', 'M', 'N', 'O'], 18)
parser_base.create_strip_file('2_normal_wild_reels.strip_set', server_strips_sheet, ['S', 'T', 'U', 'V', 'W'], 18)
parser_base.create_strip_file('3_side_bet_paid.strip_set', server_strips_sheet, ['AA', 'AB', 'AC', 'AD', 'AE'], 18)
