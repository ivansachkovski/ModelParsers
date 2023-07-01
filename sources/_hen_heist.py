import openpyxl

import sys

import utils_array
import utils_transform
import parser_base

game_name = sys.argv[1]
skin_id = sys.argv[2]
model_path = sys.argv[3]

wb = openpyxl.load_workbook(model_path, data_only=True)

# Define sheets
server_sheet = wb['Server']
server_bonus_buy_sheet = wb['Server Bonus Buy']
server_strips_sheet = wb['Server Strips']

# Define custom files names
rs_scenarios = "/b2b-real-games/ngs/weights/" + skin_id + "/rs_scenarios.data"
rs_scenarios_odds_normal_bet = "/b2b-real-games/ngs/weights/" + skin_id + "/rs_scenarios_odds_normal_bet.data"
rs_scenarios_odds_side_bet = "/b2b-real-games/ngs/weights/" + skin_id + "/rs_scenarios_odds_side_bet.data"
rs_scenarios_odds_bonus_buy = "/b2b-real-games/ngs/weights/" + skin_id + "/rs_scenarios_odds_bonus_buy.data"

# Extract scenarios
normal_bet_scenarios = utils_array.get_complex_array_vertical(server_sheet,
                                                              ['AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
                                                               'AK'], 6)
side_bet_scenarios = utils_array.get_complex_array_vertical(server_sheet,
                                                            ['BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV',
                                                             'BW'], 6)
bonus_buy_scenarios = utils_array.get_complex_array_vertical(server_bonus_buy_sheet,
                                                             ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'], 6)

# All the scenarios must be the same
assert normal_bet_scenarios == side_bet_scenarios and normal_bet_scenarios == bonus_buy_scenarios

game_settings_model = {
    "game": {
        "type": 2,
        "id": 21552,
        "machine_id": 21552
    },
    "common": {
        "num_non_triggering_scatters": {
            "values": [0, 1, 2, 3, 4],
            "weights": [509, 250, 166, 55, 20]
        },
        "golden_egg_respins": {
            "values": [0, 1],
            "weights": [2, 1]
        },
        "respins_losing_rows": {
            "values": ["1:0:0", "1:1:0"],
            "weights": [1, 1]
        }
    },
    "normal_bet": {
        "game_modes": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_vertical(server_sheet, 'C', 5)
        },
        "mod_wild_normal": {
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'E', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'F', 6)
            }
        },
        "mod_wild_multiplier": {
            "multipliers": {
                "values": utils_array.get_array_vertical(server_sheet, 'I', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'J', 6)
            },
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'L', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'M', 6)
            }
        },
        "mod_wild_reels": {
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'P', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'Q', 6)
            }
        },
        "respins": {
            "num_triggering_scatters": {
                "values": utils_array.get_array_vertical(server_sheet, 'V', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'W', 6)
            },
            "up_booster_multipliers": {
                "values": utils_array.get_array_vertical(server_sheet, 'Y', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'Z', 6)
            },
            "scenarios": {
                "file_scenarios": rs_scenarios,
                "file_scenarios_odds": rs_scenarios_odds_normal_bet,
                "num_scenarios": len(normal_bet_scenarios)
            }
        }
    },
    "side_bet": {
        "game_modes": {
            "values": [0, 1, 2, 3],
            "weights": utils_array.get_array_vertical(server_sheet, 'AO', 6)
        },
        "mod_wild_normal": {
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'AQ', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'AR', 6)
            }
        },
        "mod_wild_multiplier": {
            "multipliers": {
                "values": utils_array.get_array_vertical(server_sheet, 'AU', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'AV', 6)
            },
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'AX', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'AY', 6)
            }
        },
        "mod_wild_reels": {
            "placements": {
                "values": utils_array.get_array_vertical(server_sheet, 'BB', 6, utils_transform.transform_remove_first),
                "weights": utils_array.get_array_vertical(server_sheet, 'BC', 6)
            }
        },
        "respins": {
            "num_triggering_scatters": {
                "values": utils_array.get_array_vertical(server_sheet, 'BH', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'BI', 6)
            },
            "up_booster_multipliers": {
                "values": utils_array.get_array_vertical(server_sheet, 'BK', 6),
                "weights": utils_array.get_array_vertical(server_sheet, 'BL', 6)
            },
            "scenarios": {
                "file_scenarios": rs_scenarios,
                "file_scenarios_odds": rs_scenarios_odds_side_bet,
                "num_scenarios": len(normal_bet_scenarios)
            }
        }
    },
    "bonus_buy": {
        "respins": {
            "num_triggering_scatters": {
                "values": utils_array.get_array_vertical(server_bonus_buy_sheet, 'D', 6),
                "weights": utils_array.get_array_vertical(server_bonus_buy_sheet, 'E', 6)
            },
            "up_booster_multipliers": {
                "values": utils_array.get_array_vertical(server_bonus_buy_sheet, 'G', 6),
                "weights": utils_array.get_array_vertical(server_bonus_buy_sheet, 'H', 6)
            },
            "scenarios": {
                "file_scenarios": rs_scenarios,
                "file_scenarios_odds": rs_scenarios_odds_bonus_buy,
                "num_scenarios": len(normal_bet_scenarios)
            }
        }
    },
    "bonus_buy_enabled": True,
    "win_cap_total_bet_mult": 0
}

parser_base.init(game_name, skin_id)

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_normal_paid.strip_set', server_strips_sheet, ['C', 'D', 'E', 'F', 'G'], 18)
parser_base.create_strip_file('1_normal_wild_mult.strip_set', server_strips_sheet, ['K', 'L', 'M', 'N', 'O'], 18)
parser_base.create_strip_file('2_normal_wild_reels.strip_set', server_strips_sheet, ['S', 'T', 'U', 'V', 'W'], 18)
parser_base.create_strip_file('3_side_bet_paid.strip_set', server_strips_sheet, ['AA', 'AB', 'AC', 'AD', 'AE'], 18)

parser_base.create_file('rs_scenarios.data', normal_bet_scenarios)
parser_base.create_file('rs_scenarios_odds_normal_bet.data', utils_array.get_array_vertical(server_sheet, 'AL', 6))
parser_base.create_file('rs_scenarios_odds_side_bet.data', utils_array.get_array_vertical(server_sheet, 'BX', 6))
parser_base.create_file('rs_scenarios_odds_bonus_buy.data',
                        utils_array.get_array_vertical(server_bonus_buy_sheet, 'T', 6))
