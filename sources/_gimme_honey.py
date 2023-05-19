import openpyxl

import utils_array
import parser_base
import utils_transform

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

    "fs_config": {
        "scatters_to_fs": [0, 0, 0, 6, 8, 10, 12],

        "fs_special": {
            "values": [1, 0],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 211, 'C', 'D'),
        },
    },

    "game_modes": {
        "paid": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 29, 'C', 'G'),
        },
        "free": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 30, 'C', 'G'),
        },
        "free_special": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 31, 'C', 'G'),
        },
    },

    "game_modes_max_mw": {
        "paid": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 37, 'C', 'F'),
        },
        "free": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 38, 'C', 'F'),
        },
        "free_special": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 39, 'C', 'F'),
        }
    },

    "basic": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 46, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 47, 'C', 'F'),
            },
            "tracker_option": [
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
            "reels_size": [
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
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 59, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 60, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 61, 'I', 'L'),
                }
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 59, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 60, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 61, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 62, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 63, 'Q', 'V'),
                }
            ]
        },
    },

    "wild_scatter": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 70, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 71, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 70, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 71, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 70, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 72, 'J', 'K'),
                }
            ],
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 78, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 79, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 78, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 79, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 78, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 80, 'J', 'K'),
                }
            ],
        },
    },

    "mystery": {
        "initial_reels_odds": [0, 50, 10, 10, 30, 30],

        "paid": {
            "symbol": {
                "values": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 100, 'C', 'L')
            },
            "num_reels": {
                "values": [3, 4, 5, 6],
                "weights": [10, 30, 28, 0]
            },
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'D')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 89, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 90, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 91, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 92, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 93, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 94, 'C', 'H'),
                }
            ]
        },

        "free": {
            "symbol": {
                "values": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 101, 'C', 'L')
            },
            "num_reels": {
                "values": [3, 4, 5, 6],
                "weights": [40, 4, 0, 0]
            },
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'M')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 89, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 90, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 91, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 92, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 93, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 88, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 94, 'L', 'Q'),
                }
            ]
        },
    },

    "max_megaways": {
        "max_reels_size": [7, 6, 6, 6, 6, 7],

        "max_ways_visual": {
            "values": [0, 1],
            "weights": [95, 5]
        },

        "paid": {
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 114, 'C', 'D')
            }
        },

        "free": {
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 115, 'C', 'D')
            }
        },
    },

    "queen_wild": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 137, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 138, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 137, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 138, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 137, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 139, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 144, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 145, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 146, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 147, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 142, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 148, 'C', 'H'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 125, 'C', 'D')
            },
            "gh_position": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 124, 'I', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 125, 'I', 'L')
            },
            "gh_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 129, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 130, 'C', 'E')
            },
            "qw_height": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 129, 'K', 'P'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 130, 'K', 'P')
            }
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 169, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 169, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 169, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 174, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 175, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 176, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 177, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 178, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 179, 'C', 'H'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 156, 'C', 'D')
            },
            "gh_position": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 155, 'I', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 156, 'I', 'L')
            },
            "gh_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 160, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 161, 'C', 'E')
            },
            "qw_height": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 160, 'K', 'P'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 161, 'K', 'P')
            }
        }
    },

    "bonus_buy_enabled": True,
    "win_cap_total_bet_mult": 5000
}

parser_base.init('ngs_gimme_the_honey', '200993')

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_paid.strip_set', strips_sheet,
                              ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                               'T',
                               'U', 'V', 'W', 'X', 'Y', 'Z'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('1_free.strip_set', strips_sheet,
                              ['AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP',
                               'AQ', 'AR',
                               'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('2_paid_second_chance.strip_set', strips_sheet,
                              ['BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('3_free_second_chance.strip_set', strips_sheet,
                              ['BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('4_mystery.strip_set', strips_sheet,
                              ['CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('5_paid_queen_wild.strip_set', strips_sheet,
                              ['DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('6_free_queen_wild.strip_set', strips_sheet,
                              ['DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('0_paid.tracker', strips_sheet,
                              ['C', 'D', 'E', 'F', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('1_free.tracker', strips_sheet,
                              ['AD', 'AE', 'AF', 'AG', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('2_paid_second_chance.tracker', strips_sheet,
                              ['BF', 'BG', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('3_free_second_chance.tracker', strips_sheet,
                              ['BU', 'BV', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('4_mystery.tracker', strips_sheet,
                              ['CK', 'CL', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('5_paid_queen_wild.tracker', strips_sheet,
                              ['DA', 'DB', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('6_free_queen_wild.tracker', strips_sheet,
                              ['DP', 'DQ', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)
