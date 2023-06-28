import openpyxl

import utils_array
import parser_base
import utils_transform

model_path = '../../../math_models/GimmeHoney_41_14.06.2023.xlsx'

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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 236, 'C', 'D'),
        },

        "fs_special_qw_params": {
            "first": {
                "free": 1,
                "bonus_buy_1": 1,
                "bonus_buy_2": 1,
            },
            "interval": {
                "free": 3,
                "bonus_buy_1": 3,
                "bonus_buy_2": 3,
            }
        },

        "num_scatters": {
            "bonus_buy_1": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 249, 'H', 'K'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 250, 'H', 'K'),
            },
            "bonus_buy_2": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 249, 'H', 'K'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 251, 'H', 'K'),
            },
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
        "bonus_buy_1": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 256, 'C', 'G'),
        },
        "bonus_buy_2": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 257, 'C', 'G'),
        },
    },

    "game_modes_max_mw_1": {
        "paid": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 38, 'C', 'F'),
        },
        "free": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 39, 'C', 'F'),
        },
        "free_special": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 40, 'C', 'F'),
        },
        "bonus_buy_1": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 264, 'C', 'F'),
        },
        "bonus_buy_2": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 265, 'C', 'F'),
        }
    },

    "game_modes_max_mw_2": {
        "free": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 47, 'C', 'F'),
        },
        "free_special": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 48, 'C', 'F'),
        },
        "bonus_buy_1": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 272, 'C', 'F'),
        },
        "bonus_buy_2": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 273, 'C', 'F'),
        }
    },

    "basic": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 56, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 56, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 57, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 59, 'I', 'L'),
                }
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 56, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 57, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 59, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 60, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 55, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 61, 'Q', 'V'),
                }
            ]
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 67, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 67, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 68, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 69, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 70, 'I', 'L'),
                }
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 67, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 68, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 69, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 70, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 71, 'Q', 'V'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 66, 'Q', 'V'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 72, 'Q', 'V'),
                }
            ]
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 277, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 278, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 283, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 284, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 283, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 285, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 283, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 286, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 283, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 287, 'C', 'F'),
                }
            ],
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 277, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 279, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 289, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 290, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 289, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 291, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 289, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 292, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 289, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 293, 'C', 'F'),
                }
            ],
        }
    },

    "wild_scatter": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 79, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 80, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 79, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 80, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 79, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 81, 'J', 'K'),
                }
            ],
            "skip": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 85, 'C', 'D'),
            }
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 91, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 92, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 91, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 92, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 91, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 93, 'J', 'K'),
                }
            ],
            "skip": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 97, 'C', 'D'),
            }
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 298, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 299, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 303, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 304, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 303, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 305, 'C', 'D'),
                }
            ],
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 298, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 300, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 307, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 308, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 307, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 309, 'C', 'D'),
                }
            ],
        }
    },

    "mystery": {
        "next_reel": {
            "values": [0, 1, 2, 3, 4, 5],
            "weights": [0, 50, 10, 10, 30, 30]
        },

        "paid": {
            "symbol": {
                "values": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 124, 'C', 'L')
            },
            "num_reels": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 116, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 117, 'C', 'F')
            },
            "must_win": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 130, 'C', 'D')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 107, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 108, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 109, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 110, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 111, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 112, 'C', 'H'),
                }
            ]
        },

        "free": {
            "symbol": {
                "values": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 125, 'C', 'L')
            },
            "num_reels": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 116, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 118, 'C', 'F')
            },
            "must_win": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 130, 'L', 'M')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 107, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 108, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 109, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 110, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 111, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 106, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 112, 'L', 'Q'),
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
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 140, 'C', 'D')
            }
        },

        "free": {
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 141, 'C', 'D')
            }
        },
    },

    "queen_wild": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 163, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 164, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 163, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 164, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 163, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 165, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 169, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 170, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 171, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 172, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 173, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 168, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 174, 'C', 'H'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 151, 'C', 'D')
            },
            "gh_position": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 150, 'I', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 151, 'I', 'L')
            },
            "gh_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 155, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 156, 'C', 'E')
            },
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 194, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 195, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 194, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 195, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 194, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 196, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 200, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 201, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 202, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 203, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 204, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 199, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 205, 'C', 'H'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 182, 'C', 'D')
            },
            "gh_position": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 181, 'I', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 182, 'I', 'L')
            },
            "gh_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 186, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 187, 'C', 'E')
            },
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 325, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 326, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 331, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 332, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 331, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 333, 'C', 'D'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 315, 'C', 'D')
            },
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 325, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 327, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 335, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 336, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 335, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 337, 'C', 'D'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 316, 'C', 'D')
            },
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
