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

game_settings_model = {
    "game": {
        "type": 2,
        "id": 26433,
        "machine_id": 26433
    },

    "fs_config": {
        "scatters_to_fs": [0, 0, 0, 6, 8, 10, 12],

        "fs_type": {
            "values": utils_array.get_array_horizontal(game_settings_sheet, 78, 'C', 'E'),
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 79, 'C', 'E'),
        },

        "fs_special": {
            "values": [1, 0],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 291, 'C', 'D'),
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
                "values": utils_array.get_array_horizontal(game_settings_sheet, 304, 'H', 'K'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 305, 'H', 'K'),
            },
            "bonus_buy_2": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 304, 'H', 'K'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 306, 'H', 'K'),
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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 311, 'C', 'G'),
        },
        "bonus_buy_2": {
            "values": [0, 1, 2, 3, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 312, 'C', 'G'),
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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 319, 'C', 'F'),
        },
        "bonus_buy_2": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 320, 'C', 'F'),
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
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 327, 'C', 'F'),
        },
        "bonus_buy_2": {
            "values": [3, 1, 2, 4],
            "weights": utils_array.get_array_horizontal(game_settings_sheet, 328, 'C', 'F'),
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
                "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 85, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 85, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 86, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 87, 'I', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 88, 'I', 'L'),
                }
            ],
            "reels_size": [
                [
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 85, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 86, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 87, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 88, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 89, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 84, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 90, 'Q', 'V'),
                    }
                ],
                [
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 93, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 94, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 95, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 96, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 97, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 92, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 98, 'Q', 'V'),
                    }
                ],
                [
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 101, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 102, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 103, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 104, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 105, 'Q', 'V'),
                    },
                    {
                        "values": utils_array.get_array_horizontal(game_settings_sheet, 100, 'Q', 'V'),
                        "weights": utils_array.get_array_horizontal(game_settings_sheet, 106, 'Q', 'V'),
                    }
                ]
            ]
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 332, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 333, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 338, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 339, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 338, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 340, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 338, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 341, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 338, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 342, 'C', 'F'),
                }
            ],
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 332, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 334, 'C', 'F'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 344, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 345, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 344, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 346, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 344, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 347, 'C', 'F'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 344, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 348, 'C', 'F'),
                }
            ],
        }
    },

    "wild_scatter": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 114, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 115, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 114, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 115, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 114, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 116, 'J', 'K'),
                }
            ],
            "skip": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 120, 'C', 'D'),
            }
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 126, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 127, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 126, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 127, 'J', 'K'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 126, 'J', 'K'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 128, 'J', 'K'),
                }
            ],
            "skip": [
                {
                    "values": [0, 1],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 132, 'C', 'D'),
                },
                {
                    "values": [0, 1],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 133, 'C', 'D'),
                },
                {
                    "values": [0, 1],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 134, 'C', 'D'),
                }
            ]
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 353, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 354, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 358, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 359, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 358, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 360, 'C', 'D'),
                }
            ],
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 353, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 355, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 362, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 363, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 362, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 364, 'C', 'D'),
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
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 161, 'C', 'L')
            },
            "num_reels": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 153, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 154, 'C', 'F')
            },
            "must_win": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 167, 'C', 'D')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 144, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 145, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 146, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 147, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 148, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 149, 'C', 'H'),
                }
            ]
        },

        "free": {
            "symbol": {
                "values": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 162, 'C', 'L')
            },
            "num_reels": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 153, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 155, 'C', 'F')
            },
            "must_win": {
                "values": [0, 1],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 167, 'L', 'M')
            },
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 144, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 145, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 146, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 147, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 148, 'L', 'Q'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 143, 'L', 'Q'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 149, 'L', 'Q'),
                }
            ]
        },
    },

    "max_megaways": {
        "max_reels_size": [7, 6, 6, 6, 6, 7],

        "paid": {
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 177, 'C', 'D')
            },

            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 182, 'C', 'F'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 183, 'C', 'F'),
            },
        },

        "free": {
            "can_lose": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 178, 'C', 'D')
            },

            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 182, 'J', 'M'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 183, 'J', 'M'),
            },
        },
    },

    "queen_wild": {
        "paid": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 205, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 206, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 205, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 206, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 205, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 207, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 211, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 212, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 213, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 214, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 215, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 210, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 216, 'C', 'H'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 193, 'C', 'D')
            },
            "gh_position": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 192, 'I', 'L'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 193, 'I', 'L')
            },
            "gh_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 197, 'C', 'E'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 198, 'C', 'E')
            },
            "rs_limit": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 197, 'P', 'Y'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 198, 'P', 'Y')
            },
        },
        "free": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 249, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 250, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 249, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 250, 'K', 'L'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 249, 'K', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 251, 'K', 'L'),
                },
            ],
            "reels_size": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 255, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 256, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 257, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 258, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 259, 'C', 'H'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 254, 'C', 'H'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 260, 'C', 'H'),
                }
            ],
            "place_gh": [
                {
                    "values": [1, 0],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 224, 'C', 'D')
                },
                {
                    "values": [1, 0],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 225, 'C', 'D')
                },
                {
                    "values": [1, 0],
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 226, 'C', 'D')
                },
            ],
            "gh_position": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 223, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 224, 'I', 'L')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 223, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 225, 'I', 'L')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 223, 'I', 'L'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 226, 'I', 'L')
                }
            ],
            "gh_limit": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'C', 'E'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 231, 'C', 'E')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'C', 'E'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 236, 'C', 'E')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'C', 'E'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 241, 'C', 'E')
                }
            ],
            "rs_limit": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'P', 'Y'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 231, 'P', 'Y')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'P', 'Y'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 232, 'P', 'Y')
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 230, 'P', 'Y'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 233, 'P', 'Y')
                }
            ],
        },
        "bonus_buy_1": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 380, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 381, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 386, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 387, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 386, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 388, 'C', 'D'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 370, 'C', 'D')
            },
        },
        "bonus_buy_2": {
            "reels_option": {
                "values": utils_array.get_array_horizontal(game_settings_sheet, 380, 'C', 'D'),
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 382, 'C', 'D'),
            },
            "tracker_option": [
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 390, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 391, 'C', 'D'),
                },
                {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 390, 'C', 'D'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 392, 'C', 'D'),
                }
            ],
            "place_gh": {
                "values": [1, 0],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 371, 'C', 'D')
            },
        }
    },

    "bonus_buy_enabled": True
}

parser_base.init(game_name, skin_id)

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
