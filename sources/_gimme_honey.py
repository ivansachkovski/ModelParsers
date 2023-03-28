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
    "common": {
        "max_reels_size": [7, 6, 6, 6, 6, 7],
        "scatters_to_fs": [0, 0, 0, 6, 8, 10, 12]
    },
    "normal_bet": {
        "paid_spins": {
            "game_modes_odds": {
                "values": [0, 1, 2, 3, 4],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 29, 'C', 'G'),
            },
            "basic": {
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
                "reels_size_odds": [
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
            "mystery": {
                "reels_size_odds": [
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
            }
        },
        "free_spins": {
            "game_modes_odds": {
                "values": [0, 1, 2, 3, 4],
                "weights": utils_array.get_array_horizontal(game_settings_sheet, 30, 'C', 'G'),
            },
            "basic": {
                "reels_odds": {
                    "values": utils_array.get_array_horizontal(game_settings_sheet, 57, 'C', 'F'),
                    "weights": utils_array.get_array_horizontal(game_settings_sheet, 58, 'C', 'F'),
                },
                "tracker_odds": [
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
                "reels_size_odds": [
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
            }
        },
        "mystery": {
            "reels_size_odds": [
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
        }
    },
    "bonus_buy_x75": {
    },
    "bonus_buy_x150": {
    },
    "bonus_buy_enabled": True,
    "win_cap_total_bet_mult": 5000
}

parser_base.init('ngs_gimme_the_honey', '200993')

parser_base.create_game_settings_file(game_settings_model)

parser_base.create_strip_file('0_paid.strip_set', strips_sheet,
                              ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                               'U', 'V', 'W', 'X', 'Y', 'Z'],
                              3, 142,
                              utils_transform.transform_inc)

parser_base.create_strip_file('1_free.strip_set', strips_sheet,
                              ['AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR',
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

parser_base.create_strip_file('0_paid.bottom_tracker', strips_sheet,
                              ['C', 'D', 'E', 'F', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('1_free.bottom_tracker', strips_sheet,
                              ['AD', 'AE', 'AF', 'AG', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('2_paid_second_chance.bottom_tracker', strips_sheet,
                              ['BF', 'BG', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('3_free_second_chance.bottom_tracker', strips_sheet,
                              ['BU', 'BV', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('4_mystery.bottom_tracker', strips_sheet,
                              ['CK', 'CL', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('5_paid_queen_wild.bottom_tracker', strips_sheet,
                              ['DA', 'DB', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)

parser_base.create_strip_file('6_free_queen_wild.bottom_tracker', strips_sheet,
                              ['DP', 'DQ', 'A', 'A', 'A', 'A'],
                              147, 266,
                              utils_transform.transform_inc)
