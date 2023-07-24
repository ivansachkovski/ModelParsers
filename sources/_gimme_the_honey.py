import openpyxl

import sys

import utils_array
import parser_base
import utils_transform

game_name = sys.argv[1]
skin_id = sys.argv[2]
model_path = sys.argv[3]

wb = openpyxl.load_workbook(model_path, data_only=True)

gs_sheet = wb['Parameters']  # game settings
strips_sheet = wb['Reels']


def create_table(
        sheet,
        col_1: str,
        col_2: str,
        values_row: int = 0,
        values=None,
        first_weights_row: int = 0,
        num_rows: int = 0):
    def create_row(row: int):
        return {
            "values": utils_array.get_array_horizontal(sheet, values_row, col_1, col_2) if values is None else values,
            "weights": utils_array.get_array_horizontal(sheet, row, col_1, col_2)
        }

    assert first_weights_row > 0 and num_rows > 0

    return [create_row(first_weights_row + row) for row in range(num_rows)]


def create_table_2(
        sheet,
        col_1: str,
        col_2: str,
        values_rows: list,
        first_weights_rows: list,
        num_rows: int = 0):
    assert num_rows > 0
    assert len(values_rows) == len(first_weights_rows)
    num_tables = len(values_rows)
    assert num_tables > 0

    return [
        create_table(
            sheet,
            col_1,
            col_2,
            values_row=values_rows[i],
            first_weights_row=first_weights_rows[i],
            num_rows=num_rows)
        for i in range(num_tables)]


def create_obj(
        sheet,
        col_1: str,
        col_2: str,
        values_row: int = 0,
        values=None,
        weights_row: int = 0):
    return {
        "values": utils_array.get_array_horizontal(sheet, values_row, col_1, col_2) if values is None else values,
        "weights": utils_array.get_array_horizontal(sheet, weights_row, col_1, col_2),
    }


game_settings_model = {
    "game": {
        "type": 2,
        "id": 26433,
        "machine_id": 26433
    },

    "fs_config": {
        "scatters_to_fs": [0, 0, 0, 6, 8, 10, 12],
        "fs_type": create_obj(gs_sheet, 'C', 'E', values_row=78, weights_row=79),
        "fs_special": create_table(gs_sheet, 'C', 'D', values=[1, 0], first_weights_row=313, num_rows=3),

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
            "bonus_buy_1": create_obj(gs_sheet, 'H', 'K', values_row=328, weights_row=329),
            "bonus_buy_2": create_obj(gs_sheet, 'H', 'K', values_row=328, weights_row=330),
        },
    },

    "game_modes": {
        "paid": create_obj(gs_sheet, 'C', 'G', values=[0, 1, 2, 3, 4], weights_row=29),
        "free": create_obj(gs_sheet, 'C', 'G', values=[0, 1, 2, 3, 4], weights_row=30),
        "free_special": create_obj(gs_sheet, 'C', 'G', values=[0, 1, 2, 3, 4], weights_row=31),
        "bonus_buy_1": create_obj(gs_sheet, 'C', 'G', values=[0, 1, 2, 3, 4], weights_row=335),
        "bonus_buy_2": create_obj(gs_sheet, 'C', 'G', values=[0, 1, 2, 3, 4], weights_row=336),
    },

    "game_modes_max_mw_1": {
        "paid": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=38),
        "free": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=39),
        "free_special": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=40),
        "bonus_buy_1": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=343),
        "bonus_buy_2": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=344),
    },

    "game_modes_max_mw_2": {
        "free": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=47),
        "free_special": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=48),
        "bonus_buy_1": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=351),
        "bonus_buy_2": create_obj(gs_sheet, 'C', 'F', values=[3, 1, 2, 4], weights_row=352),
    },

    "basic": {
        "paid": {
            "reels_option": create_obj(gs_sheet, 'C', 'F', values_row=55, weights_row=56),
            "tracker_option": create_table(gs_sheet, 'I', 'L', values_row=55, first_weights_row=56, num_rows=4),
            "reels_size": create_table(gs_sheet, 'Q', 'V', values_row=55, first_weights_row=56, num_rows=6)
        },
        "free": {
            "reels_option": create_table(gs_sheet, 'C', 'F', values_row=84, first_weights_row=85, num_rows=3),
            "tracker_option": create_table_2(gs_sheet, 'I', 'L', values_rows=[84, 90, 96],
                                             first_weights_rows=[85, 91, 97], num_rows=4),
            "reels_size": create_table_2(gs_sheet, 'Q', 'V', values_rows=[84, 92, 100],
                                         first_weights_rows=[85, 93, 101], num_rows=6),
        },
        "bonus_buy_1": {
            "reels_option": create_obj(gs_sheet, 'C', 'F', values_row=356, weights_row=357),
            "tracker_option": create_table(gs_sheet, 'C', 'F', values_row=362, first_weights_row=363, num_rows=4),
        },
        "bonus_buy_2": {
            "reels_option": create_obj(gs_sheet, 'C', 'F', values_row=356, weights_row=358),
            "tracker_option": create_table(gs_sheet, 'C', 'F', values_row=368, first_weights_row=369, num_rows=4),
        }
    },

    "wild_scatter": {
        "paid": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=114, weights_row=115),
            "tracker_option": create_table(gs_sheet, 'J', 'K', values_row=114, first_weights_row=115, num_rows=2),
            "skip": create_obj(gs_sheet, 'C', 'D', values=[0, 1], weights_row=120),
        },
        "free": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=126, weights_row=127),
            "tracker_option": create_table(gs_sheet, 'J', 'K', values_row=126, first_weights_row=127, num_rows=2),
            "skip": create_table(gs_sheet, 'C', 'D', values=[0, 1], first_weights_row=132, num_rows=3),
        },
        "bonus_buy_1": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=377, weights_row=378),
            "tracker_option": create_table(gs_sheet, 'C', 'D', values_row=382, first_weights_row=383, num_rows=2),
        },
        "bonus_buy_2": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=377, weights_row=379),
            "tracker_option": create_table(gs_sheet, 'C', 'D', values_row=386, first_weights_row=387, num_rows=2),
        }
    },

    "mystery": {
        "next_reel": {
            "values": [0, 1, 2, 3, 4, 5],
            "weights": [0, 50, 10, 10, 30, 30]
        },

        "paid": {
            "symbol": create_obj(gs_sheet, 'C', 'L', values=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], weights_row=161),
            "num_reels": create_obj(gs_sheet, 'C', 'F', values_row=153, weights_row=154),
            "must_win": create_obj(gs_sheet, 'C', 'D', values=[0, 1], weights_row=167),
            "reels_size": create_table(gs_sheet, 'C', 'H', values_row=143, first_weights_row=144, num_rows=6)
        },

        "free": {
            "symbol": create_obj(gs_sheet, 'C', 'L', values=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], weights_row=162),
            "num_reels": create_obj(gs_sheet, 'C', 'F', values_row=153, weights_row=155),
            "must_win": create_obj(gs_sheet, 'L', 'M', values=[0, 1], weights_row=167),
            "reels_size": create_table(gs_sheet, 'L', 'Q', values_row=143, first_weights_row=144, num_rows=6),
        },
    },

    "max_megaways": {
        "max_reels_size": [7, 6, 6, 6, 6, 7],

        "paid": {
            "can_lose": create_obj(gs_sheet, 'C', 'D', values=[1, 0], weights_row=177),
            "reels_option": create_obj(gs_sheet, 'C', 'F', values_row=182, weights_row=183),
        },

        "free": {
            "can_lose": create_obj(gs_sheet, 'C', 'D', values=[1, 0], weights_row=178),
            "reels_option": create_obj(gs_sheet, 'J', 'M', values_row=182, weights_row=183),
        },
    },

    "queen_wild": {
        "paid": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=212, weights_row=213),
            "tracker_option": create_table(gs_sheet, 'K', 'L', values_row=212, first_weights_row=213, num_rows=2),
            "reels_size": create_table(gs_sheet, 'C', 'H', values_row=217, first_weights_row=218, num_rows=6),
            "place_gh": create_obj(gs_sheet, 'C', 'D', values=[1, 0], weights_row=193),
            "gh_position": create_obj(gs_sheet, 'I', 'L', values_row=192, weights_row=193),
            "gh_limit": create_table(gs_sheet, 'C', 'E', values_row=197, first_weights_row=198, num_rows=4),
            "rs_limit": create_obj(gs_sheet, 'C', 'L', values_row=207, weights_row=208),
        },
        "free": {
            "reels_option": create_table(gs_sheet, 'C', 'D', values_row=264, first_weights_row=265, num_rows=3),
            "tracker_option": create_table_2(gs_sheet, 'K', 'L', values_rows=[264, 264, 264],
                                             first_weights_rows=[265, 268, 271], num_rows=2),
            "reels_size": create_table(gs_sheet, 'C', 'H', values_row=276, first_weights_row=277, num_rows=6),
            "place_gh": create_table(gs_sheet, 'C', 'D', values=[1, 0], first_weights_row=231, num_rows=3),
            "gh_position": create_table(gs_sheet, 'I', 'L', values_row=230, first_weights_row=231, num_rows=3),
            "gh_limit": create_table_2(gs_sheet, 'C', 'E', values_rows=[237, 237, 237],
                                       first_weights_rows=[238, 243, 248], num_rows=4),
            "rs_limit": create_table(gs_sheet, 'C', 'L', values_row=256, first_weights_row=257, num_rows=3)
        },
        "bonus_buy_1": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=404, weights_row=405),
            "tracker_option": create_table(gs_sheet, 'C', 'D', values_row=410, first_weights_row=411, num_rows=2),
            "place_gh": create_obj(gs_sheet, 'C', 'D', values=[1, 0], weights_row=394),
        },
        "bonus_buy_2": {
            "reels_option": create_obj(gs_sheet, 'C', 'D', values_row=404, weights_row=406),
            "tracker_option": create_table(gs_sheet, 'C', 'D', values_row=414, first_weights_row=415, num_rows=2),
            "place_gh": create_obj(gs_sheet, 'C', 'D', values=[1, 0], weights_row=395),
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
