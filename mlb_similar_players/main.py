
import datetime as dt
import numpy as np
import pandas as pd
import collections as c
import xlsxwriter


START = dt.datetime.now()
INPUT_FILE = 'mlb_2017_stats.csv'
OUTPUT_FILE = 'mlb_similar_players.xlsx'
POSITION_LIST = ['1B', '2B', '3B', 'SS', 'OF', 'C', 'DH', 'ALL']
WEIGHTS = {'G': 1, 'AB': 1, 'R': 1, 'H': 1, '2B': 1, '3B': 1, 'HR': 1, 'RBI': 1,
           'SB': 1, 'CS': 1, 'BB': 1, 'SO': 1, 'SH': 1, 'SF': 1, 'HBP': 1, 'AVG': 1, 'OBP': 1, 'SLG': 1, 'OPS': 1}
# Higher weight = more impact on the similarity score. A weight of 0 removes factor from similarity calculation.


def get_mlb_data(file_name):
    """Gets CSV file with player data"""

    xls = pd.read_csv(file_name)
    xls = xls.fillna('')
    player_data = xls.to_dict('index')

    return player_data


def create_position_dict(player_data):
    """Creates a dictionary with positions as keys and players as values. To be used for filtering purposes."""
    position_dict = {}
    for playerid, stats in player_data.items():
        position = stats['Pos']
        if position not in position_dict:
            position_dict[position] = [playerid]
        else:
            position_dict[position].append(playerid)

    return position_dict


def compare_players(player_data, position_dict):
    """Takes all players in the same position and compares them to each other. Output is a dictionary with positions
    as keys and data frames with similarity score per comparison as values."""
    final_scores = {}
    cnt = 0
    for position in POSITION_LIST:
        start2 = dt.datetime.now()
        print('Making dict of players for', position)
        if position != 'ALL':
            unordered_position_player_data_dict = {player: player_data[player] for player in position_dict[position]}
        else:
            unordered_position_player_data_dict = player_data
        position_player_data_dict = c.OrderedDict(sorted(unordered_position_player_data_dict.items()))
        for factor in WEIGHTS:
            print(factor, ': creating matrix for', position)
            arr, player_index = create_player_factor_matrix(position_player_data_dict, factor)
            print(factor, ': comparing players for', position)
            results = compute_squared_EDM_method4(arr)
            max_dist = np.amax(results)
            min_dist = np.amin(results)
            print(factor, ': calculating scaled results for', position)
            dist_range = max_dist - min_dist
            if dist_range == 0:
                scaled_results = (results + 1)/(results + 1) #avoid dividing by zero
            else:
                scaled_results = 1 - ((results - min_dist) / dist_range)
            weighted_results = scaled_results * WEIGHTS[factor]
            weighted_results_df = pd.DataFrame(data=weighted_results, index=position_player_data_dict.keys(),
                                             columns=position_player_data_dict.keys())
            print(factor, ': storing', factor, 'data for', position)
            weighted_results_df = weighted_results_df.fillna(0)
            if position not in final_scores:
                final_scores[position] = weighted_results_df
            else:
                final_scores[position] += weighted_results_df
        end2 = dt.datetime.now()
        cnt += 1
        print('{0} Runtime: {1} seconds - {2} positions done'.format(position, (end2 - start2).seconds, cnt))

    return final_scores


def create_player_factor_matrix(position_player_data_dict, factor):
    """Creates a matrix of players and their values for a given factor, the required format for the comparison
        function."""
    compare_player_dict = c.OrderedDict()
    for player in position_player_data_dict:
        compare_player = position_player_data_dict[player][factor]
        compare_player_dict[player] = compare_player
    compare_player_df = pd.DataFrame.from_dict(compare_player_dict, orient='index')
    compare_player_df.fillna(0, inplace=True)
    player_index = compare_player_df.index.values
    compare_player_npm = compare_player_df.as_matrix()

    return compare_player_npm, player_index


def compute_squared_EDM_method4(arr):
    """This is the formula that actually calculates player similarity for a set of players across a single factor."""
    m, n = arr.shape
    G = np.dot(arr, arr.T)
    H = np.tile(np.diag(G), (m, 1))
    distances = np.sqrt(H + H.T - 2 * G)

    return distances


def create_top_5_dict(all_comps, player_data, position_dict):
    """Takes data frames of similarity scores and creates a list of 5 most similar players for each valid player/position
        combo.  Output is a dictionary of player IDs with name, position, and the top 5 list the to as nested dicts."""
    top_5_dict = {}
    x = 0
    for player, stats in player_data.items():
        player_name = stats['Player Name']
        top_5_dict[player] = {}

        # For comparisons within position
        position = stats['Pos']
        player_comps = all_comps[position].ix[player]
        players_to_compare = position_dict[position]
        player_comps_data = player_comps.filter(items=players_to_compare)
        player_comps_data = player_comps_data.sort_values(axis=0, ascending=False)
        top_5_player_list = list(player_comps_data[1:6].keys())
        new_dict = {}
        for p in top_5_player_list:
            new_dict[player_data[p]['Player Name']] = player_comps_data[p]
        top_5_dict[player]['Player Name'] = player_name
        top_5_dict[player]['Position'] = position
        top_5_dict[player]['Top_5_Position'] = new_dict

        # For comparisons across all positions
        player_all_comps_data = all_comps['ALL'].ix[player]
        player_all_comps_data = player_all_comps_data.sort_values(axis=0, ascending=False)
        top_5_all_player_list = list(player_all_comps_data[1:6].keys())
        all_players_dict = {}
        for p in top_5_all_player_list:
            all_players_dict[player_data[p]['Player Name']] = player_all_comps_data[p]
        top_5_dict[player]['Top_5_All'] = all_players_dict

        x += 1
        print(x, 'players completed')

    return top_5_dict


def create_workbook(top_5_dict, position_list, position_dict):
    """Writes data to an Excel workbook, one sheet per position."""

    # Create workbook
    workbook = xlsxwriter.Workbook(OUTPUT_FILE, {'constant_memory': True})

    # Create worksheets
    for position in position_list:
        print('Creating {0} Worksheet . . . '.format(position))
        worksheet = workbook.add_worksheet('{0}'.format(position))
        row = 0
        col = 0
        worksheet.write(row, col + 1, 'Top 5 Most Similar Players')
        row += 1
        worksheet.write(row, col, 'Player')
        worksheet.write(row, col + 1, '1')
        worksheet.write(row, col + 2, '2')
        worksheet.write(row, col + 3, '3')
        worksheet.write(row, col + 4, '4')
        worksheet.write(row, col + 5, '5')
        row += 1
        if position != 'ALL':
            for player in position_dict[position]:
                name = top_5_dict[player]['Player Name']
                top_5_list = [p for p in top_5_dict[player]['Top_5_Position'].items()]
                worksheet.write(row, col, name)
                worksheet.write(row, col + 1, top_5_list[0][0])
                worksheet.write(row, col + 2, top_5_list[1][0])
                worksheet.write(row, col + 3, top_5_list[2][0])
                worksheet.write(row, col + 4, top_5_list[3][0])
                if len(top_5_list) == 5:
                    worksheet.write(row, col + 5, top_5_list[4][0])
                else:
                    continue
                row += 1
        else:
            for player in top_5_dict:
                name = top_5_dict[player]['Player Name']
                top_5_list = [p for p in top_5_dict[player]['Top_5_All'].items()]
                worksheet.write(row, col, name)
                worksheet.write(row, col + 1, top_5_list[0][0])
                worksheet.write(row, col + 2, top_5_list[1][0])
                worksheet.write(row, col + 3, top_5_list[2][0])
                worksheet.write(row, col + 4, top_5_list[3][0])
                if len(top_5_list) == 5:
                    worksheet.write(row, col + 5, top_5_list[4][0])
                else:
                    continue
                row += 1

    workbook.close()

    return None


def main():
    print('Getting player data . . .')
    player_data = get_mlb_data(INPUT_FILE)
    print('Creating dict of players by position . . .')
    position_dict = create_position_dict(player_data)
    print('Starting position loop . . .')
    all_comps = compare_players(player_data, position_dict)
    print('Starting to create top 5 dict . . .')
    top_5_dict = create_top_5_dict(all_comps, player_data, position_dict)
    print('Create Excel file with results')
    create_workbook(top_5_dict, POSITION_LIST, position_dict)


main()
