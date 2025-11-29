import pandas as pd

def read_leaderboard(file_path):
    try:
        df = pd.read_excel(file_path, skiprows=1)
        print(f"Successfully loaded data from {file_path}")
        # print(f"Columns found: {list(df.columns)}")
        
        # print("\nFirst 3 rows of data:")
        # print(df.head(3))                  #just for testing ignore
        
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def find_key_columns(df):
    column_mapping = {}
    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['player', 'name', 'team']):
            column_mapping['player_col'] = col
            break
    else:
        column_mapping['player_col'] = df.columns[1] 

    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['total', 'points', 'score']):
            column_mapping['total_points_col'] = col
            break
     
    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['spent', 'cost', 'money']):
            column_mapping['total_spent_col'] = col
            break
    
    round_columns = []
    for col in df.columns:
        col_str = str(col)
        if col_str.startswith('R') and len(col_str) <= 4:
            if col_str[1:].isdigit():  # R01 R02 etc.
                round_columns.append(col)
    
    column_mapping['round_cols'] = sorted(round_columns)
    
    # print(f"\nDetected columns:")            #just for testing
    # print(f"Player column: {column_mapping['player_col']}")
    # print(f"Total points column: {column_mapping.get('total_points_col', 'Not found')}")
    # print(f"Total spent column: {column_mapping.get('total_spent_col', 'Not found')}")
    # print(f"Round columns: {column_mapping['round_cols']}")
    
    return column_mapping

def clean_points(value):
    if pd.isna(value) or value in ['-', 'D$Q']:
        return 0
    if isinstance(value, str):
        if value.replace('.', '').isdigit() or (value.replace('-', '').replace('.', '').isdigit() and value.count('-') == 1):
            return float(value)
        else:
            return 0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0

def get_player_scores(df, player_index, round_columns):
    scores = []
    for round_col in round_columns:
        score = clean_points(df.iloc[player_index][round_col])
        if score > 0:  
            scores.append(score)
    
    scores.sort(reverse=True)
    return scores

def compare_players_countback(df, player1_idx, player2_idx, round_columns):
    p1_scores = get_player_scores(df, player1_idx, round_columns)
    p2_scores = get_player_scores(df, player2_idx, round_columns)
    
    max_comparisons = max(len(p1_scores), len(p2_scores))
    
    for i in range(max_comparisons):
        if i >= len(p1_scores) and i < len(p2_scores):
            return 1 
        if i >= len(p2_scores) and i < len(p1_scores):
            return -1  
        
        if i < len(p1_scores) and i < len(p2_scores):
            if p1_scores[i] > p2_scores[i]:
                return -1  
            elif p1_scores[i] < p2_scores[i]:
                return 1   
    
    if p1_scores and p2_scores:
        highest_score = max(p1_scores[0], p2_scores[0])
        p1_count = p1_scores.count(highest_score)
        p2_count = p2_scores.count(highest_score)
        
        if p1_count > p2_count:
            return -1
        elif p1_count < p2_count:
            return 1
    
    return 0 

def sort_tied_players(df, tied_indices, round_columns, total_spent_col):
    """
    Sort a group of tied players using all tiebreakers
    """
    if len(tied_indices) == 1:
        return tied_indices
    
    spent_groups = {}
    for idx in tied_indices:
        spent = clean_points(df.iloc[idx][total_spent_col])
        if spent not in spent_groups:
            spent_groups[spent] = []
        spent_groups[spent].append(idx)
    
    sorted_by_spent = []
    for spent in sorted(spent_groups.keys()):
        group_players = spent_groups[spent]
        
        if len(group_players) == 1:
            sorted_by_spent.append(group_players[0])
        else:
            sorted_group = sort_by_countback(df, group_players, round_columns)
            sorted_by_spent.extend(sorted_group)
    
    return sorted_by_spent

def sort_by_countback(df, player_indices, round_columns):
    if len(player_indices) <= 1:
        return player_indices
    
    players_list = player_indices.copy()
    n = len(players_list)
    
    for i in range(n):
        for j in range(0, n - i - 1):
            result = compare_players_countback(df, players_list[j], players_list[j+1], round_columns)
            if result == 1:  
                players_list[j], players_list[j+1] = players_list[j+1], players_list[j]
    
    final_sorted = []
    current_group = [players_list[0]]
    
    for i in range(1, len(players_list)):
        prev_player = current_group[0]
        curr_player = players_list[i]
        
        if compare_players_countback(df, prev_player, curr_player, round_columns) == 0:
            current_group.append(curr_player)
        else:
            if len(current_group) > 1:
                current_group.sort(key=lambda x: str(df.iloc[x][df.columns[1]]).lower())
            final_sorted.extend(current_group)
            current_group = [curr_player]
    
    if len(current_group) > 1:
        current_group.sort(key=lambda x: str(df.iloc[x][df.columns[1]]).lower())
    final_sorted.extend(current_group)
    
    return final_sorted

def calculate_total_points(df, round_columns):
    totals = []
    for idx in range(len(df)):
        total = 0
        for round_col in round_columns:
            total += clean_points(df.iloc[idx][round_col])
        totals.append(total)
    return totals

def main():

    print(" Leaderboard Ranking System ")
    
    df = read_leaderboard('leaderboard.xlsx')
    if df is None:
        return
    
    col_map = find_key_columns(df)
    
    player_col = col_map['player_col']
    round_columns = col_map['round_cols']
    total_points_col = col_map.get('total_points_col')
    total_spent_col = col_map.get('total_spent_col')
    
    df_clean = df.copy()
    for col in round_columns:
        df_clean[col] = df_clean[col].apply(clean_points)
    
    if not total_points_col:
        print("\nCalculating total points from round scores...")
        df_clean['Calculated_Total_Points'] = calculate_total_points(df_clean, round_columns)
        total_points_col = 'Calculated_Total_Points'
    else:
        df_clean[total_points_col] = df_clean[total_points_col].apply(clean_points)
    
    if not total_spent_col:
        print("No spending data found - using zeros for spending tiebreaker")
        df_clean['Calculated_Total_Spent'] = 0
        total_spent_col = 'Calculated_Total_Spent'
    else:
        df_clean[total_spent_col] = df_clean[total_spent_col].apply(clean_points)
    
    print(f"\nUsing columns:")
    print(f"Players: {player_col}")
    print(f"Total Points: {total_points_col}") 
    print(f"Total Spent: {total_spent_col}")
    print(f"Rounds: {len(round_columns)} columns")
    
    df_clean = df_clean[df_clean[player_col].notna()]
    df_clean = df_clean[~df_clean[player_col].astype(str).str.isnumeric()]
    df_clean[player_col] = df_clean[player_col].str.strip()
    df_clean = df_clean.drop_duplicates(subset=[player_col], keep='first')    #remove duplicate player names i made a mistake and missed this earlier
    df_clean = df_clean[~df_clean[player_col].str.lower().isin([
        "points totals",
        "spending totals"
    ])]                                                #to remove summary rows and gets me 25 payers as in the sheet
    df_clean = df_clean[df_clean[player_col].str.lower() != "player"]

    
    print(f"\nProcessing {len(df_clean)} players...")
    
    df_sorted = df_clean.sort_values(total_points_col, ascending=False)
    df_sorted = df_sorted.reset_index(drop=True)
    
    print("Applying tiebreakers...")
    final_order = []
    
    unique_points = df_sorted[total_points_col].unique()
    unique_points.sort()
    unique_points = unique_points[::-1]  
    
    for points in unique_points:
        tied_players = df_sorted[df_sorted[total_points_col] == points]
        
        if len(tied_players) == 1:
            final_order.append(tied_players.index[0])
        else:
            tied_indices = list(tied_players.index)
            sorted_tied = sort_tied_players(df_sorted, tied_indices, round_columns, total_spent_col)
            final_order.extend(sorted_tied)
    
    final_df = df_sorted.loc[final_order].reset_index(drop=True)
    final_df['Rank'] = range(1, len(final_df) + 1)
    
    print("\n" + "="*80)
    print("FINAL RANKINGS")
    print("="*80)
    print(f"{'Rank':<4} {'Player':<30} {'Total Points':<12} {'Total Spent':<12}")
    print("-" * 80)
    
    for idx, row in final_df.iterrows():
        player_name = str(row[player_col])
        if len(player_name) > 28:
            player_name = player_name[:28] + '..'
        
        points = row[total_points_col]
        spent = row[total_spent_col]
        
        print(f"{row['Rank']:<4} {player_name:<30} {points:<12.0f} {spent:<12.2f}")

if __name__ == "__main__":
    main()