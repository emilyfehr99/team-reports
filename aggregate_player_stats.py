
import pandas as pd
import os

def aggregate_data():
    input_path = "data/players_2025_26.csv"
    output_path = "data/player_averages_2025_26.csv"
    
    if not os.path.exists(input_path):
        print(f"Error: {input_path} not found.")
        return

    print(f"Reading {input_path}...")
    df = pd.read_csv(input_path)
    
    # Identify numeric columns for averaging
    # We exclude game_id, date, opponent, home_away as they are game-specific
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    exclude_cols = ['game_id', 'player_id']
    cols_to_avg = [c for c in numeric_cols if c not in exclude_cols]

    print("Aggregating data by Player...")
    
    # Group by Player Identity
    # We take the 'last' team and position seen (or mode), but grouping by all helps if they changed unique combos
    # For "One row per player", we strictly group by ID/Name
    
    # Define aggregation dictionary
    agg_dict = {col: 'mean' for col in cols_to_avg}
    agg_dict['team'] = lambda x: '/'.join(x.unique()) # Show all teams played for
    agg_dict['position'] = 'first' # Assume position doesn't change much, or take first
    agg_dict['game_id'] = 'count' # Count games played
    
    # Grouping
    grouped = df.groupby(['player_id', 'player_name']).agg(agg_dict).reset_index()
    
    # Rename columns for clarity
    grouped.rename(columns={'game_id': 'games_played'}, inplace=True)
    
    # Round metrics to 2 decimal places
    grouped = grouped.round(2)
    
    # Reorder columns: ID, Name, Team, Pos, GP, then stats
    cols = ['player_id', 'player_name', 'team', 'position', 'games_played'] + \
           [c for c in grouped.columns if c not in ['player_id', 'player_name', 'team', 'position', 'games_played']]
    grouped = grouped[cols]
    
    print(f"Saving aggregated data to {output_path}...")
    grouped.to_csv(output_path, index=False)
    print("Done!")

if __name__ == "__main__":
    aggregate_data()
