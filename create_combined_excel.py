import pandas as pd
from pathlib import Path

def create_excel_with_averages():
    input_file = Path('penguins_season_split_2025_26_comprehensive.csv')
    output_file = Path.home() / 'Desktop' / 'Penguins_Full_Season_Combined.xlsx'
    
    if not input_file.exists():
        print(f"Error: {input_file} not found. Please ensure export has run.")
        return

    print(f"Reading {input_file}...")
    df = pd.read_csv(input_file)
    
    # --- Calculate Averages by Period ---
    # Identify numeric columns for averaging
    # Excluding specific non-averagable columns
    exclude_cols = ['Date', 'Opponent', 'Venue', 'Result', 'Period']
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    cols_to_avg = [c for c in numeric_cols if c not in exclude_cols]
    
    # Aggregation dictionary
    agg_dict = {'Result': lambda x: (x == 'W').sum()} # Count wins
    for c in cols_to_avg:
        agg_dict[c] = 'mean'
        
    summary = df.groupby('Period').agg(agg_dict)
    
    # Rename Result -> Wins and add Games/Win%
    summary = summary.rename(columns={'Result': 'Wins'})
    summary.insert(0, 'Games', df.groupby('Period').size())
    summary.insert(2, 'Win%', (summary['Wins'] / summary['Games'] * 100).round(1))
    
    # Round all FLOAT columns to 2 decimals
    for col in summary.columns:
        if col not in ['Games', 'Wins']:
            summary[col] = summary[col].round(2)
            
    # Add a "Total Season" row
    total_games = len(df)
    total_wins = (df['Result'] == 'W').sum()
    total_row = pd.DataFrame({'Games': [total_games], 'Wins': [total_wins], 'Win%': [round(total_wins/total_games*100, 1)]}, index=['Season Total'])
    
    # Calculate averages for total season for other metrics
    for col in cols_to_avg:
        total_row[col] = round(df[col].mean(), 2)
        
    summary = pd.concat([summary, total_row])

    # --- Write to Excel ---
    print(f"Writing to {output_file}...")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Sheet 1: Averages (User likely wants this first or prominent if they asked for it specifically, but standard is Data then Pivot. I'll put Averages second as "Summary")
        # Actually user call "average worksheet" implies it's a specific sheet. 
        
        df.to_excel(writer, sheet_name='Full Season Data', index=False)
        summary.to_excel(writer, sheet_name='Period Averages')
        
        # Add some formatting to Period Averages
        workbook = writer.book
        worksheet = writer.sheets['Period Averages']
        header_format = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D3D3D3'})
        
        # Apply header format
        for col_num, value in enumerate(summary.columns.values):
            worksheet.write(0, col_num + 1, value, header_format) # +1 for Index
            
    print("Done!")

if __name__ == "__main__":
    create_excel_with_averages()
