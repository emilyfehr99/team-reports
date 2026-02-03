import pandas as pd
import sys
from pathlib import Path

def create_styled_excel():
    input_csv = Path('team-reports/data/players_2025_26.csv')
    output_xlsx = Path.home() / 'Desktop' / 'Player_Stats_Report_2025_26.xlsx'
    
    if not input_csv.exists():
        print(f"Error: {input_csv} not found")
        return

    # Load data
    df = pd.read_csv(input_csv)
    
    # Identify numeric columns for aggregation
    # We want to aggregate ALL numeric columns except 'Date'
    numeric_df = df.select_dtypes(include=['number'])
    
    # Define aggregation dictionary
    agg_dict = {'game_id': 'count'} # Count games via game_id
    
    # Add all numeric columns to aggregation
    for col in numeric_df.columns:
        if col != 'game_id' and col != 'player_id':
             agg_dict[col] = 'mean'
            
    # Create Summary by Player
    # "condensed into the per game averages... one row per player"
    summary = df.groupby('player_name').agg(agg_dict).rename(columns={'game_id': 'Games'})
    
    # Reorder columns
    # We want Identifiers (Name is index) + Games + Metrics
    
    # Create Excel Writer
    try:
        writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    except ImportError:
        print("xlsxwriter not found, cannot apply formatting.")
        return

    # Write Summary
    summary.to_excel(writer, sheet_name='Season Averages')
    
    # Write Raw Data
    df.to_excel(writer, sheet_name='Game Log', index=False)
    
    # Access Workbook / Worksheet for formatting
    workbook = writer.book
    worksheet = writer.sheets['Season Averages']
    
    # Formats
    # Blue/Green for Good
    good_props = {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFFFFF', 'max_color': '#63BE7B'}
    # Red for Bad (High is Bad)
    bad_props = {'type': '3_color_scale', 'min_color': '#63BE7B', 'mid_color': '#FFFFFF', 'max_color': '#F8696B'}
    
    # Metric Logic
    # Standard 'For' -> High is Good (Green)
    # Standard 'Against' -> High is Bad (Red)
    # Exceptions: 'Giveaways', 'Turnovers', 'PIM'
    
    # We will determine direction for each column
    summary_cols = summary.columns.tolist()
    
    for col_name in summary_cols:
        col_idx = summary_cols.index(col_name) + 1 # +1 for Index (Name)
        
        is_high_good = True
        
        if col_name.endswith('_Against'):
            is_high_good = False
        elif 'Giveaways' in col_name or 'PIM' in col_name:
            is_high_good = False  # Giveaways For is Bad
        elif col_name.endswith('_Pct') or col_name.endswith('%'):
            is_high_good = True
        elif col_name in ['GA', 'xGA', 'Shots_Against']:
            is_high_good = False
            
        if col_name == 'Games' or 'id' in col_name: continue
        
        props = good_props if is_high_good else bad_props
        worksheet.conditional_format(1, col_idx, len(summary), col_idx, props)

    # Adjust widths
    worksheet.set_column(0, 0, 15) # Index
    worksheet.set_column(1, len(summary_cols), 12) # Data columns
    
    writer.close()
    print(f"Excel report saved to: {output_xlsx}")

if __name__ == "__main__":
    create_styled_excel()
