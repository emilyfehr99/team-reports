import pandas as pd
import sys
from pathlib import Path

def create_styled_excel():
    input_csv = Path('/Users/emilyfehr8/CascadeProjects/data/players_2025_26.csv')
    output_xlsx = Path.home() / 'Desktop' / 'Player_Stats_Report_2025_26.xlsx'
    
    if not input_csv.exists():
        print(f"Error: {input_csv} not found")
        return

    # Load data
    df = pd.read_csv(input_csv)
    
    # Filter out goalies
    if 'position' in df.columns:
        df = df[df['position'] != 'G']
    
    # Define columns to exclude from mean aggregation
    exclude_cols = ['game_id', 'player_id', 'player_name', 'team', 'opponent', 'home_away', 'position', 'date']
    
    # Define aggregation dictionary
    agg_dict = {'game_id': 'count'} # Count games via game_id
    
    # Identify and clean all metric columns
    fo_cols = ['fow', 'fol']
    move_cols = ['Lateral_Move_For', 'Longitudinal_Move_For']
    
    for col in df.columns:
        if col not in exclude_cols:
            # Force to numeric and fill NaNs with 0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            if col in fo_cols:
                agg_dict[col] = 'sum'
            else:
                agg_dict[col] = 'mean'
            
    # Create Summary by Player (Group by ID, Name, and Team)
    summary = df.groupby(['player_id', 'player_name', 'team']).agg(agg_dict).reset_index()
    summary = summary.rename(columns={'game_id': 'Games'})
    
    # 1. Fix Faceoff % (Calculate from Totals)
    summary['fo_pct'] = (summary['fow'] / (summary['fow'] + summary['fol']) * 100).fillna(0).round(1)
    
    # 2. Convert TOI to Minutes
    if 'toi_seconds' in summary.columns:
        summary['TOI_Min_GP'] = (summary['toi_seconds'] / 60).round(1)
        summary = summary.drop(columns=['toi_seconds'])
    
    # 3. Movement Classifications
    def classify_lateral(avg_feet):
        if avg_feet == 0: return "Stationary"
        elif avg_feet < 10: return "Minor"
        elif avg_feet < 20: return "Cross-ice"
        elif avg_feet < 35: return "Wide-lane"
        else: return "Full-width"
    
    def classify_longitudinal(avg_feet):
        if avg_feet == 0: return "Stationary"
        elif avg_feet < 15: return "Close-range"
        elif avg_feet < 30: return "Mid-range"
        elif avg_feet < 50: return "Extended"
        else: return "Long-range"

    if 'Lateral_Move_For' in summary.columns:
        summary['Lateral_Movement'] = summary['Lateral_Move_For'].apply(classify_lateral)
        summary = summary.drop(columns=['Lateral_Move_For'])
    
    if 'Longitudinal_Move_For' in summary.columns:
        summary['Vertical_Movement'] = summary['Longitudinal_Move_For'].apply(classify_longitudinal)
        summary = summary.drop(columns=['Longitudinal_Move_For'])

    # Round all remaining numeric columns
    numeric_cols = summary.select_dtypes(include=['number']).columns
    summary[numeric_cols] = summary[numeric_cols].round(2)
    
    # Calculate Shooting % if shots > 0
    if 'goals' in summary.columns and 'shots' in summary.columns:
        summary['Sh_Pct'] = (summary['goals'] / summary['shots'] * 100).fillna(0).round(1)
    
    # Set player_name as index for better readability in Excel
    summary = summary.set_index('player_name').drop(columns=['player_id'])
    
    # Reorder columns
    # We want Identifiers (Name is index) + Games + Metrics
    
    # Create Excel Writer
    try:
        writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    except ImportError:
        print("xlsxwriter not found, cannot apply formatting.")
        return

    # Write Summary (Only sheet)
    summary.to_excel(writer, sheet_name='Season Averages')
    
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
