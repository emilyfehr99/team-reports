
import pandas as pd
import sys
from pathlib import Path

def split_csv():
    try:
        input_file = Path("penguins_season_split_2025_26_comprehensive.csv")
        if not input_file.exists():
            print(f"Error: {input_file} not found.")
            return

        print(f"Reading {input_file}...")
        df = pd.read_csv(input_file)
        
        # Split by Period column
        pre_df = df[df['Period'] == 'Pre-Jan 1']
        post_df = df[df['Period'] == 'Post-Jan 1']
        
        pre_file = "penguins_pre_jan1.csv"
        post_file = "penguins_post_jan1.csv"
        
        print(f"Saving {len(pre_df)} rows to {pre_file}...")
        pre_df.to_csv(pre_file, index=False)
        
        print(f"Saving {len(post_df)} rows to {post_file}...")
        post_df.to_csv(post_file, index=False)
        
        print("Done!")
        
    except Exception as e:
        print(f"Error splitting CSV: {e}")

if __name__ == "__main__":
    split_csv()
