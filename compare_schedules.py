
import sys
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
except ImportError:
    print("Error: Could not import NHLAPIClient")
    sys.exit(1)

def compare_schedules():
    api = NHLAPIClient()
    print("Fetching Standings for Opponent Strength...")
    standings = api.get_standings()
    
    team_p_pct = {}
    for team in standings['standings']:
        abbr = team['teamAbbrev']['default']
        team_p_pct[abbr] = team['pointPctg']
        
    targets = ['PIT', 'CAR']
    schedules = {t: [] for t in targets}
    
    print("Fetching Remaining Schedule (This may take a moment)...")
    start_date = datetime.now()
    end_date = datetime(2026, 4, 18)
    
    current_date = start_date
    while current_date <= end_date:
        date_str = current_date.strftime("%Y-%m-%d")
        try:
            schedule = api.get_game_schedule(date_str)
            if schedule and 'gameWeek' in schedule:
                for day in schedule['gameWeek']:
                    if day['date'] != date_str: continue
                    
                    for game in day.get('games', []):
                        home = game['homeTeam']['abbrev']
                        away = game['awayTeam']['abbrev']
                        
                        for t in targets:
                            if home == t:
                                schedules[t].append(away)
                            elif away == t:
                                schedules[t].append(home)
        except:
            pass
        current_date += timedelta(days=1)

    print("\n--- SCHEDULE COMPARISON ---")
    print(f"{'Metric':<20} {targets[0]:<10} {targets[1]:<10} {'Diff':<10}")
    print("-" * 50)
    
    # Calculate Metrics
    results = {}
    for t in targets:
        opps = schedules[t]
        avg_sos = sum(team_p_pct.get(o, 0.5) for o in opps) / len(opps) if opps else 0
        
        # Hard Games (> .600)
        hard = sum(1 for o in opps if team_p_pct.get(o, 0) > 0.60)
        # Easy Games (< .450)
        easy = sum(1 for o in opps if team_p_pct.get(o, 0) < 0.45)
        
        results[t] = {
            'Games Remaining': len(opps),
            'Avg Opponent P%': avg_sos,
            'Hard Games (> .600)': hard,
            'Easy Games (< .450)': easy
        }
        
    metrics = [
        ('Games Remaining', 'Games Remaining'),
        ('Avg Opponent P%', 'Avg Opponent P%'),
        ('Hard Games (> .600)', 'Hard Games (> .600)'),
        ('Easy Games (< .450)', 'Easy Games (< .450)')
    ]
    
    for label, key in metrics:
        v1 = results[targets[0]][key]
        v2 = results[targets[1]][key]
        
        if isinstance(v1, float):
            print(f"{label:<20} {v1:>6.3f}     {v2:>6.3f}     {v1-v2:>+6.3f}")
        else:
            print(f"{label:<20} {v1:>6}     {v2:>6}     {v1-v2:>+6}")

    # Verdict
    pit_sos = results['PIT']['Avg Opponent P%']
    car_sos = results['CAR']['Avg Opponent P%']
    
    print("\nVerdict:")
    if abs(pit_sos - car_sos) < 0.01:
        print("Schedules are roughly EQUAL difficulty.")
    elif pit_sos > car_sos:
        print(f"PIT has the TOUGHER schedule.")
    else:
        print(f"CAR has the TOUGHER schedule.")

if __name__ == "__main__":
    compare_schedules()
