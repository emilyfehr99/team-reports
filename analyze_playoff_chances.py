
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
except ImportError:
    print("Error: Could not import NHLAPIClient")
    sys.exit(1)

def analyze_playoff_chances():
    api = NHLAPIClient()
    
    print("Fetching Standings and Schedule...")
    
    # 1. Get Current Standings
    standings = api.get_standings()
    if not standings or 'standings' not in standings:
        print("Error: Could not fetch standings.")
        return

    # Process Standings to get P% and Current Points
    team_stats = {}
    east_teams = []
    
    for team in standings['standings']:
        abbr = team['teamAbbrev']['default']
        points = team['points']
        games_played = team['gamesPlayed']
        p_pct = team['pointPctg']
        conference = team['conferenceName']
        division = team['divisionName']
        
        team_stats[abbr] = {
            'points': points,
            'gp': games_played,
            'p_pct': p_pct,
            'conf': conference,
            'div': division,
            'name': team['teamName']['default']
        }
        
        if conference == 'Eastern':
            east_teams.append(abbr)

    # 2. Get Remaining Schedule for Penguins
    print("Calculating Remaining Strength of Schedule (SOS)...")
    
    # We need to find remaining games.
    # Start from tomorrow (or today if game not played)
    start_date = datetime.now()
    # End of Season (Approx mid-April 2026)
    end_date = datetime(2026, 4, 18) 
    
    remaining_opponents = []
    
    current_date = start_date
    while current_date <= end_date:
        date_str = current_date.strftime("%Y-%m-%d")
        schedule = api.get_game_schedule(date_str)
        
        if schedule and 'gameWeek' in schedule:
            for day in schedule['gameWeek']:
                if day['date'] != date_str: continue
                
                for game in day.get('games', []):
                    home = game['homeTeam']['abbrev']
                    away = game['awayTeam']['abbrev']
                    
                    if 'PIT' == home:
                        remaining_opponents.append(away)
                    elif 'PIT' == away:
                        remaining_opponents.append(home)
                        
        current_date += pd.Timedelta(days=1)

    # 3. Calculate SOS
    if not remaining_opponents:
        print("No remaining games found (Season over?).")
        return

    total_opp_p_pct = 0
    opp_count = 0
    
    for opp in remaining_opponents:
        if opp in team_stats:
            total_opp_p_pct += team_stats[opp]['p_pct']
            opp_count += 1
            
    avg_sos = total_opp_p_pct / opp_count if opp_count > 0 else 0
    
    print(f"\nRemaining Games: {len(remaining_opponents)}")
    print(f"Average Opponent Point % (SOS): {avg_sos:.3f}")
    
    # Compare SOS context
    # Is .550 high? .500 average?
    sos_difficulty = "Average"
    if avg_sos > 0.58: sos_difficulty = "Brutal"
    elif avg_sos > 0.54: sos_difficulty = "Tough"
    elif avg_sos < 0.48: sos_difficulty = "Easy"
    
    print(f"Schedule Difficulty: {sos_difficulty}")
    
    # 4. Playoff Projection
    print("\n--- Playoff Projection (Eastern Conference) ---")
    
    projections = []
    
    for team in east_teams:
        stats = team_stats[team]
        remaining_games = 82 - stats['gp']
        
        # Projection Method: Continue at current Pace
        proj_points_pace = stats['points'] + (remaining_games * stats['p_pct'] * 2)
        
        # Adjust for SOS? (Simplified: Just use Pace for now to determine 'Cut Line')
        
        projections.append({
            'Team': team,
            'Points': stats['points'],
            'GP': stats['gp'],
            'P%': stats['p_pct'],
            'Proj_Points': round(proj_points_pace, 1)
        })
        
    projections.sort(key=lambda x: x['Proj_Points'], reverse=True)
    
    # Determine Cut Line
    # Top 3 in Atlantic, Top 3 in Metro = 6 spots.
    # Next 2 highest = Wild Cards.
    # We need to sort by Division to be precise, but roughly top 8 in Conf usually calculate the line.
    
    print(f"{'Team':<5} {'Pts':<5} {'GP':<3} {'P%':<6} {'Proj Pts':<10}")
    print("-" * 40)
    
    for i, p in enumerate(projections):
        is_pit = (p['Team'] == 'PIT')
        marker = "<<" if is_pit else ""
        print(f"{p['Team']:<5} {p['Points']:<5} {p['GP']:<3} {p['P%']:.3f}  {p['Proj_Points']:<10} {marker}")
        
    # Find PIT rank
    pit_rank = next((i for i, p in enumerate(projections) if p['Team'] == 'PIT'), -1) + 1
    
    # Estimated Cutoff (8th place)
    cutoff_points = projections[7]['Proj_Points'] if len(projections) >= 8 else 0
    
    print(f"\nProjected Cutoff (8th Place): {cutoff_points}")
    
    pit_proj = next((p for p in projections if p['Team'] == 'PIT'), None)
    
    if pit_proj:
        gap = pit_proj['Proj_Points'] - cutoff_points
        if gap >= 0:
            print(f"\nPrediction: IN PLAYOFFS (+{gap:.1f} pts above cut)")
            print(f"The Penguins are currently on pace to make it.")
        else:
            print(f"\nPrediction: OUT OF PLAYOFFS ({gap:.1f} pts below cut)")
            print(f"They need to improve their pace significantly.")
            
            # How much better do they need to play?
            # Points needed: cutoff_points - current_points
            # PPG needed: (needed) / remaining_games
            remaining = 82 - pit_proj['GP']
            if remaining > 0:
                points_needed = cutoff_points - pit_proj['Points']
                # Add a buffer of 1 point to be safe
                target_record_pts = points_needed + 1
                ppg_needed = target_record_pts / remaining
                record_needed_p_pct = ppg_needed / 2
                
                print(f"To reach {cutoff_points} pts, they need ~{target_record_pts:.0f} pts in remaining {remaining} games.")
                print(f"Required Point %: {record_needed_p_pct:.3f} (Current: {pit_proj['P%']:.3f})")

if __name__ == "__main__":
    analyze_playoff_chances()
