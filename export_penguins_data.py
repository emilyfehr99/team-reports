
import json
import csv
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import sys

# Add current directory to path to ensure imports work
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
    from advanced_metrics_analyzer import AdvancedMetricsAnalyzer
except ImportError as e:
    print(f"CRITICAL ERROR: Could not import dependencies: {e}")
    sys.exit(1)

def get_team_id(abbr):
    # Mapping from abbreviation to ID (standard NHL IDs)
    team_ids = {
        'NJD': 1, 'NYI': 2, 'NYR': 3, 'PHI': 4, 'PIT': 5, 'BOS': 6, 'BUF': 7, 'MTL': 8, 'OTT': 9, 'TOR': 10,
        'CAR': 12, 'FLA': 13, 'TBL': 14, 'WSH': 15, 'CHI': 16, 'DET': 17, 'NSH': 18, 'STL': 19, 'CGY': 20,
        'COL': 21, 'EDM': 22, 'VAN': 23, 'ANA': 24, 'DAL': 25, 'LAK': 26, 'SJS': 28, 'CBJ': 29, 'MIN': 30,
        'WPG': 52, 'ARI': 53, 'VGK': 54, 'SEA': 55, 'UTA': 59
    }
    return team_ids.get(abbr.upper(), 0)

def safe_float(val, ndigits=2):
    if val is None: return 0.0
    try:
        return round(float(val), ndigits)
    except (ValueError, TypeError):
        return 0.0

def process_game_metrics(metrics, game_info, period_label):
    """Normalize metrics into a flat dictionary for CSV export"""
    home_team = game_info.get('home_team')
    away_team = game_info.get('away_team')
    is_home = (home_team == 'PIT')
    opponent = away_team if is_home else home_team
    
    # Existing logic to determine result
    actual_winner = game_info.get('actual_winner', '')
    if actual_winner == 'home': actual_winner = home_team
    elif actual_winner == 'away': actual_winner = away_team
    
    # Logic for API games (might not have actual_winner set same way)
    if not actual_winner and 'winner_team' in game_info:
        actual_winner = game_info['winner_team']

    # Determine Result
    result = 'L'
    if actual_winner == 'PIT':
        result = 'W'
    # Fallback if actual_winner is explicitly None but scores exist
    elif game_info.get('home_score') is not None and game_info.get('away_score') is not None:
        pit_score = game_info['home_score'] if is_home else game_info['away_score']
        opp_score = game_info['away_score'] if is_home else game_info['home_score']
        if pit_score > opp_score:
            result = 'W'

    row = {
        'Date': game_info.get('date'),
        'Period': period_label,
        'Opponent': opponent,
        'Venue': 'Home' if is_home else 'Away',
        'Result': result,
        'GF': game_info.get('home_score') if is_home else game_info.get('away_score'),
        'GA': game_info.get('away_score') if is_home else game_info.get('home_score'),
        'WinProb': safe_float(game_info.get('win_prob', 0))
    }

    # Helper to map keys
    def map_metric(csv_name, json_key_suffix, scale=1.0):
        # Local JSON uses home_xxx / away_xxx
        # We need to pick the right one based on is_home
        if is_home:
            val_for = metrics.get(f"home_{json_key_suffix}")
            val_against = metrics.get(f"away_{json_key_suffix}")
        else:
            val_for = metrics.get(f"away_{json_key_suffix}")
            val_against = metrics.get(f"home_{json_key_suffix}")
            
        if val_for is not None: row[f"{csv_name}_For"] = safe_float(val_for * scale)
        if val_against is not None: row[f"{csv_name}_Against"] = safe_float(val_against * scale)

    # Basic Shot Metrics
    map_metric('Shots', 'shots')
    map_metric('Corsi%', 'corsi_pct')
    map_metric('xG', 'xg')
    map_metric('HDC', 'hdc')
    map_metric('GameScore', 'gs')
    map_metric('EXtoEN', 'extoen')
    map_metric('ENtoS', 'entos')
    
    # Advanced / Zone
    map_metric('NZ_Turnovers', 'nzt')
    map_metric('NZ_Turnovers_Leading_to_SA', 'nztsa')
    map_metric('OZ_Shots', 'ozs')
    map_metric('NZ_Shots', 'nzs')
    map_metric('DZ_Shots', 'dzs')
    map_metric('Forecheck_Shots', 'fc')
    map_metric('Rush_Shots', 'rush')
    map_metric('Lateral_Move', 'lateral')
    map_metric('Longitudinal_Move', 'longitudinal')
    map_metric('PP_Goals', 'pp_goals')
    map_metric('PP_Attempts', 'pp_attempts')
    map_metric('FO_Wins', 'faceoff_wins')
    map_metric('FO_Total', 'faceoff_total')
    map_metric('Blocks', 'blocked_shots')
    map_metric('PIM', 'penalty_minutes')
    map_metric('Hits', 'hits')
    map_metric('Giveaways', 'giveaways')
    map_metric('Takeaways', 'takeaways')

    return row

def load_win_probs_map():
    """Load local JSON only to get the Win Probabilities map"""
    input_file = Path('win_probability_predictions_v2.json')
    win_probs = {}
    
    if not input_file.exists():
        return win_probs
    
    with open(input_file, 'r') as f:
        data = json.load(f)
    
    for game in data.get('predictions', []):
        date = game.get('date')
        if not date: continue
        
        # Get PIT probability
        if game.get('home_team') == 'PIT':
            prob = game.get('predicted_home_win_prob')
        elif game.get('away_team') == 'PIT':
            prob = game.get('predicted_away_win_prob')
        else:
            continue
            
        win_probs[date] = prob
        
    return win_probs

def fetch_all_api_data(win_probs_map):
    api = NHLAPIClient()
    rows = []
    
    # Start Date: Oct 1, 2025 (Season Start)
    start_date = datetime(2025, 10, 1)
    end_date = datetime.now() 
    
    current_date = start_date
    while current_date <= end_date:
        date_str = current_date.strftime("%Y-%m-%d")
        print(f"Checking date: {date_str}...")
        
        try:
            schedule = api.get_game_schedule(date_str)
            if schedule and 'gameWeek' in schedule:
                for day in schedule['gameWeek']:
                    # Ensure matching date (API returns week)
                    if day.get('date') != date_str: continue
                    
                    for game in day.get('games', []):
                        away_abbr = game.get('awayTeam', {}).get('abbrev')
                        home_abbr = game.get('homeTeam', {}).get('abbrev')
                        
                        if 'PIT' not in [away_abbr, home_abbr]:
                            continue
                            
                        # Found a PIT game
                        game_state = game.get('gameState')
                        if game_state not in ['FINAL', 'OFF']:
                            print(f"  Skipping non-final game {game.get('id')}")
                            continue

                        print(f"  Fetching data for game {game.get('id')} - {away_abbr} @ {home_abbr}")
                        
                        try:
                            full_data = api.get_comprehensive_game_data(str(game.get('id')))
                            if not full_data or 'play_by_play' not in full_data:
                                print("    No play-by-play data available.")
                                continue
                                
                            pbp = full_data['play_by_play']
                            
                            # ANALYZE!
                            analyzer = AdvancedMetricsAnalyzer(pbp)
                            
                            pit_id = get_team_id('PIT')
                            
                            metrics = {}
                            
                            for tid, side in [(get_team_id(home_abbr), 'home'), (get_team_id(away_abbr), 'away')]:
                                # 1. Shot Quality / xG / Corsi
                                sq = analyzer.calculate_shot_quality_metrics(tid)
                                metrics[f"{side}_xg"] = sq['expected_goals']
                                metrics[f"{side}_shots"] = sq['shots_on_goal']
                                metrics[f"{side}_hdc"] = sq['high_danger_shots']
                                
                                # 2. Pressure / Zone
                                press = analyzer.calculate_pressure_metrics(tid)
                                metrics[f"{side}_ozs"] = sq['shot_locations']['O']
                                metrics[f"{side}_nzs"] = sq['shot_locations']['N']
                                metrics[f"{side}_dzs"] = sq['shot_locations']['D']
                                
                                # 3. Movement
                                move = analyzer.calculate_pre_shot_movement_metrics(tid)
                                metrics[f"{side}_lateral"] = move['lateral_movement']['avg_delta_y']
                                metrics[f"{side}_longitudinal"] = move['longitudinal_movement']['avg_delta_x']
                                metrics[f"{side}_rush"] = move['longitudinal_movement']['attempts'] 

                                # 4. Defense
                                defense = analyzer.calculate_defensive_metrics(tid)
                                metrics[f"{side}_takeaways"] = defense['takeaways']
                                metrics[f"{side}_blocked_shots"] = defense['blocked_shots']
                                metrics[f"{side}_hits"] = defense['hits']

                                # 5. Giveaways
                                gv_count = 0
                                for play in pbp.get('plays', []):
                                    if play.get('typeDescKey') == 'giveaway' and play.get('details', {}).get('eventOwnerTeamId') == tid:
                                        gv_count += 1
                                metrics[f"{side}_giveaways"] = gv_count
                                
                                # 6. Transition & Game Score
                                trans = analyzer.calculate_transition_metrics(tid)
                                metrics[f"{side}_extoen"] = trans['extoen_exits_to_entries']
                                metrics[f"{side}_entos"] = trans['entos_entries_to_shots']
                                
                                metrics[f"{side}_gs"] = analyzer.calculate_game_score(tid)

                            # Calculate Aggregated Corsi %
                            home_cf = metrics.get('home_shots', 0) + metrics.get('home_blocked_shots', 0)
                            away_cf = metrics.get('away_shots', 0) + metrics.get('away_blocked_shots', 0)
                            total_corsi = home_cf + away_cf
                            if total_corsi > 0:
                                metrics['home_corsi_pct'] = (home_cf / total_corsi) * 100
                                metrics['away_corsi_pct'] = (away_cf / total_corsi) * 100
                            
                            # Determine Period
                            period = 'Pre-Jan 1' if date_str < '2026-01-01' else 'Post-Jan 1'

                            # Game Info
                            winner_team = home_abbr if game['homeTeam']['score'] > game['awayTeam']['score'] else away_abbr
                            
                            game_info = {
                                'date': date_str,
                                'home_team': home_abbr,
                                'away_team': away_abbr,
                                'actual_winner': None,
                                'winner_team': winner_team,
                                'home_score': game['homeTeam']['score'],
                                'away_score': game['awayTeam']['score'],
                                'win_prob': win_probs_map.get(date_str, 0.0) # Look up prob
                            }
                            
                            rows.append(process_game_metrics(metrics, game_info, period))
                            
                        except Exception as e:
                            print(f"    Error processing game {game.get('id')}: {e}")

        except Exception as e:
            print(f"Error checking date {date_str}: {e}")
        
        current_date += timedelta(days=1)
        
    return rows

def main():
    print("Exporting Penguins Data (Full Season Re-fetch)...")
    
    # 1. Load Local Win Probs
    print("Loading Win Probabilities from local JSON...")
    win_probs = load_win_probs_map()
    print(f"Loaded {len(win_probs)} predictions.")
    
    # 2. Fetch ALL Data
    print("Fetching ALL data from NHL API...")
    all_rows = fetch_all_api_data(win_probs)
    
    all_rows.sort(key=lambda x: x['Date'])
    
    output_file = Path('penguins_season_split_2025_26_comprehensive.csv')
    
    if all_rows:
        df = pd.DataFrame(all_rows)
        
        # Column ordering
        cols = list(df.columns)
        prioritized = ['Date', 'Period', 'Opponent', 'Venue', 'Result', 'GF', 'GA']
        remaining = sorted([c for c in cols if c not in prioritized])
        final_cols = prioritized + remaining
        
        df = df[final_cols]
        df.to_csv(output_file, index=False)
        print(f"\nSuccessfully exported {len(all_rows)} games to {output_file}")
        
        # Summary
        needed_cols = ['Result', 'GF', 'GA', 'xGF', 'xGA', 'Corsi%_For']
        available_cols = [c for c in needed_cols if c in df.columns]
        
        agg_dict = {'Result': lambda x: (x == 'W').sum()}
        for c in ['GF', 'GA', 'xGF', 'xGA', 'Corsi%_For']:
             if c in df.columns:
                 agg_dict[c] = 'mean'
                 
        summary = df.groupby('Period').agg(agg_dict).rename(columns={'Result': 'Wins'})
        summary['Games'] = df.groupby('Period').size()
        summary['Win%'] = (summary['Wins'] / summary['Games'] * 100).round(1)
        
        cols_to_print = ['Games', 'Wins', 'Win%'] + [c for c in ['GF', 'GA', 'xGF', 'xGA', 'Corsi%_For'] if c in summary.columns]
        print("\nSummary by Period:")
        print(summary[cols_to_print])
    else:
        print("No games found.")

if __name__ == "__main__":
    main()
