
import sys
from pathlib import Path
import pandas as pd

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
    from advanced_metrics_analyzer import AdvancedMetricsAnalyzer
except ImportError:
    print("Error: Could not import dependencies")
    sys.exit(1)

def compare_teams(team1, team2):
    api = NHLAPIClient()
    teams = [team1, team2]
    
    print(f"Comparing {team1} vs {team2}...\n")
    
    results = {}
    
    for team in teams:
        print(f"Analyzing {team} (Last 5 Games)...")
        game_ids = api.get_team_recent_games(team, limit=5)
        
        metrics_sum = {
            'xG_For': 0, 'xG_Against': 0,
            'Transition_For': 0, 'Transition_Against': 0,
            'DZS_Share': 0,
            'Games': 0
        }
        
        for gid in game_ids:
            try:
                data = api.get_comprehensive_game_data(str(gid))
                if not data or 'play_by_play' not in data: continue
                
                pbp = data['play_by_play']
                analyzer = AdvancedMetricsAnalyzer(pbp)
                
                home_abbr = data['boxscore']['homeTeam']['abbrev']
                target_id = data['boxscore']['homeTeam']['id'] if home_abbr == team else data['boxscore']['awayTeam']['id']
                opp_id = data['boxscore']['awayTeam']['id'] if home_abbr == team else data['boxscore']['homeTeam']['id']
                
                # Metrics
                sq_for = analyzer.calculate_shot_quality_metrics(target_id)
                sq_opp = analyzer.calculate_shot_quality_metrics(opp_id)
                metrics_sum['xG_For'] += sq_for['expected_goals']
                metrics_sum['xG_Against'] += sq_opp['expected_goals']
                
                trans_for = analyzer.calculate_transition_metrics(target_id)
                my_trans = trans_for['extoen_exits_to_entries'] + trans_for['entos_entries_to_shots']
                metrics_sum['Transition_For'] += my_trans
                
                press_opp = analyzer.calculate_pressure_metrics(opp_id)
                metrics_sum['DZS_Share'] += press_opp.get('zone_time', {}).get('O', 0)
                
                metrics_sum['Games'] += 1
            except:
                pass
                
        if metrics_sum['Games'] > 0:
            avg_xg_f = metrics_sum['xG_For'] / metrics_sum['Games']
            avg_xg_a = metrics_sum['xG_Against'] / metrics_sum['Games']
            xg_share = (avg_xg_f / (avg_xg_f + avg_xg_a)) * 100
            
            avg_trans = metrics_sum['Transition_For'] / metrics_sum['Games']
            avg_dzs = metrics_sum['DZS_Share'] / metrics_sum['Games']
            
            results[team] = {
                'xG%': xg_share,
                'xGF': avg_xg_f,
                'xGA': avg_xg_a,
                'Transition': avg_trans,
                'DZS': avg_dzs
            }
            
    print("\n--- HEAD TO HEAD METRICS ---")
    print(f"{'Metric':<20} {team1:<10} {team2:<10} {'Diff':<10}")
    print("-" * 50)
    
    t1 = results[team1]
    t2 = results[team2]
    
    metrics = [
        ('xG Share (%)', 'xG%'),
        ('xGoals For', 'xGF'),
        ('xGoals Against', 'xGA'),
        ('Transition Eff', 'Transition'),
        ('Def Zone Starts', 'DZS')
    ]
    
    for label, key in metrics:
        val1 = t1[key]
        val2 = t2[key]
        diff = val1 - val2
        print(f"{label:<20} {val1:>6.2f}     {val2:>6.2f}     {diff:>+6.2f}")

if __name__ == "__main__":
    compare_teams('PIT', 'CAR')
