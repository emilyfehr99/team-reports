
import sys
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import random
import time

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
    from advanced_metrics_analyzer import AdvancedMetricsAnalyzer
except ImportError:
    print("Error: Could not import dependencies")
    sys.exit(1)

class AdvancedPlayoffProjector:
    def __init__(self):
        self.api = NHLAPIClient()
        self.team_stats = {}
        self.schedule = []
        # Store teams by conference for easier processing later
        self.teams_by_conf = {'Eastern': [], 'Western': []}
        
        # Power Score Weights
        self.WEIGHTS = {
            'xG_Share': 0.35,
            'GameScore': 0.25,
            'Transition': 0.20, # EXtoEN / ENtoS
            'Defense': 0.20     # DZS / Supression
        }

    def fetch_current_standings(self):
        print("Fetching Current Standings...")
        standings = self.api.get_standings()
        if not standings or 'standings' not in standings:
            print("Error: Could not fetch standings.")
            return

        for team in standings['standings']:
            abbr = team['teamAbbrev']['default']
            conf = team['conferenceName']
            div = team['divisionName']
            
            # Store in main dict
            self.team_stats[abbr] = {
                'points': team['points'],
                'gp': team['gamesPlayed'],
                'p_pct': team['pointPctg'],
                'wins': team['wins'],
                'name': team['teamName']['default'],
                'div': div,
                'conf': conf,
                'power_score': 0.5, # Default average
                'recent_metrics': {} 
            }
            
            # Add to conference list
            if conf in self.teams_by_conf:
                self.teams_by_conf[conf].append(abbr)
                
        print(f"Found {len(self.team_stats)} teams total.")

    def calculate_team_power_scores(self):
        print("\nCalculating Team Power Scores (based on last 5 games)...")
        print("This requires analyzing play-by-play data for recent games.")
        
        # Analyze all teams
        for team in self.team_stats:
            print(f"  Analyzing {team}...", end='', flush=True)
            
            # Fetch recent games
            game_ids = self.api.get_team_recent_games(team, limit=5)
            
            metrics_sum = {
                'xG_For': 0, 'xG_Against': 0,
                'GS_For': 0, 'GS_Against': 0,
                'Transition_Diff': 0, # (EXtoEN + ENtoS) - Against
                'DZS_Share': 0 
            }
            
            games_count = 0
            
            for gid in game_ids:
                try:
                    data = self.api.get_comprehensive_game_data(str(gid))
                    if not data or 'play_by_play' not in data: continue
                    
                    pbp = data['play_by_play']
                    analyzer = AdvancedMetricsAnalyzer(pbp)
                    
                    # Determine Team ID
                    home_id = data['boxscore']['homeTeam']['id']
                    away_id = data['boxscore']['awayTeam']['id']
                    home_abbr = data['boxscore']['homeTeam']['abbrev']
                    
                    target_id = home_id if home_abbr == team else away_id
                    opp_id = away_id if home_abbr == team else home_id
                    
                    # 1. Shot Quality / xG
                    sq_for = analyzer.calculate_shot_quality_metrics(target_id)
                    sq_against = analyzer.calculate_shot_quality_metrics(opp_id)
                    
                    metrics_sum['xG_For'] += sq_for['expected_goals']
                    metrics_sum['xG_Against'] += sq_against['expected_goals']
                    
                    # 2. Game Score
                    metrics_sum['GS_For'] += analyzer.calculate_game_score(target_id)
                    
                    # 3. Transition
                    trans_for = analyzer.calculate_transition_metrics(target_id)
                    trans_against = analyzer.calculate_transition_metrics(opp_id)
                    
                    # Net Transition efficiency
                    my_trans = trans_for['extoen_exits_to_entries'] + trans_for['entos_entries_to_shots']
                    opp_trans = trans_against['extoen_exits_to_entries'] + trans_against['entos_entries_to_shots']
                    metrics_sum['Transition_Diff'] += (my_trans - opp_trans)
                    
                    # 4. Defense (Zone Starts/Activity)
                    press_against = analyzer.calculate_pressure_metrics(opp_id)
                    metrics_sum['DZS_Share'] += press_against.get('zone_time', {}).get('O', 0)
                    
                    games_count += 1
                    
                except Exception as e:
                    # print(f"x ({e})", end='')
                    pass
            
            if games_count > 0:
                # Normalize Metrics
                avg_xg_f = metrics_sum['xG_For'] / games_count
                avg_xg_a = metrics_sum['xG_Against'] / games_count
                xg_share = avg_xg_f / (avg_xg_f + avg_xg_a) if (avg_xg_f + avg_xg_a) > 0 else 0.5
                
                avg_gs = metrics_sum['GS_For'] / games_count
                gs_norm = min(max((avg_gs - 3) / 10, 0), 1) 
                
                # Transition
                avg_trans_diff = metrics_sum['Transition_Diff'] / games_count
                trans_score = 0.5 + (avg_trans_diff * 0.01) # Reduced sensitivity
                trans_score = min(max(trans_score, 0), 1)
                
                # Defense
                avg_dzs_time = metrics_sum['DZS_Share'] / games_count
                def_score = 1.0 - min(max((avg_dzs_time - 5) / 15, 0), 1)
                
                # Metric-Based Score
                metric_score = (
                    (xg_share * self.WEIGHTS['xG_Share']) +
                    (gs_norm * self.WEIGHTS['GameScore']) +
                    (trans_score * self.WEIGHTS['Transition']) +
                    (def_score * self.WEIGHTS['Defense'])
                )
                
                # NEW: Anchor in Season P%
                # This prevents teams with current low points but high xG from "teleporting" to the top
                season_p_pct = self.team_stats[team].get('p_pct', 0.5)
                
                # Weighted Blend: 60% Season Results / 40% Recent Underlying Metrics
                # This makes the "Power Score" a measure of current form RELATIVE to season baseline
                final_power_score = (season_p_pct * 0.60) + (metric_score * 0.40)
                
                self.team_stats[team]['power_score'] = round(final_power_score, 3)
                print(f" Score: {self.team_stats[team]['power_score']:.3f} (P%: {season_p_pct:.3f})")
            else:
                print(" No data.")

    def fetch_remaining_schedule(self):
        print("\nFetching Remaining Schedule (ALL Teams)...")
        start_date = datetime.now()
        end_date = datetime(2026, 4, 18)
        
        current_date = start_date
        while current_date <= end_date:
            date_str = current_date.strftime("%Y-%m-%d")
            try:
                schedule = self.api.get_game_schedule(date_str)
                if schedule and 'gameWeek' in schedule:
                    for day in schedule['gameWeek']:
                        if day['date'] != date_str: continue
                        
                        for game in day.get('games', []):
                            home = game['homeTeam']['abbrev']
                            away = game['awayTeam']['abbrev']
                            
                            # Simulate ALL games now
                            self.schedule.append({
                                'date': date_str,
                                'home': home,
                                'away': away,
                                'id': game['id']
                            })
            except:
                pass
                
            current_date += timedelta(days=1)
            
        print(f"Found {len(self.schedule)} remaining games.")

    def simulate_season(self):
        print("\nSimulating Season with Contextual Adjustments...")
        
        # Track rest days: Last played date for each team
        last_played = {team: None for team in self.team_stats}
        
        for game in self.schedule:
            home = game['home']
            away = game['away']
            g_date = datetime.strptime(game['date'], "%Y-%m-%d")
            
            # Skip if teams not in our stats (e.g. hypothetical non-NHL games?)
            if home not in self.team_stats or away not in self.team_stats:
                continue

            # 1. Base Win Probability from Power Scores
            p_home = self.team_stats[home]['power_score']
            p_away = self.team_stats[away]['power_score']
            
            base_prob = 0.5 + (p_home - p_away)
            
            # 2. Context Adjustments
            base_prob += 0.05 # Home Ice
            
            # Rest (Back-to-Back)
            if last_played.get(home) == (g_date - timedelta(days=1)):
                 base_prob -= 0.08
            
            if last_played.get(away) == (g_date - timedelta(days=1)):
                 base_prob += 0.08
            
            # Cap Prob
            win_prob = max(0.1, min(0.9, base_prob))
            
            # 3. Simulate Result (Expected Value)
            exp_pts_home = (win_prob * 2) + ((1 - win_prob) * 0.24) 
            exp_pts_away = ((1 - win_prob) * 2) + (win_prob * 0.24)
            
            # Update Standings
            self.team_stats[home]['points'] += exp_pts_home
            self.team_stats[home]['gp'] += 1
            last_played[home] = g_date
            
            self.team_stats[away]['points'] += exp_pts_away
            self.team_stats[away]['gp'] += 1
            last_played[away] = g_date

    def print_results(self):
        print("\n=== FINAL PROJECTED PLAYOFF PICTURE ===")
        print("Based on Advanced Metrics Power Score & Schedule Simulation\n")
        
        # Sort all by points for overall checking
        sorted_teams = sorted(self.team_stats.values(), key=lambda x: x['points'], reverse=True)
        
        def process_conference(conf_name, div1_name, div2_name):
            print(f"\n>> {conf_name.upper()} CONFERENCE <<")
            
            conf_teams = [t for t in sorted_teams if t['conf'] == conf_name]
            
            # Breakdown by Division
            div1 = []
            div2 = []
            
            for team in conf_teams:
                if team['div'] == div1_name:
                    div1.append(team)
                elif team['div'] == div2_name:
                    div2.append(team)
                    
            # Top 3 per Division
            div1_top3 = div1[:3]
            div2_top3 = div2[:3]
            
            # Wild Card Pool (Everyone else)
            wc_pool = div1[3:] + div2[3:]
            wc_pool.sort(key=lambda x: x['points'], reverse=True)
            
            wild_cards = wc_pool[:2]
            outside = wc_pool[2:]
            
            # Helper Print
            def print_row(rank, team, label=""):
                name = team['name']
                pts = team['points']
                ps = team['power_score']
                is_pit = ("Penguins" in name)
                marker = "<<" if is_pit else ""
                print(f"{rank:<4} {name[:3].upper():<5} {pts:.1f}      {ps:.3f}   {label:<10} {marker}")

            print(f"{'Rank':<4} {'Team':<5} {'Proj Pts':<10} {'Power Score':<12} {'Slot':<10}")
            print("-" * 60)
            
            print(f"{div1_name.upper()} DIVISION")
            for i, t in enumerate(div1_top3): print_row(i+1, t, div1_name[0] + str(i+1))
            
            print(f"\n{div2_name.upper()} DIVISION")
            for i, t in enumerate(div2_top3): print_row(i+1, t, div2_name[0] + str(i+1))
            
            print("\nWILD CARDS")
            for i, t in enumerate(wild_cards): print_row(i+1, t, "WC" + str(i+1))
            
            print("\nIN THE HUNT")
            for i, t in enumerate(outside): 
                if i < 3: print_row(i+3, t, "Out") 

        # Process East
        process_conference('Eastern', 'Atlantic', 'Metropolitan')
        
        # Process West
        process_conference('Western', 'Central', 'Pacific')
        
        # Check PIT Status
        pit_entry = next((t for t in self.team_stats.values() if "Penguins" in t['name']), None)
        if pit_entry:
            print(f"\nPenguins Status Check:")
            print(f"Projected Points: {pit_entry['points']:.1f}")
            print(f"Conf Rank: {next(i for i, t in enumerate([x for x in sorted_teams if x['conf']=='Eastern']) if t['name'] == pit_entry['name']) + 1}")


if __name__ == "__main__":
    projector = AdvancedPlayoffProjector()
    projector.fetch_current_standings()
    projector.calculate_team_power_scores()
    projector.fetch_remaining_schedule()
    projector.simulate_season()
    projector.print_results()
