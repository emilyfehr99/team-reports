import sys
import json
import re
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import random

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
except ImportError:
    print("Error: Could not import dependencies")
    sys.exit(1)

class DeepPlayoffProjector:
    def __init__(self):
        self.api = NHLAPIClient()
        self.team_stats = {}
        self.schedule = []
        self.teams_by_conf = {'Eastern': [], 'Western': []}
        
        # Injury Config
        self.INJURY_PENALTIES = {
            'PIT': 0.18 # Crosby & Malkin out (Major impact on C depth/PP)
        }
        
        # Weighting: 70% Season Baseline, 30% Recent Metrics
        self.BASE_WEIGHT = 0.70
        
    def load_season_data(self):
        print("Loading Season Data from JSON...")
        stats_path = Path("/Users/emilyfehr8/CascadeProjects/automated-post-game-reports/data/season_2025_2026_team_stats.json")
        with open(stats_path, 'r') as f:
            data = json.load(f)
            self.raw_data = data['teams']
            
    def fetch_current_standings(self):
        print("Fetching Current Standings...")
        standings = self.api.get_standings()
        for team in standings['standings']:
            abbr = team['teamAbbrev']['default']
            self.team_stats[abbr] = {
                'points': team['points'],
                'gp': team['gamesPlayed'],
                'p_pct': team['pointPctg'],
                'rw': team.get('regulationWins', 0), # Regulation Wins for tiebreaker
                'name': team['teamName']['default'],
                'div': team['divisionName'],
                'conf': team['conferenceName'],
                'base_power': team['pointPctg'] 
            }
            if team['conferenceName'] in self.teams_by_conf:
                self.teams_by_conf[team['conferenceName']].append(abbr)

    def calculate_deep_power_scores(self):
        print("Calculating Deep Power Scores (Season + Metrics + Injuries)...")
        for team in self.team_stats:
            # Get underlying and result metrics from JSON
            team_json = self.raw_data.get(team, {})
            h_xg = team_json.get('home', {}).get('xg', [0.5])
            a_xg = team_json.get('away', {}).get('xg', [0.5])
            h_gs = team_json.get('home', {}).get('gs', [5.0])
            a_gs = team_json.get('away', {}).get('gs', [5.0])
            
            # Use mean of last 10 games as "Recent Form"
            recent_xg = np.mean(h_xg[-5:] + a_xg[-5:]) if (h_xg and a_xg) else 0.5
            recent_gs = np.mean(h_gs[-5:] + a_gs[-5:]) if (h_gs and a_gs) else 5.0
            
            # Normalize Form (xG 2.5 is avg, GS 6.0 is avg)
            form_score = (recent_xg / 5.0) + (recent_gs / 12.0) 
            form_score = min(max(form_score, 0), 1)
            
            # Season Context
            base = self.team_stats[team]['p_pct']
            
            # Penalty
            penalty = self.INJURY_PENALTIES.get(team, 0)
            
            # Final Blend
            final_ps = (base * self.BASE_WEIGHT) + (form_score * (1 - self.BASE_WEIGHT))
            final_ps -= penalty
            
            self.team_stats[team]['power_score'] = round(final_ps, 3)
            # print(f"  {team:<4} | Base: {base:.3f} | Form: {form_score:.3f} | Penalty: {penalty:.2f} | Final: {final_ps:.3f}")

    def fetch_schedule(self):
        print("Fetching Remaining Schedule...")
        # Today is March 11
        d = datetime(2026, 3, 11)
        end = datetime(2026, 4, 18)
        while d <= end:
            ds = d.strftime("%Y-%m-%d")
            sched = self.api.get_game_schedule(ds)
            if sched and 'gameWeek' in sched:
                for day in sched['gameWeek']:
                    if day['date'] == ds:
                        for g in day.get('games', []):
                            self.schedule.append({
                                'date': ds,
                                'home': g['homeTeam']['abbrev'],
                                'away': g['awayTeam']['abbrev'],
                                'div_game': self.team_stats[g['homeTeam']['abbrev']]['div'] == self.team_stats[g['awayTeam']['abbrev']]['div']
                            })
            d += timedelta(days=1)

    def run_simulation(self, iterations=5000):
        print(f"Running {iterations} Season Simulations...")
        final_points = {team: [] for team in self.team_stats}
        playoff_counts = {team: 0 for team in self.team_stats}
        
        for p in range(iterations):
            temp_stats = {t: {'pts': self.team_stats[t]['points'], 'rw': self.team_stats[t]['rw']} for t in self.team_stats}
            last_played = {t: None for t in self.team_stats}
            
            for g in self.schedule:
                h, a = g['home'], g['away']
                g_date = datetime.strptime(g['date'], "%Y-%m-%d")
                
                # 1. Base Prob
                p_h = self.team_stats[h]['power_score']
                p_a = self.team_stats[a]['power_score']
                prob = 0.5 + (p_h - p_a) + 0.05 # Home Adv
                
                # 2. Fatigue (B2B or Travel)
                if last_played[h] == (g_date - timedelta(days=1)): prob -= 0.07
                if last_played[a] == (g_date - timedelta(days=1)): prob += 0.07
                
                # 3. Random Sample
                roll = random.random()
                if roll < (prob * 0.75): # Regulation Win
                    temp_stats[h]['pts'] += 2
                    temp_stats[h]['rw'] += 1
                elif roll < prob: # OT Win
                    temp_stats[h]['pts'] += 2
                    temp_stats[a]['pts'] += 1
                elif roll < (prob + (1-prob)*0.25): # OT Loss
                    temp_stats[h]['pts'] += 1
                    temp_stats[a]['pts'] += 2
                else: # Regulation Loss
                    temp_stats[a]['pts'] += 2
                    temp_stats[a]['rw'] += 1
                
                last_played[h] = last_played[a] = g_date
            
            # Record iteration results
            for t in self.team_stats:
                final_points[t].append(temp_stats[t]['pts'])
            
            # Check Playoff Status for this iteration
            # (Simple: Top 8 per Conference)
            for conf in ['Eastern', 'Western']:
                conf_teams = sorted(self.teams_by_conf[conf], key=lambda x: (temp_stats[x]['pts'], temp_stats[x]['rw']), reverse=True)
                for t in conf_teams[:8]:
                    playoff_counts[t] += (1 / iterations)
                    
        return final_points, playoff_counts

def main():
    proj = DeepPlayoffProjector()
    proj.load_season_data()
    proj.fetch_current_standings()
    proj.calculate_deep_power_scores()
    proj.fetch_schedule()
    
    final_pts, playoff_probs = proj.run_simulation(iterations=2000)
    
    print("\n" + "="*60)
    print("🏒 PENINSULA COVE / NHL DEEP PLAYOFF PROJECTION (MARCH 11)")
    print("CRISIS MODE: CROSBY & MALKIN IMPACT")
    print("="*60)
    
    # Sort East
    east_teams = sorted(proj.teams_by_conf['Eastern'], key=lambda x: np.mean(final_pts[x]), reverse=True)
    
    print(f"{'Team':<5} | {'Avg Pts':<8} | {'Playoff %':<10} | {'Status'}")
    print("-" * 60)
    for t in east_teams:
        avg = np.mean(final_pts[t])
        prob = playoff_probs[t] * 100
        marker = "<< PIT ALERT" if t == "PIT" else ""
        status = "IN" if prob > 50 else "OUT"
        print(f"{t:<5} | {avg:<8.1f} | {prob:<10.1f}% | {status:<5} {marker}")

    # Specific PIT analysis
    pit_avg = np.mean(final_pts['PIT'])
    pit_prob = playoff_probs['PIT'] * 100
    print("\n--- Penguins Special Breakdown ---")
    print(f"Projected Points (W/ Injury Penalty): {pit_avg:.1f}")
    print(f"Slam Dunk Likelihood: {pit_prob:.1f}%")
    print(f"Danger Zone: {100 - pit_prob:.1f}% chance of missing.")
    print("Factors: -18% Power Score penalty for loss of Crosby/Malkin.")

if __name__ == "__main__":
    main()
