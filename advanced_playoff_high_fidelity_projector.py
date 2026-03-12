import sys
import json
import random
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

# Add current directory to path
sys.path.append(str(Path.cwd()))

try:
    from nhl_api_client import NHLAPIClient
except ImportError:
    print("Error: Could not import dependencies")
    sys.exit(1)

# City Coordinates for travel fatigue
NHL_CITIES = {
    'ANA': (33.8, -117.9), 'UTA': (40.7, -111.9), 'BOS': (42.3, -71.0), 'BUF': (42.8, -78.8),
    'CGY': (51.0, -114.0), 'CAR': (35.8, -78.6), 'CHI': (41.8, -87.6), 'COL': (39.7, -104.9),
    'CBJ': (39.9, -83.0), 'DAL': (32.7, -96.8), 'DET': (42.3, -83.0), 'EDM': (53.5, -113.5),
    'FLA': (26.1, -80.1), 'LAK': (34.0, -118.2), 'MIN': (44.9, -93.1), 'MTL': (45.5, -73.5),
    'NSH': (36.1, -86.7), 'NJD': (40.7, -74.1), 'NYI': (40.7, -73.0), 'NYR': (40.7, -73.9),
    'OTT': (45.4, -75.7), 'PHI': (39.9, -75.1), 'PIT': (40.4, -80.0), 'SJS': (37.3, -121.8),
    'SEA': (47.6, -122.3), 'STL': (38.6, -90.2), 'TBL': (27.9, -82.4), 'TOR': (43.6, -79.3),
    'VAN': (49.2, -123.1), 'VGK': (36.1, -115.1), 'WSH': (38.8, -77.0), 'WPG': (49.8, -97.1)
}

def calculate_distance(city1, city2):
    if city1 == city2: return 0
    c1 = NHL_CITIES.get(city1)
    c2 = NHL_CITIES.get(city2)
    if not c1 or not c2: return 1000 # Penalty for unknown
    return np.sqrt((c1[0]-c2[0])**2 + (c1[1]-c2[1])**2) * 60 # Approx miles

class HighFidelityProjector:
    def __init__(self):
        self.api = NHLAPIClient()
        self.team_stats = {}
        self.schedule = []
        self.teams_by_conf = {'Eastern': [], 'Western': []}
        
        # Availability Config (Dynamic)
        # Malkin returns Mar 16. Crosby returns Mar 21.
        self.AVAILABILITY = {
            'PIT': {
                'Malkin_Back': datetime(2026, 3, 16),
                'Crosby_Back': datetime(2026, 3, 21)
            }
        }
    
    def fetch_current_standings(self):
        print("Fetching Current Standings...")
        standings = self.api.get_standings()
        for team in standings['standings']:
            abbr = team['teamAbbrev']['default']
            self.team_stats[abbr] = {
                'points': team['points'],
                'gp': team['gamesPlayed'],
                'p_pct': team['pointPctg'],
                'rw': team.get('regulationWins', 0),
                'name': team['teamName']['default'],
                'div': team['divisionName'],
                'conf': team['conferenceName']
            }
            if team['conferenceName'] in self.teams_by_conf:
                self.teams_by_conf[team['conferenceName']].append(abbr)

    def fetch_schedule(self):
        print("Fetching Full Remaining Schedule...")
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
                                'away': g['awayTeam']['abbrev']
                            })
            d += timedelta(days=1)

    def run_simulation(self, iterations=3000):
        print(f"Running {iterations} High-Fidelity Simulations...")
        # Load Recent Power Scores as Baseline (We'll use our anchored method)
        # We'll calculate a baseline Power Score based on Season P%
        for team in self.team_stats:
            self.team_stats[team]['base_ps'] = self.team_stats[team]['p_pct']
            
        final_points = {team: [] for team in self.team_stats}
        playoff_counts = {team: 0 for team in self.team_stats}
        
        for i in range(iterations):
            temp_stats = {t: {'pts': self.team_stats[t]['points'], 'rw': self.team_stats[t]['rw']} for t in self.team_stats}
            last_played = {t: None for t in self.team_stats}
            last_location = {t: t for t in self.team_stats}
            
            for g in self.schedule:
                h, a = g['home'], g['away']
                g_date = datetime.strptime(g['date'], "%Y-%m-%d")
                
                # 1. Base Prob from Power Scores
                ps_h = self.team_stats[h]['base_ps']
                ps_a = self.team_stats[a]['base_ps']
                
                # Apply Dynamic Injury Penalty (PIT SPECIFIC)
                if h == 'PIT':
                    if g_date < self.AVAILABILITY['PIT']['Malkin_Back']: 
                        ps_h -= 0.18 # Both out
                    elif g_date < self.AVAILABILITY['PIT']['Crosby_Back']:
                        ps_h -= 0.10 # Only Mallkin back
                if a == 'PIT':
                    if g_date < self.AVAILABILITY['PIT']['Malkin_Back']: 
                        ps_a -= 0.18
                    elif g_date < self.AVAILABILITY['PIT']['Crosby_Back']:
                        ps_a -= 0.10
                
                prob = 0.5 + (ps_h - ps_a) + 0.05 # Home Ice Adv
                
                # 2. Travel Fatigue
                dist_h = calculate_distance(last_location[h], h)
                dist_a = calculate_distance(last_location[a], h) # Travel to Away City
                
                if last_played[h] == (g_date - timedelta(days=1)):
                    prob -= 0.05 # B2B Penalty
                    if dist_h > 500: prob -= 0.03 # Travel Penalty
                    
                if last_played[a] == (g_date - timedelta(days=1)):
                    prob += 0.05 
                    if dist_a > 500: prob += 0.03
                
                # 3. Random Result
                roll = random.random()
                if roll < (prob * 0.75): # Reg Win
                    temp_stats[h]['pts'] += 2
                    temp_stats[h]['rw'] += 1
                elif roll < prob: # OT Win
                    temp_stats[h]['pts'] += 2
                    temp_stats[a]['pts'] += 1
                elif roll < (prob + (1-prob)*0.25): # OT Loss
                    temp_stats[h]['pts'] += 1
                    temp_stats[a]['pts'] += 2
                else: # Reg Loss
                    temp_stats[a]['pts'] += 2
                    temp_stats[a]['rw'] += 1
                
                last_played[h] = last_played[a] = g_date
                last_location[h] = h # Home
                last_location[a] = h # Away Team is now in Home city
            
            for t in self.team_stats:
                final_points[t].append(temp_stats[t]['pts'])
            
            # Check Standings
            for conf in ['Eastern', 'Western']:
                conf_teams = sorted(self.teams_by_conf[conf], key=lambda x: (temp_stats[x]['pts'], temp_stats[x]['rw']), reverse=True)
                for t in conf_teams[:8]:
                    playoff_counts[t] += (1 / iterations)
                    
        return final_points, playoff_counts

def main():
    proj = HighFidelityProjector()
    proj.fetch_current_standings()
    proj.fetch_schedule()
    
    final_pts, playoff_probs = proj.run_simulation(iterations=3000)
    
    print("\n" + "="*70)
    print("🏆 NHL HIGH-FIDELITY PLAYOFF PROJECTION (MAR 11)")
    print("RECOVERING ROSTER: CROSBY/MALKIN RETURN WINDOWS & TRAVEL ACCOUNTED")
    print("="*70)
    
    east_teams = sorted(proj.teams_by_conf['Eastern'], key=lambda x: np.mean(final_pts[x]), reverse=True)
    print(f"{'Team':<5} | {'Avg Pts':<8} | {'Playoff %':<10} | {'Status'}")
    print("-" * 70)
    for t in east_teams:
        avg = np.mean(final_pts[t])
        prob = playoff_probs[t] * 100
        marker = "<< PIT ALERT" if t == "PIT" else ""
        status = "IN" if prob > 50 else "OUT"
        print(f"{t:<5} | {avg:<8.1f} | {prob:<10.1f}% | {status:<5} {marker}")

if __name__ == "__main__":
    main()
