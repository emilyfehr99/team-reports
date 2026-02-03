#!/usr/bin/env python3
"""
Fetch Player Data Pipeline
Extracts player-level stats from NHL API boxscores and calculates advanced metrics.
"""

import sys
import os
import csv
from pathlib import Path
from datetime import datetime, timedelta
import argparse

# Add current directory to path
sys.path.append(str(Path(__file__).parent))

try:
    from nhl_api_client import NHLAPIClient
    from advanced_metrics_analyzer import AdvancedMetricsAnalyzer
except ImportError as e:
    print(f"Error importing dependencies: {e}")
    sys.exit(1)


class PlayerDataFetcher:
    """Fetches and processes player-level data for all NHL games."""
    
    HEADERS = [
        # Identifiers
        'game_id', 'date', 'player_id', 'player_name', 'team', 'opponent', 'home_away', 'position',
        # Basic Stats
        'toi_seconds', 'goals', 'assists', 'points', 'shots', 'hits', 'blocks', 
        'plus_minus', 'pim', 'fow', 'fol', 'fo_pct',
        # Advanced Metrics
        'ixg', 'game_score', 'xgf_on', 'xga_on', 'cf_on', 'ca_on',
        'oz_entries', 'dz_exits', 'turnovers', 'takeaways'
    ]
    
    def __init__(self):
        self.api = NHLAPIClient()
        self.rows = []
        
    def fetch_games_for_date(self, date_str: str) -> list:
        """Get all game IDs for a specific date."""
        game_ids = []
        try:
            schedule = self.api.get_game_schedule(date_str)
            if schedule and 'gameWeek' in schedule:
                for day in schedule['gameWeek']:
                    if day['date'] == date_str:
                        for game in day.get('games', []):
                            # Only process completed games
                            if game.get('gameState') == 'OFF':
                                game_ids.append(game['id'])
        except Exception as e:
            print(f"Error fetching schedule for {date_str}: {e}")
        return game_ids
    
    def parse_toi(self, toi_str: str) -> int:
        """Convert MM:SS to seconds."""
        try:
            parts = toi_str.split(':')
            return int(parts[0]) * 60 + int(parts[1])
        except:
            return 0
    
    def process_game(self, game_id: str):
        """Extract all player data from a single game."""
        print(f"  Processing game {game_id}...")
        
        try:
            data = self.api.get_comprehensive_game_data(str(game_id))
            if not data or 'boxscore' not in data:
                print(f"    Skipping - no boxscore data")
                return
            
            boxscore = data['boxscore']
            pbp = data.get('play_by_play', {})
            
            # Get game info
            game_date = boxscore.get('gameDate', '')[:10]
            home_team = boxscore['homeTeam']['abbrev']
            away_team = boxscore['awayTeam']['abbrev']
            home_id = boxscore['homeTeam']['id']
            away_id = boxscore['awayTeam']['id']
            
            # Initialize analyzer for advanced metrics
            analyzer = None
            if pbp:
                try:
                    analyzer = AdvancedMetricsAnalyzer(pbp)
                except:
                    pass
            
            # Process both teams
            player_by_game_stats = boxscore.get('playerByGameStats', {})
            
            for team_key in ['homeTeam', 'awayTeam']:
                team_data = boxscore.get(team_key, {})
                team_abbrev = team_data.get('abbrev', '')
                opponent = away_team if team_key == 'homeTeam' else home_team
                home_away = 'H' if team_key == 'homeTeam' else 'A'
                team_id = home_id if team_key == 'homeTeam' else away_id
                opp_id = away_id if team_key == 'homeTeam' else home_id
                
                # Get player stats from playerByGameStats (at boxscore root level)
                player_stats = player_by_game_stats.get(team_key, {})
                
                # Process forwards
                for player in player_stats.get('forwards', []):
                    self._add_player_row(
                        game_id, game_date, player, team_abbrev, opponent, 
                        home_away, 'F', analyzer, team_id, opp_id
                    )
                
                # Process defense
                for player in player_stats.get('defense', []):
                    self._add_player_row(
                        game_id, game_date, player, team_abbrev, opponent,
                        home_away, 'D', analyzer, team_id, opp_id
                    )
                
                # Process goalies
                for player in player_stats.get('goalies', []):
                    self._add_goalie_row(
                        game_id, game_date, player, team_abbrev, opponent,
                        home_away, analyzer, team_id, opp_id
                    )
                    
        except Exception as e:
            print(f"    Error processing game {game_id}: {e}")
    
    def _add_player_row(self, game_id, date, player, team, opponent, home_away, pos, analyzer, team_id, opp_id):
        """Add a skater row to the data."""
        player_id = player.get('playerId', 0)
        name = f"{player.get('name', {}).get('default', 'Unknown')}"
        
        # Basic stats
        toi = self.parse_toi(player.get('toi', '0:00'))
        goals = player.get('goals', 0)
        assists = player.get('assists', 0)
        shots = player.get('shots', 0)
        hits = player.get('hits', 0)
        blocks = player.get('blockedShots', 0)
        plus_minus = player.get('plusMinus', 0)
        pim = player.get('pim', 0)
        fow = player.get('faceoffWins', 0) or 0
        fol = player.get('faceoffLosses', 0) or 0
        fo_pct = round(fow / (fow + fol) * 100, 1) if (fow + fol) > 0 else 0
        
        # Advanced metrics (from analyzer if available)
        ixg = 0
        game_score = 0
        xgf_on = 0
        xga_on = 0
        cf_on = 0
        ca_on = 0
        oz_entries = 0
        dz_exits = 0
        turnovers = 0
        takeaways = 0
        
        if analyzer:
            try:
                # Calculate individual xG from shots
                sq = analyzer.calculate_shot_quality_metrics(team_id)
                # Approximate individual xG based on shot share
                team_shots = sq.get('total_shots', 1)
                team_xg = sq.get('expected_goals', 0)
                if team_shots > 0 and shots > 0:
                    ixg = round((shots / team_shots) * team_xg, 3)
                
                # Team-level on-ice approximations
                xgf_on = round(team_xg, 2)
                sq_opp = analyzer.calculate_shot_quality_metrics(opp_id)
                xga_on = round(sq_opp.get('expected_goals', 0), 2)
                
                # Corsi
                pm = analyzer.calculate_possession_metrics(team_id)
                cf_on = pm.get('corsi_for', 0)
                ca_on = pm.get('corsi_against', 0)
                
                # Transition
                tm = analyzer.calculate_transition_metrics(team_id)
                oz_entries = tm.get('zone_entries', 0)
                dz_exits = tm.get('zone_exits', 0)
                
                # Turnovers
                turnovers = player.get('giveaways', 0) or 0
                takeaways = player.get('takeaways', 0) or 0
                
                # Game Score (simplified formula)
                game_score = round(
                    goals * 0.75 + 
                    assists * 0.7 + 
                    shots * 0.075 + 
                    blocks * 0.05 + 
                    hits * 0.025 - 
                    turnovers * 0.15, 2
                )
            except:
                pass
        
        row = [
            game_id, date, player_id, name, team, opponent, home_away, pos,
            toi, goals, assists, goals + assists, shots, hits, blocks,
            plus_minus, pim, fow, fol, fo_pct,
            ixg, game_score, xgf_on, xga_on, cf_on, ca_on,
            oz_entries, dz_exits, turnovers, takeaways
        ]
        self.rows.append(row)
    
    def _add_goalie_row(self, game_id, date, player, team, opponent, home_away, analyzer, team_id, opp_id):
        """Add a goalie row to the data."""
        player_id = player.get('playerId', 0)
        name = f"{player.get('name', {}).get('default', 'Unknown')}"
        
        # Goalie-specific stats
        toi = self.parse_toi(player.get('toi', '0:00'))
        saves = player.get('saveShotsAgainst', '0/0').split('/')[0]
        shots_against = player.get('saveShotsAgainst', '0/0').split('/')[-1]
        
        try:
            saves = int(saves)
            shots_against = int(shots_against)
            goals_against = shots_against - saves
        except:
            saves = 0
            shots_against = 0
            goals_against = 0
        
        save_pct = round(saves / shots_against * 100, 1) if shots_against > 0 else 0
        
        # Advanced xGA
        xga_on = 0
        if analyzer:
            try:
                sq_opp = analyzer.calculate_shot_quality_metrics(opp_id)
                xga_on = round(sq_opp.get('expected_goals', 0), 2)
            except:
                pass
        
        # Simplified game score for goalies
        gsaa = round((xga_on - goals_against), 2) if xga_on > 0 else 0
        
        row = [
            game_id, date, player_id, name, team, opponent, home_away, 'G',
            toi, 0, 0, 0, 0, 0, 0,  # No offensive stats
            0, 0, 0, 0, 0,
            0, gsaa, 0, xga_on, 0, 0,
            0, 0, 0, 0
        ]
        self.rows.append(row)
    
    def fetch_date_range(self, start_date: str, end_date: str):
        """Fetch all games in a date range."""
        current = datetime.strptime(start_date, '%Y-%m-%d')
        end = datetime.strptime(end_date, '%Y-%m-%d')
        
        while current <= end:
            date_str = current.strftime('%Y-%m-%d')
            print(f"Fetching games for {date_str}...")
            
            game_ids = self.fetch_games_for_date(date_str)
            print(f"  Found {len(game_ids)} completed games")
            
            for gid in game_ids:
                self.process_game(gid)
            
            current += timedelta(days=1)
    
    def save_to_csv(self, output_path: str):
        """Save collected data to CSV."""
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        
        with open(output, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(self.HEADERS)
            writer.writerows(self.rows)
        
        print(f"\nSaved {len(self.rows)} player-game rows to {output}")


def main():
    parser = argparse.ArgumentParser(description='Fetch NHL player data')
    parser.add_argument('--date', help='Single date (YYYY-MM-DD)')
    parser.add_argument('--start', help='Start date for range')
    parser.add_argument('--end', help='End date for range')
    parser.add_argument('--yesterday', action='store_true', help='Fetch yesterday\'s games')
    parser.add_argument('--output', default='data/players_2025_26.csv', help='Output CSV path')
    
    args = parser.parse_args()
    
    fetcher = PlayerDataFetcher()
    
    if args.yesterday:
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        fetcher.fetch_date_range(yesterday, yesterday)
    elif args.date:
        fetcher.fetch_date_range(args.date, args.date)
    elif args.start and args.end:
        fetcher.fetch_date_range(args.start, args.end)
    else:
        print("Please specify --date, --start/--end, or --yesterday")
        sys.exit(1)
    
    fetcher.save_to_csv(args.output)


if __name__ == '__main__':
    main()
