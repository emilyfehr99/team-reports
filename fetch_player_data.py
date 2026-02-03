#!/usr/bin/env python3
"""
Fetch Player Data Pipeline - Comprehensive Version
Extracts player-level stats from NHL API boxscores and calculates advanced metrics.
Includes all team-level metrics attributed to each player's on-ice time.
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
        # Comprehensive Metrics - For (when player's team has possession/action)
        'GF', 'xG_For', 'Shots_For', 'HDC_For', 'Blocks_For', 'Hits_For',
        'Corsi_For', 'OZ_Shots_For', 'NZ_Shots_For', 'DZ_Shots_For', 'Rush_Shots_For',
        'ENtoS_For', 'EXtoEN_For', 'Giveaways_For', 'Takeaways_For',
        'Lateral_Move_For', 'Longitudinal_Move_For', 'GameScore_For',
        # Comprehensive Metrics - Against (when opponent has possession/action)
        'GA', 'xG_Against', 'Shots_Against', 'HDC_Against', 'Blocks_Against', 'Hits_Against',
        'Corsi_Against', 'OZ_Shots_Against', 'NZ_Shots_Against', 'DZ_Shots_Against', 'Rush_Shots_Against',
        'ENtoS_Against', 'EXtoEN_Against', 'Giveaways_Against', 'Takeaways_Against',
        'Lateral_Move_Against', 'Longitudinal_Move_Against', 'GameScore_Against',
        # Derived Ratios
        'Corsi_Pct', 'xG_Pct'
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
            home_score = boxscore['homeTeam'].get('score', 0)
            away_score = boxscore['awayTeam'].get('score', 0)
            
            # Calculate team-level metrics using analyzer
            team_metrics = {home_id: {}, away_id: {}}
            
            if pbp:
                try:
                    analyzer = AdvancedMetricsAnalyzer(pbp)
                    
                    # Process players
                    player_by_game_stats = boxscore.get('playerByGameStats', {})
                    
                    for team_key in ['homeTeam', 'awayTeam']:
                        team_data = boxscore.get(team_key, {})
                        team_abbrev = team_data.get('abbrev', '')
                        team_id = home_id if team_key == 'homeTeam' else away_id
                        opp_id = away_id if team_key == 'homeTeam' else home_id
                        opponent = away_team if team_key == 'homeTeam' else home_team
                        home_away = 'H' if team_key == 'homeTeam' else 'A'
                        team_gf = home_score if team_key == 'homeTeam' else away_score
                        team_ga = away_score if team_key == 'homeTeam' else home_score

                        # Calculate base team/opponent metrics for context
                        # We use these for "Against" columns and for goalie GSAA
                        opp_sq = analyzer.calculate_shot_quality_metrics(opp_id)
                        
                        # We also calculate "Team For" metrics just for context if needed, but we rely on player specific ones primarily
                        # team_sq = analyzer.calculate_shot_quality_metrics(team_id)
                        
                        # For Goalies, we need a "team_metrics" dict that has xG_Against
                        goalie_team_metrics = {
                            'xG_Against': opp_sq.get('expected_goals', 0)
                        }

                        # Process skaters
                        for player in player_stats.get('forwards', []) + player_stats.get('defense', []):
                            player_id = player.get('playerId')
                            
                            # Calculate INDIVIDUAL metrics
                            p_sq = analyzer.calculate_shot_quality_metrics(team_id, player_id=player_id)
                            p_tm = analyzer.calculate_transition_metrics(team_id, player_id=player_id)
                            p_mm = analyzer.calculate_pre_shot_movement_metrics(team_id, player_id=player_id)
                            p_pr = analyzer.calculate_pressure_metrics(team_id, player_id=player_id)
                            p_dm = analyzer.calculate_defensive_metrics(team_id, player_id=player_id)
                            p_gs = analyzer.calculate_game_score(team_id, player_id=player_id)
                            
                            # Opponent context metrics (Team Level)
                            opp_tm = analyzer.calculate_transition_metrics(opp_id)
                            opp_mm = analyzer.calculate_pre_shot_movement_metrics(opp_id)
                            opp_pr = analyzer.calculate_pressure_metrics(opp_id)
                            opp_dm = analyzer.calculate_defensive_metrics(opp_id)
                            opp_gs = analyzer.calculate_game_score(opp_id)
                            
                            # Package individual metrics
                            player_metrics = {
                                # For metrics (Individual)
                                'xG_For': p_sq.get('expected_goals', 0),
                                'Shots_For': p_sq.get('shots_on_goal', 0),
                                'HDC_For': p_sq.get('high_danger_shots', 0),
                                'Corsi_For': p_sq.get('total_shots', 0), # Individual Corsi (iCorsi)
                                'OZ_Shots_For': p_sq.get('shot_locations', {}).get('O', 0),
                                'NZ_Shots_For': p_sq.get('shot_locations', {}).get('N', 0),
                                'DZ_Shots_For': p_sq.get('shot_locations', {}).get('D', 0),
                                'Rush_Shots_For': p_pr.get('quick_strike_opportunities', 0),
                                'ENtoS_For': p_tm.get('entos_entries_to_shots', 0),
                                'EXtoEN_For': p_tm.get('extoen_exits_to_entries', 0),
                                'Lateral_Move_For': p_mm.get('lateral_movement', {}).get('avg_delta_y', 0) if isinstance(p_mm.get('lateral_movement'), dict) else 0,
                                'Longitudinal_Move_For': p_mm.get('longitudinal_movement', {}).get('avg_delta_x', 0) if isinstance(p_mm.get('longitudinal_movement'), dict) else 0,
                                'GameScore_For': p_gs if isinstance(p_gs, (int, float)) else 0,
                                # Defense (Individual Blocks/Hits)
                                'Blocks_For': p_dm.get('blocked_shots', 0),
                                'Hits_For': p_dm.get('hits', 0),
                                
                                # Against metrics (Team Context)
                                'xG_Against': opp_sq.get('expected_goals', 0),
                                'Shots_Against': opp_sq.get('shots_on_goal', 0),
                                'HDC_Against': opp_sq.get('high_danger_shots', 0),
                                'Corsi_Against': opp_sq.get('total_shots', 0),
                                'OZ_Shots_Against': opp_sq.get('shot_locations', {}).get('O', 0),
                                'NZ_Shots_Against': opp_sq.get('shot_locations', {}).get('N', 0),
                                'DZ_Shots_Against': opp_sq.get('shot_locations', {}).get('D', 0),
                                'Rush_Shots_Against': opp_pr.get('quick_strike_opportunities', 0),
                                'ENtoS_Against': opp_tm.get('entos_entries_to_shots', 0),
                                'EXtoEN_Against': opp_tm.get('extoen_exits_to_entries', 0),
                                'Lateral_Move_Against': opp_mm.get('lateral_movement', {}).get('avg_delta_y', 0) if isinstance(opp_mm.get('lateral_movement'), dict) else 0,
                                'Longitudinal_Move_Against': opp_mm.get('longitudinal_movement', {}).get('avg_delta_x', 0) if isinstance(opp_mm.get('longitudinal_movement'), dict) else 0,
                                'GameScore_Against': opp_gs if isinstance(opp_gs, (int, float)) else 0,
                                'Blocks_Against': opp_dm.get('blocked_shots', 0),
                                'Hits_Against': opp_dm.get('hits', 0),
                            }

                            self._add_player_row(
                                game_id, game_date, player, team_abbrev, opponent,
                                home_away, 'F' if player in player_stats.get('forwards', []) else 'D',
                                player_metrics, team_gf, team_ga, team_id
                            )
                        
                        # Process goalies
                        for player in player_stats.get('goalies', []):
                            self._add_goalie_row(
                                game_id, game_date, player, team_abbrev, opponent,
                                home_away, goalie_team_metrics, team_gf, team_ga
                            )
                    
                except Exception as e:
                    print(f"    Error processing game {game_id}: {e}")
    
    def _add_player_row(self, game_id, date, player, team, opponent, home_away, pos, team_metrics, team_gf, team_ga, team_id):
        """Add a skater row with comprehensive metrics."""
        player_id = player.get('playerId', 0)
        name = player.get('name', {}).get('default', 'Unknown')
        
        # Basic stats from boxscore
        toi = self.parse_toi(player.get('toi', '0:00'))
        goals = player.get('goals', 0)
        assists = player.get('assists', 0)
        shots = player.get('sog', 0) or player.get('shots', 0) or 0
        hits = player.get('hits', 0)
        blocks = player.get('blockedShots', 0)
        plus_minus = player.get('plusMinus', 0)
        pim = player.get('pim', 0)
        fow = player.get('faceoffWins', 0) or 0
        fol = player.get('faceoffLosses', 0) or 0
        fo_pct = round(fow / (fow + fol) * 100, 1) if (fow + fol) > 0 else 0
        giveaways = player.get('giveaways', 0) or 0
        takeaways = player.get('takeaways', 0) or 0
        
        # Get team-level metrics (attributed to all on-ice players equally for now)
        xg_for = round(team_metrics.get('xG_For', 0), 3)
        xg_against = round(team_metrics.get('xG_Against', 0), 3)
        shots_for = team_metrics.get('Shots_For', 0)
        shots_against = team_metrics.get('Shots_Against', 0)
        hdc_for = team_metrics.get('HDC_For', 0)
        hdc_against = team_metrics.get('HDC_Against', 0)
        blocks_for = blocks  # Individual
        blocks_against = team_metrics.get('Blocks_Against', 0)
        hits_for = hits  # Individual
        hits_against = team_metrics.get('Hits_Against', 0)
        corsi_for = team_metrics.get('Corsi_For', 0)
        corsi_against = team_metrics.get('Corsi_Against', 0)
        oz_shots_for = team_metrics.get('OZ_Shots_For', 0)
        oz_shots_against = team_metrics.get('OZ_Shots_Against', 0)
        nz_shots_for = team_metrics.get('NZ_Shots_For', 0)
        nz_shots_against = team_metrics.get('NZ_Shots_Against', 0)
        dz_shots_for = team_metrics.get('DZ_Shots_For', 0)
        dz_shots_against = team_metrics.get('DZ_Shots_Against', 0)
        rush_shots_for = team_metrics.get('Rush_Shots_For', 0)
        rush_shots_against = team_metrics.get('Rush_Shots_Against', 0)
        entos_for = team_metrics.get('ENtoS_For', 0)
        entos_against = team_metrics.get('ENtoS_Against', 0)
        extoen_for = team_metrics.get('EXtoEN_For', 0)
        extoen_against = team_metrics.get('EXtoEN_Against', 0)
        lat_move_for = round(team_metrics.get('Lateral_Move_For', 0), 2)
        lat_move_against = round(team_metrics.get('Lateral_Move_Against', 0), 2)
        long_move_for = round(team_metrics.get('Longitudinal_Move_For', 0), 2)
        long_move_against = round(team_metrics.get('Longitudinal_Move_Against', 0), 2)
        
        # Game Score (standard formula)
        game_score_for = round(
            goals * 0.75 + assists * 0.7 + shots * 0.075 +
            blocks * 0.05 + hits * 0.025 + takeaways * 0.1 - giveaways * 0.15, 2
        )
        game_score_against = 0  # Would need opponent player data
        
        # Derived ratios
        corsi_pct = round(corsi_for / (corsi_for + corsi_against) * 100, 1) if (corsi_for + corsi_against) > 0 else 50.0
        xg_pct = round(xg_for / (xg_for + xg_against) * 100, 1) if (xg_for + xg_against) > 0 else 50.0
        
        row = [
            game_id, date, player_id, name, team, opponent, home_away, pos,
            toi, goals, assists, goals + assists, shots, hits, blocks,
            plus_minus, pim, fow, fol, fo_pct,
            # For metrics
            team_gf, xg_for, shots_for, hdc_for, blocks_for, hits_for,
            corsi_for, oz_shots_for, nz_shots_for, dz_shots_for, rush_shots_for,
            entos_for, extoen_for, giveaways, takeaways,
            lat_move_for, long_move_for, game_score_for,
            # Against metrics
            team_ga, xg_against, shots_against, hdc_against, blocks_against, hits_against,
            corsi_against, oz_shots_against, nz_shots_against, dz_shots_against, rush_shots_against,
            entos_against, extoen_against, 0, 0,  # Opponent giveaways/takeaways
            lat_move_against, long_move_against, game_score_against,
            # Derived
            corsi_pct, xg_pct
        ]
        self.rows.append(row)
    
    def _add_goalie_row(self, game_id, date, player, team, opponent, home_away, team_metrics, team_gf, team_ga):
        """Add a goalie row with relevant metrics."""
        player_id = player.get('playerId', 0)
        name = player.get('name', {}).get('default', 'Unknown')
        toi = self.parse_toi(player.get('toi', '0:00'))
        
        # Parse goalie stats
        save_shots = player.get('saveShotsAgainst', '0/0')
        try:
            saves = int(save_shots.split('/')[0])
            shots_against = int(save_shots.split('/')[-1])
        except:
            saves = 0
            shots_against = 0
        
        xg_against = round(team_metrics.get('xG_Against', 0), 3)
        
        # Goalie game score: saves above expected
        gsaa = round(xg_against - (shots_against - saves), 2) if xg_against > 0 else 0
        
        # Build row with mostly zeros for offensive stats
        row = [
            game_id, date, player_id, name, team, opponent, home_away, 'G',
            toi, 0, 0, 0, 0, 0, 0,  # Basic stats
            0, 0, 0, 0, 0,
            # For metrics (minimal for goalies)
            team_gf, 0, 0, 0, 0, 0,
            0, 0, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, gsaa,
            # Against metrics
            team_ga, xg_against, shots_against, team_metrics.get('HDC_Against', 0), 0, 0,
            team_metrics.get('Corsi_Against', 0), 0, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0,
            # Derived
            0, 0
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
