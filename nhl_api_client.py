import requests
import json
from datetime import datetime, timedelta
import pandas as pd

class NHLAPIClient:
    def __init__(self):
        self.base_url = "https://api-web.nhle.com/v1"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
    
    def get_team_info(self, team_id):
        """Get team information by team ID"""
        url = f"{self.base_url}/teams/{team_id}"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None
    
    def get_team_roster(self, team_id):
        """Get team roster by team ID"""
        url = f"{self.base_url}/teams/{team_id}/roster"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None
    
    def get_game_schedule(self, date=None):
        """Get game schedule for a specific date"""
        if date is None:
            date = datetime.now().strftime("%Y-%m-%d")
        
        url = f"{self.base_url}/schedule/{date}"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None
    
    def get_game_center(self, game_id):
        """Get detailed game information by combining boxscore and play-by-play"""
        # Get boxscore data
        boxscore_url = f"{self.base_url}/gamecenter/{game_id}/boxscore"
        boxscore_response = self.session.get(boxscore_url)
        
        # Get play-by-play data
        pbp_url = f"{self.base_url}/gamecenter/{game_id}/play-by-play"
        pbp_response = self.session.get(pbp_url)
        
        if boxscore_response.status_code == 200 and pbp_response.status_code == 200:
            boxscore_data = boxscore_response.json()
            pbp_data = pbp_response.json()
            
            # Combine the data
            combined_data = {
                'boxscore': boxscore_data,
                'play_by_play': pbp_data
            }
            return combined_data
        return None
    
    def get_game_boxscore(self, game_id):
        """Get game boxscore"""
        url = f"{self.base_url}/gamecenter/{game_id}/boxscore"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None
    
    def get_player_stats(self, player_id):
        """Get player statistics"""
        url = f"{self.base_url}/players/{player_id}/stats"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None
    
    def find_recent_game(self, team1_abbrev, team2_abbrev, days_back=30):
        """Find the most recent game between two teams"""
        # Get team IDs from abbreviations
        team_ids = {
            'FLA': 13, 'EDM': 22, 'BOS': 6, 'TOR': 10, 'MTL': 8, 'OTT': 9,
            'BUF': 7, 'DET': 17, 'TBL': 14, 'CAR': 12, 'WSH': 15, 'PIT': 5,
            'NYR': 3, 'NYI': 2, 'NJD': 1, 'PHI': 4, 'CBJ': 29, 'NSH': 18,
            'STL': 19, 'MIN': 30, 'WPG': 52, 'COL': 21, 'ARI': 53, 'VGK': 54,
            'SJS': 28, 'LAK': 26, 'ANA': 24, 'CGY': 20, 'VAN': 23, 'SEA': 55,
            'CHI': 16, 'DAL': 25
        }
        
        team1_id = team_ids.get(team1_abbrev.upper())
        team2_id = team_ids.get(team2_abbrev.upper())
        
        if not team1_id or not team2_id:
            raise ValueError(f"Team abbreviation not found: {team1_abbrev} or {team2_id}")
        
        # Search through recent dates for the game
        for i in range(days_back):
            date = datetime.now() - timedelta(days=i)
            date_str = date.strftime("%Y-%m-%d")
            
            schedule = self.get_game_schedule(date_str)
            if schedule and 'gameWeek' in schedule:
                for day in schedule['gameWeek']:
                    for game in day.get('games', []):
                        if (game.get('awayTeam', {}).get('id') == team1_id and 
                            game.get('homeTeam', {}).get('id') == team2_id) or \
                           (game.get('awayTeam', {}).get('id') == team2_id and 
                            game.get('homeTeam', {}).get('id') == team1_id):
                            return game['id']
        
        return None
    
    def get_stanley_cup_finals_game(self):
        """Get the most recent Stanley Cup Finals game between FLA and EDM"""
        # For Stanley Cup Finals, we'll look for recent games between these teams
        # Since this is a specific request, we'll search for recent matchups
        return self.find_recent_game('FLA', 'EDM', days_back=60)
    
    def get_play_by_play(self, game_id):
        """Get play-by-play data for a game"""
        url = f"{self.base_url}/gamecenter/{game_id}/play-by-play"
        response = self.session.get(url)
        if response.status_code == 200:
            return response.json()
        return None

    def get_comprehensive_game_data(self, game_id):
        """Get comprehensive game data including boxscore and play-by-play"""
        game_center = self.get_game_center(game_id)
        boxscore = self.get_game_boxscore(game_id)
        play_by_play = self.get_play_by_play(game_id)
        
        print(f"Debug - Game Center data: {game_center is not None}")
        print(f"Debug - Boxscore data: {boxscore is not None}")
        print(f"Debug - Play-by-play data: {play_by_play is not None}")
        
        # If we have boxscore but no game_center, create a minimal game_center from boxscore
        if boxscore is not None and game_center is None:
            print("Creating minimal game_center from boxscore data...")
            game_center = {
                'game': {
                    'gameDate': '2024-03-04',  # Default date
                    'awayTeamScore': boxscore['awayTeam']['score'],
                    'homeTeamScore': boxscore['homeTeam']['score'],
                    'awayTeamScoreByPeriod': [0, 0, 0, 0],  # Default periods
                    'homeTeamScoreByPeriod': [0, 0, 0, 0]
                },
                'awayTeam': {
                    'abbrev': boxscore['awayTeam']['abbrev'],
                    'name': boxscore['awayTeam'].get('name', 'Away Team')
                },
                'homeTeam': {
                    'abbrev': boxscore['homeTeam']['abbrev'],
                    'name': boxscore['homeTeam'].get('name', 'Home Team')
                },
                'venue': {
                    'default': 'Unknown Arena'
                }
            }
        
        if game_center is None or boxscore is None:
            print("Warning: Missing game data, returning None")
            return None
        
        return {
            'game_center': game_center,
            'boxscore': boxscore,
            'play_by_play': play_by_play
        }

    def get_team_recent_games(self, team_abbr, limit=5):
        """Get recent game IDs for a team"""
        # Get team ID
        team_ids = {
            'FLA': 13, 'EDM': 22, 'BOS': 6, 'TOR': 10, 'MTL': 8, 'OTT': 9,
            'BUF': 7, 'DET': 17, 'TBL': 14, 'CAR': 12, 'WSH': 15, 'PIT': 5,
            'NYR': 3, 'NYI': 2, 'NJD': 1, 'PHI': 4, 'CBJ': 29, 'NSH': 18,
            'STL': 19, 'MIN': 30, 'WPG': 52, 'COL': 21, 'ARI': 53, 'VGK': 54,
            'SJS': 28, 'LAK': 26, 'ANA': 24, 'CGY': 20, 'VAN': 23, 'SEA': 55,
            'CHI': 16, 'DAL': 25, 'UTA': 59
        }
        team_id = team_ids.get(team_abbr.upper())
        if not team_id:
            return []

        # We need to find recent games. The schedule endpoint is by date.
        # A better way is to use the team schedule endpoint if available, 
        # or search backwards.
        # For now, we'll search backwards from today.
        
        game_ids = []
        days_back = 0
        max_days = 60 # Look back 2 months max
        
        while len(game_ids) < limit and days_back < max_days:
            date = datetime.now() - timedelta(days=days_back)
            date_str = date.strftime("%Y-%m-%d")
            
            try:
                schedule = self.get_game_schedule(date_str)
                if schedule and 'gameWeek' in schedule:
                    for day in schedule['gameWeek']:
                        for game in day.get('games', []):
                            if game.get('gameState') in ['FINAL', 'OFF']:
                                if (game.get('awayTeam', {}).get('id') == team_id or 
                                    game.get('homeTeam', {}).get('id') == team_id):
                                    game_ids.append(game['id'])
            except Exception as e:
                print(f"Error fetching schedule for {date_str}: {e}")
                
            days_back += 1
            
        return game_ids[:limit]
