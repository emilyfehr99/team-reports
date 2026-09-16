import requests
import json
from datetime import datetime, timedelta
import pandas as pd

class NHLAPIClient:
    def __init__(self, timeout: int = 10):
        self.base_url = "https://api-web.nhle.com/v1"
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
    
    def _safe_get(self, url: str, timeout: int = None):
        """Helper to perform session.get with timeout and error handling"""
        t = timeout or self.timeout
        try:
            resp = self.session.get(url, timeout=t)
            if resp.status_code == 200:
                return resp.json()
        except Exception:
            pass
        return None

    def get_team_info(self, team_id):
        """Get team information by team ID"""
        url = f"{self.base_url}/teams/{team_id}"
        return self._safe_get(url)
    
    def get_team_roster(self, team_id):
        """Get team roster by team ID"""
        url = f"{self.base_url}/teams/{team_id}/roster"
        return self._safe_get(url)
    
    def get_game_schedule(self, date=None):
        """Get game schedule for a specific date"""
        if date is None:
            date = datetime.now().strftime("%Y-%m-%d")
        
        url = f"{self.base_url}/schedule/{date}"
        return self._safe_get(url)
    
    def get_game_center(self, game_id):
        """Get detailed game information by combining boxscore and play-by-play"""
        boxscore_data = self.get_game_boxscore(game_id)
        pbp_data = self.get_play_by_play(game_id)
        
        if boxscore_data and pbp_data:
            return {
                'boxscore': boxscore_data,
                'play_by_play': pbp_data
            }
        return None
    
    def get_game_boxscore(self, game_id):
        """Get game boxscore"""
        url = f"{self.base_url}/gamecenter/{game_id}/boxscore"
        return self._safe_get(url)
    
    def get_player_stats(self, player_id):
        """Get player statistics"""
        url = f"{self.base_url}/players/{player_id}/stats"
        return self._safe_get(url)
    
    def find_recent_game(self, team1_abbrev, team2_abbrev, days_back=30):
        """Find the most recent game between two teams"""
        team_ids = {
            'FLA': 13, 'EDM': 22, 'BOS': 6, 'TOR': 10, 'MTL': 8, 'OTT': 9,
            'BUF': 7, 'DET': 17, 'TBL': 14, 'CAR': 12, 'WSH': 15, 'PIT': 5,
            'NYR': 3, 'NYI': 2, 'NJD': 1, 'PHI': 4, 'CBJ': 29, 'NSH': 18,
            'STL': 19, 'MIN': 30, 'WPG': 52, 'COL': 21, 'ARI': 53, 'VGK': 54,
            'SJS': 28, 'LAK': 26, 'ANA': 24, 'CGY': 20, 'VAN': 23, 'SEA': 55,
            'CHI': 16, 'DAL': 25, 'UTA': 59
        }
        
        team1_id = team_ids.get(team1_abbrev.upper())
        team2_id = team_ids.get(team2_abbrev.upper())
        
        if not team1_id or not team2_id:
            raise ValueError(f"Team abbreviation not found: {team1_abbrev} or {team2_id}")
        
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
        return self.find_recent_game('FLA', 'EDM', days_back=60)
    
    def get_play_by_play(self, game_id):
        """Get play-by-play data for a game"""
        url = f"{self.base_url}/gamecenter/{game_id}/play-by-play"
        return self._safe_get(url)

    def get_shift_charts(self, game_id):
        """Get shift chart data (legacy API)"""
        url = f"https://api.nhle.com/stats/rest/en/shiftcharts?cayenneExp=gameId={game_id}"
        data = self._safe_get(url)
        if data and isinstance(data, dict) and 'data' in data:
            return data['data']
        return None

    def get_comprehensive_game_data(self, game_id):
        """Get comprehensive game data including boxscore and play-by-play"""
        boxscore = self.get_game_boxscore(game_id)
        play_by_play = self.get_play_by_play(game_id)
        
        if boxscore is None:
            return None
        
        game_center = {
            'game': {
                'gameDate': boxscore.get('gameDate', '2025-10-07'),
                'awayTeamScore': boxscore.get('awayTeam', {}).get('score', 0),
                'homeTeamScore': boxscore.get('homeTeam', {}).get('score', 0),
                'awayTeamScoreByPeriod': [0, 0, 0, 0],
                'homeTeamScoreByPeriod': [0, 0, 0, 0]
            },
            'awayTeam': {
                'abbrev': boxscore.get('awayTeam', {}).get('abbrev', 'AWAY'),
                'name': boxscore.get('awayTeam', {}).get('name', 'Away Team')
            },
            'homeTeam': {
                'abbrev': boxscore.get('homeTeam', {}).get('abbrev', 'HOME'),
                'name': boxscore.get('homeTeam', {}).get('name', 'Home Team')
            },
            'venue': {
                'default': 'NHL Arena'
            }
        }
        
        return {
            'game_center': game_center,
            'boxscore': boxscore,
            'play_by_play': play_by_play
        }

    def get_team_recent_games(self, team_abbr, limit=5):
        """Get recent game IDs for a team"""
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
        
        game_ids = []
        days_back = 0
        max_days = 60
        
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

    def get_standings(self):
        """Get current league standings"""
        url = f"{self.base_url}/standings/now"
        return self._safe_get(url)
