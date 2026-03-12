from nhl_api_client import NHLAPIClient
import json

client = NHLAPIClient()
game_id = "2025020810" # Recent game
data = client.get_game_boxscore(game_id)

if data:
    # Look at one forward’s stats
    home_skaters = data.get('playerByGameStats', {}).get('homeTeam', {}).get('forwards', [])
    if home_skaters:
        player = home_skaters[0]
        print(f"Player: {player.get('name', {}).get('default')}")
        print(f"Keys: {list(player.keys())}")
        print(f"Faceoff stats: fow={player.get('faceoffWins')}, fol={player.get('faceoffLosses')}")
        # Note: New API might use different keys like 'faceoffWin' or 'faceoffsWon'
        # Let's save the whole player object to a file to inspect.
        with open('debug_player.json', 'w') as f:
            json.dump(player, f, indent=2)
else:
    print("Failed to fetch game data")
