
from nhl_api_client import NHLAPIClient
from advanced_metrics_analyzer import AdvancedMetricsAnalyzer
import json

def verify():
    client = NHLAPIClient()
    # Using a recent game ID known to exist/work from previous interaction
    game_id = '2025020876' 
    print(f"Fetching game {game_id}...")
    pbp = client.get_play_by_play(game_id)

    if not pbp:
        print("Failed to get PBP data")
        return

    analyzer = AdvancedMetricsAnalyzer(pbp)
    
    # Get players from one team
    plays = pbp.get('plays', [])
    if not plays:
        print("No plays found")
        return
        
    team_id = plays[0].get('details', {}).get('eventOwnerTeamId')
    if not team_id:
        # scan for a team id
        for p in plays:
            t = p.get('details', {}).get('eventOwnerTeamId')
            if t:
                team_id = t
                break
    
    print(f"Analyzing Team ID: {team_id}")
    
    # Find active players (shooters) to ensure we have data
    player_counts = {}
    for play in plays:
        d = play.get('details', {})
        if d.get('eventOwnerTeamId') == team_id:
            pid = d.get('shootingPlayerId') or d.get('scoringPlayerId')
            if pid:
                player_counts[pid] = player_counts.get(pid, 0) + 1
    
    # Get top 2 active players
    top_players = sorted(player_counts.items(), key=lambda x: x[1], reverse=True)[:2]
    
    print(f"\n{'METRIC':<25} | {'PLAYER A (' + str(top_players[0][0]) + ')':<20} | {'PLAYER B (' + str(top_players[1][0]) + ')':<20}")
    print("-" * 75)
    
    p1 = top_players[0][0]
    p2 = top_players[1][0]
    
    # Snapshot calculations
    metrics = {}
    for pid in [p1, p2]:
        metrics[pid] = {}
        metrics[pid]['sq'] = analyzer.calculate_shot_quality_metrics(team_id, player_id=pid)
        metrics[pid]['mm'] = analyzer.calculate_pre_shot_movement_metrics(team_id, player_id=pid)
        metrics[pid]['tm'] = analyzer.calculate_transition_metrics(team_id, player_id=pid)
        
    # CORSI (Total Shot Attempts)
    val1 = metrics[p1]['sq'].get('total_shots', 0)
    val2 = metrics[p2]['sq'].get('total_shots', 0)
    print(f"{'Corsi (Total Shots)':<25} | {val1:<20} | {val2:<20}")
    
    # xG
    val1 = metrics[p1]['sq'].get('expected_goals', 0)
    val2 = metrics[p2]['sq'].get('expected_goals', 0)
    print(f"{'Expected Goals (xG)':<25} | {val1:<20} | {val2:<20}")
    
    # Lateral Movement Avg
    try:
        val1 = round(metrics[p1]['mm'].get('lateral_movement', {}).get('avg_delta_y', 0), 2)
    except: val1 = 0
    try:
        val2 = round(metrics[p2]['mm'].get('lateral_movement', {}).get('avg_delta_y', 0), 2)
    except: val2 = 0
    print(f"{'Avg Lateral Move (ft)':<25} | {val1:<20} | {val2:<20}")

    # Longitudinal Movement Avg
    try:
        val1 = round(metrics[p1]['mm'].get('longitudinal_movement', {}).get('avg_delta_x', 0), 2)
    except: val1 = 0
    try:
        val2 = round(metrics[p2]['mm'].get('longitudinal_movement', {}).get('avg_delta_x', 0), 2)
    except: val2 = 0
    print(f"{'Avg Long. Move (ft)':<25} | {val1:<20} | {val2:<20}")
    
    # Entries Leading to Shots
    val1 = metrics[p1]['tm'].get('entos_entries_to_shots', 0)
    val2 = metrics[p2]['tm'].get('entos_entries_to_shots', 0)
    print(f"{'Entries -> Shots':<25} | {val1:<20} | {val2:<20}")

    print("-" * 75)
    print("Verification Complete: Values should be different and non-zero.")

if __name__ == "__main__":
    verify()
