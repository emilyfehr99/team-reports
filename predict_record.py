import sys
from pathlib import Path
from datetime import datetime, timedelta

sys.path.append(str(Path.cwd()))
from advanced_playoff_projector import AdvancedPlayoffProjector

projector = AdvancedPlayoffProjector()
projector.fetch_current_standings()
projector.calculate_team_power_scores()
projector.fetch_remaining_schedule()

wins = 0
losses = 0
otl = 0
total_pts = 0

print("\n--- Penguins Remaining Schedule Prediction ---")
for game in projector.schedule:
    home = game['home']
    away = game['away']
    if home != 'PIT' and away != 'PIT':
        continue
        
    p_home = projector.team_stats.get(home, {}).get('power_score', 0.5)
    p_away = projector.team_stats.get(away, {}).get('power_score', 0.5)
    
    base_prob = 0.5 + (p_home - p_away)
    base_prob += 0.05 # Home Ice Adv
    win_prob = max(0.1, min(0.9, base_prob))
    
    pit_win_prob = win_prob if home == 'PIT' else (1 - win_prob)
    opp = away if home == 'PIT' else home
    venue = "vs" if home == 'PIT' else "@"
    
    # Predict result based on thresholds
    if pit_win_prob >= 0.53:
        res = "W"
        wins += 1
        total_pts += 2
    elif pit_win_prob <= 0.47:
        res = "L"
        losses += 1
    else:
        # Toss up / Overtime games
        if pit_win_prob >= 0.50:
            res = "OTW"
            wins += 1
            total_pts += 2
        else:
            res = "OTL"
            otl += 1
            total_pts += 1
            
    print(f"{game['date']}: {venue} {opp:<4} (Win Prob: {pit_win_prob*100:>4.1f}%) -> {res}")

print("-" * 45)
print(f"Predicted Remaining Record: {wins}W - {losses}L - {otl}OTL")
print(f"Points Earned: {total_pts}")
print(f"Projected Final Points: {projector.team_stats['PIT']['points'] + total_pts}")
