import sys
import re
from pathlib import Path

sys.path.append(str(Path.cwd()))
from advanced_playoff_projector import AdvancedPlayoffProjector

class FastProjector(AdvancedPlayoffProjector):
    def calculate_team_power_scores(self):
        print("Parsing power scores from previous run...")
        with open('playoff_projection_output.txt', 'r') as f:
            text = f.read()
            # Find all instances of "Analyzing XXX... Score: 0.XXX"
            # It might have debug text in between.
            matches = re.finditer(r"Analyzing ([A-Z]{3})\.\.\..*?Score:\s*([\d\.]+)", text, re.DOTALL)
            for m in matches:
                team = m.group(1)
                score = float(m.group(2))
                if team in self.team_stats:
                    self.team_stats[team]['power_score'] = score
        
        # Ensure PIT has a score
        if self.team_stats.get('PIT', {}).get('power_score', 0) == 0.5:
             self.team_stats['PIT']['power_score'] = 0.501

def main():
    projector = FastProjector()
    projector.fetch_current_standings()
    projector.calculate_team_power_scores()
    projector.fetch_remaining_schedule()
    
    wins = 0
    losses = 0
    otls = 0
    pts_earned = 0
    
    print("\n" + "="*50)
    print("🐧 PITTSBURGH PENGUINS: REMAINING GAME PREDICTIONS")
    print("="*50)
    
    for game in projector.schedule:
        home = game['home']
        away = game['away']
        if home != 'PIT' and away != 'PIT':
            continue
            
        p_home = projector.team_stats.get(home, {}).get('power_score', 0.5)
        p_away = projector.team_stats.get(away, {}).get('power_score', 0.5)
        
        # Base win probability calculation
        base_prob = 0.5 + (p_home - p_away)
        base_prob += 0.05 # Home ice
        win_prob = max(0.1, min(0.9, base_prob))
        
        pit_prob = win_prob if home == 'PIT' else (1 - win_prob)
        opp = away if home == 'PIT' else home
        venue = "vs" if home == 'PIT' else "@ "
        
        # Determine likely outcome based on probability thresholds
        if pit_prob >= 0.54:
            res = "W "
            wins += 1
            pts_earned += 2
        elif pit_prob <= 0.46:
            res = "L "
            losses += 1
        else:
            # Toss-up games
            if pit_prob >= 0.50:
                res = "OTW"
                wins += 1
                pts_earned += 2
            else:
                res = "OTL"
                otls += 1
                pts_earned += 1
                
        print(f"{game['date']} | {venue} {opp:<4} | Win Prob: {pit_prob*100:>4.1f}% | Predicted: {res}")
        
    print("-" * 50)
    print(f"Projected Final Record (Remaining): {wins}W - {losses}L - {otls}OTL")
    print(f"Points Picked Up: {pts_earned}")
    print(f"Total Season Points: {projector.team_stats['PIT']['points'] + pts_earned}")
    print("="*50)

if __name__ == "__main__":
    main()
