
"""
Advanced NHL Metrics Analyzer
Creates custom hockey analytics from play-by-play data
"""

import json
import csv
import os
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Tuple, Optional
from improved_xg_model import ImprovedXGModel

class AdvancedMetricsAnalyzer:
    def __init__(self, play_by_play_data: dict, shifts_data: list = None):
        self.plays = play_by_play_data.get('plays', [])
        self.roster_map = self._create_roster_map(play_by_play_data)
        self.xg_model = ImprovedXGModel()
        self.shifts_data = shifts_data
        self.on_ice_cache = {} # cache parsed shifts
        
        # Pre-process on-ice metrics if shifts are available
        self.player_on_ice_stats = defaultdict(lambda: {'CF': 0, 'CA': 0, 'xGF': 0.0, 'xGA': 0.0})
        if self.shifts_data:
            self._process_on_ice_stats()

    def _process_on_ice_stats(self):
        """Calculate on-ice metrics for all players using shift data."""
        # 1. Parse shifts into efficient lookup
        # lookup[period] = [(start_sec, end_sec, player_id, team_id), ...]
        shift_lookup = defaultdict(list)
        for shift in self.shifts_data:
            try:
                period = int(shift['period'])
                start = self._time_to_seconds(shift['startTime'])
                end = self._time_to_seconds(shift['endTime'])
                pid = shift['playerId']
                tid = shift['teamId']
                shift_lookup[period].append((start, end, pid, tid))
            except:
                continue
        
        # 2. Iterate plays
        for play in self.plays:
            event_type = play.get('typeDescKey', '')
            if event_type not in ['shot-on-goal', 'missed-shot', 'blocked-shot', 'goal']:
                continue
                
            details = play.get('details', {})
            event_team_id = details.get('eventOwnerTeamId')
            time_str = play.get('timeInPeriod', '00:00')
            play_time = self._time_to_seconds(time_str)
            period = play.get('periodDescriptor', {}).get('number', 0)
            
            # Calculate xG for this event (reuse or recalc)
            x_coord = details.get('xCoord', 0)
            y_coord = details.get('yCoord', 0)
            zone = details.get('zoneCode', '')
            shot_type = details.get('shotType', 'unknown')
            
            xg_val = 0.0
            if event_type in ['shot-on-goal', 'goal']:
                xg_val = self._calculate_single_shot_xG(x_coord, y_coord, zone, shot_type, event_type)
            elif event_type == 'missed-shot':
                 xg_val = self._calculate_single_shot_xG(x_coord, y_coord, zone, shot_type, event_type)
                 # Missed shots have some xG value in this model
            
            # Find players on ice
            shifts_in_period = shift_lookup.get(period, [])
            for start, end, pid, tid in shifts_in_period:
                if start <= play_time <= end:
                    # Player is on ice
                    if tid == event_team_id:
                        # For
                        self.player_on_ice_stats[pid]['CF'] += 1
                        self.player_on_ice_stats[pid]['xGF'] += xg_val
                    else:
                        # Against
                        self.player_on_ice_stats[pid]['CA'] += 1
                        self.player_on_ice_stats[pid]['xGA'] += xg_val

    def get_on_ice_metrics_for_player(self, player_id) -> dict:
        """Return cached on-ice metrics for specific player."""
        stats = self.player_on_ice_stats.get(player_id, {'CF': 0, 'CA': 0, 'xGF': 0.0, 'xGA': 0.0})
        return {
            'OnIce_Corsi_For': stats['CF'],
            'OnIce_Corsi_Against': stats['CA'],
            'OnIce_xG_For': round(stats['xGF'], 2),
            'OnIce_xG_Against': round(stats['xGA'], 2)
        }
    
    def _create_roster_map(self, play_by_play_data: dict) -> dict:
        """Create a mapping of player IDs to player info"""
        roster_map = {}
        if 'rosterSpots' in play_by_play_data:
            for player in play_by_play_data['rosterSpots']:
                player_id = player['playerId']
                roster_map[player_id] = {
                    'firstName': player['firstName']['default'],
                    'lastName': player['lastName']['default'],
                    'sweaterNumber': player['sweaterNumber'],
                    'positionCode': player['positionCode'],
                    'teamId': player['teamId']
                }
        return roster_map
    
    def get_available_metrics(self) -> dict:
        """Get all available metrics from the play-by-play data"""
        metrics = {
            'event_types': defaultdict(int),
            'spatial_data': set(),
            'player_actions': defaultdict(int),
            'zone_activities': defaultdict(int),
            'shot_types': set(),
            'penalty_types': set()
        }
        
        for play in self.plays:
            event_type = play.get('typeDescKey', '')
            details = play.get('details', {})
            
            # Count event types
            metrics['event_types'][event_type] += 1
            
            # Collect spatial data
            if 'xCoord' in details and 'yCoord' in details:
                metrics['spatial_data'].add('coordinates')
            if 'zoneCode' in details:
                metrics['zone_activities'][details['zoneCode']] += 1
                
            # Collect player actions
            for key in details.keys():
                if 'PlayerId' in key:
                    metrics['player_actions'][key] += 1
                    
            # Collect shot types
            if 'shotType' in details:
                metrics['shot_types'].add(details['shotType'])
                
            # Collect penalty types
            if 'descKey' in details and event_type == 'penalty':
                metrics['penalty_types'].add(details['descKey'])
        
        return metrics
    
    
    def calculate_shot_quality_metrics(self, team_id: int, player_id: int = None) -> dict:
        """Calculate advanced shot quality metrics"""
        shot_quality = {
            'total_shots': 0,
            'shots_on_goal': 0,
            'goals': 0,
            'missed_shots': 0,
            'blocked_shots': 0,
            'shot_types': defaultdict(int),
            'shot_locations': defaultdict(int),
            'high_danger_shots': 0,
            'shooting_percentage': 0,
            'expected_goals': 0
        }
        
        for play in self.plays:
            details = play.get('details', {})
            event_type = play.get('typeDescKey', '')
            event_team = details.get('eventOwnerTeamId')
            
            # Filter by Team
            if event_team != team_id:
                continue
            
            # Filter by Player if provided
            if player_id:
                # Check if this player was involved in the event
                player_involved = False
                if 'scoringPlayerId' in details and details['scoringPlayerId'] == player_id:
                    player_involved = True
                elif 'shootingPlayerId' in details and details['shootingPlayerId'] == player_id:
                    player_involved = True
                elif 'blockingPlayerId' in details and details['blockingPlayerId'] == player_id:
                    player_involved = True # For blocked shots (defensive metrics handled elsewhere, but keeping logic consistent)
                
                if not player_involved:
                    continue
                
            if event_type in ['shot-on-goal', 'missed-shot', 'blocked-shot']:
                shot_quality['total_shots'] += 1
                
                if event_type == 'shot-on-goal':
                    shot_quality['shots_on_goal'] += 1
                    
                    # Calculate xG for this shot to determine if it's high danger
                    x_coord = details.get('xCoord', 0)
                    y_coord = details.get('yCoord', 0)
                    zone = details.get('zoneCode', '')
                    shot_type = details.get('shotType', 'unknown')
                    
                    xG = self._calculate_single_shot_xG(x_coord, y_coord, zone, shot_type, event_type)
                    
                    # High danger shot: xG >= 0.15 (15% or better chance of scoring)
                    if xG >= 0.15:
                        shot_quality['high_danger_shots'] += 1
                    
                    # Shot type analysis
                    shot_quality['shot_types'][shot_type] += 1
                    
                    # Zone analysis
                    shot_quality['shot_locations'][zone] += 1
                    
                elif event_type == 'missed-shot':
                    shot_quality['missed_shots'] += 1
                elif event_type == 'blocked-shot':
                    # Only count blocked shots FOR errors if we are tracking team offense or the shooter
                    # If tracking defense, this would be a block FOR. 
                    # For shot quality metrics, a blocked shot is an attempt that was blocked.
                    shot_quality['blocked_shots'] += 1
                    
            # Track goals
            if event_type == 'goal':
                shot_quality['goals'] += 1
                shot_quality['shots_on_goal'] += 1  # Goals count as shots on goal
                shot_quality['high_danger_shots'] += 1  # Goals are always high danger
                
                # Shot type analysis
                shot_type = details.get('shotType', 'unknown')
                shot_quality['shot_types'][shot_type] += 1
                
                # Location analysis
                zone = details.get('zoneCode', '')
                shot_quality['shot_locations'][zone] += 1
        
        # Calculate shooting percentage (goals / shots on goal)
        if shot_quality['shots_on_goal'] > 0:
            shot_quality['shooting_percentage'] = shot_quality['goals'] / shot_quality['shots_on_goal']
        
        # Calculate expected goals using advanced model
        shot_quality['expected_goals'] = self._calculate_expected_goals(team_id, player_id)
        
        return shot_quality
    
    def _calculate_expected_goals(self, team_id: int, player_id: int = None) -> float:
        """Calculate total expected goals for a team (or player) using the improved model"""
        total_xg = 0.0
        
        for play in self.plays:
            details = play.get('details', {})
            event_team = details.get('eventOwnerTeamId')
            
            if event_team != team_id:
                continue
            
            # Filter by Player if provided
            if player_id:
                # Check if this player was the shooter/scorer
                player_involved = False
                if 'scoringPlayerId' in details and details['scoringPlayerId'] == player_id:
                    player_involved = True
                elif 'shootingPlayerId' in details and details['shootingPlayerId'] == player_id:
                    player_involved = True
                
                if not player_involved:
                    continue
                
            event_type = play.get('typeDescKey', '')
            if event_type in ['shot-on-goal', 'missed-shot', 'blocked-shot', 'goal']:
                x_coord = details.get('xCoord', 0)
                y_coord = details.get('yCoord', 0)
                zone = details.get('zoneCode', '')
                shot_type = details.get('shotType', 'unknown')
                
                # Build shot data dictionary expected by xG model
                shot_data = {
                    'x_coord': x_coord,
                    'y_coord': y_coord,
                    'shot_type': shot_type,
                    'event_type': event_type,
                    'zone': zone
                }
                
                # Calculate xG using improved model (passing empty previous_events for now)
                xg_value = self.xg_model.calculate_xg(shot_data, [])
                total_xg += xg_value
        
        return round(total_xg, 2)
    
    def _parse_strength_state(self, situation_code: str) -> str:
        """
        Parse NHL situation code to strength state
        Format: XXYY where XX = away skaters, YY = home skaters
        Examples: 1551 = 5v5, 1541 = 5v4 (PP), 1451 = 4v5 (PK)
        """
        try:
            if len(situation_code) >= 4:
                away_skaters = int(situation_code[2])
                home_skaters = int(situation_code[3])
                return f"{away_skaters}v{home_skaters}"
        except (ValueError, IndexError):
            pass
        return "5v5"  # Default
    
    def _get_current_score(self) -> Dict:
        """Get final game score to track score state throughout game"""
        score = {'away': 0, 'home': 0}
        for play in self.plays:
            if play.get('typeDescKey') == 'goal':
                details = play.get('details', {})
                scoring_team = details.get('eventOwnerTeamId')
                # We'll build running score in _get_score_differential
        return score
    
    def _get_score_differential(self, current_play: Dict, team_id: int, game_score: Dict) -> int:
        """
        Calculate score differential at time of shot from team's perspective
        Positive = leading, Negative = trailing, 0 = tied
        """
        # Build running score up to this play
        # Simplified: return 0 for tied (we can improve this with game data)
        return 0
    
    def _calculate_single_shot_xG(self, x_coord: float, y_coord: float, zone: str, shot_type: str, event_type: str) -> float:
        """Calculate expected goal value for a single shot based on NHL analytics model"""
        
        # Base expected goal value
        base_xG = 0.0
        
        # Distance calculation (from goal line at x=89)
        distance_from_goal = ((89 - x_coord) ** 2 + (y_coord) ** 2) ** 0.5
        
        # Angle calculation (angle from goal posts)
        # Goal posts are at y = ±3 (assuming 6-foot goal width)
        angle_to_goal = self._calculate_shot_angle(x_coord, y_coord)
        
        # Zone-based adjustments
        zone_multiplier = self._get_zone_multiplier(zone, x_coord, y_coord)
        
        # Shot type adjustments
        shot_type_multiplier = self._get_shot_type_multiplier(shot_type)
        
        # Event type adjustments (shots on goal vs missed/blocked)
        event_multiplier = self._get_event_type_multiplier(event_type)
        
        # Core distance-based model (NHL standard curve)
        if distance_from_goal <= 10:
            base_xG = 0.25  # Very close to net
        elif distance_from_goal <= 20:
            base_xG = 0.15  # Close range
        elif distance_from_goal <= 35:
            base_xG = 0.08  # Medium range
        elif distance_from_goal <= 50:
            base_xG = 0.04  # Long range
        else:
            base_xG = 0.02  # Very long range
        
        # Apply angle adjustment (shots from wider angles have lower xG)
        if angle_to_goal > 45:
            angle_multiplier = 0.3  # Very wide angle
        elif angle_to_goal > 30:
            angle_multiplier = 0.5  # Wide angle
        elif angle_to_goal > 15:
            angle_multiplier = 0.8  # Moderate angle
        else:
            angle_multiplier = 1.0  # Good angle
        
        # Calculate final expected goal value
        final_xG = base_xG * zone_multiplier * shot_type_multiplier * event_multiplier * angle_multiplier
        
        # Cap at reasonable maximum
        return min(final_xG, 0.95)
    
    def _calculate_shot_angle(self, x_coord: float, y_coord: float) -> float:
        """Calculate the angle of the shot relative to the goal"""
        import math
        
        # Goal center is at (89, 0), goal posts at (89, ±3)
        distance_to_center = ((89 - x_coord) ** 2 + (y_coord) ** 2) ** 0.5
        
        if distance_to_center == 0:
            return 0
        
        # Calculate angle using law of cosines
        # Distance from shot to left post
        dist_to_left = ((89 - x_coord) ** 2 + (y_coord - 3) ** 2) ** 0.5
        # Distance from shot to right post  
        dist_to_right = ((89 - x_coord) ** 2 + (y_coord + 3) ** 2) ** 0.5
        
        # Goal width
        goal_width = 6
        
        # Use law of cosines to find angle
        if dist_to_left > 0 and dist_to_right > 0:
            cos_angle = (dist_to_left ** 2 + dist_to_right ** 2 - goal_width ** 2) / (2 * dist_to_left * dist_to_right)
            cos_angle = max(-1, min(1, cos_angle))  # Clamp to valid range
            angle = math.acos(cos_angle)
            return math.degrees(angle)
        
        return 45  # Default angle if calculation fails
    
    def _get_zone_multiplier(self, zone: str, x_coord: float, y_coord: float) -> float:
        """Get zone-based expected goal multiplier"""
        
        # High danger area (slot, crease area)
        if zone == 'O' and x_coord > 75 and abs(y_coord) < 15:
            return 1.5
        
        # Medium danger area (offensive zone, good position)
        elif zone == 'O' and x_coord > 60 and abs(y_coord) < 25:
            return 1.2
        
        # Low danger area (point shots, wide angles)
        elif zone == 'O':
            return 0.8
        
        # Neutral zone shots (rare but possible)
        elif zone == 'N':
            return 0.3
        
        # Defensive zone shots (very rare)
        elif zone == 'D':
            return 0.1
        
        return 1.0  # Default
    
    def _get_shot_type_multiplier(self, shot_type: str) -> float:
        """Get shot type-based expected goal multiplier"""
        
        shot_type = shot_type.lower()
        
        # High-danger shot types
        if shot_type in ['tip-in', 'deflection', 'backhand']:
            return 1.3
        elif shot_type in ['wrist', 'snap']:
            return 1.0
        elif shot_type in ['slap', 'slapshot']:
            return 0.9
        elif shot_type in ['wrap-around', 'wrap']:
            return 1.1
        elif shot_type in ['one-timer', 'onetime']:
            return 1.2
        
        return 1.0  # Default for unknown types
    
    def _get_event_type_multiplier(self, event_type: str) -> float:
        """Get event type-based expected goal multiplier"""
        
        if event_type == 'shot-on-goal':
            return 1.0  # Full value for shots on goal
        elif event_type == 'missed-shot':
            return 0.7  # Reduced value for missed shots
        elif event_type == 'blocked-shot':
            return 0.5  # Lower value for blocked shots
        
        return 1.0  # Default
    
    def calculate_pressure_metrics(self, team_id: int, player_id: int = None) -> dict:
        """Calculate pressure metrics based on zone time and shot sequences"""
        pressure = {
            'offensive_zone_sequences': 0,
            'sustained_pressure_events': 0,
            'quick_strike_opportunities': 0
        }
        
        # If filtering by player, we check if player was involved in the culminating action of the pressure
        
        consecutive_oz_events = 0
        last_event_time = 0
        
        for play in self.plays:
            details = play.get('details', {})
            event_team = details.get('eventOwnerTeamId')
            zone = details.get('zoneCode', '')
            time_str = play.get('timeInPeriod', '00:00')
            current_time = self._time_to_seconds(time_str)
            
            if event_team != team_id:
                consecutive_oz_events = 0
                continue
                
            # Filter by Player if Provided - only credit pressure metrics if player is involved in the event
            player_involved = True
            if player_id:
                # Check broad involvement
                player_involved = False
                for k, v in details.items():
                    if 'PlayerId' in k and v == player_id:
                        player_involved = True
                        break
            
            if not player_involved:
                 # Even if player wasn't involved, we track the *sequence* context for when they ARE involved.
                 # But we might need to reset or continue logic. 
                 # For simplicity, if we filter by player, we only count events *by* that player.
                 pass

            if zone == 'O':
                # Check for sustained pressure (events within 5 seconds)
                if current_time - last_event_time <= 5:
                    if player_involved: # Only credit if player involved in this event
                         consecutive_oz_events += 1
                else:
                    consecutive_oz_events = 1 if player_involved else 0
                    
                if consecutive_oz_events >= 3:
                     if player_involved:
                        pressure['sustained_pressure_events'] += 1
                
                # Count distinct sequences
                if player_involved:
                    pressure['offensive_zone_sequences'] += 1
            elif zone == 'N':
                # Check for quick strike (Neutral zone to Offensive action)
                # This is hard to attribute to a single player without more complex logic tracking puck carriers.
                # Simplified: If player takes a shot/action in OZ shortly after NZ event.
                consecutive_oz_events = 0
                
            last_event_time = current_time
            
        return pressure
    
    def calculate_pre_shot_movement_metrics(self, team_id: int, player_id: int = None) -> dict:
        """Calculate pre-shot movement metrics"""
        metrics = {
            'royal_road_proxy': {'attempts': 0, 'goals': 0},
            'oz_retrieval_to_shot': {'attempts': 0, 'goals': 0},
            'lateral_movement': {'attempts': 0, 'goals': 0, 'total_delta_y': 0, 'avg_delta_y': 0},
            'longitudinal_movement': {'attempts': 0, 'goals': 0, 'total_delta_x': 0, 'avg_delta_x': 0}
        }
        
        for i, play in enumerate(self.plays):
            details = play.get('details', {})
            event_type = play.get('typeDescKey', '')
            event_team = details.get('eventOwnerTeamId')
            
            # Only analyze shots/goals for this team
            if event_team != team_id:
                continue
            
            # Filter by Player if provided
            if player_id:
                # Check if this player was the shooter/scorer
                player_involved = False
                if 'scoringPlayerId' in details and details['scoringPlayerId'] == player_id:
                    player_involved = True
                elif 'shootingPlayerId' in details and details['shootingPlayerId'] == player_id:
                    player_involved = True
                
                if not player_involved:
                    continue
            
            if event_type not in ['shot-on-goal', 'missed-shot', 'blocked-shot', 'goal']:
                continue
            
            x_coord = details.get('xCoord')
            y_coord = details.get('yCoord')
            time_in_period = play.get('timeInPeriod', '')
            
            if x_coord is None or y_coord is None:
                continue
            
            is_goal = (event_type == 'goal')
            
            # Convert time to seconds
            current_time = self._time_to_seconds(time_in_period)
            
            # Look back 4 seconds for Royal Road Proxy and lateral/longitudinal movement
            royal_road_detected = False
            lateral_delta_y = 0
            longitudinal_delta_x = 0
            prev_y = None
            prev_x = None
            
            for j in range(i - 1, max(-1, i - 20), -1):  # Look back up to 20 plays
                prev_play = self.plays[j]
                prev_details = prev_play.get('details', {})
                prev_team = prev_details.get('eventOwnerTeamId')
                prev_time_str = prev_play.get('timeInPeriod', '')
                prev_time = self._time_to_seconds(prev_time_str)
                
                # Only look at events from the same team within time window
                if prev_team != team_id:
                    continue
                
                time_diff = current_time - prev_time
                if time_diff > 4:  # Beyond 4 second window
                    break
                
                prev_x_coord = prev_details.get('xCoord')
                prev_y_coord = prev_details.get('yCoord')
                
                if prev_x_coord is None or prev_y_coord is None:
                    continue
                
                # Check for Royal Road Proxy (sign change in y)
                if not royal_road_detected and prev_y is not None:
                    if (prev_y * y_coord < 0) or (prev_y_coord * y_coord < 0):  # Sign change
                        royal_road_detected = True
                
                # Track lateral movement (y-axis changes)
                if prev_y_coord is not None:
                    lateral_delta_y = max(lateral_delta_y, abs(y_coord - prev_y_coord))
                
                # Track longitudinal movement (x-axis changes)
                if prev_x_coord is not None:
                    longitudinal_delta_x = max(longitudinal_delta_x, abs(x_coord - prev_x_coord))
                
                prev_y = prev_y_coord
                prev_x = prev_x_coord
            
            # Update Royal Road Proxy
            if royal_road_detected:
                metrics['royal_road_proxy']['attempts'] += 1
                if is_goal:
                    metrics['royal_road_proxy']['goals'] += 1
            
            # Update lateral movement
            if lateral_delta_y > 0:
                metrics['lateral_movement']['attempts'] += 1
                metrics['lateral_movement']['total_delta_y'] += lateral_delta_y
                if is_goal:
                    metrics['lateral_movement']['goals'] += 1
            
            # Update longitudinal movement
            if longitudinal_delta_x > 0:
                metrics['longitudinal_movement']['attempts'] += 1
                metrics['longitudinal_movement']['total_delta_x'] += longitudinal_delta_x
                if is_goal:
                    metrics['longitudinal_movement']['goals'] += 1
            
            # Check for OZ Retrieval to Shot (5 second window)
            oz_retrieval_detected = False
            for j in range(i - 1, max(-1, i - 25), -1):  # Look back up to 25 plays
                prev_play = self.plays[j]
                prev_details = prev_play.get('details', {})
                prev_team = prev_details.get('eventOwnerTeamId')
                prev_type = prev_play.get('typeDescKey', '')
                prev_zone = prev_details.get('zoneCode', '')
                prev_time_str = prev_play.get('timeInPeriod', '')
                prev_time = self._time_to_seconds(prev_time_str)
                
                if prev_team != team_id:
                    continue
                
                time_diff = current_time - prev_time
                if time_diff > 5:  # Beyond 5 second window
                    break
                
                # Check for OZ hit or takeaway
                if prev_zone == 'O' and prev_type in ['hit', 'takeaway']:
                    oz_retrieval_detected = True
                    break
            
            if oz_retrieval_detected:
                metrics['oz_retrieval_to_shot']['attempts'] += 1
                if is_goal:
                    metrics['oz_retrieval_to_shot']['goals'] += 1
        
        # Calculate averages
        if metrics['lateral_movement']['attempts'] > 0:
            metrics['lateral_movement']['avg_delta_y'] = (
                metrics['lateral_movement']['total_delta_y'] / metrics['lateral_movement']['attempts']
            )
        
        if metrics['longitudinal_movement']['attempts'] > 0:
            metrics['longitudinal_movement']['avg_delta_x'] = (
                metrics['longitudinal_movement']['total_delta_x'] / metrics['longitudinal_movement']['attempts']
            )
        
        return metrics
    
    def _time_to_seconds(self, time_str: str) -> float:
        """Convert MM:SS time string to seconds"""
        try:
            if ':' in time_str:
                parts = time_str.split(':')
                return int(parts[0]) * 60 + int(parts[1])
            return 0
        except (ValueError, IndexError):
            return 0

    def calculate_defensive_metrics(self, team_id: int, player_id: int = None) -> dict:
        """Calculate defensive metrics"""
        defense = {
            'blocked_shots': 0,
            'takeaways': 0,
            'hits': 0,
            'defensive_zone_clears': 0,
            'penalty_kill_efficiency': 0,
            'shot_attempts_against': 0,
            'high_danger_chances_against': 0,
            'defensive_players': defaultdict(int)
        }
        
        for play in self.plays:
            details = play.get('details', {})
            event_type = play.get('typeDescKey', '')
            event_team = details.get('eventOwnerTeamId')
            zone = details.get('zoneCode', '')
            
            # Count defensive actions
            if event_team == team_id:
                player_involved = True
                if player_id:
                    player_involved = False
                    if 'blockingPlayerId' in details and details['blockingPlayerId'] == player_id:
                        player_involved = True
                    elif 'hittingPlayerId' in details and details['hittingPlayerId'] == player_id:
                        player_involved = True
                    elif 'playerId' in details and details['playerId'] == player_id: # General fallback
                        player_involved = True

                if not player_involved:
                    continue

                if event_type == 'blocked-shot':
                    defense['blocked_shots'] += 1
                elif event_type == 'takeaway':
                    defense['takeaways'] += 1
                elif event_type == 'hit':
                    defense['hits'] += 1
                    
        return defense

    def calculate_transition_metrics(self, team_id: int, player_id: int = None) -> dict:
        """Calculate transition metrics: EXtoEN (Exit to Entry) and ENtoS (Entry to Shot)"""
        transitions = {
            'extoen_exits_to_entries': 0,
            'entos_entries_to_shots': 0
        }
        
        # Logic for EXtoEN: Successful Zone Exit followed by Zone Entry
        # Logic for ENtoS: Controlled Zone Entry followed by Shot within X seconds
        
        # Simplified:
        # EXtoEN: Any event in Neutral Zone (or after DZ exit) -> Event in Offensive Zone by same team
        # ENtoS: Entry (Neutral -> Offensive) -> Shot
        
        for i, play in enumerate(self.plays):
            details = play.get('details', {})
            event_team = details.get('eventOwnerTeamId')
            zone = details.get('zoneCode', '')
            
            if event_team != team_id:
                continue

            # Filtering transitions by player is tricky because it involves sequences of events.
            # We will adhere to: The *culminating event* (shot or entry) must be by the player.
            
            # ENtoS Logic
            # Detect Entry: Previous play was Neutral/Defensive, This play is Offensive
            if zone == 'O':
                prev_idx = max(0, i-1)
                prev_play = self.plays[prev_idx]
                prev_details = prev_play.get('details', {})
                prev_zone = prev_details.get('zoneCode', '')
                
                # Check for Entry (coming from N or D)
                if prev_zone in ['N', 'D'] and prev_details.get('eventOwnerTeamId') == team_id:
                    # Look ahead for shot
                    for j in range(i, min(i+5, len(self.plays))):
                        next_play = self.plays[j]
                        next_type = next_play.get('typeDescKey', '')
                        if next_play.get('details', {}).get('eventOwnerTeamId') == team_id:
                            if next_type in ['shot-on-goal', 'missed-shot', 'blocked-shot', 'goal']:
                                # If filtering, check if player took the shot
                                if player_id:
                                    det = next_play.get('details', {})
                                    if det.get('scoringPlayerId') == player_id or det.get('shootingPlayerId') == player_id:
                                        transitions['entos_entries_to_shots'] += 1
                                else:
                                    transitions['entos_entries_to_shots'] += 1
                                break
                                
            # EXtoEN Logic
            # Transition from D -> N -> O within short sequence
            if zone == 'N': # In neutral zone
                # Check if we were just in D
                prev_idx = max(0, i-1)
                prev_play = self.plays[prev_idx]
                prev_zone = prev_play.get('details', {}).get('zoneCode', '')
                
                if prev_zone == 'D' and prev_play.get('details', {}).get('eventOwnerTeamId') == team_id:
                    # Check if we enter O next
                     for j in range(i+1, min(i+4, len(self.plays))):
                        next_play = self.plays[j]
                        next_zone = next_play.get('details', {}).get('zoneCode', '')
                        next_type = next_play.get('typeDescKey', '')
                        if next_zone == 'O' and next_type in ['shot-on-goal', 'missed-shot', 'blocked-shot', 'goal', 'hit', 'takeaway', 'giveaway']: # Some event in OZ
                            # If filtering, check if player performed the entry/event
                            if player_id:
                                det = next_play.get('details', {})
                                # Check if player is involved in this event
                                if any(v == player_id for k,v in det.items() if 'PlayerId' in k):
                                    transitions['extoen_exits_to_entries'] += 1
                            else:
                                transitions['extoen_exits_to_entries'] += 1
                            break
                            
        return transitions

    def calculate_game_score(self, team_id: int, player_id: int = None) -> float:
        """
        Calculate Dom Luszczyszyn's Game Score for the team.
        Formula (Approx):
        GS = 0.75*G + 0.7*A1 + 0.55*A2 + 0.075*SOG + 0.05*BLK + 0.15*PD - 0.15*PT - 0.01*FACE_OFF_LOSS?
        Dom's model is player level. For Team level:
        GS = Goals + 0.75*Assists? Or just Sum of player GS?
        Let's use a simplified team version derived from the weights:
        Team GS = Goals*0.75 + Shots*0.075 + Blocks*0.05 + CF*0.05 - CA*0.05?
        Standard Game Score is a player metric, but user wants 'GameScore' column.
        We will sum up major events with simplified weights:
        Goals: 1.0
        Shots: 0.1
        Penalties Taken (Opponent PP): -0.5
        Penalties Drawn (Opponent PIM): 0.5
        """
        gs = 0.0
        
        sq = self.calculate_shot_quality_metrics(team_id)
        defense = self.calculate_defensive_metrics(team_id)
        
        goals = sq['goals']
        shots = sq['shots_on_goal']
        blocks = defense['blocked_shots']
        
        # Penalties
        pim_for = 0
        pim_against = 0
        for play in self.plays:
             if play.get('typeDescKey') == 'penalty':
                 desc = play.get('details', {})
                 if desc.get('eventOwnerTeamId') == team_id:
                     pim_for += desc.get('duration', 0)
                 else:
                     pim_against += desc.get('duration', 0)
                     
        # Formula
        gs += goals * 1.0
        gs += shots * 0.1
        gs += blocks * 0.05
        gs -= (pim_for / 2) * 0.2 # Penalties hurt
        gs += (pim_against / 2) * 0.2 # Drawing penalties helps
        
        return round(gs, 2)
