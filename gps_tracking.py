"""
GPS Tracking for active clock entries
Logs GPS positions every 10-15 minutes while user is clocked in
"""

class GPSLog:
    """Model for GPS tracking logs"""
    def __init__(self):
        self.id = None
        self.entry_id = None
        self.user_id = None
        self.timestamp = None
        self.location = {"lat": 0, "lon": 0}
        self.distance_to_project = None
        self.within_radius = None
