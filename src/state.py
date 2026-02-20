"""
State Management Module
Saves/loads run state between Wednesday and Friday reports.
Wednesday saves the selected filename; Friday reads it to find the QC'd version.
"""

import json
import os
from datetime import datetime

STATE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "state")
STATE_FILE = os.path.join(STATE_DIR, "last_run.json")


def save_wednesday_state(filename, week, week_start, week_end, filepath=None):
    """Save the Wednesday file selection so Friday can find the QC'd version."""
    os.makedirs(STATE_DIR, exist_ok=True)
    state = {
        "wednesday_filename": filename,
        "wednesday_filepath": filepath,
        "week": week,
        "week_start": str(week_start.date()),
        "week_end": str(week_end.date()),
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)
    print(f"  State saved: {filename} (for Friday reuse)")


def load_wednesday_state():
    """Load the Wednesday state to find the same file for Friday."""
    if not os.path.exists(STATE_FILE):
        return None
    with open(STATE_FILE, "r") as f:
        return json.load(f)
