"""
State Management Module
Saves/loads run state between Wednesday and Friday reports.
Wednesday saves the selected filename and a snapshot of the pre-QC file;
Friday reads the snapshot for the QC audit comparison.
"""

import json
import os
import shutil
from datetime import datetime

STATE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "state")
STATE_FILE = os.path.join(STATE_DIR, "last_run.json")
SNAPSHOT_FILE = os.path.join(STATE_DIR, "wednesday_snapshot.xlsx")


def save_wednesday_state(filename, week, week_start, week_end, filepath=None):
    """Save the Wednesday file selection and a snapshot copy for Friday QC audit."""
    os.makedirs(STATE_DIR, exist_ok=True)

    # Save a snapshot of the pre-QC file so Friday can compare against the original
    snapshot_path = None
    if filepath and os.path.exists(filepath):
        try:
            shutil.copy2(filepath, SNAPSHOT_FILE)
            snapshot_path = SNAPSHOT_FILE
            print(f"  Snapshot saved: wednesday_snapshot.xlsx (for Friday QC audit)")
        except Exception as e:
            print(f"  WARNING: Could not save snapshot: {e}")

    state = {
        "wednesday_filename": filename,
        "wednesday_filepath": filepath,
        "wednesday_snapshot": snapshot_path,
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
