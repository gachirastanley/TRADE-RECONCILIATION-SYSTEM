import os
import json

HISTORY_FILE = "history.json"
HISTORY_FOLDER = "history_files"

os.makedirs(HISTORY_FOLDER, exist_ok=True)

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    return []

def save_history(data):
    with open(HISTORY_FILE, "w") as f:
        json.dump(data, f, indent=4)
