import json
import os
from pathlib import Path
from typing import Dict, Any

CONFIG_FILE = Path.home() / ".sql_runner_config.json"

DEFAULT_CONFIG: Dict[str, Any] = {
    "server": "localhost",
    "database": "master",
    "username": "",
    "password": "",
    "auth_type": "SQL Server Authentication",
    "driver": "ODBC Driver 17 for SQL Server",
    "last_script_path": "",
    "auto_connect": True
}

def load_config() -> Dict[str, Any]:
    """Load configuration from JSON file or return defaults if not existing."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                config = DEFAULT_CONFIG.copy()
                config.update(data)
                return config
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()

def save_config(config_data: Dict[str, Any]) -> bool:
    """Save configuration to JSON file."""
    try:
        current = load_config()
        current.update(config_data)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False
