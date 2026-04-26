import json
from pathlib import Path

CONFIG_DIR = Path.home() / ".auditgen"
CONFIG_FILE=CONFIG_DIR / "config.json"

DEFAULTS = {
    "api_key": None,
    "default_user": None,
    "output_dir" : None,
}

def _ensure_config_dir():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

def load_config() -> dict:
    _ensure_config_dir()
    if not CONFIG_FILE.exists():
        return dict(DEFAULTS)
    with open(CONFIG_FILE, "r") as f:
        return {**DEFAULTS,**json.load(f)}
    
# save the configuration
def save_config(data: dict):
    _ensure_config_dir()
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=2)

def get(key: str):
    return load_config().get(key)

def set_value(key:str , value:str):
    config = load_config
    config[key]=value
    save_config(config)