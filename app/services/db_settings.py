import json
import os
from pathlib import Path

APP_NAME = "CE_BOM_Tool"
DEFAULT_SETTINGS = {
    "database_path": "//gctfile.gigacomputing.intra/NR2B/NR2B6/共用資料區/BOM DATABASE",
    "pn_location": "C7",
}


def get_settings_path():
    appdata = os.getenv("APPDATA") or str(Path.home())
    config_dir = os.path.join(appdata, APP_NAME)
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "settings.json")


def load_settings():
    path = get_settings_path()
    if not os.path.exists(path):
        save_settings(DEFAULT_SETTINGS)  # 自動產生預設設定
        return DEFAULT_SETTINGS
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_settings(settings):
    path = get_settings_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)
