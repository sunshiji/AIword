import json
import os
from typing import Any, Dict, Optional


class ConfigManager:
    DEFAULT_CONFIG = {
        "version": "1.0.0",
        "hotkey": "ctrl+alt+v",
        "hotkey_enabled": True,
        "notification_enabled": True,
        "no_app_action": "copy_to_clipboard",
        "move_cursor_to_end": True,
        "html_formatting": True,
        "keep_generated_files": False,
        "save_dir": "./output",
        "log_dir": "./logs"
    }
    
    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                merged_config = self.DEFAULT_CONFIG.copy()
                merged_config.update(config)
                return merged_config
            except (json.JSONDecodeError, IOError):
                return self.DEFAULT_CONFIG.copy()
        else:
            self._save_config(self.DEFAULT_CONFIG)
            return self.DEFAULT_CONFIG.copy()
    
    def _save_config(self, config: Dict[str, Any]) -> None:
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except IOError:
            pass
    
    def get(self, key: str, default: Any = None) -> Any:
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        self.config[key] = value
        self._save_config(self.config)
    
    def reload(self) -> None:
        self.config = self._load_config()
    
    def get_all(self) -> Dict[str, Any]:
        return self.config.copy()
