from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ConfigManager:
    def __init__(
        self,
        default_path: str = "config/default_config.json",
        user_path: str = "config/user_config.json",
    ) -> None:
        self.default_path = Path(default_path)
        self.user_path = Path(user_path)
        self._config = self._load()

    def _load_json_file(self, path: Path) -> dict[str, Any]:
        if not path.exists():
            return {}
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)

    def _deep_merge(self, base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
        merged = dict(base)
        for key, value in override.items():
            if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
                merged[key] = self._deep_merge(merged[key], value)
            else:
                merged[key] = value
        return merged

    def _load(self) -> dict[str, Any]:
        defaults = self._load_json_file(self.default_path)
        user = self._load_json_file(self.user_path)
        return self._deep_merge(defaults, user)

    def reload(self) -> dict[str, Any]:
        self._config = self._load()
        return self._config

    def all(self) -> dict[str, Any]:
        return self._config

    def get(self, path: str, default: Any = None) -> Any:
        current: Any = self._config
        for token in path.split("."):
            if not isinstance(current, dict) or token not in current:
                return default
            current = current[token]
        return current

    def update(self, partial: dict[str, Any]) -> dict[str, Any]:
        self._config = self._deep_merge(self._config, partial)
        existing_user = self._load_json_file(self.user_path)
        merged_user = self._deep_merge(existing_user, partial)
        self.user_path.parent.mkdir(parents=True, exist_ok=True)
        with self.user_path.open("w", encoding="utf-8") as file:
            json.dump(merged_user, file, ensure_ascii=False, indent=2)
        return self._config
