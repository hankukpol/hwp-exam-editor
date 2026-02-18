from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ConfigManager:
    PRESETS_DIR = "config/presets"

    def __init__(
        self,
        default_path: str = "config/default_config.json",
        user_path: str = "config/user_config.json",
    ) -> None:
        self.default_path = Path(default_path)
        self.user_path = Path(user_path)
        self._active_preset_file: str = ""
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

    # ── 프리셋 관련 메서드 ──────────────────────────────────────

    def list_presets(self) -> list[dict[str, str]]:
        """사용 가능한 프리셋 목록을 반환한다."""
        presets_dir = Path(self.PRESETS_DIR)
        if not presets_dir.exists():
            return []
        result: list[dict[str, str]] = []
        for f in sorted(presets_dir.glob("*.json")):
            data = self._load_json_file(f)
            result.append({
                "name": data.get("preset_name", f.stem),
                "file": f.name,
                "description": data.get("preset_description", ""),
            })
        return result

    def load_with_preset(self, preset_filename: str) -> dict[str, Any]:
        """프리셋을 포함한 3단 병합으로 config를 로드한다."""
        defaults = self._load_json_file(self.default_path)
        preset_path = Path(self.PRESETS_DIR) / preset_filename
        preset = self._load_json_file(preset_path)
        user = self._load_json_file(self.user_path)

        # user_config에서 편집 서식 키를 제외하여 프리셋 값이 우선되도록 한다
        # - format/paragraph: 프리셋 우선
        # - style: 템플릿/스타일명은 프리셋 우선, 공통값(enabled/module_dll_path)만 사용자 유지
        user_filtered = dict(user)
        for key in ("format", "paragraph"):
            user_filtered.pop(key, None)
        style = user_filtered.get("style")
        if isinstance(style, dict):
            keep_style_keys = {"enabled", "module_dll_path"}
            filtered_style = {k: v for k, v in style.items() if k in keep_style_keys}
            if filtered_style:
                user_filtered["style"] = filtered_style
            else:
                user_filtered.pop("style", None)

        merged = self._deep_merge(defaults, preset)   # 1+2단계
        merged = self._deep_merge(merged, user_filtered)  # 3단계
        self._config = merged
        self._active_preset_file = preset_filename
        return self._config

    def get_active_preset(self) -> str:
        """user_config.json에 저장된 마지막 사용 프리셋 이름을 반환."""
        user = self._load_json_file(self.user_path)
        return str(user.get("active_preset", ""))

    def set_active_preset(self, preset_name: str) -> None:
        """마지막 사용 프리셋을 user_config.json에 저장."""
        self.update({"active_preset": preset_name})

    def get_active_preset_file(self) -> str:
        """현재 로드된 프리셋 파일명을 반환."""
        return self._active_preset_file
