from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any


class ConfigManager:
    PRESETS_DIR = "config/presets"
    APP_DATA_DIRNAME = "HWPExamEditor"

    def __init__(
        self,
        default_path: str = "config/default_config.json",
        user_path: str = "config/user_config.json",
    ) -> None:
        self._bundle_root = self._detect_bundle_root()
        self._runtime_root = self._detect_runtime_root()
        if self._is_frozen():
            self._bootstrap_runtime_files()
        self.default_path = self._resolve_runtime_path(default_path)
        self.user_path = self._resolve_runtime_path(user_path)
        self._active_preset_file: str = ""
        self._config = self._load()

    @staticmethod
    def _is_frozen() -> bool:
        return bool(getattr(sys, "frozen", False))

    @classmethod
    def _detect_bundle_root(cls) -> Path:
        if cls._is_frozen():
            meipass = getattr(sys, "_MEIPASS", "")
            if meipass:
                return Path(meipass)
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[1]

    @classmethod
    def _detect_runtime_root(cls) -> Path:
        if cls._is_frozen():
            local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
            if local_appdata:
                return Path(local_appdata) / cls.APP_DATA_DIRNAME
            return (Path.home() / "AppData" / "Local") / cls.APP_DATA_DIRNAME
        return Path(__file__).resolve().parents[1]

    def _resolve_runtime_path(self, raw_path: str | Path) -> Path:
        path = Path(raw_path).expanduser()
        if path.is_absolute():
            return path
        return self._runtime_root / path

    def get_runtime_root(self) -> Path:
        return self._runtime_root

    def get_presets_dir(self) -> Path:
        return self._resolve_runtime_path(self.PRESETS_DIR)

    def get_templates_dir(self) -> Path:
        return self._resolve_runtime_path("config/templates")

    def _bootstrap_runtime_files(self) -> None:
        src_config = self._bundle_root / "config"
        dst_config = self._runtime_root / "config"
        if not src_config.exists():
            return
        dst_config.mkdir(parents=True, exist_ok=True)
        self._copy_if_missing(src_config / "default_config.json", dst_config / "default_config.json")
        self._copy_tree_if_missing(src_config / "presets", dst_config / "presets")
        self._copy_tree_if_missing(src_config / "templates", dst_config / "templates")

    @staticmethod
    def _copy_if_missing(src: Path, dst: Path) -> None:
        if not src.exists() or dst.exists():
            return
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

    def _copy_tree_if_missing(self, src_dir: Path, dst_dir: Path) -> None:
        if not src_dir.exists():
            return
        for src_path in src_dir.rglob("*"):
            rel = src_path.relative_to(src_dir)
            dst_path = dst_dir / rel
            if src_path.is_dir():
                dst_path.mkdir(parents=True, exist_ok=True)
                continue
            self._copy_if_missing(src_path, dst_path)

    def _normalize_path_text(self, value: str) -> str:
        text = str(value or "").strip()
        if not text:
            return ""
        path = Path(text).expanduser()
        if path.is_absolute():
            return str(path)
        runtime_candidate = self._runtime_root / path
        if runtime_candidate.exists() or self._is_frozen():
            return str(runtime_candidate)
        bundle_candidate = self._bundle_root / path
        if bundle_candidate.exists():
            return str(bundle_candidate)
        return str((Path.cwd() / path).resolve())

    def _normalize_style_paths(self, config: dict[str, Any]) -> dict[str, Any]:
        style = config.get("style")
        if not isinstance(style, dict):
            return config
        template_path = style.get("template_path")
        if isinstance(template_path, str) and template_path.strip():
            style["template_path"] = self._normalize_path_text(template_path)
        style_map_source = style.get("style_map_source")
        if isinstance(style_map_source, str) and style_map_source.strip():
            style["style_map_source"] = self._normalize_path_text(style_map_source)
        return config

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
        merged = self._deep_merge(defaults, user)
        return self._normalize_style_paths(merged)

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
        self._config = self._normalize_style_paths(self._deep_merge(self._config, partial))
        existing_user = self._load_json_file(self.user_path)
        merged_user = self._deep_merge(existing_user, partial)
        self.user_path.parent.mkdir(parents=True, exist_ok=True)
        with self.user_path.open("w", encoding="utf-8") as file:
            json.dump(merged_user, file, ensure_ascii=False, indent=2)
        return self._config

    # ── 프리셋 관련 메서드 ──────────────────────────────────────

    def list_presets(self) -> list[dict[str, str]]:
        """사용 가능한 프리셋 목록을 반환한다."""
        presets_dir = self.get_presets_dir()
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
        preset_path = self.get_presets_dir() / preset_filename
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
        self._config = self._normalize_style_paths(merged)
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
