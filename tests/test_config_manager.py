import json
import shutil
import unittest
from pathlib import Path
from uuid import uuid4

from core.config_manager import ConfigManager


class ConfigManagerPresetMergeTestCase(unittest.TestCase):
    def _make_base_dir(self, prefix: str) -> Path:
        base = Path(".tmp_cm_runtime_local") / f"{prefix}_{uuid4().hex}"
        base.mkdir(parents=True, exist_ok=True)
        return base

    def _write_json(self, path: Path, payload: dict) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)

    def test_load_with_preset_keeps_preset_template_and_user_module(self) -> None:
        base = self._make_base_dir("cm_preset")
        base.mkdir(parents=True, exist_ok=True)
        try:
            defaults_path = base / "default.json"
            user_path = base / "user.json"
            presets_dir = base / "presets"
            preset_file = presets_dir / "아침모의고사.json"

            self._write_json(defaults_path, {
                "format": {"question_font_size": 9.5},
                "paragraph": {"line_spacing": 140},
                "style": {
                    "enabled": True,
                    "template_path": "default.hwp",
                    "style_map_source": "default.hwp",
                    "question_style": "DefaultQ",
                    "passage_style": "DefaultP",
                    "module_dll_path": "",
                },
            })
            self._write_json(preset_file, {
                "preset_name": "아침모의고사",
                "format": {"question_font_size": 13.0},
                "paragraph": {"line_spacing": 115},
                "style": {
                    "template_path": "config/templates/아침모의고사 템플릿.hwp",
                    "style_map_source": "config/templates/아침모의고사 템플릿.hwp",
                    "question_style": "문제",
                    "passage_style": "지문",
                },
            })
            self._write_json(user_path, {
                "format": {"question_font_size": 99.0},
                "paragraph": {"line_spacing": 999},
                "style": {
                    "enabled": False,
                    "template_path": "config/templates/정기모의고사 템플릿.hwp",
                    "style_map_source": "config/templates/정기모의고사 템플릿.hwp",
                    "question_style": "사용자문제",
                    "passage_style": "사용자지문",
                    "module_dll_path": "C:/Program Files (x86)/Hnc/HOffice9/Bin/FilePathCheckerModuleExample.dll",
                },
            })

            cm = ConfigManager(default_path=str(defaults_path), user_path=str(user_path))
            cm.PRESETS_DIR = str(presets_dir.resolve())
            merged = cm.load_with_preset("아침모의고사.json")
            expected_template = str(cm.get_runtime_root() / "config/templates/아침모의고사 템플릿.hwp")

            self.assertEqual(merged["format"]["question_font_size"], 13.0)
            self.assertEqual(merged["paragraph"]["line_spacing"], 115)
            self.assertEqual(merged["style"]["template_path"], expected_template)
            self.assertEqual(merged["style"]["style_map_source"], expected_template)
            self.assertEqual(merged["style"]["question_style"], "문제")
            self.assertEqual(merged["style"]["passage_style"], "지문")
            self.assertEqual(
                merged["style"]["module_dll_path"],
                "C:/Program Files (x86)/Hnc/HOffice9/Bin/FilePathCheckerModuleExample.dll",
            )
            self.assertEqual(merged["style"]["enabled"], False)
            self.assertEqual(cm.get_active_preset_file(), "아침모의고사.json")
        finally:
            shutil.rmtree(base, ignore_errors=True)

    def test_frozen_mode_bootstraps_bundle_config_into_runtime(self) -> None:
        base = self._make_base_dir("cm_frozen")
        bundle_root = base / "bundle"
        runtime_root = base / "runtime"
        try:
            default_cfg = bundle_root / "config" / "default_config.json"
            preset_cfg = bundle_root / "config" / "presets" / "아침모의고사.json"
            template_hwp = bundle_root / "config" / "templates" / "아침모의고사 템플릿.hwp"

            self._write_json(default_cfg, {
                "paths": {"output_directory": ""},
                "style": {"template_path": "config/templates/아침모의고사 템플릿.hwp"},
            })
            self._write_json(preset_cfg, {
                "preset_name": "아침모의고사",
                "style": {"template_path": "config/templates/아침모의고사 템플릿.hwp"},
            })
            template_hwp.parent.mkdir(parents=True, exist_ok=True)
            template_hwp.write_bytes(b"fake_hwp_binary")

            class _FrozenConfigManager(ConfigManager):
                @staticmethod
                def _is_frozen() -> bool:
                    return True

                @classmethod
                def _detect_bundle_root(cls) -> Path:
                    return bundle_root

                @classmethod
                def _detect_runtime_root(cls) -> Path:
                    return runtime_root

            cm = _FrozenConfigManager()

            self.assertTrue((runtime_root / "config" / "default_config.json").exists())
            self.assertTrue((runtime_root / "config" / "presets" / "아침모의고사.json").exists())
            self.assertTrue((runtime_root / "config" / "templates" / "아침모의고사 템플릿.hwp").exists())
            self.assertEqual(
                cm.default_path,
                runtime_root / "config" / "default_config.json",
            )
            self.assertEqual(
                cm.get_presets_dir(),
                runtime_root / "config" / "presets",
            )
            self.assertEqual(
                cm.get_templates_dir(),
                runtime_root / "config" / "templates",
            )
        finally:
            shutil.rmtree(base, ignore_errors=True)

    def test_load_with_preset_recovers_stale_absolute_template_path(self) -> None:
        base = self._make_base_dir("cm_stale_preset")
        try:
            defaults_path = base / "default.json"
            user_path = base / "user.json"
            presets_dir = base / "presets"
            preset_file = presets_dir / "아침모의고사.json"
            templates_dir = base / "config" / "templates"
            template_path = templates_dir / "아침모의고사 템플릿.hwp"
            template_path.parent.mkdir(parents=True, exist_ok=True)
            template_path.write_bytes(b"fake_hwp_binary")

            stale_absolute = "D:/앱 프로그램/아침모의고사 자동편집 프로그램/config/templates/아침모의고사 템플릿.hwp"

            self._write_json(defaults_path, {
                "style": {"enabled": True},
            })
            self._write_json(preset_file, {
                "preset_name": "아침모의고사",
                "style": {
                    "template_path": stale_absolute,
                    "style_map_source": stale_absolute,
                    "question_style": "문제",
                    "passage_style": "지문",
                },
            })
            self._write_json(user_path, {
                "style": {"enabled": True},
            })

            class _PortableConfigManager(ConfigManager):
                @staticmethod
                def _is_frozen() -> bool:
                    return False

                @classmethod
                def _detect_bundle_root(cls) -> Path:
                    return base

                @classmethod
                def _detect_runtime_root(cls) -> Path:
                    return base

            cm = _PortableConfigManager(
                default_path=str(defaults_path.resolve()),
                user_path=str(user_path.resolve()),
            )
            cm.PRESETS_DIR = str(presets_dir.resolve())
            merged = cm.load_with_preset("아침모의고사.json")

            expected = str(template_path)
            self.assertEqual(merged["style"]["template_path"], expected)
            self.assertEqual(merged["style"]["style_map_source"], expected)
        finally:
            shutil.rmtree(base, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
