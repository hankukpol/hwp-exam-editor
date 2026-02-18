import json
import shutil
import unittest
from pathlib import Path
from uuid import uuid4

from core.config_manager import ConfigManager


class ConfigManagerPresetMergeTestCase(unittest.TestCase):
    def _write_json(self, path: Path, payload: dict) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)

    def test_load_with_preset_keeps_preset_template_and_user_module(self) -> None:
        base = Path(".tmp_cm_config_test") / f"cm_preset_{uuid4().hex}"
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
            cm.PRESETS_DIR = str(presets_dir)
            merged = cm.load_with_preset("아침모의고사.json")

            self.assertEqual(merged["format"]["question_font_size"], 13.0)
            self.assertEqual(merged["paragraph"]["line_spacing"], 115)
            self.assertEqual(merged["style"]["template_path"], "config/templates/아침모의고사 템플릿.hwp")
            self.assertEqual(merged["style"]["style_map_source"], "config/templates/아침모의고사 템플릿.hwp")
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


if __name__ == "__main__":
    unittest.main()
