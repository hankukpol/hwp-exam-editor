import unittest
from pathlib import Path

import core.formatter as formatter_module
from core.formatter import HwpFormatter


def _make_formatter() -> HwpFormatter:
    config = {
        "format": {
            "question_font": "FontA",
            "passage_font": "FontB",
        },
        "style": {
            "enabled": True,
            "question_style": "Question",
            "passage_style": "Passage",
            "choice_style": "Passage",
            "sub_items_style": "Passage",
            "explanation_style": "Passage",
            "template_path": "",
            "style_map_source": "",
        },
    }
    return HwpFormatter(config)


class FormatterQuestionStyleClassificationTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.formatter = _make_formatter()

    def test_question_number_with_dot_is_question_style(self) -> None:
        self.assertEqual(self.formatter._classify_paragraph_style("1. question", 7, 9), 7)

    def test_question_number_with_parenthesis_is_question_style(self) -> None:
        self.assertEqual(self.formatter._classify_paragraph_style("12) question", 7, 9), 7)

    def test_question_prefix_mun_is_question_style(self) -> None:
        self.assertEqual(self.formatter._classify_paragraph_style("문 3. question", 7, 9), 7)

    def test_leading_zero_width_char_still_matches_question(self) -> None:
        text = "\ufeff4. question"
        self.assertEqual(self.formatter._classify_paragraph_style(text, 7, 9), 7)

    def test_non_question_line_is_passage_style(self) -> None:
        self.assertEqual(self.formatter._classify_paragraph_style("passage text", 7, 9), 9)

    def test_empty_line_is_base_style(self) -> None:
        self.assertEqual(self.formatter._classify_paragraph_style("   ", 7, 9), 0)

    def test_para_char_shape_run_count(self) -> None:
        # one run(8 bytes)
        self.assertEqual(self.formatter._para_char_shape_run_count(8), 1)
        # three runs(24 bytes)
        self.assertEqual(self.formatter._para_char_shape_run_count(24), 3)
        # invalid size
        self.assertEqual(self.formatter._para_char_shape_run_count(13), 0)

    def test_rewrite_para_char_shape_runs_single_run(self) -> None:
        # payload: one run (position=0, charshape_id=10)
        buf = bytearray(8)
        formatter_module.struct.pack_into("<I", buf, 0, 0)
        formatter_module.struct.pack_into("<I", buf, 4, 10)
        changed = self.formatter._rewrite_para_char_shape_runs(buf, 0, 8, 21)
        self.assertTrue(changed)
        self.assertEqual(formatter_module.struct.unpack_from("<I", buf, 4)[0], 21)

    def test_rewrite_para_char_shape_runs_mixed_preserves_emphasis_run(self) -> None:
        # runs: [100, 200, 100] -> rewrite base(100) to target(300), keep 200
        payload_size = 3 * 8
        buf = bytearray(payload_size)
        formatter_module.struct.pack_into("<I", buf, 0, 0)
        formatter_module.struct.pack_into("<I", buf, 4, 100)
        formatter_module.struct.pack_into("<I", buf, 8, 5)
        formatter_module.struct.pack_into("<I", buf, 12, 200)
        formatter_module.struct.pack_into("<I", buf, 16, 9)
        formatter_module.struct.pack_into("<I", buf, 20, 100)

        changed = self.formatter._rewrite_para_char_shape_runs(buf, 0, payload_size, 300)
        self.assertTrue(changed)
        self.assertEqual(formatter_module.struct.unpack_from("<I", buf, 4)[0], 300)
        self.assertEqual(formatter_module.struct.unpack_from("<I", buf, 12)[0], 200)
        self.assertEqual(formatter_module.struct.unpack_from("<I", buf, 20)[0], 300)

    def test_missing_question_style_falls_back_with_warning(self) -> None:
        if formatter_module.olefile is None:
            self.skipTest("olefile is unavailable in this environment")

        captured: dict[str, int] = {}

        def _fake_rewrite(_: Path, question_idx: int, passage_idx: int) -> bool:
            captured["question_idx"] = question_idx
            captured["passage_idx"] = passage_idx
            return True

        self.formatter.style_index_map = {"Passage": 5, "passage": 5}
        self.formatter.question_style = "QuestionMissing"
        self.formatter.passage_style = "Passage"
        self.formatter._rewrite_style_ids = _fake_rewrite  # type: ignore[method-assign]

        self.assertTrue(self.formatter.post_process_style_ids(Path("dummy.hwp")))
        self.assertEqual(captured.get("question_idx"), 5)
        self.assertEqual(captured.get("passage_idx"), 5)
        self.assertTrue(
            any("QuestionMissing" in warning for warning in self.formatter.style_runtime_warnings)
        )

    def test_apply_question_inline_char_emphasis_uses_current_face(self) -> None:
        captured: dict[str, object] = {"calls": []}

        def _fake_apply_char_shape(hwp, font_name: str, bold: bool, underline: bool) -> None:
            captured["calls"].append("char")
            captured["font_name"] = font_name
            captured["bold"] = bold
            captured["underline"] = underline

        self.formatter._apply_char_shape = _fake_apply_char_shape  # type: ignore[method-assign]
        self.formatter.apply_question_inline_char(object(), emphasize=True)

        self.assertEqual(captured["calls"], ["char"])
        self.assertEqual(captured["font_name"], self.formatter.question_font)
        self.assertEqual(captured["bold"], True)
        self.assertEqual(captured["underline"], True)

    def test_apply_question_inline_char_emphasis_can_use_bold_when_enabled(self) -> None:
        captured: dict[str, object] = {}

        def _fake_apply_char_shape(hwp, font_name: str, bold: bool, underline: bool) -> None:
            captured["font_name"] = font_name
            captured["bold"] = bold
            captured["underline"] = underline

        self.formatter.negative_emphasis_bold = True
        self.formatter._apply_char_shape = _fake_apply_char_shape  # type: ignore[method-assign]
        self.formatter.apply_question_inline_char(object(), emphasize=True)

        self.assertEqual(captured["font_name"], self.formatter.question_font)
        self.assertEqual(captured["bold"], True)
        self.assertEqual(captured["underline"], True)

    def test_rewrite_charshape_face_refs_in_docinfo_bytes(self) -> None:
        def _record(tag_id: int, payload: bytes) -> bytes:
            header = (len(payload) << 20) | tag_id
            return formatter_module.struct.pack("<I", header) + payload

        # Build DocInfo bytes with 3 CHAR_SHAPE records(tag 21), 14-byte face refs only.
        src = bytes([1] * 14)
        mid = bytes([2] * 14)
        tgt = bytes([3] * 14)
        blob = bytearray(
            _record(21, src)
            + _record(21, mid)
            + _record(21, tgt)
        )
        changed = self.formatter._rewrite_charshape_face_refs_in_docinfo_bytes(
            blob,
            source_char_id=0,
            target_char_ids={2},
        )
        self.assertTrue(changed)

        # Re-parse payloads
        payloads = []
        pos = 0
        while pos + 4 <= len(blob):
            h = formatter_module.struct.unpack_from("<I", blob, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            ds = pos + 4
            payloads.append((tag_id, bytes(blob[ds:ds + size])))
            pos = ds + size
        self.assertEqual(payloads[0][1], src)
        self.assertEqual(payloads[1][1], mid)
        self.assertEqual(payloads[2][1], src)

    def test_collect_question_charshape_mismatch_ids(self) -> None:
        def _record(tag_id: int, payload: bytes) -> bytes:
            header = (len(payload) << 20) | tag_id
            return formatter_module.struct.pack("<I", header) + payload

        def _para_header() -> bytes:
            return bytes(11)

        def _para_text(text: str) -> bytes:
            return text.encode("utf-16le")

        def _para_char_shape(runs: list[tuple[int, int]]) -> bytes:
            payload = bytearray(len(runs) * 8)
            for idx, (pos, cid) in enumerate(runs):
                base = idx * 8
                formatter_module.struct.pack_into("<I", payload, base, pos)
                formatter_module.struct.pack_into("<I", payload, base + 4, cid)
            return bytes(payload)

        body = (
            _record(66, _para_header())
            + _record(67, _para_text("1. question"))
            + _record(68, _para_char_shape([(0, 2), (10, 5), (12, 2)]))
            + _record(66, _para_header())
            + _record(67, _para_text("passage line"))
            + _record(68, _para_char_shape([(0, 1), (4, 6), (8, 1)]))
        )

        para_texts = self.formatter._collect_para_texts(body)
        mapping = self.formatter._collect_question_charshape_mismatch_ids(body, para_texts)
        self.assertEqual(mapping, {2: {5}})

    def test_post_process_style_ids_fallback_runs_question_emphasis_fix(self) -> None:
        if formatter_module.olefile is None:
            self.skipTest("olefile is unavailable in this environment")

        called: dict[str, Path] = {}

        def _fake_fallback(path: Path) -> bool:
            called["path"] = path
            return True

        self.formatter.style_index_map = {}
        self.formatter.question_style = "MissingQuestion"
        self.formatter.passage_style = "MissingPassage"
        self.formatter.post_process_question_emphasis_faces = _fake_fallback  # type: ignore[method-assign]

        target = Path("dummy.hwp")
        self.assertTrue(self.formatter.post_process_style_ids(target))
        self.assertEqual(called.get("path"), target)
        self.assertTrue(
            any("MissingQuestion" in warning for warning in self.formatter.style_runtime_warnings)
        )


if __name__ == "__main__":
    unittest.main()
