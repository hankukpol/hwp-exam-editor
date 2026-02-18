from pathlib import Path
import unittest

from core.config_manager import ConfigManager
from core.exceptions import ParseError
from core.service import ExamProcessingService


class RealSampleParsingTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.service = ExamProcessingService(ConfigManager())
        cls.base = Path("tests/sample_files")

    def _sample(self, name: str) -> str:
        path = self.base / name
        if not path.exists():
            self.skipTest(f"sample file not found: {path}")
        return str(path)

    def test_parse_police_sample_type_a(self) -> None:
        path = self._sample("아침모의고사 6회.hwp")
        doc = self.service.parse_file(path)
        self.assertEqual(doc.file_type, "TYPE_A")
        self.assertEqual(doc.total_count, 20)
        self.assertGreaterEqual(sum(1 for q in doc.questions if q.answer), 20)

    def test_parse_criminal_law_sample_type_a(self) -> None:
        path = self._sample("6주차 - 특수폭행 ~ 경매(위전착).hwp")
        doc = self.service.parse_file(path)
        self.assertEqual(doc.file_type, "TYPE_A")
        self.assertEqual(doc.total_count, 20)
        self.assertGreaterEqual(sum(1 for q in doc.questions if q.answer), 20)

    def test_parse_criminal_law_q19_boxed_case_is_preserved(self) -> None:
        path = self._sample("6주차 - 특수폭행 ~ 경매(위전착).hwp")
        doc = self.service.parse_file(path)
        q19 = next((q for q in doc.questions if q.number == 19), None)
        self.assertIsNotNone(q19)
        assert q19 is not None
        self.assertEqual(q19.question_text, "다음 사례에 대한 설명으로 가장 적절한 것은?")
        self.assertTrue(q19.has_table)
        self.assertEqual(len(q19.sub_items), 1)
        self.assertTrue(q19.sub_items[0].startswith("甲은 "))

    def test_parse_constitution_sample_without_numbers(self) -> None:
        path = self._sample("[헌법] [6주차]  26년 1-2월 헌법 아침모의고사_2.9.(월)_집회결사~재산권.hwp")
        doc = self.service.parse_file(path)
        self.assertEqual(doc.file_type, "TYPE_A")
        self.assertEqual(doc.total_count, 20)
        self.assertGreaterEqual(sum(1 for q in doc.questions if q.answer), 20)

    def test_parse_question_only_sample_type_b(self) -> None:
        path = self._sample("제6회 문제(대물영장주의예외~거증책임, 실제는 수사까지) - 26년1월.hwp")
        doc = self.service.parse_file(path)
        self.assertEqual(doc.file_type, "TYPE_B")
        self.assertEqual(doc.total_count, 20)
        self.assertEqual(sum(1 for q in doc.questions if q.answer), 0)

    def test_answer_key_only_file_reports_clear_error(self) -> None:
        path = self._sample("제6회 정답(대물영장주의예외~거증책임, 실제는 수사까지) - 26년1월.hwp")
        with self.assertRaises(ParseError):
            self.service.parse_file(path)


if __name__ == "__main__":
    unittest.main()
