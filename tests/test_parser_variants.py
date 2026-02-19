from pathlib import Path
import unittest

from core.config_manager import ConfigManager
from core.service import ExamProcessingService


class ParserVariantIntegrationTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.service = ExamProcessingService(ConfigManager())
        self.sample_dir = Path(__file__).parent / "sample_files"

    def test_type_a_variant_file(self) -> None:
        sample = self.sample_dir / "형법_변형A.txt"
        self.service.parser.type_a_threshold = 2
        document = self.service.parse_file(str(sample))

        self.assertEqual(document.file_type, "TYPE_A")
        self.assertEqual(document.total_count, 2)
        self.assertEqual(document.questions[0].answer, "③")
        self.assertEqual(document.questions[1].answer, "②")
        self.assertTrue(document.questions[0].has_negative)
        self.assertIsNotNone(document.questions[0].explanation)

    def test_type_b_variant_file(self) -> None:
        sample = self.sample_dir / "형사소송법_변형B.txt"
        document = self.service.parse_file(str(sample))

        self.assertEqual(document.file_type, "TYPE_B")
        self.assertEqual(document.total_count, 2)
        self.assertIsNone(document.questions[0].answer)
        self.assertTrue(document.questions[0].has_negative)

    def test_default_threshold_keeps_small_type_a_as_type_a(self) -> None:
        sample = self.sample_dir / "형법_변형A.txt"
        document = self.service.parse_file(str(sample))
        self.assertEqual(document.file_type, "TYPE_A")


if __name__ == "__main__":
    unittest.main()
