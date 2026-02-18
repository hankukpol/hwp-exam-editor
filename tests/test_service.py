import unittest

from core.config_manager import ConfigManager
from core.service import ExamProcessingService


class ServiceDiagnosticTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.service = ExamProcessingService(ConfigManager())

    def test_build_parse_diagnostic_message_contains_samples_and_count(self) -> None:
        blocks = [
            "머리말",
            "다음 설명 중 옳은 것은?",
            "① 보기",
            "② 보기",
            "결론",
        ]
        message = self.service._build_parse_diagnostic_message(blocks)
        self.assertIn("문제 번호를 찾지 못했습니다", message)
        self.assertIn("문제번호 패턴 일치", message)
        self.assertIn("파일 앞부분 샘플", message)
        self.assertIn("확인사항", message)


if __name__ == "__main__":
    unittest.main()

