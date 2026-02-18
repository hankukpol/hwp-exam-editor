import unittest

from core.hwp_controller import HwpController


class HwpControllerLineCleanupTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.controller = HwpController()

    def test_clean_line_preserves_leading_hanja_subject(self) -> None:
        line = "甲은 늦은 밤 귀가하던 중 자신의 뒤편에서 다가오는 사람을 오인하였다."
        cleaned = self.controller._clean_line(line)
        self.assertEqual(cleaned, line)

    def test_clean_line_keeps_single_hanja_marker(self) -> None:
        self.assertEqual(self.controller._clean_line("甲"), "甲")


if __name__ == "__main__":
    unittest.main()
