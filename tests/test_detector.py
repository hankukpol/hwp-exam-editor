import unittest

from core.detector import detect_file_type, detect_negative_keyword, extract_question_number


class DetectorTestCase(unittest.TestCase):
    def test_detect_file_type_type_a(self) -> None:
        blocks = [
            "01. 문제 본문",
            "정답 ①",
            "해설 : 설명",
            "02. 문제 본문",
            "정답 ②",
            "해설 : 설명",
            "03. 문제 본문",
            "정답 ③",
            "해설 : 설명",
            "04. 문제 본문",
            "정답 ④",
            "해설 : 설명",
            "05. 문제 본문",
            "정답 ⑤",
        ]
        self.assertEqual(detect_file_type(blocks), "TYPE_A")

    def test_detect_file_type_type_b(self) -> None:
        blocks = [
            "01. 문제 본문",
            "① 보기",
            "② 보기",
            "03. 문제 본문",
        ]
        self.assertEqual(detect_file_type(blocks), "TYPE_B")

    def test_detect_file_type_type_a_for_small_set_with_answers(self) -> None:
        blocks = [
            "30. 음주운전 관련 설명으로 옳지 않은 것은?",
            "① 보기A",
            "② 보기B",
            "정답 : ②",
            "해설 : 설명A",
            "31. 다음 설명으로 옳은 것은?",
            "① 보기C",
            "② 보기D",
            "정답 : ①",
            "참고 : 설명B",
        ]
        self.assertEqual(detect_file_type(blocks, threshold=5), "TYPE_A")

    def test_extract_question_number(self) -> None:
        self.assertEqual(extract_question_number("문 12. 테스트"), 12)
        self.assertEqual(extract_question_number("08) 테스트"), 8)
        self.assertEqual(extract_question_number("★02. 테스트"), 2)
        self.assertEqual(extract_question_number("60 아래 설명으로 옳은 것은?"), 60)

    def test_detect_file_type_with_bracketed_answer_marker(self) -> None:
        blocks = [
            "01. 문제 본문",
            "[정답] ①",
            "02. 문제 본문",
            "[정답] ②",
            "03. 문제 본문",
            "[정답] ③",
            "04. 문제 본문",
            "[정답] ④",
            "05. 문제 본문",
            "[정답] ⑤",
        ]
        self.assertEqual(detect_file_type(blocks), "TYPE_A")

    def test_detect_negative_keyword(self) -> None:
        self.assertEqual(detect_negative_keyword("다음 중 옳지 않은 것은?"), "않은")
        self.assertEqual(detect_negative_keyword("다음 중 옳지않은 것은?"), "않은")
        self.assertEqual(detect_negative_keyword("다음 중 틀린 것은?"), "틀린")
        self.assertEqual(detect_negative_keyword("다음 중 올바르지 아니한 것은?"), "아니한")
        self.assertEqual(detect_negative_keyword("다음 중 잘못된 것은?"), "잘못된")
        self.assertEqual(detect_negative_keyword("다음 중 부적절한 것은?"), "부적절한")
        self.assertEqual(detect_negative_keyword("다음 중 부적절 한 것은?"), "부적절 한")
        self.assertEqual(detect_negative_keyword("다음 중 아닌 것은?"), "아닌")
        self.assertEqual(detect_negative_keyword("다음 중 않은 것은?"), "않은")


if __name__ == "__main__":
    unittest.main()
