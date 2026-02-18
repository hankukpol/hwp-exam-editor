import unittest

from core.parser import ExamParser


class ParserTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.config = {
            "parsing": {
                "question_patterns": [
                    r"^\s*(\d{1,2})\s*[\.\)]\s*",
                    r"^\s*(0\d)\s*[\.\)]\s*",
                    r"^\s*【\s*(\d{1,2})\s*】\s*",
                    r"^\s*문\s*(\d{1,2})\s*[\.\)]\s*",
                ],
                "answer_patterns": [
                    r"^\s*정답\s+[①②③④⑤]",
                    r"^\s*정답[①②③④⑤]",
                    r"^\s*정답\s*[:：]\s*[①②③④⑤]",
                    r"^\s*정답\s*[:：]\s*\d",
                ],
                "explanation_patterns": [
                    r"^\s*해설\s*[\[【]?[×○xo]?[】\]]?",
                    r"^\s*해설\s*[:：]",
                    r"^\s*[①②③④⑤]\s*[:：]?\s*[\[\(]?[×○xo][\]\)]?",
                ],
                "type_a_threshold": 2,
            },
            "negative_keywords": ["옳지 않은", "틀린"],
        }
        self.parser = ExamParser(self.config)

    def test_parse_type_a(self) -> None:
        blocks = [
            "01. 다음 중 옳지 않은 것은?",
            "① 보기1",
            "② 보기2",
            "정답 ②",
            "해설 : 보기2가 정답이다.",
            "02. 다음 중 맞는 것은?",
            "① 보기1",
            "② 보기2",
            "정답 1",
            "해설 : 첫 번째가 정답이다.",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.file_type, "TYPE_A")
        self.assertEqual(document.total_count, 2)
        self.assertEqual(document.questions[0].answer, "②")
        self.assertTrue(document.questions[0].has_negative)
        self.assertEqual(document.questions[1].answer, "①")

    def test_parse_type_b(self) -> None:
        blocks = [
            "1. 문제 문장",
            "① 보기1",
            "② 보기2",
            "2. 다음 문제",
            "① 보기1",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="헌법")
        self.assertEqual(document.file_type, "TYPE_B")
        self.assertEqual(document.total_count, 2)
        self.assertIsNone(document.questions[0].answer)

    def test_parse_without_explicit_numbers(self) -> None:
        blocks = [
            "표현의 자유에 관한 다음 설명 중 가장 옳지 않은 것은? (다툼이 있는 경우 판례에 의함)",
            "① 보기1",
            "② 보기2",
            "③ 보기3",
            "④ 보기4",
            "④",
            "정답",
            "[×] 설명1",
            "해설",
            "집회의 자유에 대한 설명으로 옳은 것은?",
            "① 보기A",
            "② 보기B",
            "③ 보기C",
            "④ 보기D",
            "②",
            "정답",
            "[○] 설명2",
        ]
        self.parser.type_a_threshold = 2
        document = self.parser.parse_text_blocks(blocks, subject="헌법")
        self.assertEqual(document.file_type, "TYPE_A")
        self.assertEqual(document.total_count, 2)
        self.assertEqual(document.questions[0].answer, "④")
        self.assertEqual(document.questions[1].answer, "②")


if __name__ == "__main__":
    unittest.main()
