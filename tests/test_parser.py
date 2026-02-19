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

    def test_hanja_case_line_is_promoted_to_boxed_sub_item(self) -> None:
        blocks = [
            "19. 다음 사례에 대한 설명으로 가장 적절한 것은?",
            "氠瑢",
            "甲은 늦은 밤 귀가하던 중 자신의 뒤편에서 다가오는 사람을 A로 오인하였다.",
            "① 보기1",
            "② 보기2",
            "③ 보기3",
            "④ 보기4",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertEqual(question.question_text, "다음 사례에 대한 설명으로 가장 적절한 것은?")
        self.assertEqual(
            question.sub_items,
            ["甲은 늦은 밤 귀가하던 중 자신의 뒤편에서 다가오는 사람을 A로 오인하였다."],
        )
        self.assertTrue(question.has_table)

    def test_parse_star_prefixed_question_and_bracketed_answer(self) -> None:
        self.parser.type_a_threshold = 1
        blocks = [
            "★02. 「도로교통법」상 용어의 정의이다. 올바른 것은?",
            "① 보기1",
            "② 보기2",
            "[정답] (②)",
            "[해설] ② (○) 설명",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="도로교통법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertEqual(question.number, 2)
        self.assertEqual(question.question_text, "「도로교통법」상 용어의 정의이다. 올바른 것은?")
        self.assertEqual(question.answer, "②")
        self.assertEqual(question.answer_line, "[정답] (②)")

    def test_marked_case_and_view_blocks_are_preserved_as_table_lines(self) -> None:
        blocks = [
            "11. 다음 사례를 읽고 아래 <보기>에서 적절한 설명을 모두 고른 것은?",
            "<사례>",
            "甲은 을을 도우려다 폭행했다.",
            "<보기>",
            "㉠ 첫 번째 설명",
            "㉡ 두 번째 설명",
            "① ㉠",
            "② ㉡",
            "정답②",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertTrue(question.has_table)
        self.assertEqual(question.question_text, "다음 사례를 읽고 아래 <보기>에서 적절한 설명을 모두 고른 것은?")
        self.assertIn("<사례>", question.sub_items)
        self.assertIn("<보기>", question.sub_items)
        self.assertIn("甲은 을을 도우려다 폭행했다.", question.sub_items)
        self.assertIn("", question.sub_items)

    def test_parenthesized_rows_and_view_number_are_preserved_as_two_table_blocks(self) -> None:
        blocks = [
            "60. 아래 (가)와 (나)에 관련된 설명으로 적절하지 않은 것을 <보기1>에서 모두 고른 것은?",
            "(가) 공범이 성립하려면 정범의 실행행위가 있어야 하나, 공범 성립요건에 정범의 행위가 포함되지는 않는다.",
            "(나) 공범행위는 그 자체가 반사회적인 범죄실행행위로서의 실질을 가지므로 공범은 정범의 실행행위와 관계없이 독립하여 성립한다.",
            "<보기1>",
            "㉠ 첫 번째 설명",
            "㉡ 두 번째 설명",
            "㉢ 세 번째 설명",
            "㉣ 네 번째 설명",
            "① ㉠㉡",
            "② ㉠㉣",
            "정답②",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertTrue(question.has_table)
        self.assertEqual(
            question.sub_items,
            [
                "(가) 공범이 성립하려면 정범의 실행행위가 있어야 하나, 공범 성립요건에 정범의 행위가 포함되지는 않는다.",
                "(나) 공범행위는 그 자체가 반사회적인 범죄실행행위로서의 실질을 가지므로 공범은 정범의 실행행위와 관계없이 독립하여 성립한다.",
                "",
                "<보기1>",
                "㉠ 첫 번째 설명",
                "㉡ 두 번째 설명",
                "㉢ 세 번째 설명",
                "㉣ 네 번째 설명",
            ],
        )
        self.assertEqual(question.answer, "②")

    def test_matrix_choice_table_lines_are_reconstructed_as_choice_rows(self) -> None:
        blocks = [
            "34. 책임의 근거와 본질에 관한 학설의 설명으로 옳고 그름의 표시(O, X)가 바르게 된 것은?",
            "(가) 설명",
            "(나) 설명",
            "(다) 설명",
            "(라) 설명",
            "①",
            "㉠ (O)",
            "㉡ (O)",
            "㉢ (O)",
            "㉣ (O)",
            "②",
            "㉠ (O)",
            "㉡ (X)",
            "㉢ (O)",
            "㉣ (X)",
            "③",
            "㉠ (X)",
            "㉡ (O)",
            "㉢ (X)",
            "㉣ (O)",
            "④",
            "㉠ (X)",
            "㉡ (X)",
            "㉢ (X)",
            "㉣ (X)",
            "정답④",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertEqual(len(question.choices), 4)
        self.assertTrue(all("\t" in choice for choice in question.choices))
        self.assertEqual(len(question.sub_items), 0)
        self.assertEqual(question.answer, "④")

    def test_long_parenthesized_rows_are_promoted_to_sub_items_before_choices(self) -> None:
        blocks = [
            "34. 책임의 근거와 본질에 관한 학설의 설명으로 옳고 그름의 표시(O, X)가 바르게 된 것은?",
            "(가) 책임은 자유의사를 가진 자가 그 의사에 의하여 적법한 행위를 할 수 있었음에도 위법행위를 선택했다는 점에서 윤리적 비난을 받는다.",
            "(나) 인간의 행위는 자유의사가 아니라 환경과 소질에 의해 결정된다고 보며 책임의 근거를 반사회적 성격에서 찾는다.",
            "(다) 책임은 행위 당시의 고의 과실이라는 심리적 관계로 이해하여 그 심리적 사실의 유무로 판단된다고 본다.",
            "(라) 책임을 심리적 사실관계가 아닌 규범적 평가관계로 보며 적법행위를 기대할 수 있었는지로 판단한다.",
            "① ㉠ (O) ㉡ (O) ㉢ (O) ㉣ (O)",
            "② ㉠ (O) ㉡ (X) ㉢ (O) ㉣ (X)",
            "③ ㉠ (X) ㉡ (O) ㉢ (X) ㉣ (O)",
            "④ ㉠ (X) ㉡ (X) ㉢ (X) ㉣ (X)",
            "정답④",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertTrue(question.has_table)
        self.assertEqual(len(question.sub_items), 4)
        self.assertTrue(question.sub_items[0].startswith("(가)"))
        self.assertEqual(len(question.choices), 4)
        self.assertTrue(all("\t" not in choice for choice in question.choices))
        self.assertEqual(question.answer, "④")

    def test_compound_choice_line_is_split_into_multiple_choices(self) -> None:
        blocks = [
            "59. 법죄참가형태와 관련된 내용으로 옳은 것(O)과 틀린 것(X)을 바르게 나열한 것은?",
            "① ㉠ (O), ㉡ (O), ㉢ (O), ㉣ (X)",
            "② ㉠ (X), ㉡ (O), ㉢ (O), ㉣ (X) ③ ㉠ (X), ㉡ (O), ㉢ (X), ㉣ (O)",
            "④ ㉠ (O), ㉡ (O), ㉢ (O), ㉣ (O)",
            "정답②",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertEqual(len(question.choices), 4)
        self.assertTrue(question.choices[1].startswith("②"))
        self.assertTrue(question.choices[2].startswith("③"))
        self.assertEqual(question.answer, "②")

    def test_bare_number_question_start_after_choice_block_starts_new_question(self) -> None:
        blocks = [
            "59. 법죄참가형태와 관련된 내용으로 옳은 것(O)과 틀린 것(X)을 바르게 나열한 것은?",
            "① ㉠ (O), ㉡ (O), ㉢ (O), ㉣ (X)",
            "② ㉠ (X), ㉡ (O), ㉢ (O), ㉣ (X)",
            "정답②",
            "60 아래 (가)와 [나]에 관련된 설명으로 적절하지 않은 것을 <보기>에서 모두 고른 것은?",
            "(가) 구성요건 설명",
            "(나) 공범 설명",
            "① ㄱㄴ",
            "② ㄴㄷ",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.total_count, 2)
        self.assertEqual(document.questions[0].number, 59)
        self.assertEqual(document.questions[1].number, 60)
        self.assertTrue(document.questions[1].question_text.startswith("아래 (가)와 [나]에 관련된"))

    def test_answer_line_switches_to_explanation_and_keeps_reference_block_out_of_choices(self) -> None:
        parser = ExamParser(
            {
                "parsing": {
                    "question_patterns": self.config["parsing"]["question_patterns"],
                    "answer_patterns": self.config["parsing"]["answer_patterns"],
                    "explanation_patterns": self.config["parsing"]["explanation_patterns"],
                    "type_a_threshold": 99,
                },
                "negative_keywords": self.config["negative_keywords"],
            }
        )
        blocks = [
            "30. 음주운전 관련 설명으로 옳지 않은 것은?",
            "① 첫 번째 보기",
            "② 두 번째 보기",
            "정답 : ②",
            "해설 : (x) 두 번째 보기",
            "참고 : 음주운전 처벌기준",
            "0.03% 이상",
            "운전면허 취소 기준 강화",
            "31. 다음 설명으로 옳은 것은?",
            "① 세 번째 보기",
            "② 네 번째 보기",
        ]
        document = parser.parse_text_blocks(blocks, subject="형법")
        self.assertEqual(document.file_type, "TYPE_A")
        self.assertEqual(document.total_count, 2)
        question = document.questions[0]
        self.assertEqual(question.question_text, "음주운전 관련 설명으로 옳지 않은 것은?")
        self.assertEqual(question.choices, ["① 첫 번째 보기", "② 두 번째 보기"])
        self.assertEqual(question.answer, "②")
        self.assertIn("해설 : (x) 두 번째 보기", question.explanation or "")
        self.assertIn("참고 : 음주운전 처벌기준", question.explanation or "")
        self.assertIn("0.03% 이상", question.explanation or "")

    def test_numbered_explanation_rows_do_not_start_new_question(self) -> None:
        blocks = [
            "37. 확성기 사용 제한에 관한 설명으로 옳고 그름을 바르게 나열한 것은?",
            "① ㉠(X) ㉡(O) ㉢(O) ㉣(X)",
            "② ㉠(O) ㉡(X) ㉢(X) ㉣(O)",
            "③ ㉠(X) ㉡(O) ㉢(X) ㉣(O)",
            "④ ㉠(O) ㉡(X) ㉢(O) ㉣(X)",
            "정답 : ③",
            "해설 : ㉠ (x) 설명",
            "60. 이하",
            "50. 이하",
            "45. 이하",
            "38. 다음 설명 중 옳은 것은?",
            "① 보기A",
            "② 보기B",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="행정법")
        self.assertEqual(document.total_count, 2)
        q37 = document.questions[0]
        self.assertEqual(q37.number, 37)
        self.assertIn("60. 이하", q37.explanation or "")
        self.assertIn("50. 이하", q37.explanation or "")
        self.assertEqual(q37.answer, "③")
        self.assertEqual(document.questions[1].number, 38)

    def test_sub_item_rows_after_first_choice_do_not_merge_into_choice(self) -> None:
        blocks = [
            "37. 확성기 사용 제한에 관한 설명으로 옳고 그름을 바르게 나열한 것은?",
            "① ㉠(X) ㉡(O) ㉢(O) ㉣(X)",
            "㉠ 첫 번째 설명",
            "㉡ 두 번째 설명",
            "㉢ 세 번째 설명",
            "㉣ 네 번째 설명",
            "② ㉠(O) ㉡(X) ㉢(X) ㉣(O)",
            "③ ㉠(X) ㉡(O) ㉢(X) ㉣(O)",
            "④ ㉠(O) ㉡(X) ㉢(O) ㉣(X)",
            "정답 : ③",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="행정법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertTrue(question.has_table)
        self.assertEqual(len(question.sub_items), 4)
        self.assertTrue(question.sub_items[0].startswith("㉠ "))
        self.assertEqual(len(question.choices), 4)
        self.assertNotIn("첫 번째 설명", question.choices[0])

    def test_compound_sub_item_line_is_split_into_multiple_rows(self) -> None:
        blocks = [
            "37. 확성기 사용 제한에 관한 설명으로 옳고 그름을 바르게 나열한 것은?",
            "㉠ 첫 번째 설명 ㉡ 두 번째 설명 ㉢ 세 번째 설명 ㉣ 네 번째 설명",
            "① ㉠(X) ㉡(O) ㉢(O) ㉣(X)",
            "② ㉠(O) ㉡(X) ㉢(X) ㉣(O)",
            "정답 : ①",
        ]
        document = self.parser.parse_text_blocks(blocks, subject="행정법")
        self.assertEqual(document.total_count, 1)
        question = document.questions[0]
        self.assertEqual(len(question.sub_items), 4)
        self.assertEqual(question.sub_items[0], "㉠ 첫 번째 설명")
        self.assertEqual(question.sub_items[3], "㉣ 네 번째 설명")


if __name__ == "__main__":
    unittest.main()
