import unittest

from core.generator import OutputGenerator
from core.models import ExamQuestion


def _make_generator(sub_items_table: bool = True) -> OutputGenerator:
    return OutputGenerator(
        {
            "format": {
                "sub_items_table": sub_items_table,
            },
            "style": {
                "enabled": False,
            },
        }
    )


class GeneratorSubItemsTablePolicyTestCase(unittest.TestCase):
    def test_long_block_is_skipped_without_table_preference(self) -> None:
        generator = _make_generator(True)
        sub_items = [
            "item_a " + ("A" * 180),
            "item_b " + ("B" * 120),
        ]
        self.assertFalse(generator._should_use_sub_items_table(sub_items, prefer_table=False))

    def test_has_table_preference_forces_table_for_long_block(self) -> None:
        generator = _make_generator(True)
        sub_items = [
            "item_a " + ("A" * 180),
            "item_b " + ("B" * 120),
        ]
        self.assertTrue(generator._should_use_sub_items_table(sub_items, prefer_table=True))

    def test_has_table_preference_allows_single_line_box(self) -> None:
        generator = _make_generator(True)
        sub_items = ["甲은 늦은 밤 귀가하던 중 A로 오인하였다."]
        self.assertTrue(generator._should_use_sub_items_table(sub_items, prefer_table=True))
        self.assertFalse(generator._should_use_sub_items_table(sub_items, prefer_table=False))

    def test_global_table_toggle_still_disables_table(self) -> None:
        generator = _make_generator(False)
        sub_items = [
            "item_a test",
            "item_b test",
        ]
        self.assertFalse(generator._should_use_sub_items_table(sub_items, prefer_table=True))

    def test_inline_choice_matrix_is_not_detected_as_table(self) -> None:
        generator = _make_generator(True)
        choices = [
            "① ㉠ (X), ㉡ (O), ㉢ (X), ㉣ (O)",
            "② ㉠ (O), ㉡ (X), ㉢ (O), ㉣ (X)",
            "③ ㉠ (X), ㉡ (X), ㉢ (O), ㉣ (O)",
        ]
        self.assertFalse(generator._should_render_choices_as_table(choices))

    def test_tab_separated_choice_matrix_is_detected_as_table(self) -> None:
        generator = _make_generator(True)
        choices = [
            "①\t㉠ (X)\t㉡ (O)\t㉢ (X)\t㉣ (O)",
            "②\t㉠ (O)\t㉡ (X)\t㉢ (O)\t㉣ (X)",
            "③\t㉠ (X)\t㉡ (X)\t㉢ (O)\t㉣ (O)",
        ]
        self.assertTrue(generator._should_render_choices_as_table(choices))

    def test_plain_choices_are_not_detected_as_table(self) -> None:
        generator = _make_generator(True)
        choices = [
            "① 첫 번째 보기",
            "② 두 번째 보기",
            "③ 세 번째 보기",
            "④ 네 번째 보기",
        ]
        self.assertFalse(generator._should_render_choices_as_table(choices))

    def test_split_line_blocks_for_multiple_tables(self) -> None:
        generator = _make_generator(True)
        lines = ["<사례>", "사례 본문", "", "<보기>", "㉠ 설명", "㉡ 설명"]
        self.assertEqual(
            generator._split_line_blocks(lines),
            [["<사례>", "사례 본문"], ["<보기>", "㉠ 설명", "㉡ 설명"]],
        )

    def test_prefer_table_fallback_does_not_insert_box_characters(self) -> None:
        generator = _make_generator(True)
        generator.formatter.apply_sub_items_format = lambda hwp: None
        generator._leave_table_context = lambda hwp: True
        generator._force_table_context_cleanup = lambda hwp: None
        generator._create_single_cell_sub_items_table = lambda hwp: False

        inserted: list[str] = []
        generator._insert_text = lambda hwp, text: inserted.append(text)

        sub_items = ["㉠ alpha", "㉡ beta"]
        generator._insert_sub_items_block(object(), sub_items, question_number=1, prefer_table=True)

        self.assertEqual(inserted, ["㉠ alpha\r\n", "㉡ beta\r\n"])
        merged = "".join(inserted)
        self.assertNotIn("┌", merged)
        self.assertNotIn("└", merged)
        self.assertNotIn("│", merged)
        self.assertNotIn("+-", merged)

    def test_table_success_moves_caret_before_trailing_newline(self) -> None:
        generator = _make_generator(True)
        generator.formatter.apply_sub_items_format = lambda hwp: None
        generator._create_single_cell_sub_items_table = lambda hwp: True
        generator._set_current_table_treat_as_char = lambda hwp: True
        generator._leave_table_context = lambda hwp: True
        generator._force_table_context_cleanup = lambda hwp: None

        call_order: list[str] = []
        generator._move_caret_past_recent_table = lambda hwp: call_order.append("move")

        inserted: list[str] = []
        generator._insert_text = lambda hwp, text: inserted.append(text)

        generator._insert_sub_items_block(object(), ["a", "b"], question_number=1, prefer_table=True)

        self.assertEqual(call_order, ["move"])
        self.assertEqual(inserted, ["a", "\r\n", "b", "\r\n"])

    def test_table_followed_by_choices_skips_extra_newline(self) -> None:
        generator = _make_generator(True)
        generator.formatter.apply_sub_items_format = lambda hwp: None
        generator._create_single_cell_sub_items_table = lambda hwp: True
        generator._set_current_table_treat_as_char = lambda hwp: True
        generator._leave_table_context = lambda hwp: True
        generator._force_table_context_cleanup = lambda hwp: None
        generator._move_caret_past_recent_table = lambda hwp: None

        inserted: list[str] = []
        generator._insert_text = lambda hwp, text: inserted.append(text)

        generator._insert_sub_items_block(
            object(),
            ["a", "b"],
            question_number=1,
            prefer_table=True,
            has_following_choices=True,
        )
        self.assertEqual(inserted, ["a", "\r\n", "b"])

    def test_table_applies_sub_item_style_per_line(self) -> None:
        generator = _make_generator(True)
        style_hits: list[str] = []
        generator.formatter.apply_sub_items_format = lambda hwp: style_hits.append("hit")
        generator._create_single_cell_sub_items_table = lambda hwp: True
        generator._set_current_table_treat_as_char = lambda hwp: True
        generator._leave_table_context = lambda hwp: True
        generator._force_table_context_cleanup = lambda hwp: None
        generator._move_caret_past_recent_table = lambda hwp: None
        generator._insert_text = lambda hwp, text: None

        generator._insert_sub_items_block(
            object(),
            ["line1", "line2", "line3", "line4"],
            question_number=1,
            prefer_table=True,
            has_following_choices=True,
        )
        self.assertGreaterEqual(len(style_hits), 4)

    def test_table_runs_cell_wide_sub_item_format_pass(self) -> None:
        generator = _make_generator(True)
        calls: list[str] = []
        generator.formatter.apply_sub_items_format = lambda hwp: calls.append("style")

        class _HwpStub:
            class _Action:
                @staticmethod
                def Run(name: str) -> bool:
                    calls.append(name)
                    return True

            HAction = _Action()

        generator._apply_sub_items_format_to_current_table_cell(_HwpStub())

        self.assertIn("TableCellBlock", calls)
        self.assertIn("TableCellBlockExtend", calls)
        self.assertIn("style", calls)
        self.assertIn("Cancel", calls)

    def test_choice_table_is_suppressed_when_sub_item_table_exists(self) -> None:
        generator = _make_generator(True)
        generator.formatter.apply_question_format = lambda hwp, emphasize=False: None
        generator.formatter.apply_question_inline_char = lambda hwp, emphasize=False: None
        generator.formatter.apply_choice_format = lambda hwp: None
        generator._insert_text = lambda hwp, text: None
        generator._insert_question_text_with_emphasis = lambda hwp, text, keyword: None
        generator._leave_table_context = lambda hwp: True

        calls: list[dict[str, object]] = []

        def _capture(
            hwp,
            sub_items,
            question_number=None,
            prefer_table=False,
            has_following_choices=False,
            style_applier=None,
        ):
            calls.append(
                {
                    "sub_items": list(sub_items),
                    "prefer_table": bool(prefer_table),
                    "has_following_choices": bool(has_following_choices),
                }
            )

        generator._insert_sub_items_block = _capture

        question = ExamQuestion(
            number=1,
            question_text="질문",
            sub_items=["(가) 설명", "(나) 설명"],
            has_table=True,
            choices=[
                "①\t㉠ (O)\t㉡ (O)\t㉢ (O)\t㉣ (O)",
                "②\t㉠ (O)\t㉡ (X)\t㉢ (O)\t㉣ (X)",
            ],
        )

        generator._insert_question_block(object(), question)

        self.assertEqual(len(calls), 1)
        self.assertTrue(calls[0]["prefer_table"])
        self.assertTrue(calls[0]["has_following_choices"])

    def test_inline_choice_spacing_is_fixed_to_nine_spaces(self) -> None:
        generator = _make_generator(True)
        source = "\u2460 1개  \u2461 2개 \u2462 3개   \u2463 4개"
        expected = (
            "\u2460 1개"
            + (" " * 9)
            + "\u2461 2개"
            + (" " * 9)
            + "\u2462 3개"
            + (" " * 9)
            + "\u2463 4개"
        )
        self.assertEqual(generator._normalize_inline_choice_spacing(source), expected)

    def test_separate_short_choices_are_compacted_with_nine_spaces(self) -> None:
        generator = _make_generator(True)
        source = ["\u2460 1개", "\u2461 2개", "\u2462 3개", "\u2463 4개"]
        expected = (
            "\u2460 1개"
            + (" " * 9)
            + "\u2461 2개"
            + (" " * 9)
            + "\u2462 3개"
            + (" " * 9)
            + "\u2463 4개"
        )
        self.assertEqual(generator._build_choice_lines(source), [expected])

    def test_separate_long_choices_are_not_compacted(self) -> None:
        generator = _make_generator(True)
        source = [
            "\u2460 ㄱ, ㄴ, ㄷ 중에서 옳은 설명을 고르시오",
            "\u2461 ㄱ, ㄴ, ㄹ 중에서 옳은 설명을 고르시오",
            "\u2462 ㄴ, ㄷ, ㄹ 중에서 옳은 설명을 고르시오",
            "\u2463 ㄱ, ㄷ, ㄹ 중에서 옳은 설명을 고르시오",
        ]
        lines = generator._build_choice_lines(source)
        self.assertEqual(len(lines), 4)
        self.assertEqual(lines[0], source[0])

    def test_separate_short_kiueuk_choices_are_compacted(self) -> None:
        generator = _make_generator(True)
        source = ["\u2460 ㄱ,ㄴ", "\u2461 ㄴ,ㄷ", "\u2462 ㄱ,ㄷ", "\u2463 ㄱ,ㄴ,ㄷ"]
        expected = (
            "\u2460 ㄱ,ㄴ"
            + (" " * 9)
            + "\u2461 ㄴ,ㄷ"
            + (" " * 9)
            + "\u2462 ㄱ,ㄷ"
            + (" " * 9)
            + "\u2463 ㄱ,ㄴ,ㄷ"
        )
        self.assertEqual(generator._build_choice_lines(source), [expected])


    def test_inline_choice_spacing_strips_extraction_noise_suffix(self) -> None:
        generator = _make_generator(True)
        source = "\u2460 \uccad\uad6c\uc2dc\uae30 \u3c72 \u2461 \uccad\uad6c\uad8c\uc790"
        expected = (
            "\u2460 \uccad\uad6c\uc2dc\uae30"
            + (" " * 9)
            + "\u2461 \uccad\uad6c\uad8c\uc790"
        )
        self.assertEqual(generator._normalize_inline_choice_spacing(source), expected)

    def test_inline_choice_spacing_strips_noise_in_count_choices(self) -> None:
        generator = _make_generator(True)
        source = "\u2460 0\uac1c \u4546 \u2461 1\uac1c"
        expected = "\u2460 0\uac1c" + (" " * 9) + "\u2461 1\uac1c"
        self.assertEqual(generator._normalize_inline_choice_spacing(source), expected)

    def test_insert_explanation_block_uses_original_answer_line(self) -> None:
        generator = _make_generator(True)
        generator.formatter.apply_question_format = lambda hwp, emphasize=False: None
        generator.formatter.apply_explanation_format = lambda hwp: None

        inserted: list[str] = []
        generator._insert_text = lambda hwp, text: inserted.append(text)

        question = ExamQuestion(
            number=1,
            question_text="문제",
            answer="②",
            answer_line="[정답] (②)",
            explanation=None,
        )
        generator._insert_explanation_block(object(), question)
        self.assertEqual(inserted[0], "1. [정답] (②)\r\n")


if __name__ == "__main__":
    unittest.main()
