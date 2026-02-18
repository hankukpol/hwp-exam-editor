import unittest

from core.generator import OutputGenerator


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

    def test_global_table_toggle_still_disables_table(self) -> None:
        generator = _make_generator(False)
        sub_items = [
            "item_a test",
            "item_b test",
        ]
        self.assertFalse(generator._should_use_sub_items_table(sub_items, prefer_table=True))

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
        self.assertEqual(inserted, ["a\r\nb", "\r\n"])

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
        self.assertEqual(inserted, ["a\r\nb"])

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


if __name__ == "__main__":
    unittest.main()
