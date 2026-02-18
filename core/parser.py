from __future__ import annotations

import re
from typing import Any

from .detector import (
    DEFAULT_ANSWER_PATTERNS,
    DEFAULT_EXPLANATION_PATTERNS,
    DEFAULT_NEGATIVE_KEYWORDS,
    DEFAULT_QUESTION_PATTERNS,
    detect_file_type,
    detect_negative_keyword,
    extract_question_number,
    is_line_matching,
)
from .models import ExamDocument, ExamQuestion


CIRCLE_FROM_DIGIT = {
    "1": "①",
    "2": "②",
    "3": "③",
    "4": "④",
    "5": "⑤",
}


class ExamParser:
    def __init__(self, config: dict[str, Any]) -> None:
        parsing = config.get("parsing", {})
        self.question_patterns: list[str] = parsing.get(
            "question_patterns", DEFAULT_QUESTION_PATTERNS
        )
        self.answer_patterns: list[str] = parsing.get("answer_patterns", DEFAULT_ANSWER_PATTERNS)
        self.explanation_patterns: list[str] = parsing.get(
            "explanation_patterns", DEFAULT_EXPLANATION_PATTERNS
        )
        self.type_a_threshold: int = parsing.get("type_a_threshold", 5)
        self.negative_keywords: list[str] = config.get(
            "negative_keywords", DEFAULT_NEGATIVE_KEYWORDS
        )

    def parse_text_blocks(self, text_blocks: list[str], subject: str = "") -> ExamDocument:
        clean_blocks = self._normalize_blocks(text_blocks)
        file_type = detect_file_type(
            clean_blocks,
            answer_patterns=self.answer_patterns,
            threshold=self.type_a_threshold,
        )

        document = self._parse_with_explicit_numbers(clean_blocks, subject, file_type)
        if document.total_count == 0:
            document = self._parse_without_explicit_numbers(clean_blocks, subject, file_type)
        document.refresh_total_count()
        return document

    def _parse_with_explicit_numbers(
        self,
        clean_blocks: list[str],
        subject: str,
        file_type: str,
    ) -> ExamDocument:
        document = ExamDocument(file_type=file_type, subject=subject, questions=[])
        current_number: int | None = None
        question_lines: list[str] = []
        explanation_lines: list[str] = []
        current_answer: str | None = None
        mode = "question"

        def flush_current() -> None:
            nonlocal current_number, question_lines, explanation_lines, current_answer, mode
            if current_number is None:
                return
            question = self._build_question(
                number=current_number,
                question_lines=question_lines,
                answer=current_answer,
                explanation_lines=explanation_lines,
            )
            document.questions.append(question)
            current_number = None
            question_lines = []
            explanation_lines = []
            current_answer = None
            mode = "question"

        for index, block in enumerate(clean_blocks):
            next_block = clean_blocks[index + 1] if index + 1 < len(clean_blocks) else ""
            number = extract_question_number(block, self.question_patterns)
            if number is not None:
                flush_current()
                current_number = number
                stripped = self._strip_question_prefix(block)
                if stripped:
                    question_lines.append(stripped)
                continue

            if current_number is None:
                continue

            if self._is_answer_line(block, next_block):
                extracted_answer = self._extract_answer(block)
                if extracted_answer:
                    current_answer = extracted_answer
                if file_type == "TYPE_A":
                    mode = "explanation"
                continue

            if file_type == "TYPE_A" and self._is_explanation_marker(block):
                mode = "explanation"
                if self._is_pure_explanation_marker(block):
                    continue

            if mode == "question":
                question_lines.append(block)
            else:
                explanation_lines.append(block)

        flush_current()
        document.refresh_total_count()
        return document

    def _parse_without_explicit_numbers(
        self,
        clean_blocks: list[str],
        subject: str,
        file_type: str,
    ) -> ExamDocument:
        document = ExamDocument(file_type=file_type, subject=subject, questions=[])
        current_number = 0
        question_lines: list[str] = []
        explanation_lines: list[str] = []
        current_answer: str | None = None
        mode = "question"
        has_started = False

        def flush_current() -> None:
            nonlocal current_number, question_lines, explanation_lines, current_answer, mode
            if not question_lines and not explanation_lines and not current_answer:
                return
            current_number += 1
            question = self._build_question(
                number=current_number,
                question_lines=question_lines,
                answer=current_answer,
                explanation_lines=explanation_lines,
            )
            document.questions.append(question)
            question_lines = []
            explanation_lines = []
            current_answer = None
            mode = "question"

        for index, block in enumerate(clean_blocks):
            next_block = clean_blocks[index + 1] if index + 1 < len(clean_blocks) else ""
            if self._is_probable_question_start(block):
                if has_started:
                    flush_current()
                has_started = True
                mode = "question"
                question_lines.append(block)
                continue

            if not has_started:
                continue

            if self._is_answer_line(block, next_block):
                extracted_answer = self._extract_answer(block)
                if extracted_answer:
                    current_answer = extracted_answer
                if file_type == "TYPE_A":
                    mode = "explanation"
                continue

            if file_type == "TYPE_A" and self._is_explanation_marker(block):
                mode = "explanation"
                if self._is_pure_explanation_marker(block):
                    continue

            if mode == "question":
                question_lines.append(block)
            else:
                explanation_lines.append(block)

        flush_current()
        document.refresh_total_count()
        return document

    def _is_answer_line(self, text: str, next_text: str = "") -> bool:
        if is_line_matching(text, self.answer_patterns):
            return True
        if re.match(r"^\s*정답\s*[:：]?\s*[①②③④⑤1-5](\s*[\(\[][^\)\]]+[\)\]])?\s*$", text):
            return True
        if re.match(r"^\s*[①②③④⑤1-5]\s*[\(\[][^\)\]]+[\)\]]\s*$", text):
            return True
        if re.match(r"^\s*[①②③④⑤1-5]\s*$", text) and self._is_explanation_marker(next_text):
            return True
        return False

    def _is_explanation_marker(self, text: str) -> bool:
        if text.strip() in {"정답", "해설"}:
            return True
        return is_line_matching(text, self.explanation_patterns)

    def _is_pure_explanation_marker(self, text: str) -> bool:
        stripped = text.strip()
        return stripped in {"정답", "해설", "정답:", "정답：", "해설:", "해설："}

    def _is_probable_question_start(self, text: str) -> bool:
        stripped = text.strip()
        if len(stripped) < 8:
            return False
        if not re.search(r"[?？]", stripped):
            return False
        if re.match(r"^\s*[①②③④⑤㉠㉡㉢㉣㉤㉥]", stripped):
            return False
        if stripped.startswith(("정답", "해설", "[○]", "[×]")):
            return False
        hints = ("다음", "것은", "설명", "관련", "내용", "기술", "입장", "자유", "죄")
        return any(hint in stripped for hint in hints)

    def _normalize_blocks(self, text_blocks: list[str]) -> list[str]:
        normalized: list[str] = []
        for block in text_blocks:
            for line in re.split(r"\r?\n", block):
                text = line.strip()
                if text:
                    normalized.append(text)
        return normalized

    def _strip_question_prefix(self, text: str) -> str:
        for pattern in self.question_patterns:
            try:
                stripped = re.sub(pattern, "", text, count=1).strip()
            except re.error:
                continue
            if stripped != text:
                return stripped
        return text.strip()

    def _extract_answer(self, text: str) -> str | None:
        circle = re.search(r"[①②③④⑤]", text)
        if circle:
            return circle.group(0)
        digit = re.search(r"(?<!\d)([1-5])(?!\d)", text)
        if digit:
            return CIRCLE_FROM_DIGIT.get(digit.group(1))
        return None

    def _build_question(
        self,
        number: int,
        question_lines: list[str],
        answer: str | None,
        explanation_lines: list[str],
    ) -> ExamQuestion:
        choices: list[str] = []
        sub_items: list[str] = []
        body_lines: list[str] = []

        for line in question_lines:
            choice_match = re.match(r"^\s*([①②③④⑤])\s*(.*)$", line)
            if choice_match:
                choices.append(f"{choice_match.group(1)} {choice_match.group(2).strip()}".strip())
                continue
            sub_item_match = re.match(r"^\s*(㉠|㉡|㉢|㉣|㉤|㉥)\s*(.*)$", line)
            if sub_item_match:
                sub_items.append(f"{sub_item_match.group(1)} {sub_item_match.group(2).strip()}".strip())
                continue
            body_lines.append(line)

        question_text = "\n".join(body_lines).strip()
        explanation = "\n".join(explanation_lines).strip() or None
        negative_keyword = detect_negative_keyword(question_text, self.negative_keywords)
        has_table = len(sub_items) >= 2

        return ExamQuestion(
            number=number,
            question_text=question_text,
            choices=choices,
            sub_items=sub_items,
            has_table=has_table,
            has_negative=bool(negative_keyword),
            negative_keyword=negative_keyword,
            answer=answer,
            explanation=explanation,
        )
