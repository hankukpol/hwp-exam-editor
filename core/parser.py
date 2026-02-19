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
    _BOXED_PASSAGE_START_RE = re.compile(r"^\s*[甲乙丙丁戊己庚辛壬癸]\s*[은는이가을를의]")
    _QUESTION_PROMPT_RE = re.compile(r"[?？]\s*$")
    _BOXED_PASSAGE_NOISE_RE = re.compile(r"^[\u3400-\u9FFF\uF900-\uFAFF]{2}$")
    _LEADING_NUMBER_DECORATION_RE = re.compile(r"^\s*[★☆※＊*]+\s*")
    _MARKER_LINE_RE = re.compile(r"^\s*[\[\(【]?\s*(정답|해설)\s*[\]\)】]?\s*[:：]?\s*$")
    _TABLE_BLOCK_MARKER_RE = re.compile(
        r"^\s*[<\[〈「【]?\s*(사례|보기|자료|표)\s*(?:\d+)?\s*[>\]〉」】]?\s*$"
    )
    _LETTERED_TABLE_ROW_RE = re.compile(
        r"^\s*[\(\[]\s*([가나다라마바사아자차카타파하])\s*[\)\]]\s*(.*)$"
    )
    _CHOICE_MARKER_RE = re.compile(r"^\s*([①②③④⑤])\s*(.*)$")
    _SUB_ITEM_MARKER_RE = re.compile(r"^\s*(㉠|㉡|㉢|㉣|㉤|㉥)\s*(.*)$")
    _TABLE_CHOICE_TOKEN_RE = re.compile(
        r"^\s*(㉠|㉡|㉢|㉣|㉤|㉥)\s*(?:[\(\[]?\s*[xXoO○×]\s*[\)\]]?)?\s*$"
    )
    _TABLE_CHOICE_ROW_TOKEN_EXTRACT_RE = re.compile(
        r"(㉠|㉡|㉢|㉣|㉤|㉥)\s*([\(\[]?\s*[xXoO○×]\s*[\)\]]?)"
    )

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
            question_patterns=self.question_patterns,
            explanation_patterns=self.explanation_patterns,
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
        current_answer_line: str | None = None
        mode = "question"

        def flush_current() -> None:
            nonlocal current_number, question_lines, explanation_lines, current_answer, current_answer_line, mode
            if current_number is None:
                return
            question = self._build_question(
                number=current_number,
                question_lines=question_lines,
                answer=current_answer,
                answer_line=current_answer_line,
                explanation_lines=explanation_lines,
            )
            document.questions.append(question)
            current_number = None
            question_lines = []
            explanation_lines = []
            current_answer = None
            current_answer_line = None
            mode = "question"

        for index, block in enumerate(clean_blocks):
            next_block = clean_blocks[index + 1] if index + 1 < len(clean_blocks) else ""
            number = extract_question_number(block, self.question_patterns)
            if number is not None:
                if (
                    current_number is not None
                    and mode == "explanation"
                    and not self._is_probable_explicit_question_header(block, next_block)
                ):
                    explanation_lines.append(block)
                    continue
                flush_current()
                current_number = number
                stripped = self._strip_question_prefix(block)
                if stripped:
                    question_lines.append(stripped)
                continue

            if current_number is None:
                continue

            if (mode == "question" or current_answer is None) and self._is_answer_line(block, next_block):
                extracted_answer = self._extract_answer(block)
                if extracted_answer:
                    current_answer = extracted_answer
                current_answer_line = block.strip() or None
                mode = "explanation"
                continue

            if self._is_explanation_marker(block):
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
        current_answer_line: str | None = None
        mode = "question"
        has_started = False

        def flush_current() -> None:
            nonlocal current_number, question_lines, explanation_lines, current_answer, current_answer_line, mode
            if not question_lines and not explanation_lines and not current_answer and not current_answer_line:
                return
            current_number += 1
            question = self._build_question(
                number=current_number,
                question_lines=question_lines,
                answer=current_answer,
                answer_line=current_answer_line,
                explanation_lines=explanation_lines,
            )
            document.questions.append(question)
            question_lines = []
            explanation_lines = []
            current_answer = None
            current_answer_line = None
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

            if (mode == "question" or current_answer is None) and self._is_answer_line(block, next_block):
                extracted_answer = self._extract_answer(block)
                if extracted_answer:
                    current_answer = extracted_answer
                current_answer_line = block.strip() or None
                mode = "explanation"
                continue

            if self._is_explanation_marker(block):
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
        if re.match(
            r"^\s*[\[\(【]?\s*정답\s*[\]\)】]?\s*[:：]?\s*[\(\[]?\s*[①②③④⑤1-5]\s*[\)\]]?\s*(?:번)?(\s*[\(\[][^\)\]]+[\)\]])?\s*$",
            text,
        ):
            return True
        if re.match(r"^\s*[①②③④⑤1-5]\s*[\(\[][^\)\]]+[\)\]]\s*$", text):
            return True
        if re.match(r"^\s*[\(\[]?\s*[①②③④⑤1-5]\s*[\)\]]?\s*$", text) and self._is_explanation_marker(next_text):
            return True
        return False

    def _is_explanation_marker(self, text: str) -> bool:
        stripped = text.strip()
        if self._MARKER_LINE_RE.match(stripped):
            return True
        if re.match(
            r"^\s*[\[\(【]?\s*(해설|참고|핵심정리|관련\s*판례)\s*[\]\)】]?\s*[:：]?",
            stripped,
        ):
            return True
        return is_line_matching(text, self.explanation_patterns)

    def _is_pure_explanation_marker(self, text: str) -> bool:
        return bool(self._MARKER_LINE_RE.match(text.strip()))

    def _is_probable_question_start(self, text: str) -> bool:
        stripped = text.strip()
        if len(stripped) < 8:
            return False
        if not re.search(r"[?？]", stripped):
            return False
        if re.match(r"^\s*[①②③④⑤㉠㉡㉢㉣㉤㉥]", stripped):
            return False
        if stripped.startswith(("[○]", "[×]")):
            return False
        if self._MARKER_LINE_RE.match(stripped):
            return False
        hints = ("다음", "것은", "설명", "관련", "내용", "기술", "입장", "자유", "죄")
        return any(hint in stripped for hint in hints)

    def _is_probable_explicit_question_header(self, text: str, next_text: str = "") -> bool:
        stripped = (text or "").strip()
        if not stripped:
            return False
        if extract_question_number(stripped, self.question_patterns) is None:
            return False

        payload = self._strip_question_prefix(stripped).strip()
        if not payload:
            return False
        if re.search(r"[?？]", payload):
            return True
        if self._MARKER_LINE_RE.match(payload):
            return False
        if self._CHOICE_MARKER_RE.match(payload):
            return False
        if self._SUB_ITEM_MARKER_RE.match(payload):
            return False

        next_stripped = (next_text or "").strip()
        if next_stripped and self._CHOICE_MARKER_RE.match(next_stripped):
            return True

        hints = ("다음", "것은", "설명", "관련", "내용", "판례", "옳은", "틀린", "적절")
        if len(payload) >= 8 and any(hint in payload for hint in hints):
            return True
        return False

    def _normalize_blocks(self, text_blocks: list[str]) -> list[str]:
        normalized: list[str] = []
        for block in text_blocks:
            for line in re.split(r"\r?\n", block):
                text = line.strip()
                if text:
                    normalized.append(text)
        return normalized

    def _strip_question_prefix(self, text: str) -> str:
        candidate = self._LEADING_NUMBER_DECORATION_RE.sub("", text, count=1)
        for pattern in self.question_patterns:
            try:
                stripped = re.sub(pattern, "", candidate, count=1).strip()
            except re.error:
                continue
            if stripped != candidate:
                return stripped
        fallback = re.sub(r"^\s*문\s*(\d{1,2})\s*[\.\)]\s*", "", candidate, count=1).strip()
        if fallback != candidate.strip():
            return fallback
        fallback = re.sub(r"^\s*(\d{1,2})\s*[\.\)]\s*", "", candidate, count=1).strip()
        if fallback != candidate.strip():
            return fallback
        fallback = re.sub(r"^\s*(\d{1,2})\s+(?=[가-힣A-Za-z<\[\(])", "", candidate, count=1).strip()
        if fallback != candidate.strip():
            return fallback
        return candidate.strip()

    def _extract_answer(self, text: str) -> str | None:
        circle = re.search(r"[①②③④⑤]", text)
        if circle:
            return circle.group(0)
        digit = re.search(r"(?<!\d)([1-5])(?!\d)", text)
        if digit:
            return CIRCLE_FROM_DIGIT.get(digit.group(1))
        return None

    def _split_boxed_passage_lines(
        self, body_lines: list[str], choices: list[str],
    ) -> tuple[list[str], list[str]]:
        """사례형(박스) 지문을 문제 본문에서 분리한다."""
        if choices is None or len(choices) == 0:
            return body_lines, []
        if len(body_lines) < 2:
            return body_lines, []

        stem = (body_lines[0] or "").strip()
        raw_passage_lines = [(line or "").strip() for line in body_lines[1:] if (line or "").strip()]
        passage_lines = [
            line for line in raw_passage_lines
            if not self._BOXED_PASSAGE_NOISE_RE.fullmatch(line)
        ]
        if not passage_lines:
            return body_lines, []
        if not self._QUESTION_PROMPT_RE.search(stem):
            return body_lines, []

        first_line = passage_lines[0]
        if self._BOXED_PASSAGE_START_RE.match(first_line):
            return [stem], passage_lines
        return body_lines, []

    def _split_lettered_table_rows(
        self, body_lines: list[str], choices: list[str],
    ) -> tuple[list[str], list[str]]:
        """(가)/(나)/(다)/(라)처럼 이어지는 지문을 표 블록으로 승격한다."""
        if choices is None or len(choices) == 0:
            return body_lines, []
        if len(body_lines) < 3:
            return body_lines, []

        stem = (body_lines[0] or "").strip()
        if not stem:
            return body_lines, []

        rows: list[str] = []
        payload_lengths: list[int] = []
        index = 1
        while index < len(body_lines):
            line = (body_lines[index] or "").strip()
            if not line:
                break
            match = self._LETTERED_TABLE_ROW_RE.match(line)
            if not match:
                break
            payload = (match.group(2) or "").strip()
            if not payload:
                break
            rows.append(line)
            payload_lengths.append(len(payload))
            index += 1

        if len(rows) < 2:
            return body_lines, []

        # 너무 짧은 "(가) 설명"류는 일반 본문으로 유지한다.
        if max(payload_lengths, default=0) < 12 and sum(payload_lengths) < 40:
            return body_lines, []

        trailing = [line for line in body_lines[index:] if (line or "").strip()]
        if trailing:
            return body_lines, []

        return [stem], rows

    def _extract_marked_table_blocks(
        self, question_lines: list[str],
    ) -> tuple[list[list[str]], set[int]]:
        blocks: list[list[str]] = []
        consumed_indices: set[int] = set()
        index = 0

        while index < len(question_lines):
            line = (question_lines[index] or "").strip()
            starts_with_marker = bool(self._TABLE_BLOCK_MARKER_RE.match(line))
            starts_with_lettered_row = bool(self._LETTERED_TABLE_ROW_RE.match(line))
            if starts_with_lettered_row and not starts_with_marker:
                if not self._has_following_table_block_marker_before_choice(question_lines, index + 1):
                    index += 1
                    continue
            if not starts_with_marker and not starts_with_lettered_row:
                index += 1
                continue

            start = index
            block: list[str] = [line]
            index += 1
            while index < len(question_lines):
                current = (question_lines[index] or "").strip()
                if self._TABLE_BLOCK_MARKER_RE.match(current):
                    break
                if re.match(r"^\s*[①②③④⑤]\s*", current):
                    break
                if self._is_answer_line(current):
                    break
                if self._is_explanation_marker(current):
                    break
                block.append(current)
                index += 1

            block = [item for item in block if item]
            if len(block) >= 2:
                blocks.append(block)
                consumed_indices.update(range(start, index))
            else:
                index = start + 1

        return blocks, consumed_indices

    def _has_following_table_block_marker_before_choice(
        self, question_lines: list[str], start_index: int,
    ) -> bool:
        index = start_index
        while index < len(question_lines):
            current = (question_lines[index] or "").strip()
            if self._TABLE_BLOCK_MARKER_RE.match(current):
                return True
            if re.match(r"^\s*[①②③④⑤]\s*", current):
                return False
            if self._is_answer_line(current):
                return False
            if self._is_explanation_marker(current):
                return False
            index += 1
        return False

    def _build_question(
        self,
        number: int,
        question_lines: list[str],
        answer: str | None,
        answer_line: str | None,
        explanation_lines: list[str],
    ) -> ExamQuestion:
        choices: list[str] = []
        sub_items: list[str] = []
        body_lines: list[str] = []
        marked_table_blocks, consumed_indices = self._extract_marked_table_blocks(question_lines)
        detected_lettered_sub_items: list[str] = []
        current_choice_marker: str | None = None
        current_choice_parts: list[str] = []

        def flush_choice() -> None:
            nonlocal current_choice_marker, current_choice_parts
            if current_choice_marker is None:
                return

            compact_parts = [part.strip() for part in current_choice_parts if part.strip()]
            if not compact_parts:
                choices.append(current_choice_marker)
            elif len(compact_parts) == 1:
                single = compact_parts[0]
                normalized_row = self._normalize_table_choice_row(single)
                if normalized_row is not None:
                    choices.append(f"{current_choice_marker}\t{normalized_row}")
                else:
                    choices.append(f"{current_choice_marker} {single}".strip())
            elif len(compact_parts) >= 2 and all(self._is_table_choice_token(part) for part in compact_parts):
                choices.append(f"{current_choice_marker}\t" + "\t".join(compact_parts))
            else:
                choices.append(f"{current_choice_marker} {' '.join(compact_parts)}".strip())

            current_choice_marker = None
            current_choice_parts = []

        for idx, line in enumerate(question_lines):
            if idx in consumed_indices:
                continue
            stripped_line = (line or "").strip()
            choice_match = self._CHOICE_MARKER_RE.match(stripped_line)
            if choice_match:
                segments = self._split_compound_choice_segments(stripped_line)
                if len(segments) > 1:
                    flush_choice()
                    for marker, payload in segments[:-1]:
                        normalized_row = self._normalize_table_choice_row(payload)
                        if normalized_row is not None:
                            choices.append(f"{marker}\t{normalized_row}")
                        else:
                            payload_text = payload.strip()
                            choices.append(f"{marker} {payload_text}".strip() if payload_text else marker)
                    current_choice_marker = segments[-1][0]
                    current_choice_parts = []
                    last_payload = segments[-1][1].strip()
                    if last_payload:
                        current_choice_parts.append(last_payload)
                    continue

                flush_choice()
                current_choice_marker = choice_match.group(1)
                payload = choice_match.group(2).strip()
                if payload:
                    current_choice_parts.append(payload)
                continue

            if current_choice_marker is not None:
                stripped = (line or "").strip()
                if self._is_table_choice_token(stripped):
                    current_choice_parts.append(stripped)
                    continue
                sub_item_segments = self._split_compound_sub_item_segments(stripped)
                if sub_item_segments:
                    flush_choice()
                    sub_items.extend(sub_item_segments)
                    continue
                if self._looks_like_choice_continuation(stripped):
                    current_choice_parts.append(stripped)
                    continue
                flush_choice()

            sub_item_segments = self._split_compound_sub_item_segments(stripped_line)
            if sub_item_segments:
                sub_items.extend(sub_item_segments)
                continue
            body_lines.append(line)

        flush_choice()

        detected_boxed_sub_items: list[str] = []
        if not sub_items:
            body_lines, detected_boxed_sub_items = self._split_boxed_passage_lines(body_lines, choices)
            if detected_boxed_sub_items:
                sub_items.extend(detected_boxed_sub_items)
        if not sub_items and not marked_table_blocks:
            body_lines, detected_lettered_sub_items = self._split_lettered_table_rows(body_lines, choices)
            if detected_lettered_sub_items:
                sub_items.extend(detected_lettered_sub_items)

        if marked_table_blocks:
            flattened: list[str] = []
            for block in marked_table_blocks:
                if flattened:
                    flattened.append("")
                flattened.extend(block)
            if sub_items:
                flattened.append("")
                flattened.extend(sub_items)
            sub_items = flattened

        question_text = "\n".join(body_lines).strip()
        explanation = "\n".join(explanation_lines).strip() or None
        negative_keyword = detect_negative_keyword(question_text, self.negative_keywords)
        has_table = (
            len(sub_items) >= 2
            or bool(detected_boxed_sub_items)
            or bool(detected_lettered_sub_items)
            or bool(marked_table_blocks)
        )

        return ExamQuestion(
            number=number,
            question_text=question_text,
            choices=choices,
            sub_items=sub_items,
            has_table=has_table,
            has_negative=bool(negative_keyword),
            negative_keyword=negative_keyword,
            answer_line=answer_line,
            answer=answer,
            explanation=explanation,
        )

    def _is_table_choice_token(self, text: str) -> bool:
        stripped = (text or "").strip()
        if not stripped:
            return False
        if "\t" in stripped:
            parts = [part.strip() for part in stripped.split("\t") if part.strip()]
            return bool(parts) and all(self._TABLE_CHOICE_TOKEN_RE.match(part) for part in parts)
        return bool(self._TABLE_CHOICE_TOKEN_RE.match(stripped))

    def _normalize_table_choice_row(self, text: str) -> str | None:
        stripped = (text or "").strip()
        if not stripped:
            return None
        if "\t" in stripped:
            parts = [part.strip() for part in stripped.split("\t") if part.strip()]
            if len(parts) >= 2 and all(self._is_table_choice_token(part) for part in parts):
                return "\t".join(parts)
        return None

    def _split_compound_choice_segments(self, text: str) -> list[tuple[str, str]]:
        stripped = (text or "").strip()
        if not stripped:
            return []
        if not self._CHOICE_MARKER_RE.match(stripped):
            return []

        matches = list(re.finditer(r"[①②③④⑤]", stripped))
        if len(matches) <= 1:
            marker = stripped[0]
            payload = stripped[1:].strip()
            return [(marker, payload)]

        segments: list[tuple[str, str]] = []
        for index, match in enumerate(matches):
            start = match.start()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(stripped)
            chunk = stripped[start:end].strip()
            if not chunk:
                continue
            marker = chunk[0]
            payload = chunk[1:].strip()
            segments.append((marker, payload))
        return segments

    def _split_compound_sub_item_segments(self, text: str) -> list[str]:
        stripped = (text or "").strip()
        if not stripped:
            return []
        if not self._SUB_ITEM_MARKER_RE.match(stripped):
            return []

        marker_matches = list(re.finditer(r"[㉠㉡㉢㉣㉤㉥]", stripped))
        if not marker_matches:
            return []

        chunks: list[str] = []
        for index, match in enumerate(marker_matches):
            start = match.start()
            end = marker_matches[index + 1].start() if index + 1 < len(marker_matches) else len(stripped)
            chunk = stripped[start:end].strip()
            if not chunk:
                continue
            parts = self._SUB_ITEM_MARKER_RE.match(chunk)
            if not parts:
                continue
            chunks.append(f"{parts.group(1)} {parts.group(2).strip()}".strip())

        if len(chunks) >= 2 and all(self._is_table_choice_token(chunk) for chunk in chunks):
            return []
        return chunks

    def _looks_like_choice_continuation(self, text: str) -> bool:
        stripped = (text or "").strip()
        if not stripped:
            return False
        if self._CHOICE_MARKER_RE.match(stripped):
            return False
        if self._MARKER_LINE_RE.match(stripped):
            return False
        if self._is_answer_line(stripped):
            return False
        if self._is_explanation_marker(stripped):
            return False
        if extract_question_number(stripped, self.question_patterns) is not None:
            return False
        return True
