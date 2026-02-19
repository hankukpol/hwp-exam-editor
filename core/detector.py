from __future__ import annotations

import re
from typing import Optional


DEFAULT_QUESTION_PATTERNS = [
    r"^\s*[★☆※＊*]+\s*(\d{1,2})\s*[\.\)]\s*",
    r"^\s*(\d{1,2})\s*[\.\)]\s*",
    r"^\s*(\d{1,2})\s+(?=[가-힣A-Za-z<\[\(])",
    r"^\s*(0\d)\s*[\.\)]\s*",
    r"^\s*【\s*(\d{1,2})\s*】\s*",
    r"^\s*문\s*(\d{1,2})\s*[\.\)]\s*",
]

DEFAULT_ANSWER_PATTERNS = [
    r"^\s*정답\s+[①②③④⑤]",
    r"^\s*정답[①②③④⑤]",
    r"^\s*정답\s*[:：]\s*[①②③④⑤]",
    r"^\s*정답\s*[:：]\s*\d",
    r"^\s*[\[\(【]?\s*정답\s*[\]\)】]?\s*[:：]?\s*[\(\[]?\s*[①②③④⑤1-5]\s*[\)\]]?\s*(?:번)?\s*$",
]

DEFAULT_EXPLANATION_PATTERNS = [
    r"^\s*해설\s*[\[【]?[×○xo]?[】\]]?",
    r"^\s*해설\s*[:：]",
    r"^\s*[①②③④⑤]\s*[:：]?\s*[\[\(]?[×○xo][\]\)]?",
    r"^\s*★\s*관련\s*판례",
]

DEFAULT_NEGATIVE_KEYWORDS = [
    "옳지 않은",
    "옳지않은",
    "틀린",
    "적절하지 않은",
    "적절하지않은",
    "올바르지 아니한",
    "올바르지아니한",
    "올바르지 않은",
    "올바르지않은",
    "가장 적절하지 않은",
    "가장 옳지 않은",
    "잘못된",
    "부적절한",
    "아닌 것",
    "않은 것",
]


def _compile_patterns(patterns: list[str]) -> list[re.Pattern[str]]:
    compiled: list[re.Pattern[str]] = []
    for pattern in patterns:
        try:
            compiled.append(re.compile(pattern))
        except re.error:
            continue
    return compiled


def detect_file_type(
    text_blocks: list[str],
    answer_patterns: Optional[list[str]] = None,
    threshold: int = 5,
    question_patterns: Optional[list[str]] = None,
    explanation_patterns: Optional[list[str]] = None,
) -> str:
    patterns = _compile_patterns(answer_patterns or DEFAULT_ANSWER_PATTERNS)
    question_compiled = _compile_patterns(question_patterns or DEFAULT_QUESTION_PATTERNS)
    explanation_compiled = _compile_patterns(explanation_patterns or DEFAULT_EXPLANATION_PATTERNS)
    fallback_answer_pattern = re.compile(
        r"^\s*[\[\(【]?\s*정답\s*[\]\)】]?\s*[:：]?\s*[\(\[]?\s*[①②③④⑤1-5]\s*[\)\]]?\s*(?:번)?(\s*[\(\[][^\)\]]+[\)\]])?\s*$"
    )
    grouped_answer_pattern = re.compile(r"^\s*[①②③④⑤1-5]\s*[\(\[][^\)\]]+[\)\]]\s*$")
    plain_answer_pattern = re.compile(r"^\s*[\(\[]?\s*[①②③④⑤1-5]\s*[\)\]]?\s*$")
    marker_pattern = re.compile(r"^\s*[\[\(【]?\s*(정답|해설)\s*[\]\)】]?\s*[:：]?\s*$")
    explanation_hint_pattern = re.compile(
        r"^\s*[\[\(【]?\s*(해설|참고|핵심정리|관련\s*판례)\s*[\]\)】]?\s*[:：]?"
    )
    answer_count = 0
    explanation_count = 0
    question_count = 0
    for index, block in enumerate(text_blocks):
        stripped = (block or "").strip()
        if not stripped:
            continue

        if any(pattern.match(stripped) for pattern in question_compiled) and re.search(r"[?？]", stripped):
            question_count += 1

        if any(pattern.match(block) for pattern in patterns):
            answer_count += 1
            continue
        if fallback_answer_pattern.match(block) or grouped_answer_pattern.match(block):
            answer_count += 1
            continue
        if plain_answer_pattern.match(block):
            next_block = text_blocks[index + 1] if index + 1 < len(text_blocks) else ""
            if marker_pattern.match(next_block):
                answer_count += 1

        if any(pattern.match(stripped) for pattern in explanation_compiled) or explanation_hint_pattern.match(stripped):
            explanation_count += 1

    adaptive_threshold = max(1, int(threshold))
    if question_count > 0:
        # 작은 문항 묶음(예: 2~4문항)도 TYPE_A로 인식되도록 문항 수 기반 하한을 둔다.
        adaptive_threshold = min(adaptive_threshold, max(1, (question_count + 1) // 2))

    if answer_count >= adaptive_threshold:
        return "TYPE_A"
    if answer_count >= 1 and explanation_count >= 1:
        return "TYPE_A"
    return "TYPE_B"


def extract_question_number(
    text: str,
    question_patterns: Optional[list[str]] = None,
) -> Optional[int]:
    normalized_text = re.sub(r"^\s*[★☆※＊*]+\s*", "", text)
    patterns = _compile_patterns(question_patterns or DEFAULT_QUESTION_PATTERNS)
    for pattern in patterns:
        match = pattern.match(normalized_text)
        if not match:
            continue
        remainder = normalized_text[match.end():]
        if remainder and remainder[:1].isdigit():
            # e.g. "0.03%" should not be interpreted as question number "0."
            continue
        for group in match.groups():
            if group and group.isdigit():
                number = int(group)
                if number <= 0:
                    continue
                return number
        digits = re.search(r"\d{1,2}", match.group(0))
        if digits:
            number = int(digits.group(0))
            if number <= 0:
                continue
            return number

    # Defensive fallback for custom/legacy pattern lists in user config.
    fallback_patterns = [
        r"^\s*(\d{1,2})\s*[\.\)]\s*",
        r"^\s*(\d{1,2})\s+(?=[가-힣A-Za-z<\[\(])",
        r"^\s*문\s*(\d{1,2})\s*[\.\)]?\s*",
    ]
    for raw in fallback_patterns:
        try:
            match = re.match(raw, normalized_text)
        except re.error:
            match = None
        if not match:
            continue
        digits = match.group(1)
        if digits and digits.isdigit():
            number = int(digits)
            if number <= 0:
                continue
            return number
    return None


def is_line_matching(text: str, patterns: list[str]) -> bool:
    for pattern in _compile_patterns(patterns):
        if pattern.match(text):
            return True
    return False


def detect_negative_keyword(text: str, keywords: Optional[list[str]] = None) -> str:
    token = _detect_negative_token_by_rule(text)
    if token:
        return token

    for keyword in keywords or DEFAULT_NEGATIVE_KEYWORDS:
        index = text.find(keyword)
        if index < 0:
            continue
        return _map_negative_emphasis_token(text, keyword, index)
    return ""


def _map_negative_emphasis_token(text: str, keyword: str, index: int) -> str:
    compact = re.sub(r"\s+", "", keyword)

    def _segment() -> str:
        end = index + len(keyword)
        return text[index:end]

    def _pick_in_segment(token: str) -> str:
        segment = _segment()
        return token if token in segment else token

    if compact in {
        "옳지않은",
        "적절하지않은",
        "올바르지않은",
        "가장적절하지않은",
        "가장옳지않은",
        "않은것",
    }:
        return _pick_in_segment("않은")
    if compact in {"올바르지아니한"}:
        return _pick_in_segment("아니한")
    if compact in {"아닌것"}:
        return _pick_in_segment("아닌")
    if compact in {"틀린"}:
        return "틀린"
    if compact in {"잘못된"}:
        return "잘못된"
    if compact in {"부적절한", "부절절"}:
        return _pick_in_segment("부적절한" if "부적절" in keyword else keyword)
    return keyword


def _detect_negative_token_by_rule(text: str) -> str:
    rules = [
        (r"(아닌)\s*것", 1),
        (r"(않은)\s*것", 1),
        (r"(?:옳지|적절하지|올바르지|가장\s*적절하지|가장\s*옳지)\s*(않은)", 1),
        (r"(아니한)", 1),
        (r"(잘못된)", 1),
        (r"(부적절\s*한|부절절)", 1),
        (r"(틀린)", 1),
    ]

    best: tuple[int, str] | None = None
    for pattern, group in rules:
        match = re.search(pattern, text)
        if not match:
            continue
        token = match.group(group)
        if not token:
            continue
        if best is None or match.start(group) < best[0]:
            best = (match.start(group), token)

    if best is None:
        return ""
    return best[1]
