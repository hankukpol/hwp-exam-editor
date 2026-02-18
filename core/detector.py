from __future__ import annotations

import re
from typing import Optional


DEFAULT_QUESTION_PATTERNS = [
    r"^\s*(\d{1,2})\s*[\.\)]\s*",
    r"^\s*(0\d)\s*[\.\)]\s*",
    r"^\s*【\s*(\d{1,2})\s*】\s*",
    r"^\s*문\s*(\d{1,2})\s*[\.\)]\s*",
]

DEFAULT_ANSWER_PATTERNS = [
    r"^\s*정답\s+[①②③④⑤]",
    r"^\s*정답[①②③④⑤]",
    r"^\s*정답\s*[:：]\s*[①②③④⑤]",
    r"^\s*정답\s*[:：]\s*\d",
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
) -> str:
    patterns = _compile_patterns(answer_patterns or DEFAULT_ANSWER_PATTERNS)
    fallback_answer_pattern = re.compile(
        r"^\s*정답\s*[:：]?\s*[①②③④⑤1-5](\s*[\(\[][^\)\]]+[\)\]])?\s*$"
    )
    grouped_answer_pattern = re.compile(r"^\s*[①②③④⑤1-5]\s*[\(\[][^\)\]]+[\)\]]\s*$")
    plain_answer_pattern = re.compile(r"^\s*[①②③④⑤1-5]\s*$")
    marker_pattern = re.compile(r"^\s*(정답|해설)\s*[:：]?\s*$")
    answer_count = 0
    for index, block in enumerate(text_blocks):
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
    return "TYPE_A" if answer_count >= threshold else "TYPE_B"


def extract_question_number(
    text: str,
    question_patterns: Optional[list[str]] = None,
) -> Optional[int]:
    patterns = _compile_patterns(question_patterns or DEFAULT_QUESTION_PATTERNS)
    for pattern in patterns:
        match = pattern.match(text)
        if not match:
            continue
        for group in match.groups():
            if group and group.isdigit():
                return int(group)
        digits = re.search(r"\d{1,2}", match.group(0))
        if digits:
            return int(digits.group(0))
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
