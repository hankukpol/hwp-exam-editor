from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ExamQuestion:
    number: int
    question_text: str
    choices: list[str] = field(default_factory=list)
    sub_items: list[str] = field(default_factory=list)
    has_table: bool = False
    has_negative: bool = False
    negative_keyword: str = ""
    answer: Optional[str] = None
    explanation: Optional[str] = None


@dataclass
class ExamDocument:
    file_type: str
    subject: str
    questions: list[ExamQuestion] = field(default_factory=list)
    total_count: int = 0

    def refresh_total_count(self) -> None:
        self.total_count = len(self.questions)


@dataclass
class OutputConfig:
    output_directory: str
