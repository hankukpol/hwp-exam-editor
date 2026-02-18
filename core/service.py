from __future__ import annotations

from pathlib import Path
import re
from typing import Callable, Optional

from .config_manager import ConfigManager
from .detector import extract_question_number
from .exceptions import ParseError
from .generator import OutputGenerator
from .hwp_controller import HwpController
from .models import ExamDocument
from .parser import ExamParser


class ExamProcessingService:
    def __init__(self, config_manager: Optional[ConfigManager] = None) -> None:
        self.config_manager = config_manager or ConfigManager()
        self.last_warning = ""
        self._refresh_dependencies()

    def _refresh_dependencies(self) -> None:
        config = self.config_manager.all()
        self.hwp_controller = HwpController()
        self.parser = ExamParser(config)
        self.generator = OutputGenerator(config)

    def reload_config(self) -> None:
        self.config_manager.reload()
        self._refresh_dependencies()

    def parse_file(self, file_path: str) -> ExamDocument:
        text_blocks = self.hwp_controller.extract_text_blocks(file_path)
        if not text_blocks:
            raise ParseError("문서에서 추출된 텍스트가 없습니다.")

        subject = self._infer_subject(file_path)
        document = self.parser.parse_text_blocks(text_blocks, subject=subject)
        if document.total_count == 0:
            if self._looks_like_answer_key_sheet(text_blocks):
                raise ParseError(
                    "문제 본문이 없는 정답표 형식 파일로 보입니다. 문제지가 포함된 원본 파일을 선택해주세요."
                )
            raise ParseError(self._build_parse_diagnostic_message(text_blocks))
        return document

    def generate_outputs(
        self, document: ExamDocument, source_file: str,
        on_progress: Optional[Callable[[int, str], None]] = None,
    ) -> list[str]:
        output_dir = self.config_manager.get("paths.output_directory", "") or "output"
        stem = Path(source_file).stem
        output_files = self.generator.generate(document, output_dir, stem, on_progress=on_progress)
        self.last_warning = self.generator.last_warning
        return output_files

    def _infer_subject(self, file_path: str) -> str:
        name = Path(file_path).stem
        keywords = ["경찰학", "형법", "형사소송법", "헌법", "행정법"]
        for keyword in keywords:
            if keyword in name:
                return keyword
        return ""

    def _looks_like_answer_key_sheet(self, text_blocks: list[str]) -> bool:
        number_only = sum(1 for line in text_blocks if re.match(r"^\s*\d{1,2}\s*$", line))
        answer_only = sum(
            1
            for line in text_blocks
            if re.match(r"^\s*[①②③④⑤1-5](\s*[\(\[][^\)\]]+[\)\]])?\s*$", line)
        )
        question_like = sum(1 for line in text_blocks if re.search(r"[?？]", line))
        return number_only >= 10 and answer_only >= 10 and question_like == 0

    def _build_parse_diagnostic_message(self, text_blocks: list[str]) -> str:
        non_empty = [line.strip() for line in text_blocks if line and line.strip()]
        probe_lines = non_empty[:20]
        match_count = 0
        for line in probe_lines:
            if extract_question_number(line, self.parser.question_patterns) is not None:
                match_count += 1

        sample = non_empty[:5]
        sample_text = "\n".join(f"- {line[:80]}" for line in sample) if sample else "- (샘플 없음)"
        return (
            "문제 번호를 찾지 못했습니다.\n"
            f"인식 점검: 앞 {len(probe_lines)}줄 중 문제번호 패턴 일치 {match_count}개\n\n"
            "파일 앞부분 샘플:\n"
            f"{sample_text}\n\n"
            "확인사항:\n"
            "- 문제 시작 줄이 '1.' 또는 '01)' 형식인지 확인해 주세요.\n"
            "- 번호 앞에 특수기호/표 컨트롤 문자가 붙어 있지 않은지 확인해 주세요.\n"
            "- 스캔/PDF 변환본이면 원본 HWP로 다시 시도해 주세요."
        )
