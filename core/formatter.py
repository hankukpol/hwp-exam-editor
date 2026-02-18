from __future__ import annotations

import re
import struct
import zlib
from pathlib import Path
from typing import Any

try:
    import olefile
except ImportError:  # pragma: no cover
    olefile = None

try:
    import pythoncom
except ImportError:  # pragma: no cover
    pythoncom = None


TAG_STYLE = 26
TAG_ID_MAPPINGS = 17


def _safe_set_attr(target: Any, name: str, value: Any) -> None:
    # 1) direct COM property set
    try:
        setattr(target, name, value)
        return
    except Exception:
        pass

    # 2) pythoncom direct invoke
    if pythoncom is not None:
        try:
            ole = target._oleobj_
            dispid = ole.GetIDsOfNames(0, name)
            ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
            return
        except Exception:
            pass

    # 3) HSet / SetItem fallback
    try:
        hset = getattr(target, "HSet", None)
        if hset is not None and hasattr(hset, "SetItem"):
            hset.SetItem(name, value)
            return
    except Exception:
        pass
    try:
        if hasattr(target, "SetItem"):
            target.SetItem(name, value)
    except Exception:
        return


class HwpFormatter:
    def __init__(self, config: dict[str, Any]) -> None:
        self.format_config = config.get("format", {})
        self.paragraph_config = config.get("paragraph", {})
        self.page_config = config.get("page", {})
        self.style_config = config.get("style", {})

        self.question_font = str(self.format_config.get("question_font", "중고딕")).strip() or "중고딕"
        self.passage_font = str(self.format_config.get("passage_font", "휴먼명조")).strip() or "휴먼명조"
        self.symbol_font = "바탕"
        self.font_size = float(self.format_config.get("font_size", 9.5))
        self.char_width = int(self.format_config.get("char_width", 95))
        self.char_spacing = int(self.format_config.get("char_spacing", -5))
        self.columns = int(self.format_config.get("columns", 2))
        self.negative_emphasis_bold = bool(self.format_config.get("negative_emphasis_bold", True))

        self.use_styles = bool(self.style_config.get("enabled", True))
        self.question_style = str(self.style_config.get("question_style", "문제")).strip() or "문제"
        self.passage_style = str(self.style_config.get("passage_style", "지문")).strip() or "지문"
        self.choice_style = str(self.style_config.get("choice_style", self.passage_style)).strip() or self.passage_style
        self.sub_items_style = (
            str(self.style_config.get("sub_items_style", self.passage_style)).strip() or self.passage_style
        )
        self.explanation_style = (
            str(self.style_config.get("explanation_style", self.passage_style)).strip() or self.passage_style
        )
        self.style_map_source = str(self.style_config.get("style_map_source", "")).strip()
        if not self.style_map_source:
            self.style_map_source = str(self.style_config.get("template_path", "")).strip()
        if not self.style_map_source:
            self.style_map_source = self._auto_detect_style_source()
        self.style_index_map = self._load_style_index_map(self.style_map_source)
        self._seed_builtin_style_aliases()
        self.style_runtime_warnings: list[str] = []
        self._style_runtime_warning_seen: set[str] = set()
        self._last_applied_style_index: int = -1

    def _seed_builtin_style_aliases(self) -> None:
        builtins = {
            "바탕글": 0,
            "Normal": 0,
            "본문": 1,
            "Body": 1,
        }
        for name, idx in builtins.items():
            self.style_index_map.setdefault(name, idx)
            self.style_index_map.setdefault(name.lower(), idx)

    def _auto_detect_style_source(self) -> str:
        base = Path("tests/sample_files")
        if not base.exists():
            return ""
        candidates = sorted(base.glob("*.hwp"))
        for path in candidates:
            name = path.name
            if ("결과" in name and "샘플" in name) or ("result" in name.lower() and "sample" in name.lower()):
                return str(path)
        return str(candidates[0]) if candidates else ""

    def _load_style_index_map(self, source_hwp: str) -> dict[str, int]:
        if olefile is None:
            return {}
        path = Path(source_hwp).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        if not path.exists():
            return {}

        try:
            with olefile.OleFileIO(str(path)) as ole:
                if not ole.exists("DocInfo") or not ole.exists("FileHeader"):
                    return {}
                header = ole.openstream("FileHeader").read()
                flags = struct.unpack("<I", header[36:40])[0] if len(header) >= 40 else 0
                compressed = bool(flags & 0x01)
                raw = ole.openstream("DocInfo").read()
                data = zlib.decompress(raw, -15) if compressed else raw
        except Exception:
            return {}

        mapping: dict[str, int] = {}
        style_index = 0
        index = 0
        length = len(data)

        while index + 4 <= length:
            header = struct.unpack("<I", data[index:index + 4])[0]
            index += 4

            tag_id = header & 0x3FF
            size = (header >> 20) & 0xFFF
            if size == 0xFFF:
                if index + 4 > length:
                    break
                size = struct.unpack("<I", data[index:index + 4])[0]
                index += 4
            if index + size > length:
                break

            payload = data[index:index + size]
            index += size

            if tag_id != TAG_STYLE:
                continue

            local_name, eng_name = self._parse_style_names(payload)
            if local_name:
                mapping[local_name] = style_index
                mapping[local_name.lower()] = style_index
            if eng_name:
                mapping[eng_name] = style_index
                mapping[eng_name.lower()] = style_index
            style_index += 1

        return mapping

    def _parse_style_names(self, payload: bytes) -> tuple[str, str]:
        # STYLE record: [u16 local_len][local UTF-16LE][u16 eng_len][eng UTF-16LE]...
        if len(payload) < 4:
            return "", ""
        try:
            local_len = struct.unpack("<H", payload[0:2])[0]
            local_end = 2 + local_len * 2
            if local_end > len(payload):
                return "", ""
            local_name = payload[2:local_end].decode("utf-16le", errors="ignore").strip("\x00").strip()
            if local_end + 2 > len(payload):
                return local_name, ""
            eng_len = struct.unpack("<H", payload[local_end:local_end + 2])[0]
            eng_start = local_end + 2
            eng_end = min(len(payload), eng_start + eng_len * 2)
            eng_name = payload[eng_start:eng_end].decode("utf-16le", errors="ignore").strip("\x00").strip()
            return local_name, eng_name
        except Exception:
            return "", ""

    def _resolve_style_index(self, style_name: str) -> int | None:
        if not style_name:
            return None
        text = style_name.strip()
        if not text:
            return None
        if text.isdigit():
            return int(text)
        if text in self.style_index_map:
            return self.style_index_map[text]
        lowered = text.lower()
        if lowered in self.style_index_map:
            return self.style_index_map[lowered]
        return None

    def has_style(self, style_name: str) -> bool:
        text = (style_name or "").strip()
        if not text:
            return False
        # If style-map probing failed, defer validation to runtime apply.
        if not self.style_index_map:
            return True
        return self._resolve_style_index(text) is not None

    def reset_style_runtime_warnings(self) -> None:
        self.style_runtime_warnings = []
        self._style_runtime_warning_seen = set()
        self._last_applied_style_index = -1

    def _record_style_warning(self, message: str) -> None:
        text = message.strip()
        if text and text not in self._style_runtime_warning_seen:
            self._style_runtime_warning_seen.add(text)
            self.style_runtime_warnings.append(text)

    def _get_font_for_style(self, style_name: str) -> str | None:
        """스타일 이름에 해당하는 글꼴을 반환한다."""
        if style_name == self.question_style:
            return self.question_font
        if style_name in (self.passage_style, self.choice_style,
                          self.sub_items_style, self.explanation_style):
            return self.passage_font
        return None

    def apply_style(self, hwp: Any, style_name: str) -> bool:
        """스타일에 해당하는 CharShape/ParaShape를 직접 적용한다.

        HAction.Execute("Style")은 HWP COM 영구 행(hang) 버그가 있어
        사용하지 않는다. 대신 CharShape/ParaShape를 직접 적용하고,
        파일 저장 후 바이너리 후처리로 style_id를 설정한다.
        """
        if not self.use_styles:
            return False
        text = (style_name or "").strip()
        if not text:
            return False
        font_name = self._get_font_for_style(text)
        if font_name is None:
            self._record_style_warning(f"스타일에 대응하는 글꼴을 찾지 못했습니다: {text}")
            return False

        self.apply_paragraph_format(hwp)
        self._apply_char_shape(hwp, font_name=font_name, bold=False, underline=False)
        return True

    def setup_page(self, hwp: Any) -> None:
        try:
            hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
            sec = hwp.HParameterSet.HSecDef
            page = sec.PageDef

            _safe_set_attr(page, "Landscape", 0)
            _safe_set_attr(page, "PaperWidth", hwp.MiliToHwpUnit(210.0))
            _safe_set_attr(page, "PaperHeight", hwp.MiliToHwpUnit(297.0))
            _safe_set_attr(page, "TopMargin", hwp.MiliToHwpUnit(float(self.page_config.get("top_margin", 15.0))))
            _safe_set_attr(page, "BottomMargin", hwp.MiliToHwpUnit(float(self.page_config.get("bottom_margin", 10.0))))
            _safe_set_attr(page, "LeftMargin", hwp.MiliToHwpUnit(float(self.page_config.get("left_margin", 15.0))))
            _safe_set_attr(page, "RightMargin", hwp.MiliToHwpUnit(float(self.page_config.get("right_margin", 15.0))))
            _safe_set_attr(page, "HeaderLen", hwp.MiliToHwpUnit(float(self.page_config.get("header_height", 0.0))))
            _safe_set_attr(page, "FooterLen", hwp.MiliToHwpUnit(float(self.page_config.get("footer_height", 5.0))))
            _safe_set_attr(page, "GutterLen", hwp.MiliToHwpUnit(float(self.page_config.get("gutter", 0.0))))
            _safe_set_attr(page, "GutterType", 0)

            hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)
        except Exception as exc:
            self._record_style_warning(f"페이지 설정 적용 실패: {type(exc).__name__}: {exc}")
            return

    def setup_columns(self, hwp: Any) -> None:
        if self.columns <= 1:
            return
        try:
            hwp.HAction.GetDefault("MultiColumn", hwp.HParameterSet.HColDef.HSet)
            col = hwp.HParameterSet.HColDef
            _safe_set_attr(col, "type", 0)
            _safe_set_attr(col, "Type", 0)
            _safe_set_attr(col, "Count", self.columns)
            _safe_set_attr(col, "SameSize", 1)
            _safe_set_attr(col, "SameGap", hwp.MiliToHwpUnit(8.0))
            hwp.HAction.Execute("MultiColumn", hwp.HParameterSet.HColDef.HSet)
        except Exception as exc:
            self._record_style_warning(f"다단 설정 적용 실패: {type(exc).__name__}: {exc}")
            return

    def apply_question_format(self, hwp: Any, emphasize: bool = False) -> None:
        if not emphasize and self.apply_style(hwp, self.question_style):
            return
        self.apply_paragraph_format(hwp)
        self._apply_char_shape(
            hwp,
            font_name=self.question_font,
            bold=emphasize,
            underline=emphasize,
        )

    def apply_question_inline_char(self, hwp: Any, emphasize: bool = False) -> None:
        bold = self.negative_emphasis_bold if emphasize else False
        self._apply_char_shape(
            hwp,
            font_name=self.question_font,
            bold=bold,
            underline=emphasize,
        )

    def apply_passage_format(self, hwp: Any) -> None:
        if self.apply_style(hwp, self.passage_style):
            return
        self.apply_paragraph_format(hwp)
        self._apply_char_shape(
            hwp,
            font_name=self.passage_font,
            bold=False,
            underline=False,
        )

    def apply_choice_format(self, hwp: Any) -> None:
        if self.apply_style(hwp, self.choice_style):
            return
        self.apply_passage_format(hwp)

    def apply_sub_items_format(self, hwp: Any) -> None:
        if self.apply_style(hwp, self.sub_items_style):
            return
        self.apply_passage_format(hwp)

    def apply_explanation_format(self, hwp: Any) -> None:
        if self.apply_style(hwp, self.explanation_style):
            return
        self.apply_passage_format(hwp)

    def apply_paragraph_format(self, hwp: Any) -> None:
        try:
            hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet)
            ps = hwp.HParameterSet.HParaShape
            indent_value = hwp.PointToHwpUnit(float(self.paragraph_config.get("indent_value", 13.8)))

            _safe_set_attr(ps, "AlignType", 2)
            _safe_set_attr(ps, "AlignmentType", 2)
            _safe_set_attr(ps, "TextAlignment", 2)
            _safe_set_attr(ps, "LeftMargin", 0)
            _safe_set_attr(ps, "RightMargin", 0)
            _safe_set_attr(ps, "IndentType", 2)
            _safe_set_attr(ps, "Indent", indent_value)
            _safe_set_attr(ps, "Indentation", indent_value)
            _safe_set_attr(ps, "LineSpacingType", 0)
            _safe_set_attr(ps, "LineSpacing", int(self.paragraph_config.get("line_spacing", 140)))
            _safe_set_attr(ps, "SpaceBeforePara", 0)
            _safe_set_attr(ps, "SpaceAfterPara", 0)
            _safe_set_attr(ps, "PrevSpacing", 0)
            _safe_set_attr(ps, "NextSpacing", 0)
            _safe_set_attr(ps, "UseGrid", 1 if self.paragraph_config.get("use_grid", True) else 0)
            _safe_set_attr(ps, "SnapToGrid", 1 if self.paragraph_config.get("use_grid", True) else 0)
            hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet)
            hwp.HAction.Run("ParagraphShapeAlignJustify")
        except Exception as exc:
            self._record_style_warning(f"문단 서식 적용 실패: {type(exc).__name__}: {exc}")
            return

    def _apply_char_shape(
        self,
        hwp: Any,
        font_name: str,
        bold: bool,
        underline: bool,
    ) -> None:
        try:
            candidates = self._font_candidates(font_name)
            if not candidates:
                candidates = [font_name]

            for idx, candidate_font in enumerate(candidates):
                hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                cs = hwp.HParameterSet.HCharShape

                for face_attr in (
                    "FaceNameHangul",
                    "FaceNameHanja",
                    "FaceNameJapanese",
                    "FaceNameLatin",
                    "FaceNameOther",
                ):
                    _safe_set_attr(cs, face_attr, candidate_font)
                for face_attr in ("FaceNameSymbol", "FaceNameUser"):
                    _safe_set_attr(cs, face_attr, self.symbol_font)

                _safe_set_attr(cs, "Height", hwp.PointToHwpUnit(self.font_size))
                for attr in (
                    "RatioHangul",
                    "RatioHanja",
                    "RatioJapanese",
                    "RatioLatin",
                    "RatioOther",
                    "RatioSymbol",
                    "RatioUser",
                ):
                    _safe_set_attr(cs, attr, self.char_width)
                for attr in (
                    "SpacingHangul",
                    "SpacingHanja",
                    "SpacingJapanese",
                    "SpacingLatin",
                    "SpacingOther",
                    "SpacingSymbol",
                    "SpacingUser",
                ):
                    _safe_set_attr(cs, attr, self.char_spacing)
                for attr in (
                    "SizeHangul",
                    "SizeHanja",
                    "SizeJapanese",
                    "SizeLatin",
                    "SizeOther",
                    "SizeSymbol",
                    "SizeUser",
                ):
                    _safe_set_attr(cs, attr, 100)
                for attr in (
                    "OffsetHangul",
                    "OffsetHanja",
                    "OffsetJapanese",
                    "OffsetLatin",
                    "OffsetOther",
                    "OffsetSymbol",
                    "OffsetUser",
                ):
                    _safe_set_attr(cs, attr, 0)

                _safe_set_attr(cs, "Bold", 1 if bold else 0)
                _safe_set_attr(cs, "Italic", 0)
                _safe_set_attr(cs, "UseKerning", 1)
                if underline:
                    _safe_set_attr(cs, "UnderlineType", 1)
                    # Use plain-solid underline (Alt+Shift+U behavior).
                    # Some builds interpret non-zero shape as dotted/dashed.
                    _safe_set_attr(cs, "UnderlineShape", 0)
                    _safe_set_attr(cs, "UnderlineColor", 0)
                else:
                    _safe_set_attr(cs, "UnderlineType", 0)
                    _safe_set_attr(cs, "UnderlineShape", 0)

                hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
                if idx == len(candidates) - 1 or self._current_hangul_face_matches(hwp, candidate_font):
                    break
        except Exception as exc:
            self._record_style_warning(f"글자 모양 적용 실패({font_name}): {type(exc).__name__}: {exc}")
            return

    def _font_candidates(self, font_name: str) -> list[str]:
        text = (font_name or "").strip()
        if not text:
            return []
        family_aliases = {
            "중고딕": ["중고딕", "한양중고딕", "HY중고딕"],
            "견고딕": ["견고딕", "한양견고딕", "HY견고딕"],
            "신명조": ["신명조", "한양신명조", "HY신명조"],
        }

        normalized = self._normalize_font_name(text)
        for _family, names in family_aliases.items():
            normalized_names = {self._normalize_font_name(name) for name in names}
            if normalized in normalized_names:
                # Keep caller-provided name first, then other aliases.
                ordered = [text]
                for name in names:
                    if name not in ordered:
                        ordered.append(name)
                return ordered

        return [text]

    @staticmethod
    def _normalize_font_name(font_name: str) -> str:
        text = (font_name or "").strip()
        if text.startswith("한양") and len(text) > 2:
            text = text[2:]
        if text.startswith("HY") and len(text) > 2:
            text = text[2:]
        return text

    def _current_hangul_face_matches(self, hwp: Any, expected_font: str) -> bool:
        try:
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            cs = hwp.HParameterSet.HCharShape
            current = str(getattr(cs, "FaceNameHangul", "")).strip()
            if not current:
                return False
            return self._normalize_font_name(current) == self._normalize_font_name(expected_font)
        except Exception:
            return False

    def _get_current_hangul_face(self, hwp: Any) -> str:
        try:
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            cs = hwp.HParameterSet.HCharShape
            return str(getattr(cs, "FaceNameHangul", "")).strip()
        except Exception:
            return ""

    def _apply_font_face_only(self, hwp: Any, font_name: str) -> None:
        try:
            candidates = self._font_candidates(font_name)
            if not candidates:
                candidates = [font_name]

            for idx, candidate_font in enumerate(candidates):
                hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                cs = hwp.HParameterSet.HCharShape
                for face_attr in (
                    "FaceNameHangul",
                    "FaceNameHanja",
                    "FaceNameJapanese",
                    "FaceNameLatin",
                    "FaceNameOther",
                ):
                    _safe_set_attr(cs, face_attr, candidate_font)
                for face_attr in ("FaceNameSymbol", "FaceNameUser"):
                    _safe_set_attr(cs, face_attr, self.symbol_font)
                hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
                if idx == len(candidates) - 1 or self._current_hangul_face_matches(hwp, candidate_font):
                    break
        except Exception as exc:
            self._record_style_warning(f"글꼴 Face 적용 실패({font_name}): {type(exc).__name__}: {exc}")
            return

    # ── 바이너리 후처리: 저장된 HWP 파일의 style_id 설정 ──────────

    _QUESTION_NUMBER_RE = re.compile(r"^(?:문\s*)?\d{1,3}\s*[\.\)]\s*")

    def post_process_style_ids(self, file_path: Path) -> bool:
        """저장된 HWP 파일을 열어 PARA_HEADER의 style_id를 교정한다.

        HAction.Execute("Style")이 COM 행(hang)을 유발하므로
        직접 서식만 적용한 뒤, 파일 저장 후 이 메서드로 style_id를
        올바르게 설정한다. HWP UI에서 스타일 이름이 표시된다.
        """
        if olefile is None:
            return False
        question_idx = self._resolve_style_index(self.question_style)
        passage_idx = self._resolve_style_index(self.passage_style)
        if question_idx is None and passage_idx is None:
            self._record_style_warning(
                f"문제/지문 스타일을 찾지 못했습니다: {self.question_style}, {self.passage_style}"
            )
            # Fallback: keep inline-emphasis face consistent even when style ids are unavailable.
            try:
                return self.post_process_question_emphasis_faces(file_path)
            except Exception as exc:
                self._record_style_warning(f"강조 폰트 후처리 실패: {type(exc).__name__}: {exc}")
                return False
        if question_idx is None:
            fallback = passage_idx if passage_idx is not None else 0
            self._record_style_warning(
                f"문제 스타일을 찾지 못해 대체 스타일로 처리합니다: {self.question_style}"
            )
            question_idx = fallback
        if passage_idx is None:
            fallback = question_idx if question_idx is not None else 0
            self._record_style_warning(
                f"지문 스타일을 찾지 못해 대체 스타일로 처리합니다: {self.passage_style}"
            )
            passage_idx = fallback

        try:
            # 저장된 파일의 DocInfo에 템플릿 스타일을 이식 (문제/지문 등)
            if not self._transplant_template_styles(file_path):
                self._record_style_warning(
                    "템플릿 스타일 이식에 실패했습니다. 출력 파일의 기본 스타일이 사용됩니다."
                )
            result = self._rewrite_style_ids(file_path, question_idx, passage_idx)
            # Safety pass: some HWP builds keep zero-face inline charshapes.
            try:
                self.post_process_question_emphasis_faces(file_path)
            except Exception as exc:
                self._record_style_warning(f"강조 폰트 안전 후처리 실패: {type(exc).__name__}: {exc}")
            return result
        except Exception as exc:
            self._record_style_warning(f"스타일 후처리 실패: {type(exc).__name__}: {exc}")
            return False

    def _rewrite_style_ids(
        self, file_path: Path, question_idx: int, passage_idx: int,
    ) -> bool:
        ole = olefile.OleFileIO(str(file_path), write_mode=True)
        try:
            hdr = ole.openstream("FileHeader").read()
            flags = struct.unpack("<I", hdr[36:40])[0] if len(hdr) >= 40 else 0
            compressed = bool(flags & 0x01)

            raw = ole.openstream("BodyText/Section0").read()
            body = zlib.decompress(raw, -15) if compressed else raw

            # 1단계: 각 문단의 텍스트 수집 (PARA_TEXT 파싱)
            para_texts = self._collect_para_texts(body)
            style_char_ids = self._collect_style_char_ids(ole, compressed)
            question_char_id = style_char_ids.get(question_idx)
            passage_char_id = style_char_ids.get(passage_idx)
            question_emphasis_char_ids: set[int] = set()

            # 2단계: PARA_HEADER의 style_id 수정
            modified = bytearray(body)
            para_idx = 0
            pos = 0
            changed = False
            question_para_hits = 0
            active_char_id: int | None = None
            active_style_id: int | None = None
            while pos + 4 <= len(modified):
                h = struct.unpack_from("<I", modified, pos)[0]
                tag_id = h & 0x3FF
                size = (h >> 20) & 0xFFF
                data_start = pos + 4
                if size == 0xFFF:
                    if data_start + 4 > len(modified):
                        break
                    size = struct.unpack_from("<I", modified, data_start)[0]
                    data_start += 4
                if data_start + size > len(modified):
                    break

                if tag_id == 66 and size >= 11:  # PARA_HEADER
                    text = para_texts.get(para_idx, "")
                    new_style = self._classify_paragraph_style(
                        text, question_idx, passage_idx,
                    )
                    active_style_id = new_style
                    active_char_id = question_char_id if new_style == question_idx else passage_char_id
                    if new_style == question_idx and text.strip():
                        question_para_hits += 1
                    if modified[data_start + 10] != new_style:
                        modified[data_start + 10] = new_style
                        changed = True
                    para_idx += 1
                elif tag_id == 68 and active_char_id is not None:
                    # PARA_CHAR_SHAPE record: repeated [position(uint32), charshape_id(uint32)].
                    # Preserve inline emphasis runs, but normalize base runs so
                    # question paragraphs keep question font.
                    if active_style_id == question_idx and size >= 8 and (size % 8 == 0):
                        for off in range(0, size, 8):
                            cid = int(struct.unpack_from("<I", modified, data_start + off + 4)[0])
                            if cid != int(active_char_id):
                                question_emphasis_char_ids.add(cid)
                    if self._rewrite_para_char_shape_runs(modified, data_start, size, int(active_char_id)):
                        changed = True

                pos = data_start + size

            if question_idx != passage_idx and question_para_hits == 0:
                self._record_style_warning(
                    "문제 번호 문단을 찾지 못했습니다. 문제 번호 형식 또는 스타일 설정을 확인해 주세요."
                )

            if not changed:
                if (
                    question_char_id is not None
                    and question_emphasis_char_ids
                ):
                    self._rewrite_docinfo_charshape_face_refs(
                        ole,
                        compressed,
                        int(question_char_id),
                        question_emphasis_char_ids,
                    )
                return True  # 수정 불필요

            # 3단계: 재압축 후 스트림 교체 (원본과 동일 크기 보장)
            if compressed:
                new_raw = self._recompress_to_exact_size(
                    bytes(modified), len(raw),
                )
                if new_raw is None:
                    return False  # 크기 맞추기 실패
            else:
                new_raw = bytes(modified)

            ole.write_stream("BodyText/Section0", new_raw)
            if (
                question_char_id is not None
                and question_emphasis_char_ids
                and self._rewrite_docinfo_charshape_face_refs(
                    ole,
                    compressed,
                    int(question_char_id),
                    question_emphasis_char_ids,
                )
            ):
                changed = True
            return True
        finally:
            ole.close()

    def _rewrite_docinfo_charshape_face_refs(
        self,
        ole: "olefile.OleFileIO",
        compressed: bool,
        source_char_id: int,
        target_char_ids: set[int],
    ) -> bool:
        if not target_char_ids or source_char_id in target_char_ids:
            target_char_ids = {cid for cid in target_char_ids if cid != source_char_id}
        if not target_char_ids:
            return False
        if not ole.exists("DocInfo"):
            return False

        raw = ole.openstream("DocInfo").read()
        data = bytearray(zlib.decompress(raw, -15) if compressed else raw)
        changed = self._rewrite_charshape_face_refs_in_docinfo_bytes(
            data,
            source_char_id,
            target_char_ids,
        )
        if not changed:
            return False

        if compressed:
            new_raw = self._recompress_to_exact_size(bytes(data), len(raw))
            if new_raw is None:
                return False
        else:
            new_raw = bytes(data)
        ole.write_stream("DocInfo", new_raw)
        return True

    def post_process_question_emphasis_faces(self, file_path: Path) -> bool:
        if olefile is None:
            return False

        ole = olefile.OleFileIO(str(file_path), write_mode=True)
        try:
            if not ole.exists("FileHeader") or not ole.exists("BodyText/Section0") or not ole.exists("DocInfo"):
                return False

            hdr = ole.openstream("FileHeader").read()
            flags = struct.unpack("<I", hdr[36:40])[0] if len(hdr) >= 40 else 0
            compressed = bool(flags & 0x01)

            raw_body = ole.openstream("BodyText/Section0").read()
            body = zlib.decompress(raw_body, -15) if compressed else raw_body
            para_texts = self._collect_para_texts(body)
            source_to_targets = self._collect_question_charshape_mismatch_ids(body, para_texts)
            if not source_to_targets:
                return False

            raw_doc = ole.openstream("DocInfo").read()
            doc = bytearray(zlib.decompress(raw_doc, -15) if compressed else raw_doc)
            face_map = self._collect_docinfo_charshape_face_bytes(doc)
            if not face_map:
                return False

            zero_face = b"\x00" * 14
            changed = False
            for source_char_id, target_char_ids in source_to_targets.items():
                source_face = face_map.get(source_char_id)
                if source_face is None or source_face == zero_face:
                    continue

                targets = {
                    cid
                    for cid in target_char_ids
                    if cid != source_char_id and face_map.get(cid) == zero_face
                }
                if not targets:
                    continue

                if self._rewrite_charshape_face_refs_in_docinfo_bytes(doc, source_char_id, targets):
                    changed = True
                    for cid in targets:
                        face_map[cid] = source_face

            if not changed:
                return False

            if compressed:
                new_raw = self._recompress_to_exact_size(bytes(doc), len(raw_doc))
                if new_raw is None:
                    return False
            else:
                new_raw = bytes(doc)
            ole.write_stream("DocInfo", new_raw)
            return True
        finally:
            ole.close()

    @staticmethod
    def _collect_question_charshape_mismatch_ids(
        body: bytes,
        para_texts: dict[int, str],
    ) -> dict[int, set[int]]:
        source_to_targets: dict[int, set[int]] = {}
        para_idx = 0
        pos = 0
        active_is_question = False
        while pos + 4 <= len(body):
            h = struct.unpack_from("<I", body, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            data_start = pos + 4
            if size == 0xFFF:
                if data_start + 4 > len(body):
                    break
                size = struct.unpack_from("<I", body, data_start)[0]
                data_start += 4
            if data_start + size > len(body):
                break

            if tag_id == 66:
                text = para_texts.get(para_idx, "")
                candidate = text.strip().lstrip("\ufeff\u200b\u2060\xa0")
                active_is_question = bool(HwpFormatter._QUESTION_NUMBER_RE.match(candidate))
                para_idx += 1
            elif tag_id == 68 and active_is_question and size >= 16 and (size % 8 == 0):
                run_ids: list[int] = []
                for off in range(4, size, 8):
                    run_ids.append(int(struct.unpack_from("<I", body, data_start + off)[0]))
                if len(run_ids) >= 2:
                    counts: dict[int, int] = {}
                    first_pos: dict[int, int] = {}
                    for idx, cid in enumerate(run_ids):
                        counts[cid] = counts.get(cid, 0) + 1
                        first_pos.setdefault(cid, idx)
                    base_char_id = min(
                        counts.keys(),
                        key=lambda cid: (-counts[cid], first_pos[cid]),
                    )
                    for cid in run_ids:
                        if cid == base_char_id:
                            continue
                        source_to_targets.setdefault(base_char_id, set()).add(cid)

            pos = data_start + size
        return source_to_targets

    @staticmethod
    def _collect_docinfo_charshape_face_bytes(data: bytes | bytearray) -> dict[int, bytes]:
        faces: dict[int, bytes] = {}
        charshape_index = 0
        pos = 0
        length = len(data)
        while pos + 4 <= length:
            h = struct.unpack_from("<I", data, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            data_start = pos + 4
            if size == 0xFFF:
                if data_start + 4 > length:
                    break
                size = struct.unpack_from("<I", data, data_start)[0]
                data_start += 4
            if data_start + size > length:
                break

            if tag_id == 21:
                if size >= 14:
                    faces[charshape_index] = bytes(data[data_start:data_start + 14])
                charshape_index += 1

            pos = data_start + size
        return faces

    @staticmethod
    def _rewrite_charshape_face_refs_in_docinfo_bytes(
        data: bytearray,
        source_char_id: int,
        target_char_ids: set[int],
    ) -> bool:
        # CHAR_SHAPE(tag 21): leading 14 bytes are 7 font face ids(uint16).
        source_face_bytes: bytes | None = None
        target_offsets: list[int] = []
        charshape_index = 0
        pos = 0
        length = len(data)
        while pos + 4 <= length:
            h = struct.unpack_from("<I", data, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            data_start = pos + 4
            if size == 0xFFF:
                if data_start + 4 > length:
                    break
                size = struct.unpack_from("<I", data, data_start)[0]
                data_start += 4
            if data_start + size > length:
                break

            if tag_id == 21:
                if size >= 14:
                    if charshape_index == source_char_id:
                        source_face_bytes = bytes(data[data_start:data_start + 14])
                    if charshape_index in target_char_ids:
                        target_offsets.append(data_start)
                charshape_index += 1

            pos = data_start + size

        if source_face_bytes is None or not target_offsets:
            return False
        changed = False
        for offset in target_offsets:
            current = bytes(data[offset:offset + 14])
            if current == source_face_bytes:
                continue
            data[offset:offset + 14] = source_face_bytes
            changed = True
        return changed

    @staticmethod
    def _collect_style_char_ids(
        ole: "olefile.OleFileIO", compressed: bool,
    ) -> dict[int, int]:
        if not ole.exists("DocInfo"):
            return {}
        raw = ole.openstream("DocInfo").read()
        data = zlib.decompress(raw, -15) if compressed else raw

        style_char_ids: dict[int, int] = {}
        style_index = 0
        pos = 0
        while pos + 4 <= len(data):
            h = struct.unpack_from("<I", data, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            data_start = pos + 4
            if size == 0xFFF:
                if data_start + 4 > len(data):
                    break
                size = struct.unpack_from("<I", data, data_start)[0]
                data_start += 4
            if data_start + size > len(data):
                break

            if tag_id == TAG_STYLE:
                payload = data[data_start:data_start + size]
                char_id = HwpFormatter._parse_style_char_id(payload)
                if char_id is not None:
                    style_char_ids[style_index] = char_id
                style_index += 1

            pos = data_start + size
        return style_char_ids

    @staticmethod
    def _parse_style_char_id(payload: bytes) -> int | None:
        # STYLE record tail includes ParaShapeId(uint16), CharShapeId(uint16).
        # Layout: [u16 local_len][local][u16 eng_len][eng][tail...]
        if len(payload) < 4:
            return None
        try:
            local_len = struct.unpack_from("<H", payload, 0)[0]
            local_end = 2 + local_len * 2
            if local_end + 2 > len(payload):
                return None
            eng_len = struct.unpack_from("<H", payload, local_end)[0]
            eng_end = local_end + 2 + eng_len * 2
            tail = payload[eng_end:]
            if len(tail) < 8:
                return None
            return int(struct.unpack_from("<H", tail, 6)[0])
        except Exception:
            return None

    @staticmethod
    def _recompress_to_exact_size(data: bytes, target_size: int) -> bytes | None:
        """DEFLATE 재압축 후 target_size에 맞게 패딩한다.

        DEFLATE 디코더는 스트림 종료 마커 이후 바이트를 무시하므로
        null 패딩이 안전하다.
        """
        for level in (9, 6, 3, 1):
            compressor = zlib.compressobj(level=level, wbits=-15)
            compressed = compressor.compress(data) + compressor.flush()
            if len(compressed) <= target_size:
                return compressed.ljust(target_size, b"\x00")
        # 모든 레벨에서 원본보다 큼 → stored(level=0) 시도
        compressor = zlib.compressobj(level=0, wbits=-15)
        compressed = compressor.compress(data) + compressor.flush()
        if len(compressed) <= target_size:
            return compressed.ljust(target_size, b"\x00")
        return None

    @staticmethod
    def _collect_para_texts(body: bytes) -> dict[int, str]:
        """바이너리에서 각 문단의 텍스트를 추출한다.

        인덱싱은 0-based: 첫 번째 PARA_HEADER → para_idx=0.
        _rewrite_style_ids()와 동일한 순서를 사용한다.
        """
        texts: dict[int, str] = {}
        para_idx = -1
        pos = 0
        while pos + 4 <= len(body):
            h = struct.unpack_from("<I", body, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            pos += 4
            if size == 0xFFF:
                if pos + 4 > len(body):
                    break
                size = struct.unpack_from("<I", body, pos)[0]
                pos += 4
            if pos + size > len(body):
                break
            payload = body[pos:pos + size]

            if tag_id == 66:  # PARA_HEADER
                para_idx += 1

            if tag_id == 67 and len(payload) >= 2 and para_idx >= 0:  # PARA_TEXT
                chars: list[str] = []
                i = 0
                while i < len(payload) - 1:
                    code = struct.unpack_from("<H", payload, i)[0]
                    i += 2
                    if code < 32:
                        if 1 <= code <= 23:
                            i += 12
                        elif 24 <= code <= 31:
                            i += 8
                    else:
                        chars.append(chr(code))
                texts[para_idx] = "".join(chars)

            pos += size
        return texts

    # ── 템플릿 스타일 이식 ─────────────────────────────────────

    @staticmethod
    def _parse_record_positions(data: bytes) -> list[tuple[int, int, int]]:
        """HWP 태그 레코드 스트림의 각 레코드 위치를 파싱한다.

        Returns: [(record_start, record_end, tag_id), ...]
        """
        records: list[tuple[int, int, int]] = []
        pos = 0
        length = len(data)
        while pos + 4 <= length:
            rec_start = pos
            h = struct.unpack_from("<I", data, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            pos += 4
            if size == 0xFFF:
                if pos + 4 > length:
                    break
                size = struct.unpack_from("<I", data, pos)[0]
                pos += 4
            if pos + size > length:
                break
            pos += size
            records.append((rec_start, pos, tag_id))
        return records

    def _extract_style_records_bytes(self, data: bytes) -> tuple[bytes, int]:
        """DocInfo 바이너리에서 TAG_STYLE 레코드들의 원시 바이트와 개수를 추출한다."""
        records = self._parse_record_positions(data)
        chunks: list[bytes] = []
        count = 0
        for start, end, tag_id in records:
            if tag_id == TAG_STYLE:
                chunks.append(data[start:end])
                count += 1
        return b"".join(chunks), count

    def _transplant_template_styles(self, file_path: Path) -> bool:
        """출력 파일의 DocInfo에 템플릿의 스타일 정의를 이식한다.

        HWP COM이 저장 시 템플릿의 커스텀 스타일(문제, 지문 등)을
        기본 스타일로 대체하는 문제를 해결한다.
        템플릿의 TAG_STYLE 레코드로 교체하고 TAG_ID_MAPPINGS의
        styleCount를 갱신한다.
        """
        if olefile is None:
            return False
        template_path = Path(self.style_map_source).expanduser()
        if not template_path.is_absolute():
            template_path = Path.cwd() / template_path
        if not template_path.exists():
            return False

        try:
            # 1) 템플릿 DocInfo에서 스타일 레코드 추출
            with olefile.OleFileIO(str(template_path)) as t_ole:
                if not t_ole.exists("DocInfo") or not t_ole.exists("FileHeader"):
                    return False
                t_hdr = t_ole.openstream("FileHeader").read()
                t_flags = struct.unpack("<I", t_hdr[36:40])[0] if len(t_hdr) >= 40 else 0
                t_compressed = bool(t_flags & 0x01)
                t_raw = t_ole.openstream("DocInfo").read()
                t_docinfo = zlib.decompress(t_raw, -15) if t_compressed else t_raw

            t_style_bytes, t_style_count = self._extract_style_records_bytes(t_docinfo)
            if t_style_count == 0:
                return False

            # 2) 출력 파일 DocInfo 수정
            ole = olefile.OleFileIO(str(file_path), write_mode=True)
            try:
                o_hdr = ole.openstream("FileHeader").read()
                o_flags = struct.unpack("<I", o_hdr[36:40])[0] if len(o_hdr) >= 40 else 0
                o_compressed = bool(o_flags & 0x01)
                o_raw = ole.openstream("DocInfo").read()
                o_docinfo = zlib.decompress(o_raw, -15) if o_compressed else o_raw

                o_records = self._parse_record_positions(o_docinfo)
                o_style_ranges = [
                    (start, end) for start, end, tag_id in o_records
                    if tag_id == TAG_STYLE
                ]
                if not o_style_ranges:
                    return False
                o_style_count = len(o_style_ranges)

                first_start = o_style_ranges[0][0]
                last_end = o_style_ranges[-1][1]

                # 기존 스타일 영역을 템플릿의 스타일로 교체
                new_docinfo = bytearray()
                new_docinfo.extend(o_docinfo[:first_start])
                new_docinfo.extend(t_style_bytes)
                new_docinfo.extend(o_docinfo[last_end:])

                # TAG_ID_MAPPINGS의 styleCount 갱신
                if t_style_count != o_style_count:
                    self._update_id_mappings_style_count(
                        new_docinfo, o_style_count, t_style_count,
                    )

                # 재압축 후 저장
                if o_compressed:
                    new_raw = self._recompress_to_exact_size(
                        bytes(new_docinfo), len(o_raw),
                    )
                    if new_raw is None:
                        return False
                else:
                    new_raw = bytes(new_docinfo)

                ole.write_stream("DocInfo", new_raw)
                return True
            finally:
                ole.close()
        except Exception:
            return False

    @staticmethod
    def _update_id_mappings_style_count(
        data: bytearray, old_count: int, new_count: int,
    ) -> None:
        """TAG_ID_MAPPINGS 레코드의 스타일 수 필드를 갱신한다."""
        pos = 0
        length = len(data)
        while pos + 4 <= length:
            h = struct.unpack_from("<I", data, pos)[0]
            tag_id = h & 0x3FF
            size = (h >> 20) & 0xFFF
            data_start = pos + 4
            if size == 0xFFF:
                if data_start + 4 > length:
                    break
                size = struct.unpack_from("<I", data, data_start)[0]
                data_start += 4
            if data_start + size > length:
                break

            if tag_id == TAG_ID_MAPPINGS:
                # styleCount는 ID_MAPPINGS 페이로드의 15번째 UINT32 (offset 56)
                offset = data_start + 56
                if offset + 4 <= data_start + size:
                    current = struct.unpack_from("<I", data, offset)[0]
                    if current == old_count:
                        struct.pack_into("<I", data, offset, new_count)
                return

            pos = data_start + size

    # ── 문단 분류 ─────────────────────────────────────────────

    def _classify_paragraph_style(
        self, text: str, question_idx: int, passage_idx: int,
    ) -> int:
        """문단 텍스트를 기반으로 적용할 style_id를 결정한다."""
        stripped = text.strip()
        if not stripped:
            return 0  # 빈 문단 → 바탕글
        candidate = stripped.lstrip("\ufeff\u200b\u2060\xa0")
        if self._QUESTION_NUMBER_RE.match(candidate):
            return question_idx
        return passage_idx

    @staticmethod
    def _para_char_shape_run_count(payload_size: int) -> int:
        # payload layout: N * [position(uint32), charshape_id(uint32)]
        if payload_size < 8:
            return 0
        if payload_size % 8 != 0:
            return 0
        return payload_size // 8

    @staticmethod
    def _rewrite_para_char_shape_runs(
        buffer: bytearray, data_start: int, payload_size: int, target_char_id: int,
    ) -> bool:
        run_count = HwpFormatter._para_char_shape_run_count(payload_size)
        if run_count <= 0:
            return False

        offsets: list[int] = []
        run_ids: list[int] = []
        for off in range(4, payload_size, 8):
            offsets.append(off)
            run_ids.append(int(struct.unpack_from("<I", buffer, data_start + off)[0]))
        if not run_ids:
            return False

        changed = False
        if run_count == 1:
            current_char_id = run_ids[0]
            if current_char_id != target_char_id:
                struct.pack_into("<I", buffer, data_start + offsets[0], int(target_char_id))
                changed = True
            return changed

        counts: dict[int, int] = {}
        first_pos: dict[int, int] = {}
        for idx, char_id in enumerate(run_ids):
            counts[char_id] = counts.get(char_id, 0) + 1
            first_pos.setdefault(char_id, idx)

        base_char_id = min(
            counts.keys(),
            key=lambda cid: (-counts[cid], first_pos[cid]),
        )

        for idx, current_char_id in enumerate(run_ids):
            if current_char_id != base_char_id:
                continue
            if current_char_id == target_char_id:
                continue
            struct.pack_into("<I", buffer, data_start + offsets[idx], int(target_char_id))
            changed = True
        return changed
