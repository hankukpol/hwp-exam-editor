from __future__ import annotations

import os
import re
import struct
import zlib
from pathlib import Path

from .com_utils import ensure_clean_dispatch as _ensure_clean_dispatch
from .exceptions import HwpNotAvailableError, UnsupportedFileError

try:
    import win32com.client as win32
except ImportError:  # pragma: no cover - depends on local Windows setup
    win32 = None

try:
    import olefile
except ImportError:  # pragma: no cover - depends on local environment
    olefile = None


TAG_PARA_TEXT = 67


class HwpController:
    def extract_text_blocks(self, file_path: str) -> list[str]:
        path = Path(file_path)
        suffix = path.suffix.lower()
        if suffix == ".txt":
            return self._read_text_file(path)
        if suffix != ".hwp":
            raise UnsupportedFileError("HWP 파일(.hwp)만 지원합니다.")
        return self._extract_from_hwp(path)

    def _read_text_file(self, path: Path) -> list[str]:
        with path.open("r", encoding="utf-8") as file:
            return [line.rstrip("\n") for line in file]

    def _extract_from_hwp(self, path: Path) -> list[str]:
        ole_error: Exception | None = None
        try:
            blocks = self._extract_from_hwp_ole(path)
            if blocks:
                return blocks
        except Exception as exc:
            ole_error = exc

        # COM 읽기는 환경에 따라 무한 대기할 수 있어 명시적으로 켠 경우만 사용
        if os.environ.get("HWP_USE_COM_READ", "0") == "1":
            try:
                return self._extract_from_hwp_com(path)
            except Exception as exc:
                ole_message = f" OLE 오류: {ole_error}" if ole_error else ""
                raise HwpNotAvailableError(
                    f"HWP 텍스트 추출에 실패했습니다.{ole_message} COM 오류: {exc}"
                ) from exc

        if ole_error:
            raise HwpNotAvailableError(f"HWP 텍스트 추출에 실패했습니다: {ole_error}")
        raise HwpNotAvailableError(
            "HWP 텍스트 추출에 실패했습니다. COM 읽기는 HWP_USE_COM_READ=1 환경변수로 활성화할 수 있습니다."
        )

    def _extract_from_hwp_ole(self, path: Path) -> list[str]:
        if olefile is None:
            raise HwpNotAvailableError("olefile 패키지가 없어 HWP OLE 추출을 수행할 수 없습니다.")

        with olefile.OleFileIO(str(path)) as ole:
            if not ole.exists("FileHeader"):
                raise HwpNotAvailableError("HWP FileHeader 스트림이 없습니다.")

            if self._is_hwp_encrypted(ole):
                raise HwpNotAvailableError("파일이 암호로 보호되어 있습니다.")

            compressed = self._is_hwp_compressed(ole)
            sections = self._list_body_sections(ole)
            if not sections:
                raise HwpNotAvailableError("BodyText 섹션을 찾지 못했습니다.")

            lines: list[str] = []
            for section in sections:
                data = ole.openstream(section).read()
                lines.extend(self._extract_para_text_lines(data, compressed))

            cleaned = [line for line in (self._clean_line(line) for line in lines) if line]
            if not cleaned:
                raise HwpNotAvailableError("BodyText에서 추출된 텍스트가 없습니다.")
            return cleaned

    def _is_hwp_encrypted(self, ole: "olefile.OleFileIO") -> bool:
        header = ole.openstream("FileHeader").read()
        if len(header) < 40:
            return False
        flags = struct.unpack("<I", header[36:40])[0]
        return bool(flags & 0x02)

    def _is_hwp_compressed(self, ole: "olefile.OleFileIO") -> bool:
        header = ole.openstream("FileHeader").read()
        if len(header) < 40:
            return False
        flags = struct.unpack("<I", header[36:40])[0]
        return bool(flags & 0x01)

    def _list_body_sections(self, ole: "olefile.OleFileIO") -> list[str]:
        sections: list[tuple[int, list[str]]] = []
        for entry in ole.listdir(streams=True, storages=False):
            if len(entry) >= 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
                try:
                    index = int(entry[1].replace("Section", ""))
                except ValueError:
                    index = 0
                sections.append((index, entry))
        sections.sort(key=lambda item: item[0])
        return ["/".join(entry) for _, entry in sections]

    def _extract_para_text_lines(self, stream_data: bytes, compressed: bool) -> list[str]:
        if compressed:
            stream_data = zlib.decompress(stream_data, -15)

        output: list[str] = []
        index = 0
        length = len(stream_data)
        while index + 4 <= length:
            header = struct.unpack("<I", stream_data[index:index + 4])[0]
            index += 4

            tag_id = header & 0x3FF
            size = (header >> 20) & 0xFFF
            if size == 0xFFF:
                if index + 4 > length:
                    break
                size = struct.unpack("<I", stream_data[index:index + 4])[0]
                index += 4

            if index + size > length:
                break

            payload = stream_data[index:index + size]
            index += size

            if tag_id != TAG_PARA_TEXT or not payload:
                continue

            text = payload.decode("utf-16le", errors="ignore")
            for line in text.splitlines():
                if line.strip():
                    output.append(line)
        return output

    def _clean_line(self, line: str) -> str:
        text = line.replace("\x00", "")
        # Keep tab separators so table-like rows are not flattened.
        text = re.sub(r"[\u0001-\u0008\u000B-\u001F]", " ", text)
        text = text.replace("⋅", "·")
        text = text.replace("･", "·")

        # OLE 디코딩 과정에서 섞이는 비정상 문자 제거
        text = re.sub(r"[\u0100-\u024F\u0300-\u036F\u0400-\u052F\u0590-\u05FF]", "", text)
        text = re.sub(r"[\uFFF0-\uFFFF]", "", text)

        if "\t" in text:
            chunks = [re.sub(r"[ ]+", " ", chunk).strip() for chunk in text.split("\t")]
            text = "\t".join(chunks)
            text = re.sub(r"\t{2,}", "\t", text).strip()
        else:
            text = re.sub(r"\s+", " ", text).strip()
        if not text:
            return ""

        hanja_subject = bool(re.match(r"^[甲乙丙丁戊己庚辛壬癸]\s*[은는이가을를의]", text))
        hanja_marker_only = bool(re.fullmatch(r"[甲乙丙丁戊己庚辛壬癸]", text))

        # 한글/영문/숫자/문항 기호가 나오기 전의 깨진 선행 문자를 제거
        # (단, 사례문 시작의 甲/乙 계열 라벨은 보존한다.)
        if not hanja_subject and not hanja_marker_only:
            text = re.sub(
                r"^[^0-9A-Za-z가-힣①②③④⑤㉠㉡㉢㉣㉤㉥★<\[\(]+",
                "",
                text,
            )
        if not text:
            return ""

        # 추출 과정에서 반복되는 무의미 1~2글자 노이즈 제거
        if hanja_marker_only:
            return text
        if len(text) <= 2 and not re.search(
            r"[0-9A-Za-z가-힣①②③④⑤]",
            text,
        ):
            return ""
        if re.fullmatch(r"[-.=·•▪▫◦※*]+", text):
            return ""
        return text

    def _extract_from_hwp_com(self, path: Path) -> list[str]:
        if win32 is None:
            raise HwpNotAvailableError(
                "pywin32가 설치되어 있지 않아 HWP 텍스트 추출을 수행할 수 없습니다."
            )

        hwp = None
        lines: list[str] = []
        try:
            hwp = _ensure_clean_dispatch("HWPFrame.HwpObject")
            hwp.XHwpWindows.Item(0).Visible = False
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass
            if not hwp.Open(str(path)):
                raise HwpNotAvailableError("HWP 파일을 열지 못했습니다.")
            hwp.InitScan()

            while True:
                state, text = hwp.GetText()
                if state == 0:
                    break
                if text:
                    lines.extend(text.replace("\r", "").split("\n"))
            hwp.ReleaseScan()
            hwp.Clear(3)
            return [line.strip() for line in lines if line and line.strip()]
        except Exception as exc:  # pragma: no cover - depends on COM state
            raise HwpNotAvailableError(
                f"HWP 텍스트 추출에 실패했습니다: {exc}"
            ) from exc
        finally:
            if hwp is not None:
                try:
                    hwp.Quit()
                except Exception:
                    pass
                try:
                    del hwp
                except Exception:
                    pass
