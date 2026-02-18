# 개발 문서: 시험별 멀티 프리셋 확장 가이드

> **문서 버전**: v1.0
> **작성일**: 2026-02-18
> **대상 프로젝트**: HWP 모의고사 자동 편집 프로그램 v2.0
> **목적**: 시험마다 다른 편집 양식을 프리셋으로 관리하고, 학원 전용 한글 스타일과 연동하는 확장 기능 개발 가이드

---

## 1. 배경 및 목적

### 1.1 현재 상태

현재 프로그램은 **단일 편집 설정**만 지원한다.

- `config/default_config.json` → `config/user_config.json` 2단 병합 구조
- 글자 크기(9.5pt), 줄간격(140%), 장평(95%), 자간(-5%) 등이 고정
- 용지(A4 210×297mm), 여백, 단수(2단) 등도 단일 값
- 시험이 바뀌면 매번 설정 창에서 수동으로 값을 고쳐야 함

### 1.2 문제점

학원에서 운영하는 시험별로 편집 양식이 다르다:

| 시험 유형 | 용지 | 단수 | 글자 크기 | 줄간격 | 비고 |
|-----------|------|------|-----------|--------|------|
| 아침모의고사 | A4 | 2단 | 9.5pt | 140% | 현재 기본값 |
| 전국연합 모의고사 | B4 | 3단 | 10.0pt | 135% | 별도 양식 필요 |
| 과목별 실전모의 | A4 | 2단 | 9.0pt | 145% | 여백/내어쓰기 다름 |

매번 수동 변경은 실수 가능성이 높고, 비개발자 직원이 사용하기 어렵다.

### 1.3 실제 워크플로우

프로그램의 출력물은 **최종 시험지가 아니다.**
관리자가 생성된 내용을 복사하여 **원래 시험지 양식 HWP에 붙여넣기**하는 방식이다.

```
프로그램 출력 (서식이 적용된 문제/해설)
       │
       │  관리자가 내용 복사 (Ctrl+C)
       ▼
원래 시험지 양식에 붙여넣기 (Ctrl+V)
       │
       ▼
최종 시험지 완성 (헤더, 용지, 단 배치는 원래 양식 그대로)
```

따라서 **용지 크기(A4/B4), 여백, 단수는 프리셋에서 중요하지 않다.**
붙여넣기 시 따라가는 것은 **글자 서식(글꼴, 크기, 장평, 자간)과 문단 서식(줄간격, 내어쓰기)**이므로
프리셋은 이 서식 값을 정확히 맞추는 것이 핵심이다.

### 1.4 목표

**"시험 이름을 선택하면 그에 맞는 글자/문단 서식이 자동 적용된다"**

- 관리자가 시험별 프리셋을 미리 정의
- 각 프리셋은 학원 전용 한글(HWP) 스타일 템플릿과 연동
- 사용자는 드롭다운에서 시험만 선택하면 끝
- 출력물은 A4 고정이어도 무관 (최종 양식에 복사-붙여넣기할 것이므로)

---

## 2. 핵심 개념: 프리셋과 한글 스타일의 관계

### 2.1 한글 스타일이란

한글(HWP)에서 **스타일**은 글꼴, 크기, 줄간격, 문단 설정 등을 하나의 이름으로 묶어둔 것이다.
예를 들어 "문제" 스타일에는 중고딕 9.5pt, 내어쓰기 13.8pt가 정의되어 있고,
"지문" 스타일에는 휴먼명조 9.5pt가 정의되어 있다.

### 2.2 학원 전용 스타일 만들기

학원에서 시험별로 HWP 스타일 세트를 만드는 과정:

```
1. 한글(HWP)을 연다
2. [서식] → [스타일] → [스타일 만들기]
3. "문제" 스타일 정의 (글꼴: 중고딕, 크기: 9.5pt, 내어쓰기: 13.8pt ...)
4. "지문" 스타일 정의 (글꼴: 휴먼명조, 크기: 9.5pt ...)
5. 이 파일을 "아침모의고사_템플릿.hwp"로 저장
```

시험마다 이런 템플릿 HWP 파일을 하나씩 만든다.

### 2.3 프리셋 = 세팅값 + 템플릿 연결

```
프리셋 "아침모의고사"
├── 글자 서식: 중고딕/휴먼명조, 9.5pt, 장평 95%, 자간 -5%
├── 문단 서식: 줄간격 140%, 내어쓰기 13.8pt
└── 템플릿: 아침모의고사_템플릿.hwp (문제/지문 스타일 포함)

프리셋 "전국연합 사회"
├── 글자 서식: 중고딕/휴먼명조, 10.0pt, 장평 97%, 자간 -3%
├── 문단 서식: 줄간격 135%, 내어쓰기 15.0pt
└── 템플릿: 전국연합_템플릿.hwp (문제/지문 스타일 포함)
```

---

## 3. 전체 동작 흐름

### 3.1 관리자 작업 (1회성)

```
┌──────────────────────────────────────────────────────────┐
│ [관리자] 한글(HWP)에서 시험별 스타일 정의                   │
│  → "문제", "지문" 스타일이 포함된 템플릿.hwp 저장            │
│                                                          │
│ [관리자] 프로그램 설정에서 프리셋 생성                       │
│  → 프리셋 이름: "전국연합 사회"                             │
│  → 글자 크기: 10.0pt / 줄간격: 135% / 장평: 97% / 자간: -3% │
│  → 템플릿 경로: templates/전국연합_템플릿.hwp 지정           │
│  → 저장 → config/presets/전국연합_사회.json 생성             │
└──────────────────────────────────────────────────────────┘
```

### 3.2 직원 사용 (매일 반복)

```
┌──────────────────────────────────────────────────────────┐
│ 1. 프로그램 실행                                          │
│ 2. 시험 유형 선택: [전국연합 사회 ▼]   ← 드롭다운 클릭      │
│ 3. HWP 파일 드래그 앤 드롭                                 │
│ 4. 미리보기 확인                                          │
│ 5. [문제지/해설지 생성] 클릭                                │
│ 6. → 전국연합 사회 서식(10pt, 135%, 장평97%)으로 자동 출력    │
└──────────────────────────────────────────────────────────┘
```

### 3.3 내부 처리 흐름 (프로그램 관점)

```
사용자가 프리셋 선택
       │
       ▼
ConfigManager.load_with_preset("전국연합_사회")
       │
       ├─ 1단계: default_config.json 로드 (기본값)
       ├─ 2단계: presets/전국연합_사회.json 병합 (시험별 세팅 덮어쓰기)
       └─ 3단계: user_config.json 병합 (공통 경로/모듈 유지)
       │
       ▼
최종 config 생성
       │
       ├─ format.font_size = 10.0       (← 프리셋에서 옴)
       ├─ format.char_width = 97        (← 프리셋에서 옴)
       ├─ paragraph.line_spacing = 135  (← 프리셋에서 옴)
       ├─ style.template_path = "templates/전국연합_템플릿.hwp"  (← 프리셋)
       └─ style.module_dll_path = "C:/.../Module.dll"  (← user_config 유지)
       │
       ▼
HwpFormatter(config) 초기화
       │
       ├─ 템플릿에서 스타일 인덱스 맵 로드 ("문제"→2, "지문"→3)
       ├─ setup_page(): A4 고정 (복사-붙여넣기 워크플로우이므로 무관)
       └─ 본문 포맷팅: 10pt, 장평97%, 135% 줄간격으로 적용
       │
       ▼
OutputGenerator.generate()
       │
       ├─ 템플릿.hwp 복사 → 출력 파일
       ├─ COM으로 내용 삽입 (프리셋 세팅으로 포맷팅)
       ├─ 저장
       └─ post_process_style_ids() (템플릿 스타일 ID 적용)
       │
       ▼
서식 적용된 문제지/해설지 HWP 출력
       │
       ▼
관리자가 내용 복사 → 원래 시험지 양식에 붙여넣기 → 최종 완성
```

---

## 4. 프리셋 파일 규격

### 4.1 저장 위치

```
config/presets/         ← 프리셋 JSON 파일 저장 디렉토리
config/templates/       ← 프리셋과 연결되는 HWP 템플릿 저장 디렉토리
```

### 4.2 프리셋 JSON 구조

프리셋 파일에는 **기본값과 다른 값만** 기록한다.
기록하지 않은 항목은 `default_config.json`의 기본값이 자동 적용된다.

```json
{
  "preset_name": "아침모의고사",
  "preset_description": "한국경찰학원 아침모의고사 기본 양식",

  "format": {
    "question_font": "중고딕",
    "passage_font": "휴먼명조",
    "font_size": 9.5,
    "char_width": 95,
    "char_spacing": -5
  },

  "paragraph": {
    "line_spacing": 140,
    "indent_value": 13.8
  },

  "style": {
    "template_path": "config/templates/아침모의고사_템플릿.hwp",
    "question_style": "문제",
    "passage_style": "지문"
  }
}
```

### 4.3 프리셋이 제어하는 항목 목록

> **참고**: 출력물은 원래 시험지 양식에 복사-붙여넣기할 중간 결과물이다.
> 붙여넣기 시 따라가는 것은 **글자/문단 서식**이므로, 용지/단 설정은 보조적이다.

| 우선순위 | 분류 | 항목 | 키 경로 | 설명 |
|----------|------|------|---------|------|
| **핵심** | 글자 | 문제 폰트 | `format.question_font` | 문제 번호/본문 글꼴 |
| **핵심** | | 지문 폰트 | `format.passage_font` | 보기/지문 글꼴 |
| **핵심** | | 글자 크기 | `format.font_size` | pt 단위 |
| **핵심** | | 장평 | `format.char_width` | % (100=기본) |
| **핵심** | | 자간 | `format.char_spacing` | % (-는 좁게) |
| **핵심** | 문단 | 줄간격 | `paragraph.line_spacing` | % |
| **핵심** | | 내어쓰기 | `paragraph.indent_value` | pt 단위 |
| **핵심** | | 줄 격자 사용 | `paragraph.use_grid` | true/false |
| **핵심** | 스타일 | 템플릿 경로 | `style.template_path` | 스타일 정의 HWP |
| **핵심** | | 문제/지문 스타일명 | `style.*_style` | 템플릿 내 스타일 이름 |
| 보조 | 용지 | 용지 크기 | `page.paper_type` | 복사-붙여넣기 워크플로우에서는 무관 |
| 보조 | | 여백 | `page.*_margin` | 동일 |
| 보조 | 단 | 단 수 / 단 간격 | `format.columns` | 동일 |
| **스타일** | 템플릿 경로 | `style.template_path` | HWP 파일 경로 |
| | 문제 스타일명 | `style.question_style` | 템플릿 내 스타일 이름 |
| | 지문 스타일명 | `style.passage_style` | 템플릿 내 스타일 이름 |

---

## 5. 설정 병합 우선순위

### 5.1 3단 병합 구조

```
[1] default_config.json   공통 기본값 (모든 프리셋의 base)
         ↓ deep merge
[2] presets/XXX.json       시험별 세팅 (차이값만)
         ↓ deep merge
[3] user_config.json       사용자 공통 설정 (출력 폴더, 보안 모듈 DLL 등)
         ↓
    === 최종 config ===
```

### 5.2 병합 예시

```
default_config.json:   format.font_size = 9.5
전국연합_사회.json:    format.font_size = 10.0   ← 프리셋이 덮어씀
user_config.json:      (font_size 없음)          ← 변경 없음
─────────────────────────────────────────────────
최종 결과:             format.font_size = 10.0
```

```
default_config.json:   style.module_dll_path = ""
전국연합_사회.json:    (module_dll_path 없음)     ← 프리셋에 없음
user_config.json:      style.module_dll_path = "C:/.../Module.dll"  ← 사용자 값 유지
─────────────────────────────────────────────────
최종 결과:             style.module_dll_path = "C:/.../Module.dll"
```

### 5.3 user_config.json의 역할 변경

현재 `user_config.json`에는 글자 크기, 줄간격 등 편집 값이 저장되어 있다.
프리셋 도입 후에는 **편집 값은 프리셋이 담당**하고, `user_config.json`에는 시험과 무관한 공통 설정만 남긴다:

```json
// user_config.json (프리셋 도입 후)
{
  "active_preset": "아침모의고사",
  "paths": {
    "output_directory": "D:/모의고사_출력"
  },
  "style": {
    "enabled": true,
    "module_dll_path": "C:/Program Files (x86)/Hnc/HOffice9/Bin/FilePathCheckerModuleExample.dll"
  }
}
```

`active_preset` 키로 마지막에 사용한 프리셋을 기억하여, 프로그램 재시작 시 자동 선택한다.

---

## 6. 한글(HWP) 스타일 템플릿 연동 상세

### 6.1 템플릿 HWP 파일의 역할

템플릿 파일은 **스타일 정의를 담는 그릇**이다. 내용(본문)은 비어 있어도 된다.
중요한 것은 "문제", "지문" 등의 **스타일이 정의되어 있는 것**이다.

```
아침모의고사_템플릿.hwp
├── [DocInfo]
│   ├── TAG_STYLE: "바탕글" (index 0)
│   ├── TAG_STYLE: "본문"   (index 1)
│   ├── TAG_STYLE: "문제"   (index 2)  ← 중고딕 9.5pt 내어쓰기 13.8pt
│   └── TAG_STYLE: "지문"   (index 3)  ← 휴먼명조 9.5pt
└── [BodyText/Section0]
    └── (비어 있거나 샘플 문장)
```

### 6.2 프로그램이 템플릿을 사용하는 과정

```
1. [초기화] 프리셋의 style.template_path에서 템플릿 경로를 읽음
     │
2. [스타일 맵 로드] 템플릿을 OLE로 열어 DocInfo의 TAG_STYLE 레코드 파싱
     │   → style_index_map = {"문제": 2, "지문": 3, ...}
     │
3. [문서 생성] 템플릿.hwp를 출력 경로로 복사
     │   → 출력 파일이 템플릿의 스타일 정의를 그대로 갖게 됨
     │
4. [COM 편집] 복사된 파일을 열어 문제/보기 텍스트 삽입
     │   → apply_question_format(): CharShape를 중고딕 9.5pt로 직접 설정
     │   → apply_passage_format(): CharShape를 휴먼명조 9.5pt로 직접 설정
     │   ※ HAction.Execute("Style")은 COM 행(hang) 버그로 사용 불가
     │
5. [저장] HWP 파일 저장
     │
6. [후처리] post_process_style_ids()
     │   → 저장된 파일을 바이너리로 열어 각 문단의 PARA_HEADER에서
     │     style_id를 "문제"(2) 또는 "지문"(3)으로 교정
     │   → 템플릿의 TAG_STYLE 레코드를 출력 파일에 이식
     │   → HWP에서 열면 스타일 이름이 올바르게 표시됨
```

### 6.3 시험별 템플릿이 다른 이유

각 시험에 맞는 템플릿을 따로 만드는 이유:

1. **스타일 속성이 다름**: 같은 "문제" 스타일이라도 아침모의고사는 9.5pt, 전국연합은 10pt
2. **스타일 종류가 다를 수 있음**: 특정 시험에만 "보기설명" 같은 추가 스타일이 필요할 수 있음
3. **용지 설정 내장**: 템플릿 자체에 A4/B4 기본 페이지 설정이 포함되어 있음

### 6.4 학원 전용 스타일 제작 가이드 (관리자용)

#### 새 시험용 템플릿 만들기

```
1. 한글(HWP)을 실행하고 새 문서를 연다

2. [서식] → [스타일] → [스타일 편집] 또는 [새 스타일]

3. "문제" 스타일 만들기:
   - 스타일 이름: 문제
   - 글꼴: 중고딕 (또는 원하는 글꼴)
   - 크기: 10pt (시험에 맞게 조정)
   - 문단 속성:
     - 정렬: 양쪽 정렬
     - 내어쓰기: 15pt
     - 줄간격: 글자에 따라 135%

4. "지문" 스타일 만들기:
   - 스타일 이름: 지문
   - 글꼴: 휴먼명조 (또는 원하는 글꼴)
   - 크기: 10pt
   - 문단 속성: "문제"와 동일하되 내어쓰기만 다르게

5. [파일] → [다른 이름으로 저장]
   - 경로: config/templates/전국연합_템플릿.hwp
   - ※ 본문 내용은 비워도 됨 (스타일 정의만 있으면 됨)
```

#### 기존 결과물에서 템플릿 추출하기

이미 잘 편집된 결과 HWP 파일이 있다면 그것을 템플릿으로 사용할 수 있다:

```
1. 잘 편집된 결과 HWP 파일을 연다
2. 본문 내용을 모두 선택(Ctrl+A) → 삭제
3. 스타일 정의만 남은 상태에서 다른 이름으로 저장
4. 이 파일을 config/templates/ 폴더에 복사
```

---

## 7. 파일 수정 상세

### 7.1 config_manager.py 수정

**변경 목적**: 프리셋 로딩 및 3단 병합 지원

```python
# === 추가할 메서드들 ===

class ConfigManager:
    PRESETS_DIR = "config/presets"

    def list_presets(self) -> list[dict]:
        """사용 가능한 프리셋 목록을 반환한다.
        Returns:
            [{"name": "아침모의고사", "file": "아침모의고사.json",
              "description": "..."}, ...]
        """
        presets_dir = Path(self.PRESETS_DIR)
        if not presets_dir.exists():
            return []
        result = []
        for f in sorted(presets_dir.glob("*.json")):
            data = self._load_json_file(f)
            result.append({
                "name": data.get("preset_name", f.stem),
                "file": f.name,
                "description": data.get("preset_description", ""),
            })
        return result

    def load_with_preset(self, preset_filename: str) -> dict:
        """프리셋을 포함한 3단 병합으로 config를 로드한다."""
        defaults = self._load_json_file(self.default_path)
        preset = self._load_json_file(
            Path(self.PRESETS_DIR) / preset_filename
        )
        user = self._load_json_file(self.user_path)

        merged = self._deep_merge(defaults, preset)  # 1+2단계
        merged = self._deep_merge(merged, user)       # 3단계
        self._config = merged
        return self._config

    def get_active_preset(self) -> str:
        """user_config.json에 저장된 마지막 사용 프리셋을 반환."""
        user = self._load_json_file(self.user_path)
        return user.get("active_preset", "")

    def set_active_preset(self, preset_name: str) -> None:
        """마지막 사용 프리셋을 user_config.json에 저장."""
        self.update({"active_preset": preset_name})
```

### 7.2 formatter.py 수정

**변경 목적**: 용지 크기 동적 적용 + 단 간격 설정값 사용

```python
# === setup_page() 수정 ===

PAPER_SIZES = {
    "A4": (210.0, 297.0),
    "B4": (257.0, 364.0),
    "B5": (182.0, 257.0),
    "Letter": (215.9, 279.4),
}

def setup_page(self, hwp):
    # ...기존 코드...
    paper_type = self.page_config.get("paper_type", "A4")
    width, height = PAPER_SIZES.get(paper_type, (210.0, 297.0))
    _safe_set_attr(page, "PaperWidth", hwp.MiliToHwpUnit(width))
    _safe_set_attr(page, "PaperHeight", hwp.MiliToHwpUnit(height))
    # ...나머지는 기존과 동일 (이미 config에서 읽고 있음)...


# === setup_columns() 수정 ===

def setup_columns(self, hwp):
    # ...기존 코드...
    column_gap = float(self.page_config.get("column_gap", 8.0))
    _safe_set_attr(col, "SameGap", hwp.MiliToHwpUnit(column_gap))
    # ...나머지 동일...
```

### 7.3 main_window.py 수정

**변경 목적**: 프리셋 선택 드롭다운 UI 추가

```python
# === initUI() 내부, drop_area 위에 추가 ===

# 프리셋 선택 영역
preset_group = QHBoxLayout()
preset_label = QLabel("시험 유형:")
self.preset_combo = QComboBox()
self.preset_combo.setMinimumWidth(250)
self._load_preset_list()
self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
preset_group.addWidget(preset_label)
preset_group.addWidget(self.preset_combo)
preset_group.addStretch()
main_layout.addLayout(preset_group)

# === 프리셋 관련 메서드 ===

def _load_preset_list(self):
    """config/presets/ 폴더의 프리셋 목록을 콤보박스에 로드."""
    self.preset_combo.clear()
    presets = self.config_manager.list_presets()
    if not presets:
        self.preset_combo.addItem("(기본 설정)", "")
        return
    for preset in presets:
        self.preset_combo.addItem(preset["name"], preset["file"])
    # 마지막 사용 프리셋 자동 선택
    active = self.config_manager.get_active_preset()
    if active:
        idx = self.preset_combo.findText(active)
        if idx >= 0:
            self.preset_combo.setCurrentIndex(idx)

def _on_preset_changed(self, index):
    """프리셋이 변경되면 config를 다시 로드한다."""
    preset_file = self.preset_combo.currentData()
    if preset_file:
        self.config_manager.load_with_preset(preset_file)
    else:
        self.config_manager.reload()
    preset_name = self.preset_combo.currentText()
    self.config_manager.set_active_preset(preset_name)
    self.service._refresh_dependencies()
```

### 7.4 settings_window.py 수정

**변경 목적**: 프리셋 저장 기능 추가

```python
# === 기존 저장 버튼 옆에 추가 ===

self.save_as_preset_btn = QPushButton("프리셋으로 저장")
self.save_as_preset_btn.clicked.connect(self.saveAsPreset)

# === 프리셋 저장 메서드 ===

def saveAsPreset(self):
    """현재 설정값을 새 프리셋 파일로 저장한다."""
    from PyQt5.QtWidgets import QInputDialog
    name, ok = QInputDialog.getText(self, "프리셋 이름", "프리셋 이름을 입력하세요:")
    if not ok or not name.strip():
        return
    preset = {
        "preset_name": name.strip(),
        "preset_description": "",
        "format": {
            "question_font": self.q_font_combo.currentText(),
            "passage_font": self.p_font_combo.currentText(),
            "font_size": self.size_spin.value(),
            "char_width": self.char_width_spin.value(),
            "char_spacing": self.char_spacing_spin.value(),
        },
        "paragraph": {
            "line_spacing": self.line_spacing_spin.value(),
            "indent_value": self.indent_spin.value(),
        },
        "style": {
            "template_path": self.style_template_edit.text().strip(),
            "question_style": self.question_style_edit.text().strip(),
            "passage_style": self.passage_style_edit.text().strip(),
        },
    }
    # 안전한 파일명 생성
    safe_name = re.sub(r'[\\/:*?"<>|]+', '_', name.strip())
    preset_dir = Path("config/presets")
    preset_dir.mkdir(parents=True, exist_ok=True)
    preset_path = preset_dir / f"{safe_name}.json"
    with preset_path.open("w", encoding="utf-8") as f:
        json.dump(preset, f, ensure_ascii=False, indent=2)
    QMessageBox.information(self, "저장 완료", f"프리셋 저장: {preset_path}")
```

---

## 8. 디렉토리 구조 변경

### 8.1 추가되는 파일/폴더

```
config/
├── default_config.json        (기존 유지)
├── user_config.json           (기존 유지, active_preset 키 추가)
├── presets/                   ← 신규 디렉토리
│   ├── 아침모의고사.json       ← 프리셋 파일
│   ├── 전국연합_사회.json
│   ├── 전국연합_국어.json
│   └── 경찰학_실전모의.json
└── templates/                 ← 신규 디렉토리
    ├── 아침모의고사_템플릿.hwp  ← 학원 전용 스타일 HWP
    ├── 전국연합_템플릿.hwp
    └── 경찰학_실전_템플릿.hwp
```

### 8.2 기존 파일과의 호환

- 프리셋이 없으면(config/presets/ 비어있으면) 기존과 동일하게 동작
- 프리셋 콤보박스에 "(기본 설정)" 항목이 표시됨
- 기존 user_config.json의 설정값도 그대로 유효

---

## 9. 구현 순서 (단계별)

### Phase 1: 기반 작업 (코드 변경 최소)

1. `config/presets/` 디렉토리 생성
2. `config/templates/` 디렉토리 생성
3. 현재 설정값으로 `presets/아침모의고사.json` 샘플 파일 작성
4. `formatter.py`의 `setup_page()`에서 용지 크기 하드코딩 제거 → `PAPER_SIZES` 딕셔너리 적용
5. `formatter.py`의 `setup_columns()`에서 단 간격(8mm) 하드코딩 제거 → config 값 사용

### Phase 2: ConfigManager 확장

6. `config_manager.py`에 `list_presets()`, `load_with_preset()` 메서드 추가
7. `config_manager.py`에 `get_active_preset()`, `set_active_preset()` 메서드 추가
8. `service.py`에서 프리셋 적용 시 `_refresh_dependencies()` 호출 보장

### Phase 3: GUI 연동

9. `main_window.py`에 프리셋 선택 콤보박스 추가
10. 콤보박스 변경 시 config 재로드 + service 갱신 로직 연결
11. 프로그램 시작 시 마지막 프리셋 자동 선택 (`active_preset`)

### Phase 4: 프리셋 관리 UI

12. `settings_window.py`에 "프리셋으로 저장" 버튼 추가
13. 프리셋 삭제/이름 변경 기능 (선택적)

### Phase 5: 실전 프리셋 세팅

14. 실제 시험별 양식 값을 조사하여 프리셋 JSON 작성
15. 시험별 한글 스타일 템플릿 HWP 제작
16. 테스트 및 검증

---

## 10. 주의사항

### 10.1 프리셋과 템플릿의 값은 반드시 일치시켜야 한다

프리셋 JSON의 `format.font_size: 10.0`과 템플릿 HWP의 "문제" 스타일 글자 크기가 **동일해야** 한다.
불일치하면 COM으로 적용하는 서식(프리셋 값)과 후처리로 적용하는 스타일(템플릿 값)이 충돌한다.

```
[올바른 예]
프리셋: font_size = 10.0
템플릿 "문제" 스타일: 글자 크기 10pt
→ 결과 일관성 보장

[잘못된 예]
프리셋: font_size = 10.0
템플릿 "문제" 스타일: 글자 크기 9.5pt
→ COM 적용과 스타일 정의가 충돌, 예측 불가능한 결과
```

### 10.2 기본 프리셋은 반드시 제공

프로그램 배포 시 최소 1개의 기본 프리셋(`아침모의고사.json`)을 포함해야 한다.
프리셋이 하나도 없으면 기존 동작(default + user)으로 폴백하지만, 사용자에게 혼란을 줄 수 있다.

### 10.3 templates 폴더 내 HWP 파일은 수정 금지

템플릿은 읽기 전용으로 관리해야 한다. 프로그램이 템플릿을 출력 경로로 복사한 뒤 편집하므로, 원본 템플릿이 변경되면 모든 이후 출력에 영향을 준다.

### 10.4 프리셋 JSON은 UTF-8 인코딩

한글 프리셋 이름, 설명, 스타일명을 위해 반드시 UTF-8로 저장한다.
기존 `ConfigManager._load_json_file()`이 이미 `encoding="utf-8"`을 사용하므로 추가 작업 불필요.

---

## 11. 구현 시 반드시 확인해야 할 함정 (Critical)

> 이 섹션은 다른 AI 또는 개발자가 이 문서를 참고하여 구현할 때
> **빠뜨리기 쉬운 핵심 포인트**를 모아놓은 것이다.

### 11.1 서브프로세스에 프리셋 전달 (가장 중요)

현재 HWP 출력 생성은 **별도 서브프로세스**에서 실행된다.

```
main_window.py (GUI 프로세스)
    │
    │  GenerationWorker → subprocess.Popen(...)
    ▼
subprocess_generation.py (별도 프로세스)
    │
    │  service = ExamProcessingService(ConfigManager())  ← 여기서 config를 새로 로드
    ▼
generator.py → formatter.py
```

**문제**: `subprocess_generation.py:174`에서 `ConfigManager()`를 새로 생성한다.
이 시점에 GUI에서 선택한 프리셋 정보가 전달되지 않으면, 서브프로세스는 기본 설정으로 동작한다.

**해결 방법 (택 1)**:

**(A) request.json에 프리셋 파일명 전달 (권장)**

```python
# main_window.py - GenerationWorker.run()에서
request = {
    "document": _document_to_payload(self.document),
    "source_file": self.source_file,
    "active_preset": "전국연합_사회.json",  # ← 추가
}

# subprocess_generation.py - main()에서
preset_file = request.get("active_preset", "")
cm = ConfigManager()
if preset_file:
    cm.load_with_preset(preset_file)
service = ExamProcessingService(cm)
```

**(B) user_config.json의 active_preset을 읽어서 자동 로드**

```python
# subprocess_generation.py
cm = ConfigManager()
active = cm.get_active_preset()
if active:
    # active_preset 이름으로 프리셋 파일 찾아 로드
    cm.load_with_preset(f"{active}.json")
service = ExamProcessingService(cm)
```

방법 (A)가 명시적이고 안전하다. (B)는 user_config.json 저장 타이밍에 의존하므로 경합 조건 가능.

### 11.2 settings_window.py의 저장 동작 변경

현재 `settings_window.py`의 `saveConfig()`는 글자 크기, 줄간격 등을 **user_config.json에 직접 저장**한다.

```python
# 현재 코드 (settings_window.py:305-331)
partial = {
    "format": {
        "question_font": ...,
        "font_size": ...,      # ← user_config.json에 저장됨
    },
    "paragraph": {
        "line_spacing": ...,   # ← user_config.json에 저장됨
    },
}
self.config_manager.update(partial)
```

**문제**: 3단 병합에서 user_config.json이 **최우선**이므로, 여기에 font_size가 저장되면 프리셋의 font_size를 항상 덮어쓴다.

**해결 방법**:

프리셋 도입 후 설정 창의 저장 동작을 분리해야 한다:

```
[설정 창에서 "저장"] → user_config.json (공통 설정만: 출력 폴더, 보안 모듈 등)
[설정 창에서 "프리셋으로 저장"] → presets/XXX.json (글자/문단 서식)
```

즉, `saveConfig()`에서 `format`, `paragraph` 키를 user_config.json에 저장하지 않도록 수정하거나,
프리셋이 활성화된 상태에서는 서식 관련 값을 user_config 대신 프리셋 파일에 저장해야 한다.

### 11.3 build.spec (PyInstaller 배포) 수정

프리셋 관련 폴더를 배포 패키지에 포함해야 한다.

```python
# build.spec의 datas에 추가
datas=[
    ('config/presets', 'config/presets'),      # ← 추가
    ('config/templates', 'config/templates'),  # ← 추가
    # ... 기존 항목 유지
],
```

### 11.4 프리셋이 0개일 때 폴백 동작

`config/presets/` 디렉토리가 비어있거나 없는 경우:
- 콤보박스에 "(기본 설정)" 1개만 표시
- `ConfigManager.load_with_preset()`을 호출하지 않음
- 기존 2단 병합(default + user)으로 동작
- **모든 기존 기능이 그대로 동작해야 한다** (하위 호환 필수)

### 11.5 현재 코드의 하드코딩 위치 정리

프리셋 값이 제대로 적용되려면 아래 하드코딩을 반드시 config 참조로 변경해야 한다:

| 파일 | 라인 | 현재 코드 | 변경 |
|------|------|-----------|------|
| `formatter.py` | 265-266 | `MiliToHwpUnit(210.0)`, `MiliToHwpUnit(297.0)` | `PAPER_SIZES[paper_type]` 참조 |
| `formatter.py` | 290 | `MiliToHwpUnit(8.0)` (단 간격) | `page_config.get("column_gap", 8.0)` |
| `formatter.py` | 66 | `self.symbol_font = "바탕"` | 프리셋에서 변경할 일은 없지만 참고 |

### 11.6 기존 파일 참조 맵 (구현 시 필독)

다른 AI가 코드를 수정할 때 각 파일의 역할과 연결 관계:

```
[설정 흐름]
config_manager.py ──→ service.py ──→ generator.py ──→ formatter.py
     │                    │
     │                    └─→ parser.py
     │
     └──→ subprocess_generation.py (별도 프로세스에서 독립 로드)

[GUI 흐름]
main_window.py
     ├─→ settings_window.py (설정 편집)
     ├─→ preview_window.py (파싱 결과 미리보기)
     └─→ GenerationWorker (서브프로세스로 생성 위임)
              └─→ subprocess_generation.py

[수정 필요 파일 요약]
config_manager.py    : 프리셋 로드/목록/병합 메서드 추가
formatter.py         : 용지 크기/단 간격 하드코딩 제거
main_window.py       : 프리셋 콤보박스 + 프리셋→request.json 전달
subprocess_generation.py : request에서 프리셋 읽어 config 로드
settings_window.py   : 저장 동작 분리 + 프리셋 저장 버튼
service.py           : 프리셋 변경 시 dependencies 갱신
build.spec           : presets/, templates/ 디렉토리 포함
```
