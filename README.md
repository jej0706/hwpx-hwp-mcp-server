# hwpx-hwp-mcp

Claude에 붙여서 **HWP와 HWPX 파일을 동시에 편집**할 수 있는 MCP 서버입니다. 한글 프로그램의 COM 자동화(`pyhwpx`)를 백엔드로 사용해 두 포맷 모두 완전 편집이 가능합니다.

## 특징

- 🟢 **HWP · HWPX 둘 다 완전 편집** — 포맷 감지와 변환이 확장자 기반으로 자동 처리
- 🟢 **렌더링 충실도 보장** — 한글 엔진이 직접 열고 저장하므로 실제 출력과 동일
- 🟢 **얇은 레이어** — `pyhwpx`의 150+ 메서드 위에 FastMCP 도구만 얹음
- 🟢 **38개 도구** — 세션 관리 · 읽기/분석 · 템플릿 채우기 · 문서 생성 · 표 조작 · 페이지 레이아웃 · 대량 변환

## 요구 사항

- Windows 10/11
- **32-bit Python 3.11+** ⚠️ 중요: 64-bit Python 아님
- **한글(Hancom Office)** 설치 (2018 이후 권장)
- `pyhwpx`, `pywin32`, `mcp[cli]`, `pydantic`

### 왜 32-bit Python 이어야 하나요?

한글은 2024 버전까지도 **32-bit COM 서버**입니다 (`C:\Program Files (x86)\Hnc\...\Hwp.exe`). 64-bit Python 에서 `HWPFrame.HwpObject` 를 Dispatch 하면 Windows 의 COM 브리지가 한글 LocalServer32 를 못 띄워서 `CO_E_SERVER_EXEC_FAILURE (0x80080005)` 로 실패합니다.

**해결**: https://www.python.org/downloads/windows/ 에서 **"Windows installer (32-bit)"** 를 받아 기존 64-bit Python 과 **별도 경로**에 설치하세요 (예: `C:\Python313-32`). 이후 해당 32-bit Python 을 이 MCP 서버 전용으로만 쓰면 됩니다. 시스템 기본 Python 에는 영향 없습니다.

확인 방법:
```powershell
C:\Python313-32\python.exe -c "import struct; print(struct.calcsize('P')*8, 'bit')"
# 출력: 32 bit
```

### pandas 는 왜 없나요?

`pyhwpx` 가 `pandas` 를 선언적 의존성으로 올리지만, pandas 는 **Python 3.10+ 의 32-bit Windows wheel 을 제공하지 않습니다**. 마지막 32-bit Windows wheel 은 pandas 1.5.3/Python 3.9 이었습니다. 32-bit Python 3.13 에서 pandas 를 설치하려고 하면 소스 빌드 모드로 떨어져서 MSVC/Cython/Meson 을 요구합니다.

이 서버는 pandas 에 의존하지 않는 pyhwpx 경로만 사용하므로, 내장된 **pandas stub** (`backend/pandas_stub.py`) 을 `sys.modules['pandas']` 에 미리 주입해서 `import pandas as pd` 를 만족시킵니다. 실제 pandas 가 설치되어 있으면 stub 은 비활성화되고 real pandas 를 우선 사용합니다.

- `fill_fields` → `put_field_text(dict, "")` 의 dict 분기는 pandas 를 타지 않음 ✓
- `insert_table` → `table_from_data` 대신 `create_table` + 셀 순회 방식으로 구현 ✓
- `get_table_as_csv` → `table_to_df` 대신 셀별 텍스트 수집 + `csv` 모듈 ✓

필요 시 `pip install pandas` 로 real pandas 를 추가하면 (64-bit Python 3.9 등) stub 이 자동으로 양보합니다.

## 설치

```powershell
# 1) 저장소 클론 또는 이 폴더로 이동
cd path\to\hwpx-hwp-mcp-server

# 2) 비-pandas 의존성을 wheel 로만 설치
C:\Python313-32\python.exe -m pip install --only-binary=:all: `
    "mcp[cli]" numpy pywin32 pydantic Pillow pyperclip openpyxl

# 3) pyhwpx 를 --no-deps 로 설치 (pandas 소스 빌드 회피)
C:\Python313-32\python.exe -m pip install --no-deps pyhwpx

# 4) 우리 패키지를 --no-deps 로 설치
C:\Python313-32\python.exe -m pip install --no-deps -e .

# 5) (선택) 개발용 테스트 도구
C:\Python313-32\python.exe -m pip install pytest pytest-asyncio
```

설치가 잘 됐는지 (COM을 건드리지 않는) 가벼운 확인:

```powershell
python tests/smoke/list_tools.py
```

24개 도구 이름이 출력되면 OK.

## Claude Desktop 연동

`%APPDATA%\Claude\claude_desktop_config.json`에 아래 항목을 추가하세요:

```json
{
  "mcpServers": {
    "hwpx-hwp": {
      "command": "C:\\Python313-32\\python.exe",
      "args": ["-m", "hwpx_hwp_mcp"],
      "env": { "PYTHONUTF8": "1" }
    }
  }
}
```

⚠️ `command` 는 반드시 **32-bit** Python 실행 파일을 가리켜야 합니다. 기본 64-bit Python 을 사용하면 시작할 때 아래와 같은 경고가 나오고, 첫 도구 호출에서 COM 기동 에러가 납니다:

```
WARNING: 64-bit Python detected, but Hancom HWP appears to be installed as 32-bit only.
```

`uv` 사용 시:

```json
{
  "mcpServers": {
    "hwpx-hwp": {
      "command": "uv",
      "args": [
        "--directory", "path\\to\\hwpx-hwp-mcp-server",
        "run", "hwpx-hwp-mcp"
      ],
      "env": { "PYTHONUTF8": "1" }
    }
  }
}
```

Claude Desktop을 재시작하면 도구가 로드됩니다.

## Claude Code 연동

```powershell
claude mcp add hwpx-hwp -- python -m hwpx_hwp_mcp
```

## 도구 목록 (24개)

### A. 세션 관리
| 도구 | 설명 |
|---|---|
| `open_document` | HWP/HWPX 파일 열기 → `doc_id` 반환 |
| `create_new_document` | 빈 문서 생성 |
| `save_document` | 현재 경로에 저장 (기본 `.bak` 백업) |
| `save_as` | 다른 경로/포맷으로 저장 (확장자 자동 분기) |
| `close_document` | 문서 닫기 |
| `list_open_documents` | 열려 있는 모든 문서와 `doc_id` 확인 |

### B. 읽기 / 분석
| 도구 | 설명 |
|---|---|
| `get_document_text` | 전체 텍스트 추출 |
| `get_document_info` | 제목/경로/페이지 수/필드 수 등 메타데이터 |
| `get_structure` | 표·이미지·필드 구조 아웃라인 |
| `search_text` | 본문 검색 (정규식 지원) |
| `export_document` | text/html/pdf/docx로 내보내기 |
| `get_table_as_csv` | 특정 표를 CSV로 |

### C. 템플릿 채우기 (가장 가치 높은 기능)
| 도구 | 설명 |
|---|---|
| `list_fields` | 누름틀(필드) 목록과 현재 값 |
| `fill_fields` | 누름틀 일괄 채우기 — 서식 보존 |
| `create_field` | 현재 위치에 누름틀 생성 |
| `replace_text` | 본문 찾기/바꾸기 (플레이스홀더용) |
| `fill_table_by_path` | `"이름: > right"` DSL로 표 셀 채우기 |

### D. 문서 생성 (from scratch)
| 도구 | 설명 |
|---|---|
| `insert_paragraph` | 문단 삽입 (스타일/정렬 지정 가능) |
| `insert_table` | 표 삽입 (data로 즉시 채우기) |
| `insert_image` | 이미지 삽입 |
| `insert_page_break` | 페이지 나누기 |
| `set_font` | 글꼴/크기/굵기/색상 변경 |

### F. 표 조작 고도화 (Sprint 1)
| 도구 | 설명 |
|---|---|
| `set_cell_shade` | 셀 배경색 (hex, 행/열/범위 선택) |
| `set_cell_border` | 셀 테두리 (all/outside/inside/상하좌우/대각선/없음) |
| `set_cell_alignment` | 셀 내 텍스트 정렬 (수평/수직 9방향) |
| `merge_cells` | 셀 병합 (범위 선택 후 병합) |
| `split_cell` | 단일 셀을 rows×cols 로 분할 |
| `set_column_width` | 특정 열 너비 (mm) |
| `set_row_height` | 특정 행 높이 (mm) |
| `insert_table_row` | 표에 행 추가 (끝 또는 지정 위치 아래) |
| `delete_table_row` | 표의 특정 행 삭제 |
| `insert_table_column` | 표 오른쪽에 열 추가 |
| `delete_table_column` | 표의 특정 열 삭제 |

### G. 페이지 레이아웃 (Sprint 1)
| 도구 | 설명 |
|---|---|
| `set_page_settings` | 용지 크기(A3~B5/Letter/Legal/custom), 방향, 여백 (한 번에) |
| `insert_page_number` | 쪽 번호 필드 삽입 |
| `insert_section_break` | 구역 나누기 (섹션별 다른 레이아웃) |

### E. 대량 / 일괄
| 도구 | 설명 |
|---|---|
| `batch_replace_in_files` | 여러 파일에 일괄 find/replace |
| `convert_files` | 여러 파일을 hwp/hwpx/pdf/html/docx로 일괄 변환 |

### 셀 선택자 (`cells` 파라미터 공통 문법)
표 조작 도구 대부분이 동일한 `cells` 셀렉터를 받습니다:

| 형식 | 의미 | 예 |
|---|---|---|
| `"all"` | 표 전체 | `"all"` |
| `"row:N"` | N 번째 행 (1-based) | `"row:1"` (헤더 행) |
| `"col:N"` 또는 `"col:L"` | N 번째 열 (1-based 숫자 or A,B,C...) | `"col:1"`, `"col:A"` |
| `"A1"` | 단일 셀 (Excel 표기) | `"B2"` |
| `"A1:C3"` | 직사각형 범위 | `"A1:J1"` |

## 사용 예 (Claude에 시키는 자연어)

- **템플릿 채우기**: "이 `invoice_template.hwpx` 양식의 고객명에 '홍길동', 날짜에 '2026-04-14', 금액에 '1,000,000원'을 채워서 `invoice_홍길동.hwpx`로 저장해줘"
- **포맷 변환**: "`contracts/` 폴더의 모든 .hwp 파일을 .hwpx로 변환해서 `contracts_hwpx/`에 저장해줘"
- **대량 수정**: "`reports/` 안의 모든 문서에서 '2025년'을 '2026년'으로 바꿔줘"
- **문서 생성**: "회의록 양식을 새로 만들어줘. 제목, 일시, 참석자 표, 안건, 결정 사항 섹션이 있어야 해. 저장은 `meeting_minutes_template.hwpx`로"
- **분석**: "`annual_report.hwpx`를 열어서 그 안의 모든 표를 CSV로 꺼내줘"
- **스타일 있는 표 생성**: "5행 10열 표 만들고, 헤더 행은 하늘색 배경에 가운데 정렬, 외곽선만 둘러줘"
- **A4 landscape 보고서**: "새 문서 A4 가로로 설정하고 위아래 여백 15mm, 좌우 20mm. 본문 입력 후 쪽 번호 넣어줘"

## 아키텍처 요약

```
┌────────────────┐
│   Claude       │ (stdio JSON-RPC)
└───────┬────────┘
        │
┌───────▼────────┐
│  FastMCP       │  (async tool handlers)
│  24 tools      │
└───────┬────────┘
        │  await session.call(lambda hwp: ...)
┌───────▼────────┐
│ HancomSession  │  (singleton, single-thread executor)
│   ┌──────────┐ │
│   │ worker T │ │  ← pythoncom.CoInitialize 는 이 스레드에서만
│   │  (STA)   │ │
│   └────┬─────┘ │
└────────┼───────┘
         │
┌────────▼──────────────┐
│ pyhwpx.Hwp (COM)     │
│ HWPFrame.HwpObject   │
└────────┬──────────────┘
         │
┌────────▼──────────────┐
│  한글 프로그램        │
│  (headless, invisible)│
└───────────────────────┘
```

### 왜 단일-스레드 executor?

한글 COM은 **STA(Single-Threaded Apartment)**입니다. `pythoncom.CoInitialize()`가 호출된 그 스레드에서만 COM 호출이 유효합니다. `asyncio.to_thread`는 기본 스레드풀을 공유해서 워커가 요청마다 달라질 수 있기 때문에, 전용 `ThreadPoolExecutor(max_workers=1)`을 `HancomSession`에 박아두고 **모든 pyhwpx 호출을 이 스레드 하나로 직렬화**합니다. 부수 효과로 동시성 제어까지 공짜로 해결됩니다.

## 테스트

### 단위 테스트 (COM 불필요)
```powershell
pip install -e .[dev]
pytest
```

### 수동 smoke 테스트 (한글 설치 필요)
```powershell
python tests/smoke/list_tools.py          # 도구 목록만 — COM 안 씀
python tests/smoke/cycle.py               # 실제 생성/저장/재오픈
python tests/smoke/format_roundtrip.py    # hwp/hwpx/pdf/docx 변환
```

## 주의사항

- **한글 설치 필수** — 이 서버는 한글을 COM으로 제어합니다. 한글이 없으면 시작 시 친숙한 에러와 함께 종료됩니다.
- **이미지 경로** — `insert_image`에는 반드시 **절대 경로**를 넘기세요. 상대 경로는 한글이 다른 디렉토리에서 찾습니다.
- **동시 요청** — stdio MCP 서버는 도구 호출을 직렬로 처리합니다. 긴 작업 중에는 다른 도구가 대기합니다.
- **메모리** — 장시간 실행 시 한글 프로세스가 수백 MB를 사용합니다. 필요하면 Claude Desktop을 재시작하면 됩니다.
- **`mcp__hwpx-mcp`와의 차이** — 이 서버는 COM 백엔드(실제 한글 엔진)이고, 기존 `mcp__hwpx-mcp`는 XML 기반입니다. 렌더링 충실도가 중요할 때 이 서버를, 한글 없이 돌아가야 할 때 저 서버를 쓰세요. 두 서버는 동시 사용 가능합니다.

## 한글 창 공유 동작 (중요)

한글의 `HWPFrame.HwpObject` COM 서버는 **MultipleUse LocalServer32** 로 등록되어 있어요. 이 말은:

- 시스템 전체에 **Hwp.exe 프로세스는 하나만** 실행됩니다.
- 모든 Python/외부 클라이언트 (`pyhwpx`, `win32com.client` 등) 는 **같은 프로세스에 연결**됩니다.
- 즉 사용자가 한글을 실행해둔 상태에서 MCP 도구를 호출하면, MCP 서버는 **사용자가 쓰고 있는 그 한글 인스턴스를 공유** 합니다.

### MCP 서버의 대응

`HancomSession` 은 시작 시 Hwp.exe 프로세스가 이미 떠있는지 확인하고 두 경우를 구분합니다:

| 시나리오 | 동작 |
|---|---|
| **Fresh start** (사용자 한글 없음) | MCP 가 자체 인스턴스를 띄우고 `set_visible(False)` 로 창 숨김. 백그라운드에서 작업. |
| **Shared with user** (사용자가 이미 한글 열어둠) | **visibility 를 건드리지 않음.** 사용자 창은 그대로 보이는 상태 유지. MCP 종료 시에도 `hwp.quit()` 이나 `taskkill` 을 호출하지 않아서 사용자 작업이 보존됨. |

### 그래도 주의할 점

Shared 시나리오에서도 MCP 도구 실행 중 다음과 같은 일이 일어날 수 있습니다:

- **포커스 이동**: `create_new_document` 가 새 문서 탭을 생성하면 사용자가 작업 중이던 탭에서 새 탭으로 활성 창이 바뀝니다.
- **사용자 문서 위에 작업**: MCP 가 `open_document` 로 다른 파일을 열면 한글 창에 새 탭이 추가됩니다.
- **진행 상황 노출**: 저장/수정 작업이 사용자 화면에 그대로 보입니다.

병렬 작업을 원한다면 사용자가 한글을 최소화해두거나, MCP 작업이 끝난 뒤에 본인 작업을 재개하는 걸 권장합니다.

## 라이선스

MIT
