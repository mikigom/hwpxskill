---
name: hwpx
description: "한글(HWPX) 문서 생성/읽기/편집 스킬. .hwpx 파일, 한글 문서, Hancom, OWPML 관련 요청 시 사용."
---

# HWPX 문서 스킬

한글(Hancom Office)의 HWPX 파일을 **XML 직접 작성** 중심으로 생성, 편집, 읽기할 수 있는 스킬.
HWPX는 ZIP 기반 XML 컨테이너(OWPML 표준)이다. python-hwpx API의 서식 버그를 완전히 우회하며, 세밀한 서식 제어가 가능하다.

## Critical Rules

1. **HWPX만 지원**: `.hwp`(바이너리) 파일은 지원하지 않는다. 사용자가 `.hwp` 파일을 제공하면 **한글 오피스에서 `.hwpx`로 다시 저장**하도록 안내할 것
2. **secPr 필수**: section0.xml 첫 문단의 첫 run에 반드시 secPr + colPr 포함
3. **mimetype 순서**: HWPX 패키징 시 mimetype은 첫 번째 ZIP 엔트리, ZIP_STORED
4. **네임스페이스 보존**: XML 편집 시 `hp:`, `hs:`, `hh:`, `hc:` 접두사 유지
5. **itemCnt 정합성**: header.xml의 charProperties/paraProperties/borderFills itemCnt가 실제 자식 수와 일치
6. **ID 참조 정합성**: section0.xml의 charPrIDRef/paraPrIDRef가 header.xml 정의와 일치
7. **venv 사용**: 프로젝트의 `.venv/bin/python3` (lxml, fpdf2 패키지 필요)
8. **검증**: 생성 후 반드시 `validate.py`로 무결성 확인
9. **레퍼런스**: 상세 XML 구조는 `$SKILL_DIR/references/hwpx-format.md` 참조
10. **build_hwpx.py 우선**: 새 문서 생성은 build_hwpx.py 사용 (python-hwpx API 직접 호출 지양)
11. **빈 줄**: `<hp:t/>` 사용 (self-closing tag)
12. **레퍼런스 우선 강제**: 사용자가 HWPX를 첨부하면 반드시 `analyze_template.py` + 추출 XML 기반으로 복원/재작성할 것
13. **examples 폴더 미사용**: 작업 중 `.claude/skills/hwpx/examples/*` 파일은 읽기/참조/복사에 사용하지 말 것
14. **쪽수 동일 필수**: 레퍼런스 기반 작업에서는 최종 결과의 쪽수를 레퍼런스와 동일하게 유지할 것
15. **무단 페이지 증가 금지**: 사용자 명시 요청/승인 없이 쪽수 증가를 유발하는 구조 변경 금지
16. **구조 변경 제한**: 사용자 요청이 없는 한 문단/표의 추가·삭제·분할·병합 금지 (치환 중심 편집)
17. **page_guard 필수 통과**: `validate.py`와 별개로 `page_guard.py`를 반드시 통과해야 완료 처리
18. **Mixed Content 주의**: `<hp:t>`는 텍스트와 자식 요소(`<hp:fwSpace/>`, `<hp:lineBreak/>` 등)가 혼재하는 Mixed Content 요소이다. 수동 편집 시에도 내부 공백·줄바꿈을 임의로 변경하지 말 것

## 워크플로우 선택 (작업 전 해당 파일을 반드시 Read할 것)

| 작업 유형 | 읽을 참조 파일 |
|-----------|---------------|
| 사용자가 .hwpx 첨부 → 복원/수정 | `$SKILL_DIR/references/workflow-reference-restore.md` |
| 새 문서 생성 (레퍼런스 없음) | `$SKILL_DIR/references/workflow-xml-first.md` |
| 기존 .hwpx 편집 (unpack/pack) | `$SKILL_DIR/references/workflow-edit.md` |
| 텍스트 추출/읽기 | `$SKILL_DIR/references/workflow-read.md` |
| HWPX 검증 | `$SKILL_DIR/references/workflow-validate.md` |
| HWPX → PDF 변환 | `$SKILL_DIR/references/workflow-pdf.md` |
| section0.xml 작성 시 | `$SKILL_DIR/references/section0-guide.md` |
| header.xml 수정/스타일 추가 시 | `$SKILL_DIR/references/header-guide.md` |
| 템플릿 스타일 ID 확인 | `$SKILL_DIR/references/style-maps.md` |
| OWPML XML 요소 레퍼런스 | `$SKILL_DIR/references/hwpx-format.md` |

## 환경

```
SKILL_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VENV="<프로젝트>/.venv/bin/activate"
```

모든 Python 실행 시:
```bash
source "$VENV"
```

## 스크립트 요약

| 스크립트 | 용도 |
|----------|------|
| `scripts/build_hwpx.py` | **핵심** — 템플릿 + XML → HWPX 조립 |
| `scripts/analyze_template.py` | HWPX 심층 분석 (레퍼런스 기반 생성의 청사진) |
| `scripts/office/unpack.py` | HWPX → 디렉토리 (XML pretty-print) |
| `scripts/office/pack.py` | 디렉토리 → HWPX (mimetype first) |
| `scripts/validate.py` | HWPX 파일 구조 검증 |
| `scripts/page_guard.py` | 레퍼런스 대비 페이지 드리프트 위험 검사 (필수 게이트) |
| `scripts/text_extract.py` | HWPX 텍스트 추출 |
| `scripts/hwpx2pdf.py` | HWPX → PDF 변환 (standalone/hancom/word) |

## 단위 변환

| 값 | HWPUNIT | 의미 |
|----|---------|------|
| 1pt | 100 | 기본 단위 |
| 10pt | 1000 | 기본 글자크기 |
| 1mm | 283.5 | 밀리미터 |
| 1cm | 2835 | 센티미터 |
| A4 폭 | 59528 | 210mm |
| A4 높이 | 84186 | 297mm |
| 좌우여백 | 8504 | 30mm |
| 본문폭 | 42520 | 150mm (A4-좌우여백) |

## 디렉토리 구조

```
.claude/skills/hwpx/
├── SKILL.md                              # 이 파일 (축약 — 상세는 references/ 참조)
├── scripts/
│   ├── office/
│   │   ├── unpack.py
│   │   └── pack.py
│   ├── build_hwpx.py
│   ├── analyze_template.py
│   ├── validate.py
│   ├── page_guard.py
│   ├── text_extract.py
│   └── hwpx2pdf.py
├── templates/
│   ├── base/                             # 베이스 템플릿
│   ├── gonmun/                           # 공문 오버레이
│   ├── report/                           # 보고서 오버레이
│   ├── minutes/                          # 회의록 오버레이
│   └── proposal/                         # 제안서/사업개요 오버레이
└── references/
    ├── hwpx-format.md                    # OWPML XML 요소 레퍼런스
    ├── workflow-reference-restore.md     # 레퍼런스 기반 복원 워크플로우
    ├── workflow-xml-first.md             # XML-first 새 문서 생성
    ├── workflow-edit.md                  # 기존 문서 편집 (unpack/pack)
    ├── workflow-read.md                  # 텍스트 추출/읽기
    ├── workflow-validate.md              # HWPX 검증
    ├── workflow-pdf.md                   # PDF 변환
    ├── section0-guide.md                 # section0.xml 작성 가이드
    ├── header-guide.md                   # header.xml 수정 가이드
    └── style-maps.md                     # 템플릿별 스타일 ID 맵
```
