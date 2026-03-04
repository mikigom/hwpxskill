# hwpxskill

한컴오피스 HWPX 문서를 AI 코딩 에이전트에서 다룰 수 있게 해주는 스킬입니다.

python-hwpx API를 쓰면 버그가 많아서, XML을 직접 건드리는 방식을 택했습니다. 덕분에 기존 문서의 서식이나 구조를 거의 그대로 유지하면서 내용만 갈아끼울 수 있습니다.

## 뭘 할 수 있나

원본 HWPX 파일을 넣으면 스타일, 표 구조, 셀 병합, 여백까지 분석해서 구조를 보존한 채 내용만 바꿔줍니다. 원본이 없으면 공문, 보고서 같은 내장 템플릿으로 새 문서를 만들 수도 있고요. 다 만들고 나면 `page_guard.py`가 원본 대비 페이지 수가 달라졌는지 자동으로 잡아냅니다.

OWPML 표준 XML을 직접 다루기 때문에 charPr, paraPr 단위로 서식을 제어할 수 있습니다. Claude Code, Cursor, Codex CLI에서 모두 동작합니다.

## 설치

Agent Skills 표준을 따르고 있어서, 어떤 도구든 스킬 디렉토리에 넣기만 하면 됩니다.

```bash
git clone https://github.com/Canine89/hwpxskill.git
```

### Claude Code

```bash
# 이 프로젝트에서만 쓸 때
cp -r hwpxskill .claude/skills/hwpxskill

# 어디서든 쓸 때
cp -r hwpxskill ~/.claude/skills/hwpxskill
```

넣어두면 HWPX 관련 작업할 때 알아서 불러옵니다.

### Cursor

```bash
# 이 프로젝트에서만 쓸 때
cp -r hwpxskill .cursor/skills/hwpxskill

# 어디서든 쓸 때
cp -r hwpxskill ~/.cursor/skills/hwpxskill
```

`.hwpx` 파일을 열 때 자동으로 활성화되게 하려면 rule 파일을 하나 추가하면 됩니다.

```yaml
# .cursor/rules/hwpx.mdc
---
description: "HWPX 문서 작업 시 hwpxskill 사용"
globs: ["*.hwpx"]
---
```

### Codex CLI

```bash
# 이 프로젝트에서만 쓸 때
cp -r hwpxskill .agents/skills/hwpxskill

# 어디서든 쓸 때
cp -r hwpxskill ~/.agents/skills/hwpxskill
```

Codex 세션 안에서 `$skill-installer`로 설치할 수도 있습니다.

## 빠른 시작

### 1. 새 문서 만들기

템플릿 골라서 바로 생성. 원본 파일 없을 때 씁니다.

```bash
python3 scripts/build_hwpx.py --template gonmun --output result.hwpx
```

### 2. 기존 문서 편집

HWPX를 풀고, XML 고치고, 다시 묶습니다.

```bash
python3 scripts/office/unpack.py document.hwpx ./unpacked/
# XML 수정
python3 scripts/office/pack.py ./unpacked/ edited.hwpx
```

### 3. 텍스트 추출

문서에서 텍스트만 뽑습니다. 표도 포함되고, 마크다운으로도 뽑을 수 있습니다.

```bash
python3 scripts/text_extract.py document.hwpx --format markdown
```

### 4. 문서 검증

ZIP 구조, XML 유효성, mimetype 위치 같은 걸 점검합니다.

```bash
python3 scripts/validate.py result.hwpx
```

### 5. 레퍼런스 기반 복원

이게 핵심입니다. 원본 문서를 분석해서 스타일과 구조를 통째로 가져온 뒤, 내용만 갈아끼웁니다. HWPX 파일을 첨부하면 이 흐름이 자동으로 돌아갑니다.

```bash
# 분석
python3 scripts/analyze_template.py reference.hwpx \
  --extract-header /tmp/ref_header.xml \
  --extract-section /tmp/ref_section.xml

# 빌드
python3 scripts/build_hwpx.py \
  --header /tmp/ref_header.xml \
  --section /tmp/new_section0.xml \
  --output result.hwpx

# 검증 + 페이지 가드
python3 scripts/validate.py result.hwpx
python3 scripts/page_guard.py --reference reference.hwpx --output result.hwpx
```

## 템플릿

| 템플릿 | 용도 | 특징 |
|--------|------|------|
| base | 기본 골격 | 최소 스타일, 빈 문서 시작점 |
| gonmun | 공문서 | 기관명, 수신처, 시행일자, 연락처 |
| report | 보고서 | 섹션 헤더, 들여쓰기, 체크박스 |
| minutes | 회의록 | 섹션 라벨, 테두리 구분 |
| proposal | 제안서 | 색상 헤더, 번호 뱃지 |

## 요구사항

- Python 3.6 이상
- lxml (`pip install lxml`)
- 가상환경 권장

## 스크립트

| 스크립트 | 하는 일 |
|----------|---------|
| `build_hwpx.py` | 템플릿 + XML 조합해서 HWPX 생성 |
| `analyze_template.py` | 레퍼런스 HWPX 분석 |
| `office/unpack.py` | HWPX를 디렉토리로 풀기 |
| `office/pack.py` | 디렉토리를 HWPX로 묶기 |
| `validate.py` | HWPX 구조 검증 |
| `page_guard.py` | 원본 대비 페이지 수 변동 감지 |
| `text_extract.py` | 텍스트 추출 |

## 자세한 사용법

스타일 ID 체계, XML 구조 규칙, 템플릿별 charPr/paraPr 매핑 같은 건 [SKILL.md](./SKILL.md)에 다 정리되어 있습니다.
