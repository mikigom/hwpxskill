# HWPX → PDF 변환 워크플로우 (visual inspection)

생성/수정한 HWPX를 PDF로 변환하여 시각적으로 확인하거나, 기존 PDF 관련 스킬에 전달할 때 사용한다.

## 엔진

| 엔진 | 필요 조건 | 장점 | 단점 |
|------|-----------|------|------|
| standalone | lxml, fpdf2 | 크로스플랫폼, 오피스 불필요 | 레이아웃 근사치 (visual inspection용) |
| hancom | Windows + 한컴오피스 | 완벽한 레이아웃 | Windows + 한컴오피스 필수 |
| word | Windows + MS Word + simple-hwp2pdf | MS Word 레이아웃 | Windows + MS Word 필수 |

## 사용법

```bash
source "$VENV"

# 1. standalone 변환 (기본, 크로스플랫폼)
python3 "$SKILL_DIR/scripts/hwpx2pdf.py" result.hwpx -o result.pdf --engine standalone

# 2. auto 모드 (standalone → hancom → word 순으로 시도)
python3 "$SKILL_DIR/scripts/hwpx2pdf.py" result.hwpx -o result.pdf

# 3. 특정 엔진 강제 지정
python3 "$SKILL_DIR/scripts/hwpx2pdf.py" result.hwpx -o result.pdf --engine hancom
```

## Python API

```python
from hwpx2pdf import convert

# standalone (기본 auto: standalone → hancom → word)
pdf_path = convert("result.hwpx", "result.pdf")

# 엔진 지정
pdf_path = convert("result.hwpx", "result.pdf", engine="standalone")
```

## 활용 시점

- **검증 단계**: `validate.py` + `page_guard.py` 통과 후, 시각적 확인이 필요할 때
- **PDF 스킬 연계**: HWPX 내용을 PDF 관련 도구(읽기, 병합, 워터마크 등)에서 사용해야 할 때
- **사용자 전달**: HWPX를 열 수 없는 환경에서 결과물을 확인할 때
