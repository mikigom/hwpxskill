# 레퍼런스 기반 문서 복원/재작성 워크플로우

사용자가 `.hwpx`를 첨부한 경우 **기본으로** 이 워크플로우를 사용한다.

## 흐름

1. **레퍼런스 확보**: 첨부된 HWPX를 기준 문서로 사용
2. **심층 분석/추출**: `analyze_template.py`로 `header.xml`, `section0.xml` 추출
3. **구조 복원**: header 스타일 ID/표 구조/셀 병합/여백/문단 흐름을 최대한 동일하게 유지
4. **요청 반영 재작성**: 사용자가 요구한 텍스트/데이터만 교체하고 구조는 보존
5. **빌드/검증**: `build_hwpx.py` + `validate.py`로 결과 산출 및 무결성 확인
6. **쪽수 가드(필수)**: `page_guard.py`로 레퍼런스 대비 페이지 드리프트 위험 검사

## 99% 근접 복원 기준 (실무 체크리스트)

- `charPrIDRef`, `paraPrIDRef`, `borderFillIDRef` 참조 체계 동일
- 표의 `rowCnt`, `colCnt`, `colSpan`, `rowSpan`, `cellSz`, `cellMargin` 동일
- 문단 순서, 문단 수, 주요 빈 줄/구획 위치 동일
- 페이지/여백/섹션(secPr) 동일
- 변경은 사용자 요청 범위(본문 텍스트, 값, 항목명 등)로 제한

## 쪽수 동일(100%) 필수 기준

- 사용자가 레퍼런스를 제공한 경우 **결과 문서의 최종 쪽수는 레퍼런스와 동일해야 한다**
- 쪽수가 늘어날 가능성이 보이면 먼저 텍스트를 압축/요약해서 기존 레이아웃에 맞춘다
- 사용자 명시 요청 없이 `hp:p`, `hp:tbl`, `rowCnt`, `colCnt`, `pageBreak`, `secPr`를 변경하지 않는다
- `validate.py` 통과만으로 완료 처리하지 않는다. `edit_guard.py`와 `page_guard.py`도 모두 통과해야 한다
- `edit_guard.py` 실패 시 결과를 완료로 제출하지 않고, 원인(구조 변경/텍스트 손상/IDRef 불일치)을 수정 후 재빌드한다
- `page_guard.py` 실패 시 결과를 완료로 제출하지 않고, 원인(길이 과다/구조 변경)을 수정 후 재빌드한다
- 가능하면 한글(또는 사용자의 확인) 기준 최종 쪽수 값을 확인하고 레퍼런스와 일치 여부를 재확인한다

## 실행 명령

```bash
source "$VENV"

# 1) 레퍼런스 분석 + XML 추출
python3 "$SKILL_DIR/scripts/analyze_template.py" reference.hwpx \
  --extract-header /tmp/ref_header.xml \
  --extract-section /tmp/ref_section.xml

# 2) /tmp/ref_section.xml을 복제해 /tmp/new_section0.xml 작성
#    (구조 유지, 텍스트/데이터만 요청에 맞게 수정)

# 3) 복원 빌드
python3 "$SKILL_DIR/scripts/build_hwpx.py" \
  --header /tmp/ref_header.xml \
  --section /tmp/new_section0.xml \
  --output result.hwpx

# 4) 기본 + 심층 검증
python3 "$SKILL_DIR/scripts/validate.py" result.hwpx --deep

# 5) 편집 전후 무결성 비교 (필수)
python3 "$SKILL_DIR/scripts/edit_guard.py" \
  --before reference.hwpx \
  --after result.hwpx

# 6) 쪽수 드리프트 가드 (필수)
python3 "$SKILL_DIR/scripts/page_guard.py" \
  --reference reference.hwpx \
  --output result.hwpx
```

## 분석 출력 항목

| 항목 | 설명 |
|------|------|
| 폰트 정의 | hangul/latin 폰트 매핑 |
| borderFill | 테두리 타입/두께 + 배경색 (각 면별 상세) |
| charPr | 글꼴 크기(pt), 폰트명, 색상, 볼드/이탤릭/밑줄/취소선, fontRef |
| paraPr | 정렬, 줄간격, 여백(left/right/prev/next/intent), heading, borderFillIDRef |
| 문서 구조 | 페이지 크기, 여백, 페이지 테두리, 본문폭 |
| 본문 상세 | 모든 문단의 id/paraPr/charPr + 텍스트 내용 |
| 표 상세 | 행×열, 열너비 배열, 셀별 span/margin/borderFill/vertAlign + 내용 |

## 핵심 원칙

- **charPrIDRef/paraPrIDRef를 그대로 사용**: 추출한 header.xml의 스타일 ID를 변경하지 말 것
- **열 너비 합계 = 본문폭**: 분석 결과의 열너비 배열을 그대로 복제
- **rowSpan/colSpan 패턴 유지**: 분석된 셀 병합 구조를 정확히 재현
- **cellMargin 보존**: 분석된 셀 여백 값을 동일하게 적용
- **페이지 증가 금지**: 사용자 명시 승인 없이 결과 쪽수를 늘리지 말 것
- **치환 우선 편집**: 새 문단/표 추가보다 기존 텍스트 노드 치환을 우선할 것
