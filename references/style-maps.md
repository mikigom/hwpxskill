# 템플릿별 스타일 ID 맵

## base (기본)

| ID | 유형 | 설명 |
|----|------|------|
| charPr 0 | 글자 | 10pt 함초롬바탕, 기본 |
| charPr 1 | 글자 | 10pt 함초롬돋움 |
| charPr 2~6 | 글자 | Skeleton 기본 스타일 |
| paraPr 0 | 문단 | JUSTIFY, 160% 줄간격 |
| paraPr 1~19 | 문단 | Skeleton 기본 (개요, 각주 등) |
| borderFill 1 | 테두리 | 없음 (페이지 보더) |
| borderFill 2 | 테두리 | 없음 + 투명배경 (참조용) |

## gonmun (공문) — base + 추가

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 22pt 볼드 함초롬바탕 (기관명/제목) |
| charPr 8 | 글자 | 16pt 볼드 함초롬바탕 (서명자) |
| charPr 9 | 글자 | 8pt 함초롬바탕 (하단 연락처) |
| charPr 10 | 글자 | 10pt 볼드 함초롬바탕 (표 헤더) |
| paraPr 20 | 문단 | CENTER, 160% 줄간격 |
| paraPr 21 | 문단 | CENTER, 130% (표 셀) |
| paraPr 22 | 문단 | JUSTIFY, 130% (표 셀) |
| borderFill 3 | 테두리 | SOLID 0.12mm 4면 |
| borderFill 4 | 테두리 | SOLID 0.12mm + #D6DCE4 배경 |

## report (보고서) — base + 추가

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 20pt 볼드 (문서 제목) |
| charPr 8 | 글자 | 14pt 볼드 (소제목) |
| charPr 9 | 글자 | 10pt 볼드 (표 헤더) |
| charPr 10 | 글자 | 10pt 볼드+밑줄 (강조 텍스트) |
| charPr 11 | 글자 | 9pt 함초롬바탕 (소형/각주) |
| charPr 12 | 글자 | 16pt 볼드 함초롬바탕 (1줄 제목) |
| charPr 13 | 글자 | 12pt 볼드 함초롬돋움 (섹션 헤더) |
| paraPr 20~22 | 문단 | CENTER/JUSTIFY 변형 |
| paraPr 23 | 문단 | RIGHT 정렬, 160% 줄간격 |
| paraPr 24 | 문단 | JUSTIFY, left 600 (□ 체크항목 들여쓰기) |
| paraPr 25 | 문단 | JUSTIFY, left 1200 (하위항목 ①②③ 들여쓰기) |
| paraPr 26 | 문단 | JUSTIFY, left 1800 (깊은 하위항목 - 들여쓰기) |
| paraPr 27 | 문단 | LEFT, 상하단 테두리선 (섹션 헤더용), prev 400 |
| borderFill 3 | 테두리 | SOLID 0.12mm 4면 |
| borderFill 4 | 테두리 | SOLID 0.12mm + #DAEEF3 배경 |
| borderFill 5 | 테두리 | 상단 0.4mm 굵은선 + 하단 0.12mm 얇은선 (섹션 헤더) |

**들여쓰기 규칙**: 공백 문자가 아닌 반드시 paraPr의 left margin 사용. □ 항목은 paraPr 24, 하위 ①②③ 는 paraPr 25, 깊은 - 항목은 paraPr 26.

**섹션 헤더 규칙**: paraPr 27 + charPr 13 조합. 문단 테두리(borderFillIDRef="5")로 상단 굵은선 + 하단 얇은선 자동 표시.

## minutes (회의록) — base + 추가

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 18pt 볼드 (제목) |
| charPr 8 | 글자 | 12pt 볼드 (섹션 라벨) |
| charPr 9 | 글자 | 10pt 볼드 (표 헤더) |
| paraPr 20~22 | 문단 | CENTER/JUSTIFY 변형 |
| borderFill 3 | 테두리 | SOLID 0.12mm 4면 |
| borderFill 4 | 테두리 | SOLID 0.12mm + #E2EFDA 배경 |

## proposal (제안서/사업개요) — base + 추가

시각적 구분이 필요한 공식 문서용. 색상 배경 헤더바와 번호 배지를 표(table) 기반 레이아웃으로 구현.

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 20pt 볼드 함초롬바탕 (문서 제목) |
| charPr 8 | 글자 | 14pt 볼드 함초롬바탕 (소제목) |
| charPr 9 | 글자 | 10pt 볼드 함초롬바탕 (표 헤더) |
| charPr 10 | 글자 | 14pt 볼드 흰색 함초롬돋움 (대항목 번호, 녹색 배경) |
| charPr 11 | 글자 | 11pt 볼드 흰색 함초롬돋움 (소항목 번호, 파란 배경) |
| paraPr 20 | 문단 | CENTER, 160% 줄간격 |
| paraPr 21 | 문단 | CENTER, 130% (표 셀) |
| paraPr 22 | 문단 | JUSTIFY, 130% (표 셀) |
| borderFill 3 | 테두리 | SOLID 0.12mm 4면 |
| borderFill 4 | 테두리 | SOLID 0.12mm + #DAEEF3 배경 |
| borderFill 5 | 테두리 | 올리브녹색 배경 #7B8B3D (대항목 번호 셀) |
| borderFill 6 | 테두리 | 연한 회색 배경 #F2F2F2 + 회색 테두리 (대항목 제목 셀) |
| borderFill 7 | 테두리 | 파란색 배경 #4472C4 (소항목 번호 배지) |
| borderFill 8 | 테두리 | 하단 테두리만 #D0D0D0 (소항목 제목 영역) |

### proposal 레이아웃 패턴

**대항목 헤더** (2셀 표: 번호 + 제목):
```xml
<!-- borderFillIDRef="5" + charPrIDRef="10" → 녹색배경 흰색 로마숫자 -->
<!-- borderFillIDRef="6" + charPrIDRef="8"  → 회색배경 검정 볼드 제목 -->
```

**소항목 헤더** (2셀 표: 번호배지 + 제목):
```xml
<!-- borderFillIDRef="7" + charPrIDRef="11" → 파란배경 흰색 아라비아숫자 -->
<!-- borderFillIDRef="8" + charPrIDRef="8"  → 하단선만 검정 볼드 제목 -->
```
