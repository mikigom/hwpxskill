# XML-first 문서 생성 워크플로우

레퍼런스 파일이 없을 때만 사용하는 보조 워크플로우.

> 원칙: 사용자가 레퍼런스 HWPX를 제공한 경우에는 이 워크플로우 대신 `workflow-reference-restore.md`를 사용한다.

## 흐름

1. **템플릿 선택** (base/gonmun/report/minutes/proposal)
2. **section0.xml 작성** (본문 내용) → `section0-guide.md` 참조
3. **(선택) header.xml 수정** (새 스타일 추가 필요 시) → `header-guide.md` 참조
4. **build_hwpx.py로 빌드**
5. **validate.py로 검증**

## 기본 사용법

```bash
source "$VENV"

# 빈 문서 (base 템플릿)
python3 "$SKILL_DIR/scripts/build_hwpx.py" --output result.hwpx

# 템플릿 사용
python3 "$SKILL_DIR/scripts/build_hwpx.py" --template gonmun --output result.hwpx

# 커스텀 section0.xml 오버라이드
python3 "$SKILL_DIR/scripts/build_hwpx.py" --template gonmun --section my_section0.xml --output result.hwpx

# header도 오버라이드
python3 "$SKILL_DIR/scripts/build_hwpx.py" --header my_header.xml --section my_section0.xml --output result.hwpx

# 메타데이터 설정
python3 "$SKILL_DIR/scripts/build_hwpx.py" --template report --section my.xml \
  --title "제목" --creator "작성자" --output result.hwpx
```

## 실전 패턴: section0.xml을 인라인 작성 → 빌드

```bash
# 1. section0.xml을 임시파일로 작성
SECTION=$(mktemp /tmp/section0_XXXX.xml)
cat > "$SECTION" << 'XMLEOF'
<?xml version='1.0' encoding='UTF-8'?>
<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
  <!-- secPr 포함 첫 문단 (base/section0.xml에서 복사) -->
  <!-- ... -->
  <hp:p id="1000000002" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
    <hp:run charPrIDRef="0">
      <hp:t>본문 내용</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>
XMLEOF

# 2. 빌드
python3 "$SKILL_DIR/scripts/build_hwpx.py" --section "$SECTION" --output result.hwpx

# 3. 정리
rm -f "$SECTION"
```
