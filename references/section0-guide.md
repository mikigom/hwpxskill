# section0.xml 작성 가이드

## 필수 구조

section0.xml의 첫 문단(`<hp:p>`)의 첫 런(`<hp:run>`)에 반드시 `<hp:secPr>`과 `<hp:colPr>` 포함:

```xml
<hp:p id="1000000001" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:secPr ...>
      <!-- 페이지 크기, 여백, 각주/미주 설정 등 -->
    </hp:secPr>
    <hp:ctrl>
      <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/>
    </hp:ctrl>
  </hp:run>
  <hp:run charPrIDRef="0"><hp:t/></hp:run>
</hp:p>
```

**Tip**: `templates/base/Contents/section0.xml` 의 첫 문단을 그대로 복사하면 된다.

## 문단

```xml
<hp:p id="고유ID" paraPrIDRef="문단스타일ID" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="글자스타일ID">
    <hp:t>텍스트 내용</hp:t>
  </hp:run>
</hp:p>
```

## 빈 줄

```xml
<hp:p id="고유ID" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0"><hp:t/></hp:run>
</hp:p>
```

## 서식 혼합 런 (한 문단에 여러 스타일)

```xml
<hp:p id="고유ID" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0"><hp:t>일반 텍스트 </hp:t></hp:run>
  <hp:run charPrIDRef="7"><hp:t>볼드 텍스트</hp:t></hp:run>
  <hp:run charPrIDRef="0"><hp:t> 다시 일반</hp:t></hp:run>
</hp:p>
```

## 표 작성법

```xml
<hp:p id="고유ID" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:tbl id="고유ID" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM"
            textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL"
            repeatHeader="0" rowCnt="행수" colCnt="열수" cellSpacing="0"
            borderFillIDRef="3" noAdjust="0">
      <hp:sz width="42520" widthRelTo="ABSOLUTE" height="전체높이" heightRelTo="ABSOLUTE" protect="0"/>
      <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0"
              holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP"
              horzAlign="LEFT" vertOffset="0" horzOffset="0"/>
      <hp:outMargin left="0" right="0" top="0" bottom="0"/>
      <hp:inMargin left="0" right="0" top="0" bottom="0"/>
      <hp:tr>
        <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="1" borderFillIDRef="4">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                     linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"
                     hasTextRef="0" hasNumRef="0">
            <hp:p paraPrIDRef="21" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" id="고유ID">
              <hp:run charPrIDRef="9"><hp:t>헤더 셀</hp:t></hp:run>
            </hp:p>
          </hp:subList>
          <hp:cellAddr colAddr="0" rowAddr="0"/>
          <hp:cellSpan colSpan="1" rowSpan="1"/>
          <hp:cellSz width="열너비" height="행높이"/>
          <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
        </hp:tc>
        <!-- 나머지 셀... -->
      </hp:tr>
    </hp:tbl>
  </hp:run>
</hp:p>
```

## 표 크기 계산

- **A4 본문폭**: 42520 HWPUNIT = 59528(용지) - 8504×2(좌우여백)
- **열 너비 합 = 본문폭** (42520)
- 예: 3열 균등 → 14173 + 14173 + 14174 = 42520
- 예: 2열 (라벨:내용 = 1:4) → 8504 + 34016 = 42520
- **행 높이**: 셀당 보통 2400~3600 HWPUNIT

## ID 규칙

- 문단 id: `1000000001`부터 순차 증가
- 표 id: `1000000099` 등 별도 범위 사용 권장
- 모든 id는 문서 내 고유해야 함
