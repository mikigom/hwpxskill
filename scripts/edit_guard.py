#!/usr/bin/env python3
"""edit_guard.py - HWPX 편집 전후 무결성 비교 검증

before.hwpx와 after.hwpx를 비교하여 의도하지 않은 구조 변경이나
콘텐츠 손상(garbling, duplication, mixed-content corruption)을 탐지한다.

검사 항목:
  1. header.xml itemCnt ↔ 실제 자식 수 정합성
  2. section0.xml의 charPrIDRef/paraPrIDRef/borderFillIDRef → header.xml 존재 여부
  3. 편집 전후 XML 구조 비교 (문단/표/run 수 변화)
  4. Mixed-content 손상 탐지 (hp:t 내부 비정상 공백)
  5. 텍스트 중복/소실 탐지 (before 텍스트와 after 텍스트 비교)
  6. secPr/colPr 보존 확인

Usage:
    # 편집 전후 비교 (full guard)
    python edit_guard.py --before before.hwpx --after after.hwpx

    # after만 단독 검증 (구조 정합성만)
    python edit_guard.py --after result.hwpx
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

# itemCnt를 가지는 header.xml 컨테이너 요소 (local name → 자식 local name)
ITEMCNT_CONTAINERS = {
    "fontfaces": "fontface",
    "charProperties": "charPr",
    "paraProperties": "paraPr",
    "borderFills": "borderFill",
    "tabProperties": "tabPr",
    "numberingProperties": "numbering",
    "bulletProperties": "bullet",
    "styles": "style",
    "memoProperties": "memoPr",
    "trackChangeAuthors": "trackChangeAuthor",
    "trackChanges": "trackChange",
}


@dataclass
class GuardResult:
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return len(self.errors) == 0

    def error(self, msg: str) -> None:
        self.errors.append(msg)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)


def _read_xml(hwpx_path: Path, inner_path: str) -> Optional[etree._Element]:
    """Read and parse an XML file from inside an HWPX ZIP."""
    try:
        with zipfile.ZipFile(hwpx_path, "r") as zf:
            if inner_path not in zf.namelist():
                return None
            data = zf.read(inner_path)
            return etree.parse(BytesIO(data)).getroot()
    except (zipfile.BadZipFile, etree.XMLSyntaxError):
        return None


def _collect_text(root: etree._Element) -> str:
    """Collect all text from hp:t elements."""
    texts = []
    for t in root.xpath(".//hp:t", namespaces=NS):
        texts.append("".join(t.itertext()))
    return "".join(texts)


def _collect_paragraph_texts(root: etree._Element) -> List[str]:
    """Collect text per top-level paragraph."""
    paragraphs = root.xpath(".//hs:sec/hp:p", namespaces=NS)
    if not paragraphs:
        paragraphs = root.xpath(".//hp:p", namespaces=NS)
    result = []
    for p in paragraphs:
        ptexts = []
        for t in p.xpath(".//hp:t", namespaces=NS):
            ptexts.append("".join(t.itertext()))
        result.append("".join(ptexts))
    return result


# ---------------------------------------------------------------------------
# Check 1: header.xml itemCnt consistency
# ---------------------------------------------------------------------------
def check_itemcnt(header: etree._Element, result: GuardResult) -> None:
    """Verify itemCnt attributes match actual child counts in header.xml."""
    for elem in header.iter():
        local = etree.QName(elem.tag).localname
        if local in ITEMCNT_CONTAINERS:
            declared = elem.get("itemCnt")
            if declared is None:
                continue
            try:
                declared_int = int(declared)
            except ValueError:
                result.error(f"header.xml: {local}의 itemCnt가 숫자가 아님: '{declared}'")
                continue
            actual = len(list(elem))
            if declared_int != actual:
                result.error(
                    f"header.xml: {local}의 itemCnt={declared_int}이나 "
                    f"실제 자식 수={actual} (불일치)"
                )


# ---------------------------------------------------------------------------
# Check 2: section → header ID reference consistency
# ---------------------------------------------------------------------------
def _collect_header_ids(header: etree._Element, tag_local: str) -> Set[int]:
    """Collect all id values for a given element type in header.xml."""
    ids: Set[int] = set()
    for elem in header.iter():
        if etree.QName(elem.tag).localname == tag_local:
            id_val = elem.get("id")
            if id_val is not None:
                try:
                    ids.add(int(id_val))
                except ValueError:
                    pass
    return ids


def check_id_refs(
    section: etree._Element, header: etree._Element, result: GuardResult
) -> None:
    """Verify all IDRef values in section0.xml exist in header.xml."""
    char_ids = _collect_header_ids(header, "charPr")
    para_ids = _collect_header_ids(header, "paraPr")
    border_ids = _collect_header_ids(header, "borderFill")

    ref_checks = [
        ("charPrIDRef", char_ids, "charPr"),
        ("paraPrIDRef", para_ids, "paraPr"),
        ("borderFillIDRef", border_ids, "borderFill"),
    ]

    for attr_name, valid_ids, target_name in ref_checks:
        missing: Counter[int] = Counter()
        for elem in section.iter():
            val = elem.get(attr_name)
            if val is not None:
                try:
                    ref_id = int(val)
                except ValueError:
                    continue
                if ref_id not in valid_ids:
                    missing[ref_id] += 1

        for ref_id, count in missing.items():
            result.error(
                f"section0.xml: {attr_name}={ref_id}이 header.xml의 "
                f"{target_name}에 없음 ({count}회 참조)"
            )


# ---------------------------------------------------------------------------
# Check 3: Mixed-content corruption detection
# ---------------------------------------------------------------------------
def check_mixed_content(section: etree._Element, result: GuardResult) -> None:
    """Detect mixed-content corruption in hp:t elements.

    If a hp:t element has children (like hp:fwSpace, hp:lineBreak) AND
    the text/tail contains suspicious patterns (leading/trailing whitespace
    with newlines), it's likely been corrupted by pretty-printing.
    """
    corruption_count = 0
    for t_elem in section.xpath(".//hp:t", namespaces=NS):
        children = list(t_elem)
        if not children:
            continue  # No mixed content, safe

        # Check elem.text for indentation injection
        if t_elem.text is not None and re.search(r"\n\s+", t_elem.text):
            corruption_count += 1
            continue

        # Check child tails for indentation injection
        for child in children:
            if child.tail is not None and re.search(r"\n\s+", child.tail):
                corruption_count += 1
                break

    if corruption_count > 0:
        result.error(
            f"section0.xml: Mixed-content 손상 탐지 — "
            f"hp:t 요소 {corruption_count}개에서 비정상 공백/개행 발견 "
            f"(pretty-print로 인한 텍스트 손상 가능성)"
        )


# ---------------------------------------------------------------------------
# Check 4: secPr/colPr preservation
# ---------------------------------------------------------------------------
def check_secpr_preserved(section: etree._Element, result: GuardResult) -> None:
    """Verify the first paragraph contains secPr and colPr."""
    paragraphs = section.xpath(".//hs:sec/hp:p", namespaces=NS)
    if not paragraphs:
        paragraphs = section.xpath(".//hp:p", namespaces=NS)
    if not paragraphs:
        result.error("section0.xml: 문단이 하나도 없음")
        return

    first_p = paragraphs[0]
    secpr = first_p.xpath(".//hp:secPr", namespaces=NS)
    if not secpr:
        result.error("section0.xml: 첫 문단에 secPr이 없음 (페이지 설정 누락)")

    colpr = first_p.xpath(".//hp:colPr", namespaces=NS)
    if not colpr:
        result.warn("section0.xml: 첫 문단에 colPr이 없음 (단 설정 누락 가능)")


# ---------------------------------------------------------------------------
# Check 5: Before↔After structural comparison
# ---------------------------------------------------------------------------
@dataclass
class StructureSnapshot:
    paragraph_count: int
    table_count: int
    run_count: int
    page_break_count: int
    column_break_count: int
    table_shapes: List[Tuple[str, str]]  # (rowCnt, colCnt)


def _snapshot(section: etree._Element) -> StructureSnapshot:
    paragraphs = section.xpath(".//hs:sec/hp:p", namespaces=NS)
    if not paragraphs:
        paragraphs = section.xpath(".//hp:p", namespaces=NS)

    tables = section.xpath(".//hp:tbl", namespaces=NS)
    runs = section.xpath(".//hp:run", namespaces=NS)

    return StructureSnapshot(
        paragraph_count=len(paragraphs),
        table_count=len(tables),
        run_count=len(runs),
        page_break_count=sum(1 for p in paragraphs if p.get("pageBreak") == "1"),
        column_break_count=sum(1 for p in paragraphs if p.get("columnBreak") == "1"),
        table_shapes=[
            (t.get("rowCnt", "?"), t.get("colCnt", "?")) for t in tables
        ],
    )


def check_structure_diff(
    before_section: etree._Element,
    after_section: etree._Element,
    result: GuardResult,
) -> None:
    """Compare structural snapshots of before and after section0.xml."""
    b = _snapshot(before_section)
    a = _snapshot(after_section)

    if b.paragraph_count != a.paragraph_count:
        result.error(
            f"구조 변경: 문단 수 {b.paragraph_count} → {a.paragraph_count} "
            f"(사용자 요청 없이 문단 추가/삭제 금지)"
        )
    if b.table_count != a.table_count:
        result.error(
            f"구조 변경: 표 수 {b.table_count} → {a.table_count} "
            f"(사용자 요청 없이 표 추가/삭제 금지)"
        )
    if b.page_break_count != a.page_break_count:
        result.error(
            f"구조 변경: pageBreak 수 {b.page_break_count} → {a.page_break_count}"
        )
    if b.column_break_count != a.column_break_count:
        result.error(
            f"구조 변경: columnBreak 수 {b.column_break_count} → {a.column_break_count}"
        )
    if b.table_shapes != a.table_shapes:
        result.error("구조 변경: 표 구조(rowCnt/colCnt)가 변경됨")

    # Run count changes can indicate structural corruption
    if b.run_count > 0:
        run_ratio = abs(a.run_count - b.run_count) / b.run_count
        if run_ratio > 0.3:
            result.warn(
                f"run 수 급변: {b.run_count} → {a.run_count} "
                f"(변화율 {run_ratio:.0%}, 구조 손상 가능성)"
            )


# ---------------------------------------------------------------------------
# Check 6: Text duplication / loss detection
# ---------------------------------------------------------------------------
def check_text_integrity(
    before_section: etree._Element,
    after_section: etree._Element,
    result: GuardResult,
) -> None:
    """Detect text duplication or significant loss between before and after."""
    before_texts = _collect_paragraph_texts(before_section)
    after_texts = _collect_paragraph_texts(after_section)

    # Check for exact paragraph duplication in after (a paragraph appearing
    # identically where it shouldn't)
    if len(before_texts) == len(after_texts):
        for idx, (bt, at) in enumerate(zip(before_texts, after_texts), start=1):
            bt_stripped = "".join(bt.split())
            at_stripped = "".join(at.split())
            if not bt_stripped and not at_stripped:
                continue

            # Detect if text was doubled (content appears twice)
            if at_stripped and len(at_stripped) >= 2 * len(bt_stripped) and bt_stripped:
                half = len(at_stripped) // 2
                first_half = at_stripped[:half]
                second_half = at_stripped[half:]
                if first_half == second_half:
                    result.error(
                        f"{idx}번째 문단: 텍스트가 중복됨 "
                        f"(길이 {len(bt_stripped)} → {len(at_stripped)}, 동일 텍스트 반복)"
                    )
                    continue

            # Detect if text was completely replaced with something very different
            # (could indicate garbling)
            if bt_stripped and at_stripped:
                # Check for character-level overlap
                common = sum(1 for c in at_stripped if c in bt_stripped)
                overlap = common / max(len(at_stripped), 1)
                if overlap < 0.3 and len(bt_stripped) > 10:
                    result.warn(
                        f"{idx}번째 문단: 텍스트가 크게 변경됨 "
                        f"(원본과의 문자 겹침 {overlap:.0%})"
                    )

    # Global text length check
    before_total = len("".join("".join(t.split()) for t in before_texts))
    after_total = len("".join("".join(t.split()) for t in after_texts))
    if before_total > 0:
        change_ratio = abs(after_total - before_total) / before_total
        if change_ratio > 0.5:
            result.error(
                f"전체 텍스트 길이 급변: {before_total} → {after_total} "
                f"(변화율 {change_ratio:.0%}, 50% 초과 — 콘텐츠 손상 가능성)"
            )
        elif change_ratio > 0.25:
            result.warn(
                f"전체 텍스트 길이 변화: {before_total} → {after_total} "
                f"(변화율 {change_ratio:.0%})"
            )


# ---------------------------------------------------------------------------
# Check 7: header.xml preservation (before vs after)
# ---------------------------------------------------------------------------
def check_header_preserved(
    before_header: etree._Element,
    after_header: etree._Element,
    result: GuardResult,
) -> None:
    """Warn if header.xml was significantly modified (styles should be preserved)."""
    def count_children(root: etree._Element, local_name: str) -> int:
        return sum(1 for e in root.iter() if etree.QName(e.tag).localname == local_name)

    for tag in ["charPr", "paraPr", "borderFill", "fontface"]:
        bc = count_children(before_header, tag)
        ac = count_children(after_header, tag)
        if bc != ac:
            result.warn(
                f"header.xml: {tag} 수 변경 {bc} → {ac} "
                f"(레퍼런스 스타일 변경은 주의 필요)"
            )


# ---------------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------------
def run_guard(
    before_path: Optional[Path],
    after_path: Path,
) -> GuardResult:
    result = GuardResult()

    # Load after (required)
    after_header = _read_xml(after_path, "Contents/header.xml")
    after_section = _read_xml(after_path, "Contents/section0.xml")

    if after_header is None:
        result.error(f"after({after_path}): header.xml을 읽을 수 없음")
        return result
    if after_section is None:
        result.error(f"after({after_path}): section0.xml을 읽을 수 없음")
        return result

    # Always run: structural integrity of after
    check_itemcnt(after_header, result)
    check_id_refs(after_section, after_header, result)
    check_mixed_content(after_section, result)
    check_secpr_preserved(after_section, result)

    # If before is provided, run comparative checks
    if before_path is not None:
        before_header = _read_xml(before_path, "Contents/header.xml")
        before_section = _read_xml(before_path, "Contents/section0.xml")

        if before_header is None:
            result.error(f"before({before_path}): header.xml을 읽을 수 없음")
        if before_section is None:
            result.error(f"before({before_path}): section0.xml을 읽을 수 없음")

        if before_section is not None and after_section is not None:
            check_structure_diff(before_section, after_section, result)
            check_text_integrity(before_section, after_section, result)

        if before_header is not None and after_header is not None:
            check_header_preserved(before_header, after_header, result)

    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description="HWPX 편집 전후 무결성 비교 검증 (edit guard)"
    )
    parser.add_argument(
        "--before", "-b",
        type=Path,
        help="편집 전 HWPX 경로 (생략 시 after 단독 검증)",
    )
    parser.add_argument(
        "--after", "-a",
        type=Path,
        required=True,
        help="편집 후 HWPX 경로",
    )
    args = parser.parse_args()

    if args.before and not args.before.exists():
        print(f"Error: before not found: {args.before}", file=sys.stderr)
        return 2
    if not args.after.exists():
        print(f"Error: after not found: {args.after}", file=sys.stderr)
        return 2

    result = run_guard(args.before, args.after)

    # Print results
    if result.warnings:
        print("WARNINGS:")
        for w in result.warnings:
            print(f"  [WARN]{w}")

    if result.errors:
        print("FAIL: edit-guard")
        for e in result.errors:
            print(f"  [ERR]{e}")
        return 1

    print("PASS: edit-guard")
    if args.before:
        print("  편집 전후 구조 정합성 및 콘텐츠 무결성 검증 통과.")
    else:
        print("  구조 정합성 검증 통과 (단독 검증 모드).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
