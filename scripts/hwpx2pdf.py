#!/usr/bin/env python3
"""HWPX → PDF 변환 스크립트.

세 가지 엔진을 지원한다:
  1. standalone — lxml + fpdf2 기반, 크로스플랫폼 (기본)
  2. hancom    — 한컴오피스 HwpCtrl COM (Windows + 한컴오피스)
  3. word      — MS Word COM via simple-hwp2pdf (Windows + MS Word)

auto 모드(기본): standalone → hancom → word 순으로 시도.

사용법:
  python hwpx2pdf.py input.hwpx                        # input.pdf 로 저장
  python hwpx2pdf.py input.hwpx -o output.pdf          # 지정 경로로 저장
  python hwpx2pdf.py input.hwpx --engine standalone    # standalone 강제 사용
  python hwpx2pdf.py input.hwpx --engine hancom        # 한컴오피스 강제 사용
  python hwpx2pdf.py input.hwpx --engine word          # MS Word 강제 사용

요구사항 (standalone):
  - Python 3.8+
  - lxml, fpdf2  (pip install lxml fpdf2)
  - 한글 TTF 폰트 (맑은 고딕, 나눔고딕 등 — 없으면 Helvetica fallback)
"""

import argparse
import os
import sys
import zipfile
from io import BytesIO

from lxml import etree


# ── 공통 유틸 ─────────────────────────────────────────────────────────


def _resolve_paths(input_path: str, output_path: str | None) -> tuple[str, str]:
    """입력/출력 경로를 절대 경로로 정규화한다."""
    input_abs = os.path.abspath(input_path)
    if not os.path.isfile(input_abs):
        print(f"[ERROR] 입력 파일을 찾을 수 없습니다: {input_abs}", file=sys.stderr)
        sys.exit(1)
    if output_path is None:
        output_abs = os.path.splitext(input_abs)[0] + ".pdf"
    else:
        output_abs = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_abs)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return input_abs, output_abs


# ── Standalone 엔진 (lxml + fpdf2) ───────────────────────────────────

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

# HWPUNIT → pt 변환 (1pt = 100 HWPUNIT)
_HU2PT = 0.01


def _hu2mm(hu: int) -> float:
    """HWPUNIT → mm."""
    return hu / 283.5


def _parse_hwpx(hwpx_path: str) -> dict:
    """HWPX에서 section0.xml과 header.xml을 파싱해 구조화된 데이터를 반환한다."""
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        names = zf.namelist()
        # section0.xml 경로 찾기
        section_path = None
        header_path = None
        for n in names:
            nl = n.lower()
            if nl.endswith("section0.xml"):
                section_path = n
            elif nl.endswith("header.xml"):
                header_path = n

        if section_path is None:
            raise ValueError("section0.xml을 찾을 수 없습니다.")

        section_xml = zf.read(section_path)
        header_xml = zf.read(header_path) if header_path else None

    # header.xml에서 charPr 정보 추출
    char_styles = {}
    if header_xml:
        header_tree = etree.fromstring(header_xml)
        for cp in header_tree.findall(".//hh:charPr", NS):
            cid = cp.get("id")
            if cid is None:
                continue
            cid = int(cid)
            height_hu = int(cp.get("height", "1000"))
            bold = cp.find("hh:bold", NS) is not None
            italic = cp.find("hh:italic", NS) is not None
            color = cp.get("textColor", "#000000")
            char_styles[cid] = {
                "size_pt": height_hu * _HU2PT,
                "bold": bold,
                "italic": italic,
                "color": color,
            }

    # paraPr에서 정렬 정보 추출
    para_styles = {}
    if header_xml:
        header_tree = etree.fromstring(header_xml)
        for pp in header_tree.findall(".//hh:paraPr", NS):
            pid = pp.get("id")
            if pid is None:
                continue
            pid = int(pid)
            align_elem = pp.find(".//hp:align", NS)
            if align_elem is None:
                align_elem = pp.find(".//hh:align", NS)
            align = "JUSTIFY"
            if align_elem is not None:
                align = align_elem.get("horizontal", "JUSTIFY")
            para_styles[pid] = {"align": align}

    # section0.xml 파싱
    section_tree = etree.fromstring(section_xml)

    # secPr에서 페이지 크기/여백 추출
    sec_pr = section_tree.find(".//hp:secPr", NS)
    page = {"width_mm": 210, "height_mm": 297,
            "left_mm": 30, "right_mm": 30, "top_mm": 25, "bottom_mm": 25}
    if sec_pr is not None:
        pg = sec_pr.find("hp:pg", NS)
        if pg is not None:
            page["width_mm"] = _hu2mm(int(pg.get("width", "59528")))
            page["height_mm"] = _hu2mm(int(pg.get("height", "84186")))
        margin = sec_pr.find("hp:margin", NS)
        if margin is not None:
            page["left_mm"] = _hu2mm(int(margin.get("left", "8504")))
            page["right_mm"] = _hu2mm(int(margin.get("right", "8504")))
            page["top_mm"] = _hu2mm(int(margin.get("top", "7087")))
            page["bottom_mm"] = _hu2mm(int(margin.get("bottom", "4252")))

    # 문단/표 순서대로 추출
    elements = []
    for child in section_tree:
        if not isinstance(child.tag, str):
            continue  # 주석, PI 등 건너뛰기
        tag = etree.QName(child).localname
        if tag == "p":
            p_data = _parse_paragraph(child, char_styles, para_styles)
            if p_data:
                elements.append(p_data)

    return {"page": page, "elements": elements, "char_styles": char_styles}


def _get_text_recursive(elem) -> str:
    """요소와 자식의 텍스트를 모두 합친다 (mixed content 대응)."""
    parts = []
    if elem.text:
        parts.append(elem.text)
    for child in elem:
        localname = etree.QName(child).localname
        if localname == "fwSpace":
            parts.append(" ")
        elif localname == "lineBreak":
            parts.append("\n")
        elif localname == "tab":
            parts.append("\t")
        # 자식 텍스트
        if child.text:
            parts.append(child.text)
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)


def _parse_paragraph(p_elem, char_styles: dict, para_styles: dict) -> dict | None:
    """<hp:p> 요소를 파싱하여 문단 데이터를 반환한다."""
    para_pr_id = int(p_elem.get("paraPrIDRef", "0"))
    align = "JUSTIFY"
    if para_pr_id in para_styles:
        align = para_styles[para_pr_id].get("align", "JUSTIFY")

    # 표 확인
    tbl = p_elem.find(".//hp:tbl", NS)
    if tbl is not None:
        table_data = _parse_table(tbl, char_styles, para_styles)
        return {"type": "table", "data": table_data}

    # 텍스트 런 추출
    runs = []
    for run in p_elem.findall(".//hp:run", NS):
        char_pr_id = int(run.get("charPrIDRef", "0"))
        style = char_styles.get(char_pr_id, {"size_pt": 10, "bold": False, "italic": False, "color": "#000000"})

        t_elem = run.find("hp:t", NS)
        if t_elem is not None:
            text = _get_text_recursive(t_elem)
        else:
            text = ""

        if text:
            runs.append({"text": text, **style})

    if not runs:
        return {"type": "empty_line", "align": align}

    return {"type": "paragraph", "runs": runs, "align": align}


def _parse_table(tbl_elem, char_styles: dict, para_styles: dict) -> dict:
    """<hp:tbl> 요소를 파싱해 표 데이터를 반환한다."""
    rows = []
    col_cnt = int(tbl_elem.get("colCnt", "1"))

    # 열 너비 비율 계산 (첫 행의 cellSz width 기준)
    col_widths = []

    for tr in tbl_elem.findall("hp:tr", NS):
        row_cells = []
        for tc in tr.findall("hp:tc", NS):
            # 셀 텍스트 추출
            texts = []
            for p in tc.findall(".//hp:p", NS):
                p_runs = []
                for run in p.findall(".//hp:run", NS):
                    t_elem = run.find("hp:t", NS)
                    if t_elem is not None:
                        text = _get_text_recursive(t_elem)
                        if text:
                            char_pr_id = int(run.get("charPrIDRef", "0"))
                            style = char_styles.get(char_pr_id, {
                                "size_pt": 10, "bold": False, "italic": False, "color": "#000000"
                            })
                            p_runs.append({"text": text, **style})
                if p_runs:
                    texts.append(p_runs)

            # 셀 span 정보
            span_elem = tc.find("hp:cellSpan", NS)
            col_span = int(span_elem.get("colSpan", "1")) if span_elem is not None else 1
            row_span = int(span_elem.get("rowSpan", "1")) if span_elem is not None else 1

            # 셀 너비
            sz_elem = tc.find("hp:cellSz", NS)
            cell_w = int(sz_elem.get("width", "0")) if sz_elem is not None else 0

            if not col_widths and cell_w > 0:
                col_widths.append(cell_w)

            row_cells.append({
                "texts": texts,
                "col_span": col_span,
                "row_span": row_span,
                "width_hu": cell_w,
            })
        rows.append(row_cells)

    return {"rows": rows, "col_cnt": col_cnt}


def _find_korean_font() -> tuple[str | None, str | None]:
    """시스템에서 한글 TTF 폰트를 찾아 (regular, bold) 경로를 반환한다."""
    font_dirs = []
    if sys.platform == "win32":
        font_dirs.append("C:/Windows/Fonts")
        local_fonts = os.path.join(os.path.expanduser("~"),
                                   "AppData/Local/Microsoft/Windows/Fonts")
        if os.path.isdir(local_fonts):
            font_dirs.append(local_fonts)
    elif sys.platform == "darwin":
        font_dirs.extend(["/Library/Fonts", os.path.expanduser("~/Library/Fonts")])
    else:
        font_dirs.extend([
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            os.path.expanduser("~/.local/share/fonts"),
        ])

    # 우선순위: 맑은고딕 > 나눔고딕 > NanumSquare
    candidates = [
        ("malgun.ttf", "malgunbd.ttf"),
        ("NanumGothic.ttf", "NanumGothicBold.ttf"),
    ]
    for regular, bold in candidates:
        for d in font_dirs:
            r_path = os.path.join(d, regular)
            b_path = os.path.join(d, bold)
            if os.path.isfile(r_path):
                return r_path, b_path if os.path.isfile(b_path) else r_path
    return None, None


def convert_with_standalone(input_path: str, output_path: str) -> None:
    """lxml + fpdf2 기반 standalone HWPX → PDF 변환.

    한컴오피스/MS Word 없이 동작한다. HWPX의 텍스트, 표, 기본 서식을
    파싱하여 PDF로 렌더링한다. 완벽한 레이아웃 재현은 아니지만
    visual inspection 용도로 충분하다.
    """
    from fpdf import FPDF

    doc = _parse_hwpx(input_path)
    pg = doc["page"]

    pdf = FPDF(orientation="P", unit="mm",
               format=(pg["width_mm"], pg["height_mm"]))
    pdf.set_auto_page_break(auto=True, margin=pg["bottom_mm"])

    # 한글 폰트 등록
    regular_font, bold_font = _find_korean_font()
    font_family = "Korean"
    if regular_font:
        pdf.add_font(font_family, "", regular_font)
        if bold_font and bold_font != regular_font:
            pdf.add_font(font_family, "B", bold_font)
        else:
            pdf.add_font(font_family, "B", regular_font)
    else:
        font_family = "Helvetica"

    pdf.add_page()
    pdf.set_margins(pg["left_mm"], pg["top_mm"], pg["right_mm"])
    pdf.set_y(pg["top_mm"])

    body_w = pg["width_mm"] - pg["left_mm"] - pg["right_mm"]

    align_map = {
        "LEFT": "L", "CENTER": "C", "RIGHT": "R", "JUSTIFY": "J",
    }

    for elem in doc["elements"]:
        if elem["type"] == "empty_line":
            pdf.set_font(font_family, "", 10)
            pdf.ln(5)

        elif elem["type"] == "paragraph":
            align = align_map.get(elem.get("align", "JUSTIFY"), "J")

            # 단일 런이면 간단히 처리
            if len(elem["runs"]) == 1:
                run = elem["runs"][0]
                style = "B" if run["bold"] else ""
                if run.get("italic"):
                    style += "I"
                size = max(run["size_pt"], 6)
                pdf.set_font(font_family, style, size)
                pdf.set_x(pg["left_mm"])
                pdf.multi_cell(body_w, size * 0.5, run["text"], align=align)
            else:
                # 멀티런: write()로 이어 쓰기
                max_size = max(r["size_pt"] for r in elem["runs"])
                line_h = max(max_size * 0.5, 5)
                pdf.set_x(pg["left_mm"])
                for i, run in enumerate(elem["runs"]):
                    style = "B" if run["bold"] else ""
                    if run.get("italic"):
                        style += "I"
                    size = max(run["size_pt"], 6)
                    pdf.set_font(font_family, style, size)
                    pdf.write(line_h, run["text"])
                pdf.ln(line_h)

        elif elem["type"] == "table":
            _render_table(pdf, elem["data"], font_family, body_w, pg["left_mm"])
            pdf.ln(2)

    pdf.output(output_path)


def _render_table(pdf, table_data: dict, font_family: str,
                  body_w: float, left_margin: float) -> None:
    """표를 PDF에 렌더링한다."""
    rows = table_data["rows"]
    if not rows:
        return

    # 열 너비 계산: 첫 행의 셀 width_hu 비율 기반
    first_row = rows[0]
    total_hu = sum(c["width_hu"] for c in first_row) or 1
    col_widths_mm = [(c["width_hu"] / total_hu) * body_w for c in first_row]

    # 표가 없는 셀이면 균등 분배
    if all(w == 0 for w in col_widths_mm):
        n = len(first_row) or 1
        col_widths_mm = [body_w / n] * n

    row_height = 7
    pdf.set_font(font_family, "", 9)

    for row in rows:
        pdf.set_x(left_margin)
        max_h = row_height

        # 셀 내용 높이 사전 계산
        cell_texts = []
        for ci, cell in enumerate(row):
            text_parts = []
            for para_runs in cell.get("texts", []):
                line = "".join(r["text"] for r in para_runs)
                text_parts.append(line)
            cell_text = "\n".join(text_parts) if text_parts else ""
            cell_texts.append(cell_text)

            # 높이 추정
            w = col_widths_mm[ci] if ci < len(col_widths_mm) else col_widths_mm[-1]
            if cell_text:
                lines = max(1, len(cell_text) / max(w / 2, 1))
                est_h = max(row_height, lines * 4.5)
                max_h = max(max_h, est_h)

        max_h = min(max_h, 50)  # 과도한 높이 제한

        for ci, cell in enumerate(row):
            w = col_widths_mm[ci] if ci < len(col_widths_mm) else col_widths_mm[-1]
            # 볼드 여부 (첫 행이거나 셀 스타일에 bold가 있으면)
            is_bold = False
            for para_runs in cell.get("texts", []):
                for r in para_runs:
                    if r.get("bold"):
                        is_bold = True
                        break
            pdf.set_font(font_family, "B" if is_bold else "", 9)
            pdf.cell(w, max_h, cell_texts[ci][:80], border=1, align="C")

        pdf.ln(max_h)


# ── 한컴오피스 COM 엔진 ──────────────────────────────────────────────


def convert_with_hancom(input_path: str, output_path: str) -> None:
    """한컴오피스 HwpCtrl COM 객체를 사용하여 HWPX → PDF 변환."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(input_path, "HWP", "forceopen:true")
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = output_path
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    finally:
        if hwp:
            hwp.Clear(1)
            hwp.Quit()
        pythoncom.CoUninitialize()


# ── MS Word COM 엔진 (simple-hwp2pdf) ────────────────────────────────


def convert_with_word(input_path: str, output_path: str) -> None:
    """simple-hwp2pdf 라이브러리를 사용하여 HWPX → PDF 변환 (MS Word COM)."""
    from simple_hwp2pdf import convert_hwp_to_pdf

    convert_hwp_to_pdf(input_path, output_path)


# ── 자동 선택 ─────────────────────────────────────────────────────────


ENGINE_MAP = {
    "standalone": ("Standalone (lxml+fpdf2)", convert_with_standalone),
    "hancom": ("한컴오피스 (HwpCtrl)", convert_with_hancom),
    "word": ("MS Word (simple-hwp2pdf)", convert_with_word),
}


def convert(input_path: str, output_path: str | None = None,
            engine: str | None = None) -> str:
    """HWPX → PDF 변환을 수행하고 출력 경로를 반환한다.

    Args:
        input_path: 변환할 HWPX 파일 경로.
        output_path: 저장할 PDF 파일 경로 (기본: 입력파일.pdf).
        engine: 'standalone' | 'hancom' | 'word' | None(자동).

    Returns:
        생성된 PDF 파일의 절대 경로.
    """
    input_abs, output_abs = _resolve_paths(input_path, output_path)

    if engine:
        label, fn = ENGINE_MAP[engine]
        print(f"[hwpx2pdf] 엔진: {label}")
        fn(input_abs, output_abs)
    else:
        # 자동: standalone → hancom → word
        errors = {}
        for key in ("standalone", "hancom", "word"):
            label, fn = ENGINE_MAP[key]
            try:
                print(f"[hwpx2pdf] {label} 시도...")
                fn(input_abs, output_abs)
                print(f"[hwpx2pdf] 엔진: {label}")
                break
            except Exception as e:
                errors[label] = e
                print(f"[hwpx2pdf] {label} 실패: {e}")
        else:
            msg = "[ERROR] 모든 엔진 실패.\n"
            for lbl, err in errors.items():
                msg += f"  {lbl}: {err}\n"
            print(msg, file=sys.stderr)
            sys.exit(1)

    if not os.path.isfile(output_abs):
        print(f"[ERROR] PDF 파일이 생성되지 않았습니다: {output_abs}", file=sys.stderr)
        sys.exit(1)

    size_kb = os.path.getsize(output_abs) / 1024
    print(f"[hwpx2pdf] 완료: {output_abs} ({size_kb:.1f} KB)")
    return output_abs


# ── CLI ───────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(
        description="HWPX 파일을 PDF로 변환합니다.",
    )
    parser.add_argument("input", help="변환할 HWPX 파일 경로")
    parser.add_argument("-o", "--output", help="출력 PDF 경로 (기본: 입력파일.pdf)")
    parser.add_argument(
        "--engine",
        choices=["standalone", "hancom", "word", "auto"],
        default="auto",
        help="변환 엔진 선택 (기본: auto → standalone 우선)",
    )
    args = parser.parse_args()
    engine = None if args.engine == "auto" else args.engine
    convert(args.input, args.output, engine)


if __name__ == "__main__":
    main()
