#!/usr/bin/env python3
"""Validate the structural integrity of an HWPX file.

Checks:
  - Valid ZIP archive
  - Required files present (mimetype, content.hpf, header.xml, section0.xml)
  - mimetype content is correct
  - mimetype is the first ZIP entry and stored without compression
  - All XML files are well-formed
  - (--deep) header.xml itemCnt ↔ 실제 자식 수 정합성
  - (--deep) section0.xml IDRef → header.xml 존재 여부
  - (--deep) Mixed-content 손상 탐지
  - (--deep) secPr/colPr 보존 확인

Usage:
    python validate.py document.hwpx
    python validate.py document.hwpx --deep
"""

import argparse
import sys
from pathlib import Path
from zipfile import ZIP_STORED, BadZipFile, ZipFile

from lxml import etree

REQUIRED_FILES = [
    "mimetype",
    "Contents/content.hpf",
    "Contents/header.xml",
    "Contents/section0.xml",
]

EXPECTED_MIMETYPE = "application/hwp+zip"


def validate(hwpx_path: str, deep: bool = False) -> list[str]:
    """Validate HWPX file and return a list of error messages (empty = valid)."""

    errors: list[str] = []
    path = Path(hwpx_path)

    if not path.is_file():
        return [f"File not found: {hwpx_path}"]

    # Check valid ZIP
    try:
        zf = ZipFile(hwpx_path, "r")
    except BadZipFile:
        return [f"Not a valid ZIP archive: {hwpx_path}"]

    with zf:
        names = zf.namelist()

        # Check required files
        for required in REQUIRED_FILES:
            if required not in names:
                errors.append(f"Missing required file: {required}")

        # Check mimetype content
        if "mimetype" in names:
            mimetype_content = zf.read("mimetype").decode("utf-8").strip()
            if mimetype_content != EXPECTED_MIMETYPE:
                errors.append(
                    f"Invalid mimetype: expected '{EXPECTED_MIMETYPE}', "
                    f"got '{mimetype_content}'"
                )

            # Check mimetype is first entry
            if names[0] != "mimetype":
                errors.append(
                    f"mimetype is not the first ZIP entry (found at index "
                    f"{names.index('mimetype')})"
                )

            # Check mimetype is stored without compression
            info = zf.getinfo("mimetype")
            if info.compress_type != ZIP_STORED:
                errors.append(
                    f"mimetype should use ZIP_STORED (0), "
                    f"got compress_type={info.compress_type}"
                )

        # Check XML well-formedness
        for name in names:
            if name.endswith(".xml") or name.endswith(".hpf"):
                try:
                    data = zf.read(name)
                    etree.fromstring(data)
                except etree.XMLSyntaxError as e:
                    errors.append(f"Malformed XML in {name}: {e}")

        # Deep validation: structural integrity checks
        if deep and not errors:
            try:
                from edit_guard import (
                    check_id_refs,
                    check_itemcnt,
                    check_mixed_content,
                    check_secpr_preserved,
                    GuardResult,
                )
            except ImportError:
                # Fallback: try importing from same directory
                import importlib.util
                spec = importlib.util.spec_from_file_location(
                    "edit_guard",
                    Path(__file__).parent / "edit_guard.py",
                )
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                check_id_refs = mod.check_id_refs
                check_itemcnt = mod.check_itemcnt
                check_mixed_content = mod.check_mixed_content
                check_secpr_preserved = mod.check_secpr_preserved
                GuardResult = mod.GuardResult

            from io import BytesIO

            header_data = zf.read("Contents/header.xml")
            section_data = zf.read("Contents/section0.xml")
            header_root = etree.parse(BytesIO(header_data)).getroot()
            section_root = etree.parse(BytesIO(section_data)).getroot()

            guard = GuardResult()
            check_itemcnt(header_root, guard)
            check_id_refs(section_root, header_root, guard)
            check_mixed_content(section_root, guard)
            check_secpr_preserved(section_root, guard)

            errors.extend(guard.errors)

    return errors


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Validate the structural integrity of an HWPX file"
    )
    parser.add_argument("input", help="Path to .hwpx file")
    parser.add_argument(
        "--deep",
        action="store_true",
        help="Enable deep validation (itemCnt, IDRef, mixed-content, secPr checks)",
    )
    args = parser.parse_args()

    errors = validate(args.input, deep=args.deep)

    if errors:
        print(f"INVALID: {args.input}", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"VALID: {args.input}")
        if args.deep:
            print(f"  All structural + deep integrity checks passed.")
        else:
            print(f"  All structural checks passed.")


if __name__ == "__main__":
    main()
