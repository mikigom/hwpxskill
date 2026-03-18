#!/usr/bin/env python3
"""Unpack an HWPX file into a directory with pretty-printed XML.

Usage:
    python unpack.py input.hwpx output_dir/
"""

import argparse
import os
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


def _save_mixed_content(root):
    """Save text/tail of mixed-content elements before indent corrupts them.

    A mixed-content element has children AND non-whitespace text or tail,
    e.g. <hp:t>TEXT<hp:fwSpace/></hp:t>.  etree.indent() inserts whitespace
    into the None/.tail slots of such elements, which Hancom Office then
    renders as visible spacing.
    """
    saved = []
    for elem in root.iter():
        if len(elem) == 0:
            continue
        has_text = elem.text is not None and elem.text.strip()
        has_tail_text = any(
            child.tail is not None and child.tail.strip() for child in elem
        )
        if has_text or has_tail_text:
            saved.append(
                (elem, elem.text, [(child, child.tail) for child in elem])
            )
    return saved


def _restore_mixed_content(saved):
    """Restore text/tail that etree.indent() overwrote."""
    for elem, text, child_tails in saved:
        elem.text = text
        for child, tail in child_tails:
            child.tail = tail


def unpack(hwpx_path: str, output_dir: str) -> None:
    """Extract HWPX archive and pretty-print all XML files."""

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)

    with ZipFile(hwpx_path, "r") as zf:
        for entry in zf.namelist():
            data = zf.read(entry)
            dest = output / entry
            dest.parent.mkdir(parents=True, exist_ok=True)

            if entry.endswith(".xml") or entry.endswith(".hpf"):
                try:
                    tree = etree.fromstring(data)
                    saved = _save_mixed_content(tree)
                    etree.indent(tree, space="  ")
                    _restore_mixed_content(saved)
                    pretty = etree.tostring(
                        tree,
                        pretty_print=True,
                        xml_declaration=True,
                        encoding="UTF-8",
                    )
                    dest.write_bytes(pretty)
                    continue
                except etree.XMLSyntaxError:
                    pass  # Fall through to raw write

            dest.write_bytes(data)

    print(f"Unpacked: {hwpx_path} -> {output_dir}")
    print(f"  Files: {len(list(output.rglob('*')))} entries")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Unpack HWPX file into a directory with pretty-printed XML"
    )
    parser.add_argument("input", help="Path to .hwpx file")
    parser.add_argument("output", help="Output directory path")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    unpack(args.input, args.output)


if __name__ == "__main__":
    main()
