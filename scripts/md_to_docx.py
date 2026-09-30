#!/usr/bin/env python3
"""Convert a Markdown report to a Word document.

Usage: python scripts/md_to_docx.py report.md -o report.docx

Pandoc is used when available. The dependency-free fallback handles the basic
Markdown used by CORE reports: headings, paragraphs, lists, tables, emphasis,
code, and visible link URLs. It does not embed images or support full Markdown.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", W)


def _w(name: str) -> str:
    return f"{{{W}}}{name}"


def _el(parent: ET.Element, name: str, **attrs: str) -> ET.Element:
    return ET.SubElement(parent, _w(name), {_w(key): value for key, value in attrs.items()})


def _run(paragraph: ET.Element, value: str, *, bold: bool = False,
         italic: bool = False, code: bool = False) -> None:
    if not value:
        return
    run = _el(paragraph, "r")
    if bold or italic or code:
        properties = _el(run, "rPr")
        if bold:
            _el(properties, "b")
        if italic:
            _el(properties, "i")
        if code:
            _el(properties, "rFonts", ascii="Consolas", hAnsi="Consolas")
    element = _el(run, "t")
    if value[0].isspace() or value[-1].isspace():
        element.set(f"{{{XML}}}space", "preserve")
    element.text = value


INLINE = re.compile(
    r"(!?\[[^\]]+\]\([^)]*\)|`[^`]+`|\*\*[^*]+\*\*|__[^_]+__|\*[^*\n]+\*|_[^_\n]+_)"
)


def _inline(paragraph: ET.Element, text: str, *, bold: bool = False) -> None:
    """Render the small inline Markdown subset used in report templates."""
    cursor = 0
    for match in INLINE.finditer(text):
        _run(paragraph, text[cursor:match.start()], bold=bold)
        token = match.group()
        if token.startswith("!["):
            label, address = token[2:-1].split("](", 1)
            _run(paragraph, f"[Изображение: {label}] ({address})", bold=bold)
        elif token.startswith("["):
            label, address = token[1:-1].split("](", 1)
            _run(paragraph, f"{label} ({address})" if address else label, bold=bold)
        elif token.startswith("`"):
            _run(paragraph, token[1:-1], bold=bold, code=True)
        elif token.startswith(("**", "__")):
            _run(paragraph, token[2:-2], bold=True)
        else:
            _run(paragraph, token[1:-1], bold=bold, italic=True)
        cursor = match.end()
    _run(paragraph, text[cursor:], bold=bold)


def _paragraph(parent: ET.Element, text: str, *, style: str | None = None,
               indent: int = 0, bold: bool = False) -> ET.Element:
    paragraph = _el(parent, "p")
    if style or indent:
        properties = _el(paragraph, "pPr")
        if style:
            _el(properties, "pStyle", val=style)
        if indent:
            _el(properties, "ind", left=str(indent))
    _inline(paragraph, text, bold=bold)
    return paragraph


def _cells(line: str) -> list[str]:
    return [cell.strip().replace(r"\|", "|") for cell in re.split(r"(?<!\\)\|", line.strip().strip("|"))]


def _is_table_separator(line: str) -> bool:
    cells = _cells(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell) for cell in cells)


def _table(parent: ET.Element, rows: list[list[str]]) -> None:
    count = max(len(row) for row in rows)
    width = 9000 // count
    table = _el(parent, "tbl")
    properties = _el(table, "tblPr")
    _el(properties, "tblW", w="0", type="auto")
    borders = _el(properties, "tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _el(borders, edge, val="single", sz="4", color="D9DEE6")
    grid = _el(table, "tblGrid")
    for _ in range(count):
        _el(grid, "gridCol", w=str(width))
    for row_number, values in enumerate(rows):
        row = _el(table, "tr")
        for value in values + [""] * (count - len(values)):
            cell = _el(row, "tc")
            cell_properties = _el(cell, "tcPr")
            _el(cell_properties, "tcW", w=str(width), type="dxa")
            if row_number == 0:
                _el(cell_properties, "shd", fill="EEF3FA")
            _paragraph(cell, value, bold=row_number == 0)


HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
LIST = re.compile(r"^(\s*)([-*+] |\d+[.)] )(.+)$")
FENCE = re.compile(r"^\s*(`{3,}|~{3,})")


def _document(markdown: str) -> bytes:
    root = ET.Element(_w("document"))
    body = _el(root, "body")
    lines = markdown.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue

        fence = FENCE.match(line)
        if fence:
            marker = fence.group(1)
            index += 1
            while index < len(lines) and not lines[index].strip().startswith(marker):
                _paragraph(body, lines[index], style="Code")
                index += 1
            index += 1
            continue

        if index + 1 < len(lines) and "|" in line and _is_table_separator(lines[index + 1]):
            rows = [_cells(line)]
            index += 2
            while index < len(lines) and lines[index].strip() and "|" in lines[index]:
                rows.append(_cells(lines[index]))
                index += 1
            _table(body, rows)
            continue

        heading = HEADING.match(line)
        if heading:
            _paragraph(body, heading.group(2), style=f"Heading{min(len(heading.group(1)), 3)}")
            index += 1
            continue

        list_item = LIST.match(line)
        if list_item:
            marker = list_item.group(2).strip()
            prefix = f"{marker} " if marker[0].isdigit() else "• "
            _paragraph(body, prefix + list_item.group(3), style="List", indent=360 + 240 * (len(list_item.group(1)) // 2))
            index += 1
            continue

        if re.fullmatch(r"(?:---+|___+|\*\*\*+)\s*", stripped):
            index += 1
            continue

        paragraph_lines = [stripped.removeprefix("> ")]
        index += 1
        while index < len(lines) and lines[index].strip():
            next_line = lines[index]
            if (HEADING.match(next_line) or LIST.match(next_line) or FENCE.match(next_line)
                    or (index + 1 < len(lines) and "|" in next_line and _is_table_separator(lines[index + 1]))):
                break
            paragraph_lines.append(next_line.strip().removeprefix("> "))
            index += 1
        _paragraph(body, " ".join(paragraph_lines))

    section = _el(body, "sectPr")
    _el(section, "pgSz", w="11906", h="16838")  # A4
    _el(section, "pgMar", top="1134", right="1134", bottom="1134", left="1134",
        header="708", footer="708", gutter="0")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
    <w:sz w:val="22"/><w:lang w:val="ru-RU"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="140"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="23"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="List"><w:name w:val="List"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="40"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/></w:rPr>
  </w:style>
</w:styles>'''.encode("utf-8")

CONTENT_TYPES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

PACKAGE_RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOCUMENT_RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def _write_fallback(source: Path, destination: Path) -> None:
    markdown = source.read_text(encoding="utf-8-sig")
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", PACKAGE_RELS)
        archive.writestr("word/document.xml", _document(markdown))
        archive.writestr("word/styles.xml", STYLES)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)


def convert(source: Path, destination: Path, engine: str = "auto") -> str:
    if not source.is_file():
        raise ValueError(f"Markdown file does not exist: {source}")
    if destination.suffix.lower() != ".docx":
        raise ValueError("Output filename must end in .docx")
    if not destination.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {destination.parent}")

    pandoc = shutil.which("pandoc") if engine != "python" else None
    if engine == "pandoc" and not pandoc:
        raise RuntimeError("Pandoc is not installed or is not on PATH")

    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{destination.stem}-", suffix=".docx", dir=destination.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        if pandoc:
            result = subprocess.run(
                [pandoc, str(source.resolve()), "--from=gfm", "--to=docx", "--output", str(temporary)],
                cwd=source.parent, capture_output=True, text=True, encoding="utf-8", errors="replace",
                check=False,
            )
            if result.returncode:
                raise RuntimeError(f"Pandoc failed: {(result.stderr or result.stdout).strip()}")
            used_engine = "pandoc"
        else:
            _write_fallback(source, temporary)
            used_engine = "python"
        os.replace(temporary, destination)
        return used_engine
    finally:
        temporary.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    # A redirected Windows console can default to cp1252 even for Cyrillic paths.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Markdown report to convert")
    parser.add_argument("-o", "--output", type=Path, help="Word output path (default: input filename with .docx)")
    parser.add_argument("--engine", choices=("auto", "pandoc", "python"), default="auto",
                        help="Conversion engine (default: Pandoc when available, otherwise Python)")
    args = parser.parse_args(argv)
    destination = args.output or args.input.with_suffix(".docx")
    try:
        engine = convert(args.input, destination, args.engine)
    except (OSError, UnicodeError, ValueError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Created {destination} ({engine})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
