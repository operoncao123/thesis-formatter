#!/usr/bin/env python3
"""
Generate two thesis-review variants from a source DOCX:

1. Inline red-note review version with issue notes inserted below affected paragraphs.
2. Conservative auto-fixed version that only applies low-risk fixes verified against
   the current thesis template used in this workspace.
"""

from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}

RED = "FF0000"
ALLOWED_TEMPLATE_FONTS = {
    "宋体",
    "SimSun",
    "黑体",
    "SimHei",
    "Times New Roman",
    "Arial",
    "Cambria Math",
    "Symbol",
    "Wingdings",
    "仿宋",
    "幼圆",
    "Courier New",
    "TimesNewRoman",
    "TimesNewRoman,Bold",
}
BAD_FONTS = {"DM Sans", "Cambria", "Calibri", "Songti SC"}
LATIN_FONT = "Times New Roman"
EAST_ASIA_FONT = "宋体"
SUBSTANTIVE_TEXT_RE = re.compile(r"[A-Za-z0-9\u3400-\u9fff]")

CHAPTER_SPACE_RE = re.compile(r"^(第[一二三四五六七八九十百千万零〇两]+章)\s+(\S)")
NUMBERED_HEADING_SPACE_RE = re.compile(r"^(\d+(?:\.\d+)+)(?=[^0-9.])(\S)")


def qn(tag: str) -> str:
    return f"{{{WORD_NS}}}{tag}"


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def normalize_heading_spacing(text: str) -> str:
    normalized = CHAPTER_SPACE_RE.sub(r"\1 \2", text)
    normalized = NUMBERED_HEADING_SPACE_RE.sub(r"\1 \2", normalized)
    return normalized


def has_visible_substantive_text(text: str) -> bool:
    return bool(text and SUBSTANTIVE_TEXT_RE.search(text))


def find_visible_font_issues(paragraph: etree._Element) -> list[str]:
    fonts = set()
    for run in paragraph.xpath("./w:r", namespaces=NS):
        run_text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        if not has_visible_substantive_text(run_text):
            continue

        rfonts = run.find("./w:rPr/w:rFonts", namespaces=NS)
        if rfonts is None:
            continue

        has_cjk = bool(re.search(r"[\u3400-\u9fff]", run_text))
        has_latin = bool(re.search(r"[A-Za-z0-9]", run_text))

        if has_cjk:
            east_asia_font = rfonts.get(qn("eastAsia"))
            if east_asia_font and east_asia_font not in ALLOWED_TEMPLATE_FONTS:
                fonts.add(east_asia_font)

        if has_latin:
            for attr in ("ascii", "hAnsi", "cs"):
                value = rfonts.get(qn(attr))
                if value and value not in ALLOWED_TEMPLATE_FONTS:
                    fonts.add(value)

    return sorted(fonts)


def normalize_visible_run_fonts(root: etree._Element) -> int:
    font_fixes = 0
    for run in root.xpath(".//w:p/w:r | .//w:tc//w:r", namespaces=NS):
        run_text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        if not has_visible_substantive_text(run_text):
            continue

        rfonts = run.find("./w:rPr/w:rFonts", namespaces=NS)
        if rfonts is None:
            continue

        changed = False
        has_cjk = bool(re.search(r"[\u3400-\u9fff]", run_text))
        has_latin = bool(re.search(r"[A-Za-z0-9]", run_text))

        if has_cjk:
            east_asia_key = qn("eastAsia")
            east_asia_value = rfonts.get(east_asia_key)
            if east_asia_value and east_asia_value not in ALLOWED_TEMPLATE_FONTS:
                rfonts.set(east_asia_key, EAST_ASIA_FONT)
                changed = True

        if has_latin:
            for attr in ("ascii", "hAnsi", "cs"):
                key = qn(attr)
                value = rfonts.get(key)
                if value and value not in ALLOWED_TEMPLATE_FONTS:
                    rfonts.set(key, LATIN_FONT)
                    changed = True

        if changed:
            font_fixes += 1

    return font_fixes


def make_note_paragraph(message: str) -> etree._Element:
    paragraph = etree.Element(qn("p"))
    ppr = etree.SubElement(paragraph, qn("pPr"))
    spacing = etree.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "120")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "exact")

    run = etree.SubElement(paragraph, qn("r"))
    rpr = etree.SubElement(run, qn("rPr"))
    rfonts = etree.SubElement(rpr, qn("rFonts"))
    rfonts.set(qn("ascii"), LATIN_FONT)
    rfonts.set(qn("hAnsi"), LATIN_FONT)
    rfonts.set(qn("cs"), LATIN_FONT)
    rfonts.set(qn("eastAsia"), EAST_ASIA_FONT)
    color = etree.SubElement(rpr, qn("color"))
    color.set(qn("val"), RED)
    size = etree.SubElement(rpr, qn("sz"))
    size.set(qn("val"), "22")
    size_cs = etree.SubElement(rpr, qn("szCs"))
    size_cs.set(qn("val"), "22")
    text = etree.SubElement(run, qn("t"))
    text.text = message
    return paragraph


def replace_paragraph_text(paragraph: etree._Element, new_text: str) -> bool:
    text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
    if not text_nodes:
        return False
    old_text = "".join(node.text or "" for node in text_nodes)
    if old_text == new_text:
        return False
    text_nodes[0].text = new_text
    for node in text_nodes[1:]:
        node.text = ""
    return True


def unzip_docx(src: Path, dest_dir: Path) -> None:
    with zipfile.ZipFile(src, "r") as archive:
        archive.extractall(dest_dir)


def zip_dir(src_dir: Path, output_docx: Path) -> None:
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(src_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(src_dir))


def update_font_table(font_table_path: Path) -> int:
    if not font_table_path.exists():
        return 0
    tree = etree.parse(str(font_table_path))
    root = tree.getroot()
    removed = 0
    for font in list(root.xpath("./w:font", namespaces=NS)):
        name = font.get(qn("name"))
        if name in BAD_FONTS:
            root.remove(font)
            removed += 1
    if removed:
        tree.write(str(font_table_path), encoding="utf-8", xml_declaration=True)
    return removed


def generate_inline_review_version(source_docx: Path, output_docx: Path) -> dict[str, int]:
    with tempfile.TemporaryDirectory(prefix="thesis-inline-notes-") as temp_dir:
        temp_root = Path(temp_dir)
        unzip_docx(source_docx, temp_root)

        document_xml = temp_root / "word" / "document.xml"
        tree = etree.parse(str(document_xml))
        root = tree.getroot()
        body = root.find(".//w:body", namespaces=NS)
        if body is None:
            raise RuntimeError("DOCX body not found")

        inserted_note_count = 0
        heading_issue_count = 0
        font_issue_count = 0

        for paragraph in list(body):
            if paragraph.tag != qn("p"):
                continue

            text = paragraph_text(paragraph)
            if not text:
                continue

            issues = []
            normalized = normalize_heading_spacing(text)
            if normalized != text:
                heading_issue_count += 1
                if text.startswith("第"):
                    issues.append("AI格式提示：本章标题中章序号与标题文字之间的空格数量不规范，建议统一为 1 个半角空格。")
                else:
                    issues.append("AI格式提示：该标题编号与标题文字之间缺少规范空格，建议统一保留 1 个半角空格。")

            bad_fonts = find_visible_font_issues(paragraph)
            if bad_fonts:
                font_issue_count += 1
                issues.append(
                    "AI格式提示：本段检测到偏离当前模板的字体："
                    + "、".join(bad_fonts)
                    + "。建议中文统一为宋体，英文、数字与符号统一为 Times New Roman。"
                )

            if not issues:
                continue

            note = make_note_paragraph(" ".join(issues))
            body.insert(body.index(paragraph) + 1, note)
            inserted_note_count += 1

        tree.write(str(document_xml), encoding="utf-8", xml_declaration=True)
        zip_dir(temp_root, output_docx)

    return {
        "inserted_note_count": inserted_note_count,
        "heading_issue_count": heading_issue_count,
        "font_issue_count": font_issue_count,
    }


def generate_auto_fixed_version(source_docx: Path, output_docx: Path) -> dict[str, int]:
    with tempfile.TemporaryDirectory(prefix="thesis-safe-fix-") as temp_dir:
        temp_root = Path(temp_dir)
        unzip_docx(source_docx, temp_root)

        document_xml = temp_root / "word" / "document.xml"
        tree = etree.parse(str(document_xml))
        root = tree.getroot()
        body = root.find(".//w:body", namespaces=NS)
        if body is None:
            raise RuntimeError("DOCX body not found")

        heading_fixes = 0
        font_fixes = 0

        for paragraph in body.xpath("./w:p", namespaces=NS):
            text = paragraph_text(paragraph)
            if not text:
                continue
            normalized = normalize_heading_spacing(text)
            if normalized != text and replace_paragraph_text(paragraph, normalized):
                heading_fixes += 1

        font_fixes = normalize_visible_run_fonts(root)

        tree.write(str(document_xml), encoding="utf-8", xml_declaration=True)
        font_table_removed = update_font_table(temp_root / "word" / "fontTable.xml")
        zip_dir(temp_root, output_docx)

    return {
        "heading_fixes": heading_fixes,
        "font_rfonts_fixed": font_fixes,
        "font_table_entries_removed": font_table_removed,
    }


def write_report(report_path: Path, source_docx: Path, inline_docx: Path, fixed_docx: Path, inline_stats: dict[str, int], fixed_stats: dict[str, int]) -> None:
    lines = [
        "# Thesis Review Variant Report",
        "",
        "## Files",
        f"- Source: {source_docx}",
        f"- Inline review version: {inline_docx}",
        f"- Conservative auto-fixed version: {fixed_docx}",
        "",
        "## Inline Review Version",
        f"- Inserted {inline_stats['inserted_note_count']} red inline notes",
        f"- Heading-spacing notes: {inline_stats['heading_issue_count']}",
        f"- Font notes: {inline_stats['font_issue_count']}",
        "",
        "## Conservative Auto-Fixes Applied",
        f"- Fixed {fixed_stats['heading_fixes']} heading-spacing issue(s)",
        f"- Rewrote {fixed_stats['font_rfonts_fixed']} run-font setting block(s)",
        f"- Removed {fixed_stats['font_table_entries_removed']} unexpected font declaration(s) from fontTable.xml",
        "",
        "## Not Auto-Fixed",
        "- TOC field refresh inside Word",
        "- Header/footer logic, section page-number transitions, and odd/even page behavior",
        "- Reference-by-reference style conformance",
        "- Anonymous-review desensitization",
        "- Image/table layout shifts that only become visible after Word repaginates the document",
        "",
        "## Notes",
        "- This conservative pass keeps the current template-matching page settings unchanged.",
        "- 仿宋、幼圆、Courier New were not treated as errors because the current official template also contains them.",
    ]
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate thesis inline-review and conservative auto-fixed variants")
    parser.add_argument("source_docx", help="Source DOCX path")
    parser.add_argument("--inline-output", required=True, help="Output DOCX path for inline red-note review version")
    parser.add_argument("--fixed-output", required=True, help="Output DOCX path for conservative auto-fixed version")
    parser.add_argument("--report", required=True, help="Output text report path")
    args = parser.parse_args()

    source_docx = Path(args.source_docx).resolve()
    inline_output = Path(args.inline_output).resolve()
    fixed_output = Path(args.fixed_output).resolve()
    report_path = Path(args.report).resolve()

    inline_output.parent.mkdir(parents=True, exist_ok=True)
    fixed_output.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    inline_stats = generate_inline_review_version(source_docx, inline_output)
    fixed_stats = generate_auto_fixed_version(source_docx, fixed_output)
    write_report(report_path, source_docx, inline_output, fixed_output, inline_stats, fixed_stats)

    print(f"inline_output={inline_output}")
    print(f"fixed_output={fixed_output}")
    print(f"report={report_path}")
    print(f"inline_stats={inline_stats}")
    print(f"fixed_stats={fixed_stats}")


if __name__ == "__main__":
    main()
