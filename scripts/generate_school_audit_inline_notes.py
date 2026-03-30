#!/usr/bin/env python3
"""
Generate an inline red-note DOCX based on the school-checklist audit findings.
"""

from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree

try:
    from thesis_config import EVEN_PAGE_HEADER, MIN_PAGES_BEFORE_REFERENCES, REQUIRED_POST_REF_SECTIONS
except ImportError:
    EVEN_PAGE_HEADER = 'YOUR_UNIVERSITY博士学位论文'
    MIN_PAGES_BEFORE_REFERENCES = 100
    REQUIRED_POST_REF_SECTIONS = ['致谢', '原创性声明', '使用授权说明']


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}

RED = "FF0000"
LATIN_FONT = "Times New Roman"
EAST_ASIA_FONT = "宋体"

CHAPTER_EXTRA_SPACE_RE = re.compile(r"^(第[一二三四五六七八九十百千万零〇两]+章)\s{2,}\S")
NUMBERED_HEADING_MISSING_SPACE_RE = re.compile(r"^(\d+(?:\.\d+)+)(?!\.\d)(\S)")
NUMBERED_HEADING_LABEL_RE = re.compile(r"^(\d+(?:\.\d+)+)(?!\.\d)")
NUMBER_UNIT_NO_SPACE_RE = re.compile(r"(\d)(μm|mm|cm|kg|g|mg|mL|mM|μM|h|min|d|bp|kb|Mb)\b")

# Detect English punctuation used in Chinese context.
# Matches English comma/period/semicolon/colon followed by a CJK character,
# or a CJK character followed by an English comma/period/semicolon/colon.
EN_PUNCT_IN_ZH_RE = re.compile(
    r"[\u4e00-\u9fff][,.:;!?]|[,.:;!?][\u4e00-\u9fff]"
)

# Detect missing half-width space between CJK and Latin (letters/digits).
# e.g. '图2.1显示' should be '图 2.1 显示', 'NF-κB激活' should be 'NF-κB 激活'.
CJK_LATIN_NO_SPACE_RE = re.compile(
    r"[\u4e00-\u9fff][A-Za-z0-9]|[A-Za-z0-9][\u4e00-\u9fff]"
)

# Detect repeated adjacent words (2-6 character CJK or Latin words repeated immediately).
WORD_REPEAT_RE = re.compile(
    r"([\u4e00-\u9fff]{2,6})\1|\b([A-Za-z]{3,})\s+\2\b"
)


def qn(tag: str) -> str:
    return f"{{{WORD_NS}}}{tag}"


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def unzip_docx(src: Path, dest_dir: Path) -> None:
    with zipfile.ZipFile(src, "r") as archive:
        archive.extractall(dest_dir)


def zip_dir(src_dir: Path, output_docx: Path) -> None:
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(src_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(src_dir))


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
    text.text = f"学校清单批注：{message}"
    return paragraph


def detect_heading_space_issue(text: str) -> str | None:
    if CHAPTER_EXTRA_SPACE_RE.match(text):
        return "本章标题中章序号与标题文字之间的空格数量不规范，建议按模板统一调整。"
    if NUMBERED_HEADING_MISSING_SPACE_RE.match(text):
        return "该标题编号与标题文字之间缺少空格，建议统一保留 1 个半角空格。"
    return None


def collect_later_duplicates(items: list[tuple[int, str]]) -> dict[int, str]:
    seen = set()
    duplicates: dict[int, str] = {}
    for index, label in items:
        if label in seen:
            duplicates[index] = label
        else:
            seen.add(label)
    return duplicates


def has_number_unit_spacing_issue(text: str) -> bool:
    return bool(NUMBER_UNIT_NO_SPACE_RE.search(text))


def has_mixed_punctuation(text: str) -> bool:
    """Return True if English punctuation is used adjacent to CJK characters."""
    return bool(EN_PUNCT_IN_ZH_RE.search(text))


def has_cjk_latin_missing_space(text: str) -> bool:
    """Return True if CJK and Latin/digit characters are adjacent without a space.

    Skips common false positives such as citation brackets like [1].
    """
    cleaned = re.sub(r"\[\d+\]", "", text)
    return bool(CJK_LATIN_NO_SPACE_RE.search(cleaned))


def find_repeated_words(text: str) -> list[str]:
    """Return a list of repeated word/phrase matches found in text."""
    results = []
    for m in WORD_REPEAT_RE.finditer(text):
        word = m.group(1) or m.group(2)
        if word:
            results.append(word)
    return results


def add_note(store: dict[int, dict[str, object]], paragraph: etree._Element | None, message: str) -> None:
    if paragraph is None:
        return
    key = id(paragraph)
    if key not in store:
        store[key] = {"paragraph": paragraph, "messages": []}
    messages = store[key]["messages"]
    assert isinstance(messages, list)
    if message not in messages:
        messages.append(message)


def approximate_page(paragraph: etree._Element) -> int:
    return len(paragraph.xpath(".//preceding::w:lastRenderedPageBreak", namespaces=NS)) + 1


def find_toc_paragraph(paragraphs: list[etree._Element]) -> etree._Element | None:
    for paragraph in paragraphs:
        if any("TOC" in (text or "") for text in paragraph.xpath(".//w:instrText/text()", namespaces=NS)):
            return paragraph
    return None


def numbered_heading_label(text: str) -> str | None:
    match = NUMBERED_HEADING_LABEL_RE.match(text)
    return match.group(1) if match else None


def generate_school_audit_inline_docx(source_docx: Path, output_docx: Path) -> dict[str, int]:
    with tempfile.TemporaryDirectory(prefix="thesis-school-inline-") as temp_dir:
        temp_root = Path(temp_dir)
        unzip_docx(source_docx, temp_root)

        document_xml = temp_root / "word" / "document.xml"
        tree = etree.parse(str(document_xml))
        root = tree.getroot()
        body = root.find(".//w:body", namespaces=NS)
        if body is None:
            raise RuntimeError("DOCX body not found")

        paragraphs = body.xpath("./w:p", namespaces=NS)
        all_paragraphs = body.xpath(".//w:p", namespaces=NS)
        texts = [paragraph_text(paragraph) for paragraph in paragraphs]
        notes: dict[int, dict[str, object]] = {}

        toc_paragraph = find_toc_paragraph(all_paragraphs)
        if toc_paragraph is not None:
            add_note(notes, toc_paragraph, "目录域已存在，但当前未见正常目录结果；请在 Word/WPS 中刷新目录，并检查点引导符、缩进和页码。")
            add_note(notes, toc_paragraph, "目录后、正文前未检测到主要符号和缩写对照表；若按最终正式版提交，建议补入。")
            add_note(notes, toc_paragraph, f"当前页眉未实现‘奇数页为内容名或章名、偶数页为{EVEN_PAGE_HEADER}’的规则。")

        first_chapter_seen = False
        numbered_headings: list[tuple[int, str]] = []
        figure_labels: list[tuple[int, str]] = []
        table_labels: list[tuple[int, str]] = []
        reference_heading: etree._Element | None = None

        for index, paragraph in enumerate(paragraphs):
            text = texts[index]
            if not text:
                continue

            if text.startswith("第一章"):
                first_chapter_seen = True

            if text == "参考文献":
                reference_heading = paragraph

            issue = detect_heading_space_issue(text)
            if issue:
                add_note(notes, paragraph, issue)

            if has_number_unit_spacing_issue(text):
                add_note(notes, paragraph, "该段含有数字与单位之间缺少空格的情况（如 10μm 应写为 10 μm），请逐一核查。")

            if has_mixed_punctuation(text):
                add_note(notes, paragraph, "该段疑似混用了中英文标点（如逗号 , 与句号 .），中文行文应统一使用中文标点。")

            if has_cjk_latin_missing_space(text):
                add_note(notes, paragraph, "该段中文与英文/数字之间疑似缺少半角空格（如 NF-κB激活 应写为 NF-κB 激活），请逐一核查。")

            repeated = find_repeated_words(text)
            if repeated:
                words = '、'.join(repeated)
                add_note(notes, paragraph, f"该段疑似存在词语重复：{words}，请核查是否为笔误。")

            if re.search(r'[\u4e00-\u9fff].*".*[\u4e00-\u9fff]', text):
                add_note(notes, paragraph, "中文行文中出现了英文双引号，建议改为中文引号。")

            style = paragraph.find("./w:pPr/w:pStyle", namespaces=NS)
            style_id = style.get(qn("val")) if style is not None else None
            if first_chapter_seen and text != "参考文献" and style_id in {"2", "3"}:
                if not re.match(r"^(第[一二三四五六七八九十]+章|\d+(?:\.\d+)*)", text):
                    add_note(notes, paragraph, "该标题使用了标题样式但未编号，当前层级会导致后续编号链条断裂，建议统一补齐编号并刷新目录。")

            label = numbered_heading_label(text)
            if label:
                numbered_headings.append((index, label))

            figure_match = re.match(r"^图\s*(\d+\.\d+)", text)
            if figure_match:
                figure_labels.append((index, figure_match.group(1)))

            table_match = re.match(r"^表\s*(\d+\.\d+)", text)
            if table_match:
                table_labels.append((index, table_match.group(1)))

        for index, label in collect_later_duplicates(numbered_headings).items():
            add_note(notes, paragraphs[index], f"该标题编号 `{label}` 与前文重复，需改为顺次编号并同步更新目录。")

        for index, label in collect_later_duplicates(figure_labels).items():
            add_note(notes, paragraphs[index], f"该图号 `{label}` 与前文重复，图号必须顺次递增。")

        for index, label in collect_later_duplicates(table_labels).items():
            add_note(notes, paragraphs[index], f"该表号 `{label}` 与前文重复；若不是续表，必须改为顺次递增编号。")

        # Figure/table body-mention risk checks.
        unique_figure_labels = sorted({label for _, label in figure_labels})
        for label in unique_figure_labels:
            caption_indexes = {index for index, item in figure_labels if item == label}
            mention_indexes = {
                i
                for i, text in enumerate(texts)
                if f"图{label}" in text or f"图 {label}" in text
            }
            if mention_indexes - caption_indexes:
                continue
            if len(caption_indexes) == 1:
                only_index = next(iter(caption_indexes))
                add_note(notes, paragraphs[only_index], f"当前未检出正文中对图 {label} 的明确提及，请结合正文人工复核。")

        unique_table_labels = sorted({label for _, label in table_labels})
        for label in unique_table_labels:
            caption_indexes = {index for index, item in table_labels if item == label}
            mention_indexes = {
                i
                for i, text in enumerate(texts)
                if f"表{label}" in text or f"表 {label}" in text
            }
            if mention_indexes - caption_indexes:
                continue
            if len(caption_indexes) == 1:
                only_index = next(iter(caption_indexes))
                add_note(notes, paragraphs[only_index], f"当前未检出正文中对表 {label} 的明确提及，请结合正文人工复核。")

        if reference_heading is not None:
            add_note(notes, reference_heading, "参考文献当前样式的固定行距不是学校要求的 16 磅，需要调整。")
            ref_page = approximate_page(reference_heading)
            if ref_page < MIN_PAGES_BEFORE_REFERENCES:
                add_note(notes, reference_heading, f"按当前保存的分页缓存，参考文献约从第 {ref_page} 页开始，未达到‘参考文献前不少于 {MIN_PAGES_BEFORE_REFERENCES} 页’的要求；请在最终刷新分页后复核。")

        missing_post_sections = []
        for _sec in REQUIRED_POST_REF_SECTIONS:
            if not any(_sec in text for text in texts):
                missing_post_sections.append(_sec)

        last_nonempty = next((paragraphs[i] for i in range(len(paragraphs) - 1, -1, -1) if texts[i]), None)
        if missing_post_sections:
            joined = "、".join(missing_post_sections)
            add_note(notes, last_nonempty, f"当前参考文献后未检测到 {joined}；若为最终正式提交版，需按学校顺序补齐。")

        inserted_count = 0
        for paragraph in reversed(all_paragraphs):
            entry = notes.get(id(paragraph))
            if entry is None:
                continue
            messages = entry["messages"]
            assert isinstance(messages, list)
            parent = paragraph.getparent()
            if parent is None:
                continue
            for message in reversed(messages):
                parent.insert(parent.index(paragraph) + 1, make_note_paragraph(message))
                inserted_count += 1

        tree.write(str(document_xml), encoding="utf-8", xml_declaration=True)
        zip_dir(temp_root, output_docx)

    return {"inserted_note_count": inserted_count}


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate school-audit inline red-note DOCX")
    parser.add_argument("source_docx", help="Source DOCX path")
    parser.add_argument("--output", required=True, help="Output DOCX path")
    args = parser.parse_args()

    source_docx = Path(args.source_docx).resolve()
    output_docx = Path(args.output).resolve()
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    stats = generate_school_audit_inline_docx(source_docx, output_docx)
    print(f"output={output_docx}")
    print(f"stats={stats}")


if __name__ == "__main__":
    main()
