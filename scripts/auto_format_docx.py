#!/usr/bin/env python3
"""
对博士论文 DOCX 执行低风险自动格式化：
- 页面设置
- 明确可识别标题的直接格式
- 正文和表格文字的基础字体字号
- 三线表转换

高风险项目仍保留人工复核。
"""

import argparse
import re
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from convert_to_three_line_table import convert_to_three_line_table
import validate_format as validator


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = {"w": WORD_NS}

ET.register_namespace("w", WORD_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
ET.register_namespace("v", "urn:schemas-microsoft-com:vml")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w10", "urn:schemas-microsoft-com:office:word")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
ET.register_namespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
ET.register_namespace("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
ET.register_namespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml")
ET.register_namespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
ET.register_namespace("wpsCustomData", "http://www.wps.cn/officeDocument/2013/wpsCustomData")
ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")


CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百千万零〇两]+章\s+\S+")
SECTION1_RE = re.compile(r"^\d+\.\d+\s+\S+")
SECTION2_RE = re.compile(r"^\d+\.\d+\.\d+\s+\S+")
SECTION3_RE = re.compile(r"^\d+\.\d+\.\d+\.\d+\s+\S+")
FIGURE_RE = re.compile(r"^图\s*\d")
TABLE_RE = re.compile(r"^表\s*\d")


def unpack_docx(input_path: Path, output_dir: Path) -> None:
    with zipfile.ZipFile(input_path, "r") as source:
        source.extractall(output_dir)


def pack_docx(input_dir: Path, output_path: Path) -> None:
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as target:
        for file_path in sorted(input_dir.rglob("*")):
            if file_path.is_file():
                target.write(file_path, file_path.relative_to(input_dir))


def ensure_child(parent, tag_name):
    child = parent.find(f"w:{tag_name}", NAMESPACES)
    if child is None:
        child = ET.SubElement(parent, f"{{{WORD_NS}}}{tag_name}")
    return child


def set_page_settings(root, stats):
    sect_pr_list = root.findall(".//w:sectPr", NAMESPACES)
    if not sect_pr_list:
        body = root.find(".//w:body", NAMESPACES)
        if body is None:
            return
        sect_pr_list = [ET.SubElement(body, f"{{{WORD_NS}}}sectPr")]

    for sect_pr in sect_pr_list:
        pg_sz = ensure_child(sect_pr, "pgSz")
        pg_sz.set(f"{{{WORD_NS}}}w", "12240")
        pg_sz.set(f"{{{WORD_NS}}}h", "15840")

        pg_mar = ensure_child(sect_pr, "pgMar")
        pg_mar.set(f"{{{WORD_NS}}}top", "1701")
        pg_mar.set(f"{{{WORD_NS}}}right", "1474")
        pg_mar.set(f"{{{WORD_NS}}}bottom", "1417")
        pg_mar.set(f"{{{WORD_NS}}}left", "1474")
        pg_mar.set(f"{{{WORD_NS}}}header", "1134")
        pg_mar.set(f"{{{WORD_NS}}}footer", "992")
        pg_mar.set(f"{{{WORD_NS}}}gutter", "0")

    stats["sections_updated"] = len(sect_pr_list)


def ensure_paragraph_properties(paragraph):
    ppr = paragraph.find("w:pPr", NAMESPACES)
    if ppr is None:
        ppr = ET.SubElement(paragraph, f"{{{WORD_NS}}}pPr")
        paragraph.insert(0, ppr)
    return ppr


def set_paragraph_style(ppr, style_id):
    pstyle = ppr.find("w:pStyle", NAMESPACES)
    if pstyle is None:
        pstyle = ET.Element(f"{{{WORD_NS}}}pStyle")
        ppr.insert(0, pstyle)
    pstyle.set(f"{{{WORD_NS}}}val", style_id)


def ensure_paragraph_run_properties(ppr):
    rpr = ppr.find("w:rPr", NAMESPACES)
    if rpr is None:
        rpr = ET.SubElement(ppr, f"{{{WORD_NS}}}rPr")
    return rpr


def ensure_run_properties(run):
    rpr = run.find("w:rPr", NAMESPACES)
    if rpr is None:
        rpr = ET.SubElement(run, f"{{{WORD_NS}}}rPr")
        run.insert(0, rpr)
    return rpr


def set_spacing(ppr, line=None, line_rule=None, before=None, after=None):
    spacing = ensure_child(ppr, "spacing")
    if line is not None:
        spacing.set(f"{{{WORD_NS}}}line", str(line))
    if line_rule is not None:
        spacing.set(f"{{{WORD_NS}}}lineRule", line_rule)
    if before is not None:
        spacing.set(f"{{{WORD_NS}}}before", str(before))
    if after is not None:
        spacing.set(f"{{{WORD_NS}}}after", str(after))


def set_alignment(ppr, value):
    jc = ensure_child(ppr, "jc")
    jc.set(f"{{{WORD_NS}}}val", value)


def set_indent(ppr, first_line_chars=None):
    ind = ensure_child(ppr, "ind")
    for attr in ["left", "right", "firstLine", "hanging", "firstLineChars"]:
        ind.attrib.pop(f"{{{WORD_NS}}}{attr}", None)
    if first_line_chars is not None:
        ind.set(f"{{{WORD_NS}}}firstLineChars", str(first_line_chars))


def set_run_format(run, east_asia, latin, size, bold=None):
    rpr = ensure_run_properties(run)
    set_rpr_format(rpr, east_asia, latin, size, bold=bold)


def set_rpr_format(rpr, east_asia, latin, size, bold=None):
    rfonts = ensure_child(rpr, "rFonts")
    rfonts.set(f"{{{WORD_NS}}}ascii", latin)
    rfonts.set(f"{{{WORD_NS}}}hAnsi", latin)
    rfonts.set(f"{{{WORD_NS}}}cs", latin)
    rfonts.set(f"{{{WORD_NS}}}eastAsia", east_asia)

    sz = ensure_child(rpr, "sz")
    sz.set(f"{{{WORD_NS}}}val", str(size))
    sz_cs = ensure_child(rpr, "szCs")
    sz_cs.set(f"{{{WORD_NS}}}val", str(size))

    if bold is not None:
        if bold:
            ensure_child(rpr, "b")
            ensure_child(rpr, "bCs")
        else:
            for tag_name in ["b", "bCs"]:
                element = rpr.find(f"w:{tag_name}", NAMESPACES)
                if element is not None:
                    rpr.remove(element)


def set_paragraph_mark_format(ppr, east_asia, latin, size, bold=None):
    rpr = ensure_paragraph_run_properties(ppr)
    set_rpr_format(rpr, east_asia, latin, size, bold=bold)


def paragraph_text(paragraph):
    return "".join(paragraph.itertext()).strip()


def is_in_table(paragraph, parent_map):
    current = paragraph
    tc_tag = f"{{{WORD_NS}}}tc"
    while current is not None:
        if current.tag == tc_tag:
            return True
        current = parent_map.get(current)
    return False


def classify_paragraph(text, in_table):
    if not text:
        return "empty"
    if in_table:
        return "table_text"
    if CHAPTER_RE.match(text):
        return "chapter"
    if SECTION3_RE.match(text):
        return "section3"
    if SECTION2_RE.match(text):
        return "section2"
    if SECTION1_RE.match(text):
        return "section1"
    if text in {"摘要", "目录", "参考文献", "致谢"}:
        return "special_heading"
    if text == "ABSTRACT":
        return "abstract_en_heading"
    if FIGURE_RE.match(text):
        return "figure_caption"
    if TABLE_RE.match(text):
        return "table_title"
    if text.startswith("关键词：") or text.startswith("KEY WORDS"):
        return "keywords"
    return "body"


def format_paragraph(paragraph, kind, stats):
    runs = paragraph.findall("w:r", NAMESPACES)
    if not runs:
        return

    ppr = ensure_paragraph_properties(paragraph)

    if kind == "chapter":
        set_paragraph_style(ppr, "Heading1")
        set_alignment(ppr, "center")
        set_spacing(ppr, before=480, after=360)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "黑体", "Times New Roman", 32, bold=True)
        for run in runs:
            set_run_format(run, "黑体", "Times New Roman", 32, bold=True)
        stats["chapter_paragraphs"] += 1
        return

    if kind == "section1":
        set_paragraph_style(ppr, "Heading2")
        set_alignment(ppr, "left")
        set_spacing(ppr, line=400, line_rule="exact", before=480, after=120)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "黑体", "Times New Roman", 28, bold=True)
        for run in runs:
            set_run_format(run, "黑体", "Times New Roman", 28, bold=True)
        stats["heading_paragraphs"] += 1
        return

    if kind == "section2":
        set_paragraph_style(ppr, "Heading3")
        set_alignment(ppr, "left")
        set_spacing(ppr, line=400, line_rule="exact", before=240, after=120)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "黑体", "Times New Roman", 26, bold=True)
        for run in runs:
            set_run_format(run, "黑体", "Times New Roman", 26, bold=True)
        stats["heading_paragraphs"] += 1
        return

    if kind == "section3":
        set_paragraph_style(ppr, "Heading4")
        set_alignment(ppr, "left")
        set_spacing(ppr, line=400, line_rule="exact", before=240, after=120)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "黑体", "Times New Roman", 24, bold=True)
        for run in runs:
            set_run_format(run, "黑体", "Times New Roman", 24, bold=True)
        stats["heading_paragraphs"] += 1
        return

    if kind == "special_heading":
        set_paragraph_style(ppr, "Heading1")
        set_alignment(ppr, "center")
        set_spacing(ppr, before=480, after=360)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "黑体", "Times New Roman", 32, bold=True)
        for run in runs:
            set_run_format(run, "黑体", "Times New Roman", 32, bold=True)
        stats["special_heading_paragraphs"] += 1
        return

    if kind == "abstract_en_heading":
        set_alignment(ppr, "center")
        set_spacing(ppr, line=400, line_rule="exact", before=160, after=120)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "Times New Roman", "Times New Roman", 24, bold=True)
        for run in runs:
            set_run_format(run, "Times New Roman", "Times New Roman", 24, bold=True)
        stats["special_heading_paragraphs"] += 1
        return

    if kind == "figure_caption":
        set_alignment(ppr, "both")
        set_spacing(ppr, before=120, after=240)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "宋体", "Times New Roman", 22, bold=None)
        for run in runs:
            set_run_format(run, "宋体", "Times New Roman", 22, bold=None)
        stats["caption_paragraphs"] += 1
        return

    if kind == "table_title":
        set_alignment(ppr, "center")
        set_spacing(ppr, before=240, after=120)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "宋体", "Times New Roman", 22, bold=None)
        for run in runs:
            set_run_format(run, "宋体", "Times New Roman", 22, bold=None)
        stats["caption_paragraphs"] += 1
        return

    if kind == "table_text":
        set_paragraph_mark_format(ppr, "宋体", "Times New Roman", 22, bold=None)
        for run in runs:
            set_run_format(run, "宋体", "Times New Roman", 22, bold=None)
        stats["table_paragraphs"] += 1
        return

    if kind == "keywords":
        set_alignment(ppr, "left")
        set_spacing(ppr, line=400, line_rule="exact", before=0, after=0)
        set_indent(ppr, first_line_chars=None)
        set_paragraph_mark_format(ppr, "宋体", "Times New Roman", 24, bold=None)
        for run in runs:
            set_run_format(run, "宋体", "Times New Roman", 24, bold=None)
        stats["body_paragraphs"] += 1
        return

    if kind == "body":
        set_alignment(ppr, "both")
        set_spacing(ppr, line=400, line_rule="exact", before=0, after=0)
        set_indent(ppr, first_line_chars=200)
        set_paragraph_mark_format(ppr, "宋体", "Times New Roman", 24, bold=None)
        for run in runs:
            set_run_format(run, "宋体", "Times New Roman", 24, bold=None)
        stats["body_paragraphs"] += 1


def normalize_pingfang_rfonts(root):
    for rfonts in root.findall(".//w:rFonts", NAMESPACES):
        changed = False
        for attr in ["ascii", "hAnsi", "cs"]:
            key = f"{{{WORD_NS}}}{attr}"
            if rfonts.get(key) == "PingFang SC":
                rfonts.set(key, "Times New Roman")
                changed = True
        east_asia_key = f"{{{WORD_NS}}}eastAsia"
        if rfonts.get(east_asia_key) == "PingFang SC":
            rfonts.set(east_asia_key, "宋体")
            changed = True
        if changed:
            continue


def format_document_xml(document_xml: Path):
    tree = ET.parse(document_xml)
    root = tree.getroot()
    parent_map = {child: parent for parent in root.iter() for child in parent}

    stats = {
        "sections_updated": 0,
        "tables_converted": 0,
        "chapter_paragraphs": 0,
        "heading_paragraphs": 0,
        "special_heading_paragraphs": 0,
        "caption_paragraphs": 0,
        "body_paragraphs": 0,
        "table_paragraphs": 0,
    }

    set_page_settings(root, stats)

    tables = root.findall(".//w:tbl", NAMESPACES)
    for table in tables:
        convert_to_three_line_table(table)
    stats["tables_converted"] = len(tables)

    for paragraph in root.findall(".//w:p", NAMESPACES):
        text = paragraph_text(paragraph)
        kind = classify_paragraph(text, is_in_table(paragraph, parent_map))
        if kind != "empty":
            format_paragraph(paragraph, kind, stats)

    normalize_pingfang_rfonts(root)

    tree.write(document_xml, encoding="utf-8", xml_declaration=True)
    return stats


def normalize_styles_xml(styles_xml: Path):
    if not styles_xml.exists():
        return
    tree = ET.parse(styles_xml)
    root = tree.getroot()
    normalize_pingfang_rfonts(root)
    tree.write(styles_xml, encoding="utf-8", xml_declaration=True)


def normalize_font_table_xml(font_table_xml: Path):
    if not font_table_xml.exists():
        return
    tree = ET.parse(font_table_xml)
    root = tree.getroot()
    removed = []
    for font in list(root.findall("w:font", NAMESPACES)):
        if font.get(f"{{{WORD_NS}}}name") == "PingFang SC":
            root.remove(font)
            removed.append(font)
    if removed:
        tree.write(font_table_xml, encoding="utf-8", xml_declaration=True)


def collect_validation_summary(output_docx: Path):
    with zipfile.ZipFile(output_docx, "r") as archive:
        with archive.open("word/document.xml") as xml_file:
            tree = ET.parse(xml_file)
            root = tree.getroot()

    issues = []
    issues.extend(validator.check_page_settings(root))
    issues.extend(validator.check_fonts(root))
    issues.extend(validator.check_headings(root))
    issues.extend(validator.check_figures_tables(root))
    issues.extend(validator.check_references(root))
    return issues


def write_report(report_path: Path, input_docx: Path, output_docx: Path, stats, validation_issues):
    lines = [
        "# 博士论文低风险自动格式化报告",
        "",
        "## 文件信息",
        f"- 原始文件: {input_docx.name}",
        f"- 输出文件: {output_docx.name}",
        "",
        "## 已自动修复",
        f"- 页面设置已统一到北大模板参数，共处理 {stats['sections_updated']} 个节属性",
        f"- 已转换 {stats['tables_converted']} 个表格为三线表边框",
        f"- 已直接格式化 {stats['chapter_paragraphs']} 个章标题段落",
        f"- 已直接格式化 {stats['heading_paragraphs']} 个数字编号小节标题段落",
        f"- 已直接格式化 {stats['special_heading_paragraphs']} 个摘要/目录/参考文献类标题段落",
        f"- 已统一 {stats['caption_paragraphs']} 个图表标题或图注段落的基础字体字号",
        f"- 已统一 {stats['body_paragraphs']} 个正文段落的基础字体字号和基本段落属性",
        f"- 已统一 {stats['table_paragraphs']} 个表格内段落的基础字体字号",
        "",
        "## 需要人工复核",
        "- 标题层级是否与论文真实结构一致，尤其是无编号标题和实验步骤标题",
        "- 奇偶页不同页眉、摘要/目录罗马数字页码、正文阿拉伯数字页码是否正确",
        "- 目录是否为 Word 域，且已在 Word 中刷新",
        "- 图注、表题、续表、分图标识是否与正文讨论一致",
        "- 匿名送审版脱敏是否完整",
        "",
        "## 未处理/不能安全自动处理",
        "- 不自动承诺 APA/BibTeX/混杂格式到 Nature 参考文献的准确转换",
        "- 不自动重排正文引用序号与参考文献列表映射",
        "- 不自动处理复杂分节页眉与 STYLEREF 章节名联动",
        "- 不自动刷新目录域、交叉引用域和分页结果",
        "- 不自动修正复杂公式、化学式、上下标和特殊符号的所有细节",
        "",
        "## 基础验证结果",
    ]

    if validation_issues:
        lines.append(f"- 基础校验发现 {len(validation_issues)} 项提示或问题：")
        for issue in validation_issues:
            lines.append(f"  - {issue}")
    else:
        lines.append("- 基础校验未发现问题")

    lines.extend(
        [
            "",
            "## 说明",
            "- 本报告只覆盖低风险自动格式化范围。",
            "- 若要对外宣称“基本符合要求”或“可提交终稿”，仍需人工完成高风险复核。",
        ]
    )

    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_default_output_path(input_docx: Path) -> Path:
    return input_docx.with_name(f"{input_docx.stem}_formatted_safe.docx")


def main():
    parser = argparse.ArgumentParser(description="对博士论文 DOCX 执行低风险自动格式化")
    parser.add_argument("input_docx", help="输入 .docx 文件")
    parser.add_argument("--output", help="输出 .docx 文件路径")
    parser.add_argument("--report", help="格式化报告路径")
    args = parser.parse_args()

    input_docx = Path(args.input_docx).resolve()
    if not input_docx.exists():
        raise SystemExit(f"输入文件不存在: {input_docx}")
    if input_docx.suffix.lower() != ".docx":
        raise SystemExit("只支持 .docx 文件")

    output_docx = Path(args.output).resolve() if args.output else build_default_output_path(input_docx)
    report_path = Path(args.report).resolve() if args.report else output_docx.with_name("format_report.txt")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="thesis-format-") as temp_dir:
        temp_root = Path(temp_dir)
        unpacked_dir = temp_root / "unpacked"
        unpack_docx(input_docx, unpacked_dir)

        document_xml = unpacked_dir / "word" / "document.xml"
        if not document_xml.exists():
            raise SystemExit("DOCX 中缺少 word/document.xml")

        stats = format_document_xml(document_xml)
        normalize_styles_xml(unpacked_dir / "word" / "styles.xml")
        normalize_font_table_xml(unpacked_dir / "word" / "fontTable.xml")
        pack_docx(unpacked_dir, output_docx)

    validation_issues = collect_validation_summary(output_docx)
    write_report(report_path, input_docx, output_docx, stats, validation_issues)

    print(f"输出文件: {output_docx}")
    print(f"报告文件: {report_path}")


if __name__ == "__main__":
    main()
