from __future__ import annotations

import importlib.util
import zipfile
from pathlib import Path

from lxml import etree


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "generate_school_audit_inline_notes.py"
SPEC = importlib.util.spec_from_file_location("generate_school_audit_inline_notes", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def test_heading_space_detector_does_not_flag_valid_numbered_heading():
    assert MODULE.detect_heading_space_issue("1.2.1 气道平滑肌增生") is None


def test_heading_space_detector_flags_missing_space_after_number():
    issue = MODULE.detect_heading_space_issue("1.2.1气道平滑肌增生")
    assert issue is not None
    assert "编号与标题文字之间缺少空格" in issue


def test_heading_space_detector_flags_extra_space_in_chapter_heading():
    issue = MODULE.detect_heading_space_issue("第三章  实验结果")
    assert issue is not None
    assert "空格数量不规范" in issue


def test_collect_later_duplicates_returns_only_repeated_occurrences():
    items = [(10, "1.3"), (11, "1.4"), (20, "1.3"), (30, "1.4"), (40, "1.4")]
    assert MODULE.collect_later_duplicates(items) == {
        20: "1.3",
        30: "1.4",
        40: "1.4",
    }


def test_number_unit_missing_space_detector():
    assert MODULE.has_number_unit_spacing_issue("比例尺：10μm。")
    assert MODULE.has_number_unit_spacing_issue("大小约为2mm。")
    assert not MODULE.has_number_unit_spacing_issue("比例尺：10 μm。")
    assert not MODULE.has_number_unit_spacing_issue("温度为37°C。")


def test_generate_school_audit_inline_docx_inserts_toc_notes_for_nested_toc_paragraph(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt>
      <w:sdtContent>
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>目录</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p><w:r><w:t>正文末尾</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml)

    MODULE.generate_school_audit_inline_docx(source, output)

    with zipfile.ZipFile(output) as archive:
        xml = archive.read("word/document.xml")
    root = etree.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    toc_paragraph = root.xpath(".//w:sdtContent/w:p[1]", namespaces=ns)[0]
    following = toc_paragraph.xpath("following-sibling::w:p", namespaces=ns)
    notes = ["".join(p.xpath(".//w:t/text()", namespaces=ns)).strip() for p in following[:3]]

    assert "目录域已存在" in notes[0]
    assert "主要符号和缩写对照表" in notes[1]
    assert "奇数页为内容名或章名" in notes[2]
