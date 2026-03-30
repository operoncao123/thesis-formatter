from __future__ import annotations

import importlib.util
from pathlib import Path

from lxml import etree


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "generate_review_variants.py"
SPEC = importlib.util.spec_from_file_location("generate_review_variants", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}


def make_paragraph(inner_xml: str):
    xml = f'<w:p xmlns:w="{WORD_NS}">{inner_xml}</w:p>'
    return etree.fromstring(xml.encode("utf-8"))


def test_ignores_bad_font_on_whitespace_run():
    paragraph = make_paragraph(
        """
        <w:r>
          <w:t>Visible text</w:t>
        </w:r>
        <w:r>
          <w:rPr><w:rFonts w:ascii="DM Sans" w:hAnsi="DM Sans"/></w:rPr>
          <w:t xml:space="preserve"> </w:t>
        </w:r>
        """
    )
    assert MODULE.find_visible_font_issues(paragraph) == []


def test_ignores_bad_font_on_paragraph_mark_only():
    paragraph = make_paragraph(
        """
        <w:pPr>
          <w:rPr><w:rFonts w:eastAsia="Songti SC"/></w:rPr>
        </w:pPr>
        <w:r><w:t>肺是呼吸器官。</w:t></w:r>
        """
    )
    assert MODULE.find_visible_font_issues(paragraph) == []


def test_ignores_cjk_text_when_only_latin_font_slots_are_set():
    paragraph = make_paragraph(
        """
        <w:r>
          <w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria" w:hint="eastAsia"/></w:rPr>
          <w:t>预实验中，</w:t>
        </w:r>
        """
    )
    assert MODULE.find_visible_font_issues(paragraph) == []


def test_ignores_punctuation_only_run():
    paragraph = make_paragraph(
        """
        <w:r>
          <w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr>
          <w:t>。</w:t>
        </w:r>
        """
    )
    assert MODULE.find_visible_font_issues(paragraph) == []


def test_reports_visible_latin_run_with_non_template_font():
    paragraph = make_paragraph(
        """
        <w:r>
          <w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/></w:rPr>
          <w:t>SMC</w:t>
        </w:r>
        """
    )
    assert MODULE.find_visible_font_issues(paragraph) == ["Cambria"]
