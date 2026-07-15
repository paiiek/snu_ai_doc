"""생성된 원고를 드라이한(꾸미지 않은) docx로 저장한다.

서식 원칙: 본문 글꼴 하나, 장 제목만 약간 키움, 색·표·테두리·머리글 없음.
한국어 A4 기준으로 분량 계산이 맞도록 글꼴/줄간격/여백을 고정한다.

공식 규격(전문가 활용비 강의료 및 원고료 책정표):
- A4 1페이지: 12Font / 35Line, 신명조
- 상하 여백 15mm, 좌우 여백 20mm, 머리말·꼬리말 15mm
"""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Cm, Mm, Pt
from docx.oxml.ns import qn

from .generate import Section

BODY_FONT = "신명조"
BODY_SIZE = 12
HEADING_SIZE = 13
TITLE_SIZE = 15
# A4(297mm) - 상하 여백 30mm = 267mm 안에 35줄이 들어가도록 줄 높이를 잡는다.
# 267mm / 35줄 ≈ 7.62mm/줄 ≈ 21.6pt. MULTIPLE 1.8 은 12pt × 1.8 = 21.6pt.
LINE_SPACING_MULTIPLE = 1.8


def _apply_rfonts(rpr, name: str) -> None:
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def _set_korean_font(run, name: str) -> None:
    run.font.name = name
    _apply_rfonts(run._element.get_or_add_rPr(), name)


def _base_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = Pt(BODY_SIZE)
    _apply_rfonts(style.element.get_or_add_rPr(), BODY_FONT)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = LINE_SPACING_MULTIPLE
    # 35줄/1페이지 정확성을 지키기 위해 문단 사이 추가 여백은 두지 않는다.
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)

    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)
        section.top_margin = Mm(15)
        section.bottom_margin = Mm(15)
        section.left_margin = Mm(20)
        section.right_margin = Mm(20)
        section.header_distance = Mm(15)
        section.footer_distance = Mm(15)


def _add_paragraph(doc: Document, text: str, size: int, bold: bool, space_before: int = 0):
    p = doc.add_paragraph()
    if space_before:
        p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    _set_korean_font(run, BODY_FONT)
    return p


def write_docx(
    out_path: str,
    doc_title: str,
    sections: list[Section],
) -> None:
    doc = Document()
    _base_style(doc)

    if doc_title:
        _add_paragraph(doc, doc_title, TITLE_SIZE, bold=True)
        doc.add_paragraph()

    for i, sec in enumerate(sections):
        space_before = 0 if (i == 0 and not doc_title) else 12
        _add_paragraph(doc, sec.title, HEADING_SIZE, bold=True, space_before=space_before)
        for para in _split_paragraphs(sec.body):
            _add_paragraph(doc, para, BODY_SIZE, bold=False)

    doc.save(out_path)


def _split_paragraphs(body: str) -> list[str]:
    """본문 텍스트를 문단 단위로 나눈다."""
    chunks = [c.strip() for c in body.replace("\r\n", "\n").split("\n\n")]
    out: list[str] = []
    for c in chunks:
        if not c:
            continue
        # 단일 줄바꿈은 문단 내 줄로 보고 공백으로 잇는다.
        out.append(" ".join(line.strip() for line in c.split("\n") if line.strip()))
    return out


def manuscript_char_count(sections: list[Section]) -> int:
    return sum(len(s.body) for s in sections)
