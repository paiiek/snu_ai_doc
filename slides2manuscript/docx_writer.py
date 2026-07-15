"""생성된 원고를 드라이한(꾸미지 않은) docx로 저장한다.

서식 원칙: 본문 글꼴 하나, 장 제목만 약간 키움, 색·표·테두리·머리글 없음.
한국어 A4 기준으로 분량 계산이 맞도록 글꼴/줄간격/여백을 고정한다.

공식 규격(전문가 활용비 강의료 및 원고료 책정표):
- A4 1페이지: 12Font / 35Line, 신명조
- 상하 여백 15mm, 좌우 여백 20mm, 머리말·꼬리말 15mm
"""

from __future__ import annotations

import os
import platform

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Cm, Mm, Pt
from docx.oxml.ns import qn

from .generate import Section


def _default_body_font() -> str:
    """생성 시점 OS에 실제로 존재하는 명조체 폰트 이름을 반환한다.

    규격은 "신명조"지만 이 이름의 폰트가 시스템에 없으면 LibreOffice·Word가
    엉뚱한 산세리프로 대체해 버려 명조체 규격이 깨진다. 그래서 각 OS의 기본
    명조체를 지정해 두고, `--font` 로 override 가능하게 한다.
    """
    system = platform.system()
    if system == "Darwin":
        # macOS 기본 제공 명조체
        if os.path.exists("/System/Library/Fonts/Supplemental/AppleMyungjo.ttf"):
            return "AppleMyungjo"
    if system == "Windows":
        # Windows 한국어 기본 명조체
        return "바탕"
    # Linux / 기타: 나눔명조가 흔함(설치돼 있어야 함)
    return "NanumMyeongjo"


BODY_FONT = _default_body_font()
BODY_SIZE = 12
HEADING_SIZE = 13
TITLE_SIZE = 15
# A4(297mm) - 상하 여백 30mm = 267mm 안에 35줄이 들어가도록 줄 높이를 EXACTLY로 고정한다.
# 267mm / 35줄 ≈ 7.62mm/줄 ≈ 21.6pt.
# 다만 LibreOffice·Word는 페이지 하단에 안전 여백을 남기고 렌더링해서 21.6pt로 두면
# 실제 34줄만 나오는 경우가 있다. 21.0pt로 살짝 좁혀 35줄이 확실히 들어가도록 한다.
# MULTIPLE 은 워드프로세서마다 해석이 달라(폰트 metrics × multiplier) 줄 수가 안 맞는 경우가 있어
# EXACT 방식으로 못박는다.
LINE_HEIGHT_PT = 21.0


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


def _base_style(doc: Document, body_font: str) -> None:
    style = doc.styles["Normal"]
    style.font.name = body_font
    style.font.size = Pt(BODY_SIZE)
    _apply_rfonts(style.element.get_or_add_rPr(), body_font)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(LINE_HEIGHT_PT)
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


def _add_paragraph(doc: Document, text: str, size: int, bold: bool, body_font: str,
                   space_before: int = 0):
    p = doc.add_paragraph()
    if space_before:
        p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    _set_korean_font(run, body_font)
    return p


def write_docx(
    out_path: str,
    doc_title: str,
    sections: list[Section],
    body_font: str | None = None,
) -> None:
    body_font = body_font or BODY_FONT
    doc = Document()
    _base_style(doc, body_font)

    if doc_title:
        _add_paragraph(doc, doc_title, TITLE_SIZE, bold=True, body_font=body_font)
        doc.add_paragraph()

    for i, sec in enumerate(sections):
        space_before = 0 if (i == 0 and not doc_title) else 12
        _add_paragraph(doc, sec.title, HEADING_SIZE, bold=True, body_font=body_font,
                       space_before=space_before)
        for para in _split_paragraphs(sec.body):
            _add_paragraph(doc, para, BODY_SIZE, bold=False, body_font=body_font)

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
