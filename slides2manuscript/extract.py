"""텍스트 기반 PDF에서 슬라이드별 텍스트를 추출한다."""

from __future__ import annotations

import re
from dataclasses import dataclass

import fitz  # PyMuPDF


@dataclass
class Slide:
    index: int          # 1부터 시작하는 슬라이드(페이지) 번호
    text: str           # 정리된 텍스트
    char_count: int     # 공백 포함 글자 수

    @property
    def is_empty(self) -> bool:
        return self.char_count == 0


def _clean(raw: str) -> str:
    """추출 텍스트의 잔여 노이즈를 정리한다."""
    # 윈도/맥 줄바꿈 통일
    text = raw.replace("\r\n", "\n").replace("\r", "\n")
    # 단어 중간에서 잘린 하이픈 줄바꿈 복원 (영문 자료 대비)
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    # 줄 단위 좌우 공백 제거
    lines = [ln.strip() for ln in text.split("\n")]
    # 페이지 번호만 있는 줄 제거
    lines = [ln for ln in lines if not re.fullmatch(r"\d{1,3}", ln)]
    # 빈 줄 연속 축약
    out: list[str] = []
    for ln in lines:
        if ln == "" and (not out or out[-1] == ""):
            continue
        out.append(ln)
    return "\n".join(out).strip()


def extract_slides(pdf_path: str) -> list[Slide]:
    """PDF를 열어 페이지(슬라이드)별 Slide 목록을 반환한다.

    텍스트 기반 PDF를 가정한다. 텍스트가 거의 없는 PDF면
    호출부에서 경고를 낼 수 있도록 char_count를 그대로 노출한다.
    """
    doc = fitz.open(pdf_path)
    slides: list[Slide] = []
    try:
        for i, page in enumerate(doc, start=1):
            text = _clean(page.get_text("text"))
            slides.append(Slide(index=i, text=text, char_count=len(text)))
    finally:
        doc.close()
    return slides


def total_source_chars(slides: list[Slide]) -> int:
    return sum(s.char_count for s in slides)


def looks_like_image_pdf(slides: list[Slide], threshold: int = 15) -> bool:
    """페이지당 평균 글자 수가 매우 적으면 이미지/스캔 PDF로 의심한다."""
    if not slides:
        return True
    avg = total_source_chars(slides) / len(slides)
    return avg < threshold
