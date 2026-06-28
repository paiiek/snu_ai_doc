"""슬라이드를 이미지로 렌더링해 비전 모델로 내용을 읽어 텍스트를 보강한다.

텍스트 추출만으로는 그림·도표·그래프·수식이 빠지므로, 각 슬라이드를 이미지로
만들어 비전 모델에게 "보이는 내용을 빠짐없이 서술"하게 한 뒤 그 결과를
슬라이드 텍스트로 삼는다. 텍스트가 적은(도식 위주) 자료에서 특히 효과가 크다.
"""

from __future__ import annotations

import concurrent.futures

import fitz  # PyMuPDF

from .extract import Slide
from .llm import LLMClient

VISION_DPI = 150
VISION_MAX_TOKENS = 2000

VISION_SYSTEM = (
    "당신은 강의 슬라이드를 정확히 옮겨 적는 조교다. "
    "슬라이드 이미지를 보고 그 내용을 한국어로 빠짐없이 서술한다."
)

VISION_USER = """\
다음은 강의 슬라이드 한 장의 이미지다. 이 슬라이드에 담긴 내용을 한국어로 빠짐없이 서술하라.

- 제목과 본문의 글자를 그대로 반영한다.
- 그림, 도표, 그래프, 표, 수식이 있으면 그것이 무엇을 보여주는지 구체적으로 설명한다
  (축과 값, 항목 간 관계, 흐름·구조, 수식의 의미 등).
- 슬라이드에 실제로 보이는 것만 서술한다. 없는 사실을 지어내지 않는다.
- 머리말이나 목록 기호 없이 설명문(줄글)만 출력한다.

참고로 자동 추출된 텍스트는 아래와 같다(누락·깨짐이 있을 수 있으니 이미지를 우선한다):
{ocr}
"""


def render_pages_png(pdf_path: str, dpi: int = VISION_DPI) -> list[bytes]:
    """PDF의 각 페이지를 PNG 바이트로 렌더링한다(페이지 순서 유지)."""
    doc = fitz.open(pdf_path)
    out: list[bytes] = []
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            out.append(pix.tobytes("png"))
    finally:
        doc.close()
    return out


def describe_slide(client: LLMClient, png: bytes, ocr_text: str) -> str:
    user = VISION_USER.format(ocr=(ocr_text.strip() or "(추출된 텍스트 없음)"))
    return client.vision_chat(VISION_SYSTEM, user, [png], VISION_MAX_TOKENS)


def enrich_slides(
    client: LLMClient,
    pdf_path: str,
    slides: list[Slide],
    workers: int = 4,
    progress=None,
) -> list[Slide]:
    """각 슬라이드를 비전 모델로 읽어 text/char_count를 보강한다(병렬)."""
    pngs = render_pages_png(pdf_path)
    n = min(len(pngs), len(slides))

    def work(i: int):
        try:
            return i, describe_slide(client, pngs[i], slides[i].text)
        except Exception:  # noqa: BLE001 - 한 장 실패해도 전체는 진행
            return i, ""

    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        for fut in concurrent.futures.as_completed([ex.submit(work, i) for i in range(n)]):
            i, desc = fut.result()
            if desc:
                slides[i].text = desc
                slides[i].char_count = len(desc)
            done += 1
            if progress:
                progress(done, n)
    return slides
