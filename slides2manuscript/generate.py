"""Anthropic API로 장 구성을 설계하고 각 장의 줄글 본문을 생성한다."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass

from . import prompts
from .extract import Slide
from .llm import LLMClient

SECTION_MAX_TOKENS = 8192
OUTLINE_MAX_TOKENS = 2048


@dataclass
class Section:
    title: str
    start: int          # 슬라이드 시작 번호(1-base, 포함)
    end: int            # 슬라이드 끝 번호(1-base, 포함)
    target_chars: int = 0
    body: str = ""


def _slides_block(slides: list[Slide], start: int, end: int) -> str:
    parts = []
    for s in slides:
        if start <= s.index <= end and s.text:
            parts.append(f"[S{s.index}]\n{s.text}")
    return "\n\n".join(parts)


def _call_text(client: LLMClient, system, user, max_tokens) -> str:
    return client.chat(system, user, max_tokens)


# ---------------------------------------------------------------------------
# 1단계: 장 구성 설계
# ---------------------------------------------------------------------------

def _fallback_sections(slides: list[Slide], target_sections: int) -> list[Section]:
    """아웃라인 호출이 실패했을 때 슬라이드를 균등 분할한다."""
    nonempty = [s.index for s in slides if s.text]
    if not nonempty:
        nonempty = [s.index for s in slides]
    lo, hi = nonempty[0], nonempty[-1]
    total = hi - lo + 1
    n = max(1, min(target_sections, total))
    size = -(-total // n)  # 올림
    sections: list[Section] = []
    cur = lo
    while cur <= hi:
        end = min(cur + size - 1, hi)
        sections.append(Section(title=f"{len(sections) + 1}장", start=cur, end=end))
        cur = end + 1
    return sections


def design_outline(
    client: LLMClient,
    slides: list[Slide],
    pages: int,
    min_sections: int,
    max_sections: int,
) -> list[Section]:
    full_block = _slides_block(slides, slides[0].index, slides[-1].index)
    user = prompts.OUTLINE_USER_TEMPLATE.format(
        pages=pages,
        min_sections=min_sections,
        max_sections=max_sections,
        slides_block=full_block,
    )
    try:
        raw = _call_text(client, prompts.OUTLINE_SYSTEM, user, OUTLINE_MAX_TOKENS)
        data = _parse_json(raw)
        sections = [
            Section(title=str(s["title"]).strip(), start=int(s["start"]), end=int(s["end"]))
            for s in data["sections"]
        ]
        sections = _repair_ranges(sections, slides)
        if sections:
            return sections
    except Exception as exc:  # noqa: BLE001 - 어떤 실패든 폴백으로 진행
        print(f"  (아웃라인 자동 생성 실패, 균등 분할로 대체: {exc})")
    target = max(min_sections, min(max_sections, max(1, len(slides) // 4)))
    return _fallback_sections(slides, target)


def _parse_json(raw: str) -> dict:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw).strip()
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    return json.loads(m.group(0) if m else raw)


def _repair_ranges(sections: list[Section], slides: list[Slide]) -> list[Section]:
    """장 범위가 슬라이드 전체를 빠짐없이 덮도록 보정한다."""
    if not sections:
        return sections
    lo, hi = slides[0].index, slides[-1].index
    sections = sorted(sections, key=lambda s: s.start)
    sections[0].start = lo
    sections[-1].end = hi
    for i in range(1, len(sections)):
        if sections[i].start <= sections[i - 1].end:
            sections[i].start = sections[i - 1].end + 1
    sections = [s for s in sections if s.start <= s.end]
    return sections


def allocate_char_budget(sections: list[Section], slides: list[Slide], total_target: int) -> None:
    """슬라이드 분량에 비례해 각 장의 목표 글자 수를 배분한다(하한 보장)."""
    by_index = {s.index: s.char_count for s in slides}
    weights = []
    for sec in sections:
        w = sum(by_index.get(i, 0) for i in range(sec.start, sec.end + 1))
        weights.append(max(w, 1))
    wsum = sum(weights)
    floor = max(800, total_target // (len(sections) * 3))
    for sec, w in zip(sections, weights):
        sec.target_chars = max(floor, round(total_target * w / wsum))


# ---------------------------------------------------------------------------
# 2단계: 장별 본문 생성
# ---------------------------------------------------------------------------

def section_source_chars(section: Section, slides: list[Slide]) -> int:
    """장이 담당하는 슬라이드 원문 글자 수(확장 여력 판단용)."""
    return sum(s.char_count for s in slides if section.start <= s.index <= section.end)


def write_section(
    client: LLMClient,
    section: Section,
    slides: list[Slide],
    prev_tail: str,
    toc_titles: list[str],
    hint: str = "",
) -> str:
    block = _slides_block(slides, section.start, section.end)
    if prev_tail:
        toc_line = ""
        if toc_titles:
            toc_line = "지금까지 다룬 장: " + ", ".join(toc_titles)
        context_block = prompts.CONTEXT_TEMPLATE.format(
            prev_tail=prev_tail[-300:], toc_line=toc_line
        )
    else:
        context_block = ""
    hint_block = ("\n" + hint + "\n") if hint else ""
    user = prompts.SECTION_USER_TEMPLATE.format(
        title=section.title,
        target_chars=section.target_chars,
        hint_block=hint_block,
        context_block=context_block,
        slides_block=block,
    )
    return _call_text(client, prompts.SECTION_SYSTEM, user, SECTION_MAX_TOKENS)
