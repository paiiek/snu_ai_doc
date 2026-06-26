"""명령행 진입점: 슬라이드 PDF -> 줄글 강의 원고(docx)."""

from __future__ import annotations

import argparse
import os
import sys

from . import __version__, generate, docx_writer, llm, prompts
from .extract import extract_slides, total_source_chars, looks_like_image_pdf


def _load_dotenv() -> None:
    """현재 폴더 또는 스크립트 폴더의 .env에서 키를 읽어 환경에 채운다.

    별도 의존성 없이 KEY=VALUE 형식만 처리한다. 이미 설정된 환경변수는 덮지 않는다.
    """
    candidates = [
        os.path.join(os.getcwd(), ".env"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"),
    ]
    seen: set[str] = set()
    for path in candidates:
        if path in seen or not os.path.isfile(path):
            continue
        seen.add(path)
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key, val = key.strip(), val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val

# A4 한 쪽에 실제로 들어가는 대략의 글자 수(맑은 고딕 11pt, 줄간격 1.5,
# 여백 2.54cm, 문단 나눔 포함 기준). 쪽수 추정과 분량 목표 산정에 함께 쓴다.
CHARS_PER_PAGE = 1200

# 모델이 장별 목표 글자 수의 약 75~80%만 쓰는 경향이 있어, 생성 목표를 그만큼
# 키워 잡는다. (실측 결과 실현율 ≈ 0.77 → 보정 계수 ≈ 1.3)
OVERSHOOT = 1.3

# 분량 수렴 보정 최대 반복 횟수(범위를 벗어났을 때만 작동).
# 분량 부족(35쪽 미만) 시 끝까지 끌어올려야 하므로 넉넉히 둔다.
MAX_CONVERGE_ROUNDS = 6


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="slides2manuscript",
        description="강의 슬라이드 PDF를 줄글 강의 원고(docx)로 변환한다.",
    )
    p.add_argument("pdf", help="입력 슬라이드 PDF 경로(텍스트 기반)")
    p.add_argument("-o", "--out", help="출력 docx 경로(기본: 입력파일명_원고.docx)")
    p.add_argument("-t", "--title", default="", help="원고 맨 위 제목(기본: 없음)")
    p.add_argument("--pages", type=int, default=38, help="목표(중앙값) A4 분량(기본: 38)")
    p.add_argument("--min-pages", type=int, default=35, help="허용 최소 분량(기본: 35, 무조건 보장)")
    p.add_argument("--max-pages", type=int, default=50, help="허용 최대 분량(기본: 50)")
    p.add_argument("--provider", choices=("auto",) + llm.PROVIDERS, default="auto",
                   help="LLM 제공자(기본: auto = 설정된 키로 자동 선택). anthropic 또는 openai")
    p.add_argument("--model", default=None,
                   help="사용할 모델(기본: provider별 기본값). "
                        f"anthropic→{llm.DEFAULT_MODELS['anthropic']}, "
                        f"openai→{llm.DEFAULT_MODELS['openai']}")
    p.add_argument("--base-url", default=None,
                   help="OpenAI 호환 엔드포인트 주소(선택). OPENAI_BASE_URL 로도 지정 가능")
    p.add_argument("--chars-per-page", type=int, default=CHARS_PER_PAGE,
                   help=f"A4 한 쪽당 글자 수 환산값(기본: {CHARS_PER_PAGE})")
    p.add_argument("--min-sections", type=int, default=8, help="최소 장 수(기본: 8)")
    p.add_argument("--max-sections", type=int, default=16, help="최대 장 수(기본: 16)")
    p.add_argument("--force", action="store_true",
                   help="이미지/스캔 PDF로 의심돼도 그대로 진행")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    _load_dotenv()

    if not os.path.isfile(args.pdf):
        print(f"입력 파일을 찾을 수 없습니다: {args.pdf}", file=sys.stderr)
        return 2

    out_path = args.out or _default_out(args.pdf)

    print(f"[1/4] PDF 추출: {args.pdf}")
    slides = extract_slides(args.pdf)
    nonempty = [s for s in slides if s.text]
    print(f"      슬라이드 {len(slides)}장, 텍스트 있는 슬라이드 {len(nonempty)}장, "
          f"원문 {total_source_chars(slides):,}자")

    if not nonempty:
        print("텍스트를 전혀 추출하지 못했습니다. 이미지/스캔 PDF로 보입니다.", file=sys.stderr)
        return 3
    if looks_like_image_pdf(slides) and not args.force:
        print("경고: 텍스트가 매우 적어 이미지/스캔 PDF로 의심됩니다.\n"
              "      이 도구는 텍스트 기반 PDF용입니다. 그래도 진행하려면 --force 를 붙이세요.",
              file=sys.stderr)
        return 4

    provider = llm.resolve_provider(args.provider)
    client = llm.LLMClient(provider, args.model, args.base_url)
    page_chars = args.pages * args.chars_per_page
    total_target = round(page_chars * OVERSHOOT)

    print(f"[2/4] 장 구성 설계 (목표 약 {args.pages}쪽 / 본문 {page_chars:,}자, "
          f"{provider} / {client.model})")
    sections = generate.design_outline(
        client, slides, args.pages, args.min_sections, args.max_sections
    )
    generate.allocate_char_budget(sections, slides, total_target)
    for i, sec in enumerate(sections, 1):
        print(f"      {i:>2}. {sec.title}  (S{sec.start}-S{sec.end}, 목표 {sec.target_chars:,}자)")

    print(f"[3/4] 본문 생성 ({len(sections)}개 장)")
    _generate_all(client, sections, slides)

    cpp = args.chars_per_page
    aim_chars = args.pages * cpp
    min_pages, max_pages = args.min_pages, args.max_pages

    # 분량 수렴: [min_pages, max_pages] 범위에 들 때까지 부족하면 확장, 넘치면 압축.
    for rnd in range(1, MAX_CONVERGE_ROUNDS + 1):
        chars = docx_writer.manuscript_char_count(sections)
        est = chars / cpp
        if min_pages <= est <= max_pages:
            break
        before = chars
        was_short = est < min_pages
        if was_short:
            print(f"      [보정 {rnd}] 분량 부족(~{est:.1f}쪽, 목표 {min_pages}쪽 이상) → 장 확장")
            _expand(client, sections, slides, aim_chars, chars)
        else:
            print(f"      [보정 {rnd}] 분량 초과(~{est:.1f}쪽, 상한 {max_pages}쪽) → 장 압축")
            _condense(client, sections, slides, aim_chars, chars)
        after = docx_writer.manuscript_char_count(sections)
        # 분량이 모자랄 때는 35쪽을 채울 때까지 라운드를 끝까지 쓴다.
        # (넘쳐서 압축하는 경우에만, 변화가 없으면 조기 종료)
        if not was_short and abs(after - before) < 0.015 * aim_chars:
            print("      분량 변화가 거의 없어 보정을 멈춥니다.")
            break

    print(f"[4/4] docx 저장: {out_path}")
    docx_writer.write_docx(out_path, args.title, sections)

    chars = docx_writer.manuscript_char_count(sections)
    est_pages = chars / cpp
    print(f"완료. 본문 {chars:,}자, 환산 약 {est_pages:.1f}쪽.")
    if est_pages < min_pages or est_pages > max_pages:
        print(f"  주의: 목표 범위({min_pages}~{max_pages}쪽)에 들지 못했습니다(추정 {est_pages:.1f}쪽). "
              "슬라이드 내용이 부족하거나 과다한 경우입니다. --min-pages/--max-pages 조정을 검토하세요.")
    return 0


def _generate_all(client, sections, slides) -> None:
    prev_tail = ""
    done_titles: list[str] = []
    for i, sec in enumerate(sections, 1):
        print(f"      ({i}/{len(sections)}) {sec.title} ...", flush=True)
        sec.body = generate.write_section(client, sec, slides, prev_tail, done_titles)
        prev_tail = sec.body
        done_titles.append(sec.title)


def _expand(client, sections, slides, aim_chars: int, cur_chars: int) -> None:
    """슬라이드 원문이 풍부한 장부터 확장해 목표(aim) 분량에 다가간다."""
    deficit = aim_chars - cur_chars
    # 분량을 끌어올릴 때는 모든 장을 슬라이드 분량에 비례해 함께 확장한다.
    chosen = list(range(len(sections)))
    wsum = sum(generate.section_source_chars(sections[i], slides) for i in chosen) or 1
    for i in sorted(chosen):
        src = generate.section_source_chars(sections[i], slides)
        add = deficit * src / wsum
        if add < 150:
            continue
        sections[i].target_chars = round((len(sections[i].body) + add) * OVERSHOOT)
        _regen(client, sections, slides, i,
               prompts.EXPAND_HINT.format(target_chars=sections[i].target_chars))


def _condense(client, sections, slides, aim_chars: int, cur_chars: int) -> None:
    """가장 긴 장부터 압축해 목표(aim) 분량으로 줄인다."""
    excess = cur_chars - aim_chars
    ranked = sorted(range(len(sections)), key=lambda i: len(sections[i].body), reverse=True)
    chosen = ranked[: max(1, len(sections) // 2)]
    wsum = sum(len(sections[i].body) for i in chosen) or 1
    for i in sorted(chosen):
        cut = excess * len(sections[i].body) / wsum
        sections[i].target_chars = max(400, round(len(sections[i].body) - cut))
        _regen(client, sections, slides, i,
               prompts.CONDENSE_HINT.format(target_chars=sections[i].target_chars))


def _regen(client, sections, slides, i: int, hint: str) -> None:
    prev_tail = sections[i - 1].body if i > 0 else ""
    toc = [s.title for s in sections[:i]]
    print(f"        · {sections[i].title} 재작성(목표 {sections[i].target_chars:,}자)", flush=True)
    sections[i].body = generate.write_section(client, sections[i], slides, prev_tail, toc, hint)


def _default_out(pdf_path: str) -> str:
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    return os.path.join(os.path.dirname(os.path.abspath(pdf_path)), f"{base}_원고.docx")


if __name__ == "__main__":
    raise SystemExit(main())
