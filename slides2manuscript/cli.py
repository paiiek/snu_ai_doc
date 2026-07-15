"""명령행 진입점: 슬라이드 PDF -> 줄글 강의 원고(docx)."""

from __future__ import annotations

import argparse
import os
import sys

from . import __version__, generate, docx_writer, llm, prompts, refs as refs_mod, pdf_export, font_setup
from .extract import extract_slides, total_source_chars, looks_like_image_pdf
from .progress import Progress


# 키를 적어 두는 파일 이름 후보. macOS Finder는 점(.)으로 시작하는 파일을
# 만들지 못하므로, 점 없는 이름(key.txt 등)을 먼저 권장한다.
KEY_FILENAMES = ("key.txt", "env.txt", ".env")


def _key_dirs() -> list[str]:
    """키 파일을 찾을 폴더: 현재 폴더 + 프로그램이 설치된 폴더."""
    project = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dirs = [os.getcwd(), project]
    out: list[str] = []
    for d in dirs:
        if d not in out:
            out.append(d)
    return out


def _load_dotenv() -> None:
    """key.txt / env.txt / .env 에서 키를 읽어 환경에 채운다.

    별도 의존성 없이 KEY=VALUE 형식만 처리한다. 이미 설정된 환경변수는 덮지 않는다.
    """
    for d in _key_dirs():
        for name in KEY_FILENAMES:
            path = os.path.join(d, name)
            if not os.path.isfile(path):
                continue
            with open(path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, val = line.partition("=")
                    key, val = key.strip(), val.strip().strip('"').strip("'")
                    if key and key not in os.environ:
                        os.environ[key] = val


def _has_key() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("OPENAI_API_KEY"))


def _ensure_myungjo_font(user_font: str | None) -> None:
    """명조체 폰트가 하나도 없으면 사용자에게 물어 나눔명조를 자동 설치한다.

    - 사용자가 `--font` 로 폰트 이름을 명시했으면, 그 폰트가 있는지만 확인하고
      없으면 경고만 띄운다(임의 대체 안 함).
    - 대화형 터미널이 아니면 자동 설치 없이 경고만.
    """
    if user_font:
        if not font_setup.is_font_installed(user_font):
            print(f"경고: 지정한 폰트 '{user_font}' 를 시스템에서 찾지 못했습니다. "
                  "PDF에서 대체 폰트로 렌더링될 수 있습니다.")
        return
    if font_setup.has_myungjo_font():
        return
    print("\n명조체 폰트가 시스템에 없습니다. 규격(신명조)에 맞추려면 명조체가 필요합니다.")
    if not sys.stdin.isatty():
        print("  자동 설치를 원하면 `--install-font` 로 다시 실행하세요.")
        return
    ans = input("나눔명조(OFL, 자유 사용)를 지금 자동 설치하시겠습니까? [Y/n]: ").strip().lower()
    if ans in ("", "y", "yes"):
        ok, msg = font_setup.install_nanum_myeongjo()
        print(("  " + msg) if ok else f"  실패: {msg}")
    else:
        print("  건너뜁니다. 원할 때 `--install-font` 로 언제든 설치할 수 있습니다.")


def _interactive_key_setup() -> None:
    """키 파일이 없을 때, 화면에서 직접 키를 입력받아 key.txt 로 저장한다.

    터미널(대화형)에서 실행할 때만 작동한다. 비개발자가 .env 파일을 손으로
    만들지 않아도 되도록 돕는 안전망이다.
    """
    if not sys.stdin.isatty():
        return
    print("\nAPI 키가 아직 설정되어 있지 않습니다. 지금 입력하면 key.txt 에 저장해 둡니다.")
    print("  1) Anthropic(클로드) 키 사용   2) OpenAI(GPT) 키 사용")
    choice = input("번호 입력(1 또는 2): ").strip()
    if choice == "2":
        var, prefix = "OPENAI_API_KEY", "sk-"
    else:
        var, prefix = "ANTHROPIC_API_KEY", "sk-ant-"
    key = input(f"{var} 값을 붙여넣고 엔터(예: {prefix}...): ").strip()
    if not key:
        return
    os.environ[var] = key
    path = os.path.join(os.getcwd(), "key.txt")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"{var}={key}\n")
        print(f"저장 완료: {path}\n")
    except OSError as exc:
        print(f"(key.txt 저장 실패: {exc} — 이번 실행에는 입력한 키를 그대로 사용합니다)\n")

# A4 한 쪽에 실제로 들어가는 대략의 글자 수.
# 새 규격: 신명조 12pt · 35줄/쪽 · 상하 15mm · 좌우 20mm.
# 계산: 좌우 여백 뺀 폭 170mm ÷ 12pt(≈4.23mm/자) ≈ 40자/줄, 35줄 ≈ 1,400자/쪽.
# 문단 끝의 미충한 줄로 인한 손실을 감안해 실효값은 그보다 낮다.
# 최소 분량을 확실히 넘기도록(수렴 루프가 est<min이면 확장) 보수적으로 조금 높게 잡는다.
# 실제 워드로 열었을 때의 쪽수는 폰트 렌더링에 따라 1~2쪽 달라질 수 있다.
CHARS_PER_PAGE = 1400

# 모델이 장별 목표 글자 수의 약 75~80%만 쓰는 경향이 있어, 생성 목표를 그만큼
# 키워 잡는다. (실측 결과 실현율 ≈ 0.77 → 보정 계수 ≈ 1.3)
OVERSHOOT = 1.3

# 분량 수렴 보정 최대 반복 횟수(범위를 벗어났을 때만 작동).
# 분량 부족 시 끝까지 끌어올려야 하므로 넉넉히 둔다(큰 목표 분량 대비).
MAX_CONVERGE_ROUNDS = 8


def _auto_sections(pages: int, given_min, given_max):
    """목표 분량에 맞춰 장 수를 자동 산정한다(사용자가 지정하면 그 값 사용).

    한 장이 약 2~3쪽(≈3,000자) 정도가 되도록 잡아, 장이 너무 길어
    모델 출력 한도에서 잘리는 것을 막는다. 52·104쪽 같은 큰 목표도 커버한다.
    """
    min_s = given_min if given_min else max(8, round(pages / 3.5))
    max_s = given_max if given_max else max(min_s + 2, round(pages / 2.2))
    max_s = min(max_s, 45)          # 동시 실행/토큰 한도 고려한 상한
    min_s = min(min_s, max_s)
    return min_s, max_s


# 과제 기준별 "최소" 분량(쪽). 결과물은 반드시 이 쪽수 이상이 되어야 한다.
# 근거: 전문가 활용비 강의료·원고료 책정표(원고료 장당 3만원).
#   내부 1시간30분: 35장 / 내부 3시간: 70장
#   외부 1시간30분: 17장 / 외부 3시간: 34장
# (강의 수, 구분) -> 최소 쪽수
VOLUME_MIN = {
    (1, "내부"): 35,
    (1, "외부"): 17,
    (2, "내부"): 70,
    (2, "외부"): 34,
}
AIM_MARGIN = 3   # 최소를 안전하게 넘기도록 목표는 최소보다 이만큼 위로 잡는다.


def _ask_volume(max_override):
    """실행 시 1강/2강·내부/외부를 물어 최소 분량을 정한다. (aim, min, max) 반환."""
    print("\n원고 분량 기준을 선택하세요. (그냥 엔터를 누르면 기본값)")
    n = input("  강의 수 —  1) 1강(기본)   2) 2강 연속  : ").strip()
    lectures = 2 if n == "2" else 1
    s = input("  구분   —  1) 내부 35/70장(기본)   2) 외부 17/34장  : ").strip()
    scope = "외부" if s == "2" else "내부"
    min_p = VOLUME_MIN[(lectures, scope)]
    aim_p = min_p + AIM_MARGIN
    max_p = max_override or (aim_p + 20)
    print(f"  → 선택: {lectures}강 {scope} · 최소 {min_p}쪽 이상 (목표 {aim_p}쪽)\n")
    return aim_p, min_p, max_p


def _resolve_volume(args):
    """분량(aim, min, max)을 결정한다.

    - 옵션(--pages/--min-pages)을 직접 주면 그것을 우선한다.
    - 아무것도 안 주고 터미널이면 1강/2강·내부/외부를 물어본다.
    - 비대화형(자동 실행 등)이면 기본값(1강 내부=최소 35쪽).
    """
    if args.pages or args.min_pages:
        min_p = args.min_pages or max(1, args.pages - AIM_MARGIN)
        aim_p = args.pages or (min_p + AIM_MARGIN)
        max_p = args.max_pages or (aim_p + 20)
        return aim_p, min_p, max_p
    if sys.stdin.isatty():
        return _ask_volume(args.max_pages)
    min_p = 35
    return min_p + AIM_MARGIN, min_p, (args.max_pages or min_p + AIM_MARGIN + 20)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="slides2manuscript",
        description="강의 슬라이드 PDF를 줄글 강의 원고(docx)로 변환한다.",
    )
    p.add_argument("pdf", nargs="?", help="입력 슬라이드 PDF 경로(텍스트 기반). "
                   "`--install-font` 만 실행할 때는 생략 가능")
    p.add_argument("-o", "--out", help="출력 docx 경로(기본: 입력파일명_원고.docx)")
    p.add_argument("-t", "--title", default="", help="원고 맨 위 제목(기본: 없음)")
    p.add_argument("--pages", type=int, default=None,
                   help="목표(중앙값) A4 분량. 미지정 시 실행할 때 1강/2강·내부/외부를 물어봄")
    p.add_argument("--min-pages", type=int, default=None, help="허용 최소 분량(무조건 보장)")
    p.add_argument("--max-pages", type=int, default=None, help="허용 최대 분량")
    p.add_argument("--provider", choices=("auto",) + llm.PROVIDERS, default="auto",
                   help="LLM 제공자(기본: auto = 설정된 키로 자동 선택). anthropic 또는 openai")
    p.add_argument("--model", default=None,
                   help="사용할 모델(기본: provider별 기본값). "
                        f"anthropic→{llm.DEFAULT_MODELS['anthropic']}, "
                        f"openai→{llm.DEFAULT_MODELS['openai']}")
    p.add_argument("--base-url", default=None,
                   help="OpenAI 호환 엔드포인트 주소(선택). OPENAI_BASE_URL 로도 지정 가능")
    p.add_argument("--vision", action="store_true",
                   help="슬라이드를 이미지로 읽어 그림·도표·수식까지 반영(텍스트 적은 자료 권장). "
                        "비전 호출이 슬라이드 수만큼 추가돼 비용·시간이 늘어남")
    p.add_argument("--vision-workers", type=int, default=4,
                   help="비전 읽기 동시 처리 수(기본: 4)")
    p.add_argument("--chars-per-page", type=int, default=CHARS_PER_PAGE,
                   help=f"A4 한 쪽당 글자 수 환산값(기본: {CHARS_PER_PAGE})")
    p.add_argument("--min-sections", type=int, default=None, help="최소 장 수(기본: 분량에 맞춰 자동)")
    p.add_argument("--max-sections", type=int, default=None, help="최대 장 수(기본: 분량에 맞춰 자동)")
    p.add_argument("--ref-file", nargs="+", default=None, metavar="PATH",
                   help="참고 자료 파일(.txt/.md/.pdf/.docx). 여러 개 지정 가능. "
                        "강사 홈페이지에서 복사한 텍스트, 수업 안내서 등 슬라이드 외 근거로 쓰인다.")
    p.add_argument("--ref-url", nargs="+", default=None, metavar="URL",
                   help="참고 자료 URL. 강사 홈페이지 등 페이지 본문 텍스트를 추출해 근거로 쓴다. "
                        "사이트 구조에 따라 추출이 완벽하지 않을 수 있어, 중요한 자료는 파일로 넘기는 편이 안전하다.")
    p.add_argument("--force", action="store_true",
                   help="이미지/스캔 PDF로 의심돼도 그대로 진행")
    p.add_argument("--no-pdf", action="store_true",
                   help="docx 저장 후 pdf도 함께 만드는 기본 동작을 끈다")
    p.add_argument("--font", default=None, metavar="NAME",
                   help="본문 폰트 이름. 미지정 시 나눔명조(설치돼 있으면) → OS별 기본 명조체 순으로 자동 선택")
    p.add_argument("--install-font", action="store_true",
                   help="나눔명조(OFL 라이선스)를 사용자 폰트 폴더에 자동 다운로드·설치하고 종료. "
                        "명조체 폰트가 없어 PDF가 산세리프로 렌더링되는 문제를 해결한다.")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    # 폰트 자동 설치 모드: PDF 인자 없이도 실행 가능. 설치 후 종료.
    if args.install_font:
        ok, msg = font_setup.install_nanum_myeongjo(force=True)
        print(("[폰트 설치 성공] " if ok else "[폰트 설치 실패] ") + msg)
        return 0 if ok else 1

    if not args.pdf:
        print("입력 PDF 경로가 필요합니다. `--install-font` 만 실행하려면 옵션만 주세요.",
              file=sys.stderr)
        return 2

    _load_dotenv()
    if not _has_key():
        _interactive_key_setup()

    # 명조체 폰트가 하나도 없으면(=PDF 규격이 깨지는 상태) 대화형으로 자동 설치 유도.
    _ensure_myungjo_font(args.font)

    if not os.path.isfile(args.pdf):
        print(f"입력 파일을 찾을 수 없습니다: {args.pdf}", file=sys.stderr)
        return 2

    out_path = args.out or _default_out(args.pdf)

    print(f"[1/4] PDF 추출: {args.pdf}")
    slides = extract_slides(args.pdf)
    nonempty = [s for s in slides if s.text]
    print(f"      슬라이드 {len(slides)}장, 텍스트 있는 슬라이드 {len(nonempty)}장, "
          f"원문 {total_source_chars(slides):,}자")

    loaded_refs = refs_mod.load_refs(args.ref_file, args.ref_url)
    if loaded_refs:
        print(f"      참고 자료 {len(loaded_refs)}건, 총 {refs_mod.total_chars(loaded_refs):,}자")
        for r in loaded_refs:
            print(f"        · [{r.label}] {r.source} ({len(r.text):,}자)")
    refs_block_text = refs_mod.refs_block(loaded_refs)

    provider = llm.resolve_provider(args.provider)
    client = llm.LLMClient(provider, args.model, args.base_url)

    if args.vision:
        print(f"[1.5/4] 비전으로 슬라이드 읽는 중 ({len(slides)}장, {provider}/{client.model})")
        from . import vision
        vbar = Progress(len(slides), "비전 읽기")
        vision.enrich_slides(client, args.pdf, slides, workers=args.vision_workers,
                             progress=lambda done, total: vbar.update(done))
        vbar.done()
        print(f"      보강 완료: 원문 {total_source_chars(slides):,}자")
    else:
        if not nonempty:
            print("텍스트를 전혀 추출하지 못했습니다. 이미지/스캔 PDF로 보입니다.\n"
                  "      --vision 을 붙이면 슬라이드 이미지를 읽어 처리할 수 있습니다.",
                  file=sys.stderr)
            return 3
        if looks_like_image_pdf(slides) and not args.force:
            print("경고: 텍스트가 매우 적어 이미지/스캔 PDF로 의심됩니다.\n"
                  "      --vision 으로 그림·도표까지 읽거나, 그대로 진행하려면 --force 를 붙이세요.",
                  file=sys.stderr)
            return 4

    pages, min_pages, max_pages = _resolve_volume(args)
    page_chars = pages * args.chars_per_page
    total_target = round(page_chars * OVERSHOOT)

    print(f"[2/4] 장 구성 설계 (최소 {min_pages}쪽 이상 / 목표 약 {pages}쪽 / 본문 {page_chars:,}자, "
          f"{provider} / {client.model})")
    min_sections, max_sections = _auto_sections(pages, args.min_sections, args.max_sections)
    print(f"      (장 수 목표: {min_sections}~{max_sections})")
    sections = generate.design_outline(
        client, slides, pages, min_sections, max_sections, refs_block_text
    )
    generate.allocate_char_budget(sections, slides, total_target)
    for i, sec in enumerate(sections, 1):
        print(f"      {i:>2}. {sec.title}  (S{sec.start}-S{sec.end}, 목표 {sec.target_chars:,}자)")

    print(f"[3/4] 본문 생성 ({len(sections)}개 장)")
    _generate_all(client, sections, slides, refs_block_text)

    cpp = args.chars_per_page
    aim_chars = pages * cpp

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
            _expand(client, sections, slides, aim_chars, chars, f"보정{rnd} 확장", refs_block_text)
            after = docx_writer.manuscript_char_count(sections)
            # 확장이 정체되면 장을 쪼개 용량을 늘려 최소 분량을 채운다.
            if after - before < 0.01 * aim_chars:
                if not _split_for_capacity(client, sections, slides, refs_block_text):
                    print("      더 늘릴 여지가 없어 보정을 멈춥니다(슬라이드 내용 한계).")
                    break
        else:
            print(f"      [보정 {rnd}] 분량 초과(~{est:.1f}쪽, 상한 {max_pages}쪽) → 장 압축")
            _condense(client, sections, slides, aim_chars, chars, f"보정{rnd} 압축", refs_block_text)
            after = docx_writer.manuscript_char_count(sections)
            if before - after < 0.015 * aim_chars:
                print("      분량 변화가 거의 없어 보정을 멈춥니다.")
                break

    print(f"[4/4] docx 저장: {out_path}")
    docx_writer.write_docx(out_path, args.title, sections, body_font=args.font)

    if not args.no_pdf:
        pdf_path, info = pdf_export.convert_to_pdf(out_path)
        if pdf_path:
            print(f"       pdf 저장: {pdf_path}  ({info})")
        else:
            print(f"       pdf 생성 건너뜀: {info}")

    chars = docx_writer.manuscript_char_count(sections)
    est_pages = chars / cpp
    print(f"완료. 본문 {chars:,}자, 환산 약 {est_pages:.1f}쪽.")
    if est_pages < min_pages or est_pages > max_pages:
        print(f"  주의: 목표 범위({min_pages}~{max_pages}쪽)에 들지 못했습니다(추정 {est_pages:.1f}쪽). "
              "슬라이드 내용이 부족하거나 과다한 경우입니다. --min-pages/--max-pages 조정을 검토하세요.")
    return 0


def _generate_all(client, sections, slides, refs_block: str = "") -> None:
    bar = Progress(len(sections), "본문 생성")
    prev_tail = ""
    done_titles: list[str] = []
    for sec in sections:
        sec.body = generate.write_section(
            client, sec, slides, prev_tail, done_titles, refs_block=refs_block
        )
        prev_tail = sec.body
        done_titles.append(sec.title)
        bar.update()
    bar.done()


def _expand(client, sections, slides, aim_chars: int, cur_chars: int, label: str,
            refs_block: str = "") -> None:
    """슬라이드 원문이 풍부한 장부터 확장해 목표(aim) 분량에 다가간다."""
    deficit = aim_chars - cur_chars
    # 분량을 끌어올릴 때는 모든 장을 슬라이드 분량에 비례해 함께 확장한다.
    chosen = list(range(len(sections)))
    wsum = sum(generate.section_source_chars(sections[i], slides) for i in chosen) or 1
    bar = Progress(len(chosen), label)
    for i in sorted(chosen):
        src = generate.section_source_chars(sections[i], slides)
        add = deficit * src / wsum
        if add >= 150:
            sections[i].target_chars = round((len(sections[i].body) + add) * OVERSHOOT)
            _regen(client, sections, slides, i,
                   prompts.EXPAND_HINT.format(target_chars=sections[i].target_chars),
                   mode="grow", refs_block=refs_block)
        bar.update()
    bar.done()


def _condense(client, sections, slides, aim_chars: int, cur_chars: int, label: str,
              refs_block: str = "") -> None:
    """가장 긴 장부터 압축해 목표(aim) 분량으로 줄인다."""
    excess = cur_chars - aim_chars
    ranked = sorted(range(len(sections)), key=lambda i: len(sections[i].body), reverse=True)
    chosen = ranked[: max(1, len(sections) // 2)]
    wsum = sum(len(sections[i].body) for i in chosen) or 1
    bar = Progress(len(chosen), label)
    for i in sorted(chosen):
        cut = excess * len(sections[i].body) / wsum
        sections[i].target_chars = max(400, round(len(sections[i].body) - cut))
        _regen(client, sections, slides, i,
               prompts.CONDENSE_HINT.format(target_chars=sections[i].target_chars),
               mode="shrink", refs_block=refs_block)
        bar.update()
    bar.done()


def _regen(client, sections, slides, i: int, hint: str, mode: str = "free",
           refs_block: str = "") -> None:
    """장을 다시 생성한다.

    mode="grow"   : 새 본문이 더 길 때만 채택(확장이 줄어들지 않게 → 단조 증가)
    mode="shrink" : 새 본문이 더 짧을 때만 채택(압축이 늘어나지 않게 → 단조 감소)
    mode="free"   : 무조건 교체(신규 장 생성 등)
    """
    old = sections[i].body
    prev_tail = sections[i - 1].body if i > 0 else ""
    toc = [s.title for s in sections[:i]]
    new = generate.write_section(
        client, sections[i], slides, prev_tail, toc, hint, refs_block=refs_block
    )
    if mode == "grow":
        sections[i].body = new if len(new) > len(old) else old
    elif mode == "shrink":
        sections[i].body = new if 200 < len(new) < len(old) else old
    else:
        sections[i].body = new


def _split_for_capacity(client, sections, slides, refs_block: str = "") -> bool:
    """확장이 정체될 때, 슬라이드 2장 이상인 가장 긴 장을 둘로 나눠 용량을 늘린다."""
    cand = [(i, s) for i, s in enumerate(sections) if s.end > s.start]
    if not cand:
        return False
    i, sec = max(cand, key=lambda t: len(t[1].body))
    mid = (sec.start + sec.end) // 2
    half = max(1800, round(len(sec.body) / 2 * 1.2))
    a = generate.Section(sec.title, sec.start, mid, target_chars=half)
    b = generate.Section(sec.title + " (계속)", mid + 1, sec.end, target_chars=half)
    sections[i:i + 1] = [a, b]
    print(f"      · 정체 → 장 분할: '{sec.title}' (S{sec.start}-S{sec.end})를 둘로 나눠 재작성")
    _regen(client, sections, slides, i, "", mode="free", refs_block=refs_block)
    _regen(client, sections, slides, i + 1, "", mode="free", refs_block=refs_block)
    return True


def _default_out(pdf_path: str) -> str:
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    return os.path.join(os.path.dirname(os.path.abspath(pdf_path)), f"{base}_원고.docx")


if __name__ == "__main__":
    raise SystemExit(main())
