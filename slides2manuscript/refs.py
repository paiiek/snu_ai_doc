"""강사 홈페이지·수업 안내 등 참고 자료를 로드해 프롬프트용 블록으로 정리한다.

예술 강사처럼 슬라이드가 이미지 위주여서 텍스트 근거가 부족한 경우,
슬라이드 외에 강사 홈페이지, 전시 스테이트먼트, 커리큘럼 문서 같은
추가 자료를 근거로 넣기 위한 모듈이다.

의존성 최소 원칙에 따라 URL 스크래핑은 표준 라이브러리(urllib)와 정규식
HTML 정리로 처리한다. 사이트 구조에 따라 결과가 완벽하지 않을 수 있으므로,
안정적인 자료는 파일(.txt/.md/.pdf/.docx)로 넘기는 것을 권장한다.
"""

from __future__ import annotations

import html
import os
import re
import urllib.request
from dataclasses import dataclass
from urllib.parse import urlparse


# 참고 자료 하나가 프롬프트를 압도하지 않도록 잘라 넣는 상한.
# 너무 크면 슬라이드 근거가 희석되고, 토큰 비용도 급증한다.
MAX_CHARS_PER_REF = 20000

_USER_AGENT = "Mozilla/5.0 (compatible; slides2manuscript/1.0)"


@dataclass
class Ref:
    label: str      # [REF-...] 태그로 프롬프트에 노출되는 표기
    source: str     # 원본 경로/URL (로그용)
    text: str       # 정리된 본문 텍스트


def _shorten(text: str, limit: int = MAX_CHARS_PER_REF) -> str:
    if len(text) <= limit:
        return text
    return text[: limit - 200] + "\n\n(... 참고 자료 일부 생략 ...)"


def _clean_whitespace(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n")]
    out: list[str] = []
    for ln in lines:
        if ln == "" and (not out or out[-1] == ""):
            continue
        out.append(ln)
    return "\n".join(out).strip()


def _strip_html(raw: str) -> str:
    raw = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", raw)
    raw = re.sub(r"(?i)</(p|div|section|article|li|tr|h[1-6])>", "\n", raw)
    raw = re.sub(r"(?i)<br\s*/?>", "\n", raw)
    raw = re.sub(r"<[^>]+>", "", raw)
    raw = html.unescape(raw)
    return _clean_whitespace(raw)


def _fetch_url(url: str, timeout: float = 15.0) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        data = resp.read()
    try:
        raw = data.decode(charset, errors="replace")
    except LookupError:
        raw = data.decode("utf-8", errors="replace")
    return _strip_html(raw)


def _read_pdf(path: str) -> str:
    import fitz  # PyMuPDF (본체에서 이미 사용 중)
    doc = fitz.open(path)
    try:
        parts = [page.get_text("text") for page in doc]
    finally:
        doc.close()
    return _clean_whitespace("\n\n".join(parts))


def _read_docx(path: str) -> str:
    from docx import Document  # python-docx (본체에서 이미 사용 중)
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    return _clean_whitespace("\n".join(parts))


def _read_text(path: str) -> str:
    with open(path, encoding="utf-8", errors="replace") as f:
        return _clean_whitespace(f.read())


def load_file(path: str, label: str | None = None) -> Ref:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"참고 파일을 찾을 수 없습니다: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        text = _read_pdf(path)
    elif ext == ".docx":
        text = _read_docx(path)
    else:
        text = _read_text(path)
    label = label or os.path.splitext(os.path.basename(path))[0]
    return Ref(label=label, source=path, text=_shorten(text))


def _label_from_url(url: str) -> str:
    try:
        host = urlparse(url).netloc or "웹"
    except Exception:
        host = "웹"
    return host


def load_url(url: str, label: str | None = None) -> Ref:
    text = _fetch_url(url)
    label = label or _label_from_url(url)
    return Ref(label=label, source=url, text=_shorten(text))


def load_refs(files: list[str] | None, urls: list[str] | None) -> list[Ref]:
    refs: list[Ref] = []
    for path in files or []:
        try:
            refs.append(load_file(path))
        except Exception as exc:  # noqa: BLE001 - 개별 실패는 건너뛰고 진행
            print(f"  (참고 파일 로드 실패, 건너뜀: {path} — {exc})")
    for url in urls or []:
        try:
            refs.append(load_url(url))
        except Exception as exc:  # noqa: BLE001
            print(f"  (참고 URL 로드 실패, 건너뜀: {url} — {exc})")
    return refs


def refs_block(refs: list[Ref]) -> str:
    """프롬프트에 붙일 참고 자료 블록. refs가 비면 빈 문자열."""
    parts = [f"[REF-{r.label}]\n{r.text}" for r in refs if r.text]
    if not parts:
        return ""
    body = "\n\n".join(parts)
    return (
        "\n\n[참고 자료]\n"
        "아래 자료는 강사 홈페이지·수업 소개 등 슬라이드 밖에서 별도로 제공된 참고 자료다. "
        "슬라이드 내용과 함께 원고의 근거로 사용할 수 있다. "
        "여기에 없는 사실은 지어내지 않는다.\n\n"
        + body
    )


def total_chars(refs: list[Ref]) -> int:
    return sum(len(r.text) for r in refs)
