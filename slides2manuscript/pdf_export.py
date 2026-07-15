"""생성된 docx를 pdf로 함께 내보낸다.

두 가지 방식을 순서대로 시도해 처음 성공한 결과를 쓴다.
1. LibreOffice (`soffice --headless --convert-to pdf`)
   - 헤드리스라 창이 뜨지 않고, macOS/Linux/Windows 모두 동작.
   - 설치돼 있으면 가장 안정적이라 우선.
2. docx2pdf 파이썬 패키지
   - macOS/Windows의 Microsoft Word를 자동화해 변환.
   - Word가 잠깐 열렸다 닫힐 수 있음. Word가 있으면 대개 이 방식으로 성공.

둘 다 실패하면 pdf 없이 docx만 남기고 사용자에게 안내 문구를 띄운다.
"""

from __future__ import annotations

import os
import shutil
import subprocess


def convert_to_pdf(docx_path: str, pdf_path: str | None = None) -> tuple[str | None, str]:
    """docx → pdf 변환. 성공 시 (경로, 사용한 방식), 실패 시 (None, 사유)."""
    if not os.path.isfile(docx_path):
        return None, f"입력 docx를 찾을 수 없습니다: {docx_path}"
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    ok, msg = _convert_with_libreoffice(docx_path, pdf_path)
    if ok:
        return pdf_path, "libreoffice"

    lo_msg = msg
    ok, msg = _convert_with_docx2pdf(docx_path, pdf_path)
    if ok:
        return pdf_path, "docx2pdf(Word)"

    return None, (
        "PDF 변환에 사용할 도구를 찾지 못했습니다. "
        "LibreOffice(`brew install --cask libreoffice`) 또는 "
        "Microsoft Word 를 설치한 뒤 다시 시도하세요. "
        f"(LibreOffice: {lo_msg} / docx2pdf: {msg})"
    )


def _convert_with_libreoffice(docx_path: str, pdf_path: str) -> tuple[bool, str]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False, "soffice 실행 파일을 PATH에서 찾지 못함"
    out_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            capture_output=True, timeout=180, text=True,
        )
    except (subprocess.TimeoutExpired, OSError) as exc:
        return False, f"soffice 실행 실패: {exc}"
    if result.returncode != 0:
        return False, f"soffice 종료코드 {result.returncode}: {result.stderr.strip()[:200]}"
    # LibreOffice가 만드는 파일명은 <docx basename>.pdf. 요청 경로와 다르면 옮긴다.
    generated = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if os.path.abspath(generated) != os.path.abspath(pdf_path) and os.path.isfile(generated):
        os.replace(generated, pdf_path)
    if os.path.isfile(pdf_path):
        return True, "ok"
    return False, "변환 후 pdf 파일을 찾지 못함"


def _convert_with_docx2pdf(docx_path: str, pdf_path: str) -> tuple[bool, str]:
    try:
        from docx2pdf import convert  # type: ignore
    except ImportError as exc:
        return False, f"docx2pdf 미설치({exc})"
    try:
        convert(docx_path, pdf_path)
    except Exception as exc:  # noqa: BLE001 - 어떤 실패든 폴백 처리
        return False, f"docx2pdf 실패: {exc}"
    if os.path.isfile(pdf_path):
        return True, "ok"
    return False, "변환 후 pdf 파일을 찾지 못함"
