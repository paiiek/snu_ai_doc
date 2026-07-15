"""나눔명조 폰트를 다운로드해 사용자 폰트 폴더에 설치한다.

규격은 "신명조"지만 이 이름의 폰트는 상용이라 기본 설치가 어렵다. 대신
동등한 한글 명조체이면서 오픈 폰트 라이선스(OFL)인 **나눔명조**를 자동으로
받아 설치해 준다. Word/LibreOffice/PDF 변환에서 명조체 규격이 안정적으로
유지되도록 하는 목적이다.

정책상 사용자가 명시적으로 요청할 때만 다운로드한다
(`--install-font` 옵션 또는 첫 실행 시 프롬프트).
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import urllib.request

# Google Fonts 저장소(google/fonts, OFL 라이선스)의 raw 파일. 안정 URL.
FONT_URLS = {
    "NanumMyeongjo-Regular.ttf":
        "https://raw.githubusercontent.com/google/fonts/main/ofl/nanummyeongjo/NanumMyeongjo-Regular.ttf",
    "NanumMyeongjo-Bold.ttf":
        "https://raw.githubusercontent.com/google/fonts/main/ofl/nanummyeongjo/NanumMyeongjo-Bold.ttf",
    "NanumMyeongjo-ExtraBold.ttf":
        "https://raw.githubusercontent.com/google/fonts/main/ofl/nanummyeongjo/NanumMyeongjo-ExtraBold.ttf",
}

_USER_AGENT = "Mozilla/5.0 (compatible; slides2manuscript/1.0)"

# 시스템에 있으면 명조체가 있는 것으로 인정하는 후보들.
_MYUNGJO_CANDIDATES = [
    "NanumMyeongjo", "나눔명조",
    "AppleMyungjo", "AppleMyongjo",
    "바탕", "Batang", "BatangChe",
    "함초롬바탕", "HYSMyeongJo", "HY신명조",
    "SM신명조",
]


def user_font_dir() -> str:
    """OS별 사용자 폰트 폴더 경로."""
    system = platform.system()
    if system == "Darwin":
        return os.path.expanduser("~/Library/Fonts")
    if system == "Windows":
        base = os.environ.get("LOCALAPPDATA", "")
        return os.path.join(base, "Microsoft", "Windows", "Fonts") if base else ""
    return os.path.expanduser("~/.local/share/fonts")


def _file_matches(name: str, needle: str) -> bool:
    return needle.lower().replace(" ", "") in name.lower().replace(" ", "")


def is_font_installed(font_family: str) -> bool:
    """이 이름의 폰트가 시스템에 있는지 대략 판정."""
    # 1) fc-list (macOS/Linux)
    if shutil.which("fc-list"):
        try:
            out = subprocess.check_output(["fc-list", ":lang=ko"], text=True, timeout=5)
            if _file_matches(out, font_family):
                return True
        except (subprocess.SubprocessError, OSError):
            pass
    # 2) 사용자 폰트 폴더의 파일 이름
    d = user_font_dir()
    if d and os.path.isdir(d):
        for n in os.listdir(d):
            if _file_matches(n, font_family):
                return True
    # 3) macOS 시스템 폰트 폴더의 몇 개는 직접 확인
    if platform.system() == "Darwin":
        for probe in ("/System/Library/Fonts/Supplemental/AppleMyungjo.ttf",):
            if os.path.exists(probe) and _file_matches(os.path.basename(probe), font_family):
                return True
    return False


def has_myungjo_font() -> bool:
    """명조체 계열이 하나라도 시스템에 있는지."""
    return any(is_font_installed(c) for c in _MYUNGJO_CANDIDATES)


def install_nanum_myeongjo(force: bool = False, verbose: bool = True) -> tuple[bool, str]:
    """나눔명조 3종을 다운로드해 사용자 폰트 폴더에 설치.

    - 이미 설치돼 있으면(파일 존재) force=False 이면 건너뜀.
    - 성공 시 (True, 메시지), 실패 시 (False, 사유).
    """
    if not force and is_font_installed("NanumMyeongjo"):
        return True, "나눔명조가 이미 설치돼 있음 (건너뜀)"

    font_dir = user_font_dir()
    if not font_dir:
        return False, "사용자 폰트 폴더를 찾지 못함 (LOCALAPPDATA 미설정 등)"
    os.makedirs(font_dir, exist_ok=True)

    saved = []
    for filename, url in FONT_URLS.items():
        dest = os.path.join(font_dir, filename)
        if os.path.exists(dest) and not force:
            saved.append(filename + "(기존)")
            continue
        try:
            if verbose:
                print(f"    다운로드: {filename}")
            req = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()
            with open(dest, "wb") as f:
                f.write(data)
            saved.append(filename)
        except Exception as exc:  # noqa: BLE001
            return False, f"{filename} 다운로드/저장 실패: {exc}"

    # 폰트 캐시 갱신 (Linux/일부 macOS 환경)
    if shutil.which("fc-cache"):
        try:
            subprocess.run(["fc-cache", "-f", font_dir], capture_output=True, timeout=30)
        except (subprocess.SubprocessError, OSError):
            pass

    msg = f"설치 완료 → {font_dir}\n    파일: {', '.join(saved)}"
    if platform.system() == "Windows":
        msg += ("\n※ Windows 는 방금 복사한 파일을 처음 사용할 때 개별로 '설치' 승인이 "
                "필요할 수 있습니다. 파일 탐색기에서 파일을 더블클릭 → '설치' 를 눌러 주세요.")
    return True, msg
