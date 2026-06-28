"""진행 막대(progress bar) + 경과/예상 남은 시간 표시.

외부 의존성 없이 동작한다.
- 터미널(대화형): 한 줄을 갱신하는 실시간 막대.
- 비대화형(백그라운드 로그 파일 등): 10%마다 한 줄씩 깔끔하게 출력.
"""

from __future__ import annotations

import sys
import time


def _fmt(sec: float) -> str:
    sec = int(max(0, sec))
    m, s = divmod(sec, 60)
    h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


class Progress:
    def __init__(self, total: int, label: str, stream=None):
        self.total = max(1, total)
        self.label = label
        self.stream = stream or sys.stdout
        self.tty = bool(getattr(self.stream, "isatty", lambda: False)())
        self.start = time.time()
        self.n = 0
        self._last_step = -1
        if not self.tty:
            print(f"  {label}: 시작 (총 {total})", file=self.stream, flush=True)

    def update(self, n: int | None = None) -> None:
        self.n = self.n + 1 if n is None else n
        if self.tty:
            self._render_tty()
            return
        step = int(self.n / self.total * 100) // 10  # 0,1,..,10 (10% 단위)
        if step > self._last_step or self.n >= self.total:
            self._last_step = step
            pct = int(self.n / self.total * 100)
            print(f"  {self.label}: {self.n}/{self.total} ({pct}%)  "
                  f"경과 {_fmt(self._elapsed())}  남음 ~{_fmt(self._eta())}",
                  file=self.stream, flush=True)

    def done(self) -> None:
        self.n = self.total
        if self.tty:
            self._render_tty()
            print(file=self.stream)
        else:
            print(f"  {self.label}: 완료 ({self.total}개, {_fmt(self._elapsed())})",
                  file=self.stream, flush=True)

    def _elapsed(self) -> float:
        return time.time() - self.start

    def _eta(self) -> float:
        return self._elapsed() / self.n * (self.total - self.n) if self.n else 0.0

    def _render_tty(self) -> None:
        frac = self.n / self.total
        width = 24
        filled = int(width * frac)
        bar = "█" * filled + "░" * (width - filled)
        print(f"\r  {self.label} [{bar}] {self.n}/{self.total} ({frac * 100:3.0f}%)  "
              f"경과 {_fmt(self._elapsed())}  남음 ~{_fmt(self._eta())}   ",
              end="", file=self.stream, flush=True)
