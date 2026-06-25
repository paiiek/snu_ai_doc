#!/usr/bin/env python3
"""단일 진입 스크립트. 사용 예:

    export ANTHROPIC_API_KEY=sk-...
    python make_manuscript.py 강의1.pdf
    python make_manuscript.py 강의1.pdf -o 강의1_원고.docx --pages 28 -t "1주차 강의"
"""

from slides2manuscript.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
