# 프로젝트: slides2manuscript

전역 `~/.claude/CLAUDE.md`(oh-my-claudecode)의 원칙을 따르되, 이 저장소에서는
아래 프로젝트 규칙이 우선한다.

## 무엇을 만드는가

조교가 받은 **강의 슬라이드 PDF**를 **줄글 강의 원고(.docx)**로 변환하는
독립 실행 파이썬 CLI. 발표 대본이 아니라 사람이 쓴 듯한 한국어 산문 원고를 만든다.
다른 조교들에게 폴더째 배포해 쓰는 것이 목표다(별도 도구 의존성 없이).

## 절대 원칙 (출력물 품질)

1. **줄글 산문만.** 글머리표·번호목록·표·마크다운 기호·이모지 금지.
   "다음과 같다", "아래 그림" 같은 슬라이드 지시 표현도 쓰지 않는다.
2. **슬라이드(및 함께 넘긴 참고 자료)에 충실.** 이 범위 밖의 사실·수치·인용·출처를
   지어내지 않는다. 내용이 빈약하면 억지로 늘리지 말고 주어진 범위에서 풀어 쓴다.
   참고 자료는 `--ref-file`/`--ref-url` 로 명시적으로 주입된 것에 한한다.
3. **AI 티 제거.** 정형화된 도입/마무리 문구, '결론적으로/요약하자면' 남발,
   모든 문단을 접속사로 여는 습관, 같은 구조·길이 문장 반복 금지. 문장 길이를 변주한다.
   강의자가 직접 쓴 평서형('-이다/-한다') 학술 강의체. 구어 추임새('자', '여러분') 금지.
4. **드라이한 서식.** 콘진원 원고료 책정표 규격을 따른다: **신명조 12pt · 35줄/쪽 ·
   상하 15mm · 좌우 20mm · 머리말/꼬리말 15mm**. 본문 글꼴은 하나만 쓰고, 장 제목만
   살짝 키운다. 색·표·테두리·장식 없음. 35줄/쪽을 지키기 위해 문단 사이 추가 여백은
   두지 않는다(space_after = 0).
5. **분량.** 과제 기준별 최소 쪽수 이상을 반드시 보장한다.
   - 내부 1강: 35쪽 / 내부 2강: 70쪽
   - 외부 1강: 17쪽 / 외부 2강: 34쪽
   `pages × chars_per_page`로 목표 글자 수를 잡고, 부족하면 수렴 루프가 모든 장을
   슬라이드 근거 안에서 확장해 최소 쪽수를 채운다. 슬라이드 내용이 적어도 목표를
   채우므로 늘어짐이 없는지 사람 검토가 필요하다.

위 1~3번의 실제 지침 문구는 `slides2manuscript/prompts.py`의 `STYLE_RULES`에 있다.
원고 톤을 바꿀 일이 생기면 코드 흩뿌리지 말고 거기서만 고친다.

## 구조

```
make_manuscript.py            # 진입 스크립트 (python make_manuscript.py 강의.pdf)
slides2manuscript/
  cli.py          # argparse, 파이프라인 오케스트레이션, 분량 질문(_resolve_volume),
                  #   장 수 자동 산정(_auto_sections), 분량 수렴(_expand/_condense), 키 설정
  extract.py      # PyMuPDF로 슬라이드별 텍스트 추출 + 이미지PDF 감지
  refs.py         # (--ref-file/--ref-url) 강사 홈페이지·수업자료 등 참고 자료 로드
                  #   → 프롬프트 뒤에 [참고 자료] 블록으로 주입. 슬라이드 부족분 보강용
  vision.py       # (--vision) 슬라이드를 이미지로 렌더링 → 비전 모델로 내용 보강(병렬)
  llm.py          # provider 추상화(anthropic/openai): chat / vision_chat
  prompts.py      # 모든 프롬프트(아웃라인/장 본문/확장·압축) + STYLE_RULES  ← 톤의 단일 출처
  generate.py     # design_outline → write_section, 분량 배분
  docx_writer.py  # 드라이 docx 출력(한글 글꼴 강제, A4/여백/줄간격 고정)
  pdf_export.py   # docx → pdf 자동 변환. LibreOffice(soffice) 우선, 없으면 docx2pdf(=Word)
```

파이프라인: 추출 → (선택)비전 보강 → 장 구성 설계(1콜) → 장별 본문 생성(장 수만큼 콜)
→ 분량 수렴 보정 → docx 저장 → (기본 동작) pdf 변환.

## 작업 규칙

- **유료 API 호출 주의.** `generate.py`의 실호출(`messages.create`)은 사용자 키로
  과금된다. 사용자가 명시적으로 요청하지 않는 한 실제 PDF로 전체 파이프라인을
  돌리지 마라. 로직 검증은 API를 타지 않는 경로(extract/docx/fallback/budget)로 한다.
- **텍스트 기반 PDF가 기본.** 텍스트가 적은 이미지/스캔 PDF는 기본적으로 막고(`--force`),
  `--vision`이면 슬라이드를 이미지로 읽어 처리한다(`vision.py`, 슬라이드당 비전 1콜 추가).
- 기본 모델은 `claude-opus-4-8`(자연스러운 한국어 우선). 비용 옵션은 `--model claude-sonnet-4-6`.
- 의존성은 최소로 유지한다(pymupdf, anthropic, openai, python-docx). 배포 단순성이 중요하다.
- 한국어 주석/메시지를 유지한다. 사용자와 다른 조교가 읽는다.

## 테스트

API 없이 확인 가능한 부분만으로 회귀를 잡는다.

```bash
pip install -r requirements.txt
python make_manuscript.py --help                 # CLI 파싱
# 추출/ docx / 폴백 / 분량배분 경로는 임시 PDF를 만들어 단위 확인 (scratchpad)
```

실제 원고 품질(줄글다움, AI 티, 충실성)은 사용자가 본인 키로 1회 돌려 육안 검토한다.
