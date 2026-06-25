"""LLM 백엔드 추상화: Anthropic 또는 OpenAI 중 가진 키로 동작한다.

생성 호출은 (system, user) -> text 한 가지 형태뿐이라 백엔드를 갈아끼우기 쉽다.
조교마다 Anthropic 키 또는 OpenAI 키 중 가진 것을 쓰면 된다.
"""

from __future__ import annotations

import os

# provider별 기본 모델. 자연스러운 한국어 산문 품질을 우선해 상위 모델을 둔다.
# 키가 지원하는 다른 모델을 쓰려면 CLI에서 --model 로 바꾼다.
DEFAULT_MODELS = {
    "anthropic": "claude-opus-4-8",
    "openai": "gpt-4o",
}

PROVIDERS = tuple(DEFAULT_MODELS.keys())


def resolve_provider(explicit: str | None) -> str:
    """provider를 결정한다. 명시값 우선, 없으면 설정된 키로 자동 판단."""
    if explicit and explicit != "auto":
        if explicit not in DEFAULT_MODELS:
            raise SystemExit(f"알 수 없는 provider: {explicit} (가능: {', '.join(PROVIDERS)})")
        return explicit
    if os.environ.get("ANTHROPIC_API_KEY"):
        return "anthropic"
    if os.environ.get("OPENAI_API_KEY"):
        return "openai"
    raise SystemExit(
        "API 키가 없습니다. 둘 중 하나를 설정하세요.\n"
        "  Anthropic:  export ANTHROPIC_API_KEY=sk-ant-...\n"
        "  OpenAI:     export OPENAI_API_KEY=sk-...\n"
        "(.env 파일에 넣어도 자동으로 읽습니다.)"
    )


class LLMClient:
    def __init__(self, provider: str, model: str | None = None, base_url: str | None = None):
        self.provider = provider
        self.model = model or DEFAULT_MODELS[provider]
        if provider == "anthropic":
            import anthropic
            if not os.environ.get("ANTHROPIC_API_KEY"):
                raise SystemExit("ANTHROPIC_API_KEY 가 설정되어 있지 않습니다.")
            self._client = anthropic.Anthropic()
        elif provider == "openai":
            try:
                from openai import OpenAI
            except ImportError:
                raise SystemExit("openai 패키지가 필요합니다:  pip install openai")
            if not os.environ.get("OPENAI_API_KEY"):
                raise SystemExit("OPENAI_API_KEY 가 설정되어 있지 않습니다.")
            kwargs = {}
            base = base_url or os.environ.get("OPENAI_BASE_URL")
            if base:
                kwargs["base_url"] = base
            self._client = OpenAI(**kwargs)
        else:
            raise SystemExit(f"알 수 없는 provider: {provider}")

    def chat(self, system: str, user: str, max_tokens: int) -> str:
        if self.provider == "anthropic":
            resp = self._client.messages.create(
                model=self.model,
                max_tokens=max_tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            )
            return "".join(b.text for b in resp.content if b.type == "text").strip()
        return self._openai_chat(system, user, max_tokens)

    def _openai_chat(self, system: str, user: str, max_tokens: int) -> str:
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]
        try:
            resp = self._client.chat.completions.create(
                model=self.model, max_tokens=max_tokens, messages=messages
            )
        except Exception as exc:  # noqa: BLE001
            # 일부 신형 모델은 max_tokens 대신 max_completion_tokens 를 요구한다.
            if "max_tokens" in str(exc) or "max_completion_tokens" in str(exc):
                resp = self._client.chat.completions.create(
                    model=self.model, max_completion_tokens=max_tokens, messages=messages
                )
            else:
                raise
        return (resp.choices[0].message.content or "").strip()
