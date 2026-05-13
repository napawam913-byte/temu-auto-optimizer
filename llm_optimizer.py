from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openai import OpenAI


TITLE_PROMPT_TEMPLATE = """You are a senior Temu US listing operator.
Rewrite the product title into a high-quality English marketplace title.

Rules:
1. Target length: about {target_length} characters. Acceptable range: {min_length}-{max_length} characters.
2. Put the main product keyword first.
3. Keep the product meaning accurate. Do not invent brand, certification, warranty, material, quantity, size, or function.
4. Make it natural for Temu/Amazon search.
5. Avoid promotional, medical, absolute, IP/brand, adult, and restricted claims.
6. Do not use emojis, decorative symbols, or explanation text.

Original Chinese title:
{source_title}

Reference English title, if any:
{reference_title}

Category:
{category}

Sensitive words to avoid:
{blocked_words}

Previous failed attempt and reason, if any:
{feedback}

Return only the rewritten English title."""

DESCRIPTION_PROMPT_TEMPLATE = """You are a senior Temu US listing copywriter.
Create an English product description for the item below.

Requirements:
1. Start with a concise value-focused paragraph.
2. Add 4-6 bullet points covering material, use case, benefits, compatibility, and gift/seasonal appeal when relevant.
3. Keep the tone natural, trustworthy, and marketplace-friendly.
4. Do not make unsupported claims about certifications, medical effects, origin, warranty, or brand.
5. Do not include emojis or special decorative symbols.

Original title: {title}
Optimized English title: {english_title}
Category: {category}

Return only the description."""

CHINESE_TITLE_PROMPT_TEMPLATE = """You are a Temu listing localization specialist.
Translate the optimized English title into a clean Chinese product title for the seller-side Excel template.

Rules:
1. Keep the meaning aligned with the English title.
2. Use concise natural Chinese suitable for an ecommerce product title.
3. Do not add unsupported brand, certification, medical, price, promotion, warranty, origin, or quantity claims.
4. Do not include emojis, decorative symbols, quotation marks, or explanation text.
5. Keep important product keywords near the front.

Optimized English title:
{english_title}

Original Chinese title for reference:
{source_title}

Category:
{category}

Return only the Chinese title."""


@dataclass(frozen=True)
class LLMConfig:
    provider: str = "openai"
    api_key: str = ""
    base_url: str = ""
    model: str = "gpt-4o-mini"
    title_prompt: str = TITLE_PROMPT_TEMPLATE
    description_prompt: str = DESCRIPTION_PROMPT_TEMPLATE


@dataclass(frozen=True)
class SensitiveWordRules:
    terms: tuple[str, ...] = ()
    allowlist: tuple[str, ...] = ()
    max_retry: int = 5
    allowed_delta: int = 5

    @classmethod
    def from_file(cls, path: Optional[Path]) -> "SensitiveWordRules":
        if not path or not path.exists():
            return cls()

        data = json.loads(path.read_text(encoding="utf-8"))
        categories = data.get("categories", {})
        terms: list[str] = []
        for values in categories.values():
            if isinstance(values, list):
                terms.extend(str(value).strip().lower() for value in values if str(value).strip())

        strategy = data.get("replacement_strategy", {})
        return cls(
            terms=tuple(dict.fromkeys(terms)),
            allowlist=tuple(str(value).strip().lower() for value in data.get("allowlist", []) if str(value).strip()),
            max_retry=max(1, int(strategy.get("max_retry", 5) or 5)),
            allowed_delta=max(0, int(strategy.get("allowed_title_length_delta", 5) or 5)),
        )

    def blocked_words_for_prompt(self) -> str:
        if not self.terms:
            return "No local sensitive word list configured."
        return ", ".join(self.terms[:120])

    def validate_title(self, title: str, target_length: int) -> list[str]:
        text = clean_title_text(title)
        issues: list[str] = []
        min_length = max(1, target_length - self.allowed_delta)
        max_length = max(min_length, target_length + self.allowed_delta)

        if len(text) < min_length:
            issues.append(f"too short: {len(text)} characters, need at least {min_length}")
        if len(text) > max_length:
            issues.append(f"too long: {len(text)} characters, need no more than {max_length}")

        blocked = self.find_blocked_words(text)
        if blocked:
            issues.append("contains sensitive words: " + ", ".join(blocked))
        return issues

    def find_blocked_words(self, title: str) -> list[str]:
        searchable = f" {title.lower()} "
        for allowed in self.allowlist:
            searchable = searchable.replace(allowed, " ")
        found: list[str] = []
        for term in self.terms:
            pattern = r"(?<![a-z0-9])" + re.escape(term) + r"(?![a-z0-9])"
            if re.search(pattern, searchable):
                found.append(term)
        return found

    def sanitize_title(self, title: str, target_length: int) -> str:
        text = clean_title_text(title)
        for term in self.find_blocked_words(text):
            text = re.sub(r"(?i)(?<![a-z0-9])" + re.escape(term) + r"(?![a-z0-9])", "", text)
        text = clean_title_text(text)
        return trim_title(text, target_length + self.allowed_delta)


class LLMOptimizer:
    """Small adapter around OpenAI-compatible chat completion APIs."""

    def __init__(self, config: LLMConfig) -> None:
        self.config = config
        self.client: Optional[OpenAI] = None
        if config.api_key:
            kwargs = {"api_key": config.api_key}
            if config.base_url:
                kwargs["base_url"] = config.base_url
            self.client = OpenAI(**kwargs)

    @property
    def enabled(self) -> bool:
        return self.client is not None

    def optimize_title(
        self,
        title: str,
        category: str = "",
        max_length: int = 100,
        source_title: str = "",
        sensitive_words_file: Optional[Path] = None,
    ) -> str:
        target_length = max(10, int(max_length or 100))
        rules = SensitiveWordRules.from_file(sensitive_words_file)
        fallback = self._fallback_title(source_title or title, target_length + rules.allowed_delta)

        if not self.enabled:
            return rules.sanitize_title(fallback, target_length)

        feedback = "None"
        best = ""
        for attempt in range(1, rules.max_retry + 1):
            prompt = TITLE_PROMPT_TEMPLATE.format(
                target_length=target_length,
                min_length=max(1, target_length - rules.allowed_delta),
                max_length=target_length + rules.allowed_delta,
                source_title=source_title or title or "",
                reference_title=title or "",
                category=category or "",
                blocked_words=rules.blocked_words_for_prompt(),
                feedback=feedback,
            )
            result = clean_title_text(self._chat(prompt, temperature=0.35))
            if result:
                best = result
            issues = rules.validate_title(result, target_length)
            if not issues:
                return result
            feedback = f"Attempt {attempt}: {result or '[empty]'} | " + "; ".join(issues)

        return rules.sanitize_title(best or fallback, target_length)

    def generate_description(self, title: str, english_title: str, category: str = "") -> str:
        if not self.enabled:
            return self._fallback_description(english_title or title)

        prompt = self.config.description_prompt.format(
            title=title or "",
            english_title=english_title or "",
            category=category or "",
        )
        return self._chat(prompt, temperature=0.55) or self._fallback_description(english_title or title)

    def translate_title_to_chinese(self, english_title: str, source_title: str = "", category: str = "") -> str:
        if not self.enabled:
            return clean_title_text(source_title or english_title)

        prompt = CHINESE_TITLE_PROMPT_TEMPLATE.format(
            english_title=english_title or "",
            source_title=source_title or "",
            category=category or "",
        )
        result = clean_title_text(self._chat(prompt, temperature=0.25))
        return result or clean_title_text(source_title or english_title)

    def _chat(self, prompt: str, temperature: float) -> str:
        assert self.client is not None
        response = self.client.chat.completions.create(
            model=self.config.model,
            messages=[
                {"role": "system", "content": "Return clean marketplace-ready copy only."},
                {"role": "user", "content": prompt},
            ],
            temperature=temperature,
        )
        return (response.choices[0].message.content or "").strip()

    @staticmethod
    def _fallback_title(title: str, max_length: int) -> str:
        text = clean_title_text(title or "Product")
        return trim_title(text, max_length)

    @staticmethod
    def _fallback_description(title: str) -> str:
        clean_title = clean_title_text(title or "Product")
        return (
            f"{clean_title} is designed for everyday use with a practical, easy-to-match style.\n\n"
            "- Suitable for daily home, office, travel, or gifting needs\n"
            "- Lightweight and convenient for regular use\n"
            "- Clean design helps the product fit different scenarios\n"
            "- A practical choice for customers looking for value and usability"
        )


def clean_title_text(title: str) -> str:
    text = str(title or "").strip()
    text = re.sub(r"^(optimized title|title|english title)\s*[:：-]\s*", "", text, flags=re.IGNORECASE)
    text = text.strip("\"'`“”‘’")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def trim_title(title: str, max_length: int) -> str:
    text = clean_title_text(title)
    if len(text) <= max_length:
        return text
    trimmed = text[:max_length].rsplit(" ", 1)[0].strip()
    return trimmed or text[:max_length].strip()
