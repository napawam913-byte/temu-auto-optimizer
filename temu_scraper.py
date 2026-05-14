from __future__ import annotations

import asyncio
import json
import random
import re
from dataclasses import dataclass
from typing import Any, Callable, Optional
from urllib.parse import parse_qs, urlparse


SCRAPER_HTML_PARSE_PROMPT = """You are a Temu product page data extraction assistant.
Extract only factual information from the provided HTML/text. Do not invent missing data.

Return strict JSON with these keys:
{
  "title": "English title if available",
  "description": "Detailed product description",
  "images": ["image url 1", "image url 2"],
  "specs": {"Color": "...", "Material": "...", "Size": "..."},
  "weight_g": null,
  "length_cm": null,
  "width_cm": null,
  "height_cm": null,
  "product_id": "",
  "sku": ""
}

Rules:
- Prefer English text when both English and Chinese are present.
- Keep image URLs exactly as found.
- Convert dimensions to centimeters and weight to grams when possible.
- If a value is missing or uncertain, return null or an empty string.
- Return JSON only, no explanation."""


LogCallback = Callable[[str], None]


@dataclass(frozen=True)
class ScraperConfig:
    max_concurrent: int = 3
    headless: bool = True
    proxy: str = ""
    timeout_ms: int = 30000
    retry_count: int = 2
    min_delay_seconds: float = 2.0
    max_delay_seconds: float = 6.0
    max_images: int = 12


class TemuScraper:
    def __init__(self, config: ScraperConfig, log: Optional[LogCallback] = None) -> None:
        self.config = config
        self.log = log or (lambda message: None)
        self.semaphore = asyncio.Semaphore(max(1, int(config.max_concurrent or 1)))
        self.playwright = None
        self.browser = None

    async def __aenter__(self) -> "TemuScraper":
        await self.start()
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        await self.close()

    async def start(self) -> None:
        try:
            from playwright.async_api import async_playwright
        except Exception as exc:
            raise RuntimeError(
                "Playwright is not installed. Run: pip install playwright playwright-stealth && playwright install chromium"
            ) from exc

        self.playwright = await async_playwright().start()
        launch_options: dict[str, Any] = {
            "headless": self.config.headless,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
            ],
        }
        if self.config.proxy:
            launch_options["proxy"] = {"server": self.config.proxy}
        self.browser = await self.playwright.chromium.launch(**launch_options)

    async def close(self) -> None:
        if self.browser:
            await self.browser.close()
            self.browser = None
        if self.playwright:
            await self.playwright.stop()
            self.playwright = None

    async def scrape_product(self, url: str) -> dict[str, Any]:
        async with self.semaphore:
            last_error = ""
            for attempt in range(self.config.retry_count + 1):
                page = None
                try:
                    if not self.browser:
                        await self.start()
                    context = await self.browser.new_context(
                        locale="en-US",
                        viewport={"width": 1365, "height": 900},
                        user_agent=(
                            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/124.0.0.0 Safari/537.36"
                        ),
                    )
                    page = await context.new_page()
                    await _apply_stealth(page)
                    await page.goto(url, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
                    await page.wait_for_timeout(int(random.uniform(1200, 2600)))
                    await _soft_scroll(page)
                    data = await self._extract_page_data(page, url)
                    await context.close()
                    await asyncio.sleep(random.uniform(self.config.min_delay_seconds, self.config.max_delay_seconds))
                    return data
                except Exception as exc:
                    last_error = str(exc)
                    self.log(f"Temu scrape failed attempt {attempt + 1}: {url} | {last_error}")
                    try:
                        if page:
                            await page.context.close()
                    except Exception:
                        pass
                    await asyncio.sleep(min(8, 1.5 * (attempt + 1)))
            return {"ok": False, "url": url, "error": last_error}

    async def _extract_page_data(self, page, url: str) -> dict[str, Any]:
        payload = await page.evaluate(
            """() => {
                const text = (el) => el ? (el.innerText || el.textContent || '').trim() : '';
                const attr = (selector, name) => {
                  const el = document.querySelector(selector);
                  return el ? (el.getAttribute(name) || '') : '';
                };
                const title =
                  text(document.querySelector('h1')) ||
                  attr('meta[property="og:title"]', 'content') ||
                  attr('meta[name="title"]', 'content') ||
                  document.title || '';
                const description =
                  attr('meta[name="description"]', 'content') ||
                  attr('meta[property="og:description"]', 'content') || '';
                const images = Array.from(document.images)
                  .map(img => img.currentSrc || img.src || img.getAttribute('data-src') || '')
                  .filter(Boolean);
                const bodyText = document.body ? document.body.innerText : '';
                const jsonScripts = Array.from(document.querySelectorAll('script[type="application/ld+json"], script'))
                  .map(script => script.textContent || '')
                  .filter(Boolean)
                  .slice(0, 20);
                const specBlocks = Array.from(document.querySelectorAll(
                  '[class*="spec"], [class*="attr"], [class*="detail"], table, dl, li'
                )).map(el => text(el)).filter(Boolean).slice(0, 120);
                return {title, description, images, bodyText, jsonScripts, specBlocks};
            }"""
        )

        title = _clean_text(payload.get("title", ""))
        description = _clean_text(payload.get("description", ""))
        body_text = _clean_text(payload.get("bodyText", ""))
        spec_text = "\n".join(payload.get("specBlocks", []))
        scripts = payload.get("jsonScripts", [])

        images = _normalise_images(payload.get("images", []), max_images=self.config.max_images)
        specs = _extract_specs(spec_text or body_text)
        dimensions = _extract_dimensions(f"{spec_text}\n{body_text}")
        product_id = _extract_product_id(url, scripts, body_text)
        sku = _extract_sku(scripts, body_text)

        if not description:
            description = _extract_description_from_text(body_text)

        return {
            "ok": True,
            "url": url,
            "title": title,
            "description": description,
            "images": images,
            "specs": specs,
            "weight_g": dimensions.get("weight_g"),
            "length_cm": dimensions.get("length_cm"),
            "width_cm": dimensions.get("width_cm"),
            "height_cm": dimensions.get("height_cm"),
            "product_id": product_id,
            "sku": sku,
        }


async def _apply_stealth(page) -> None:
    try:
        from playwright_stealth import stealth_async

        await stealth_async(page)
    except Exception:
        try:
            await page.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
            )
        except Exception:
            pass


async def _soft_scroll(page) -> None:
    for _ in range(3):
        await page.mouse.wheel(0, random.randint(500, 900))
        await page.wait_for_timeout(random.randint(500, 1000))


def _normalise_images(images: list[str], max_images: int) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for image in images:
        value = str(image or "").strip()
        if not value or value.startswith("data:"):
            continue
        if value.startswith("//"):
            value = "https:" + value
        if not value.startswith(("http://", "https://")):
            continue
        lowered = value.lower()
        if any(token in lowered for token in ("logo", "avatar", "sprite", "icon")):
            continue
        if value in seen:
            continue
        seen.add(value)
        result.append(value)
        if len(result) >= max(1, max_images):
            break
    return result


def _extract_specs(text: str) -> dict[str, str]:
    specs: dict[str, str] = {}
    pairs = re.findall(r"([A-Za-z\u4e00-\u9fff][A-Za-z0-9\u4e00-\u9fff /_-]{1,30})\s*[:：]\s*([^\n\r]{1,80})", text)
    for key, value in pairs:
        key = _clean_text(key)
        value = _clean_text(value)
        if key and value and key.lower() not in {"http", "https"}:
            specs[key] = value
    return specs


def _extract_dimensions(text: str) -> dict[str, float | None]:
    result: dict[str, float | None] = {"weight_g": None, "length_cm": None, "width_cm": None, "height_cm": None}
    weight_match = re.search(r"(\d+(?:\.\d+)?)\s*(kg|g|lb|lbs|oz)\b", text, flags=re.IGNORECASE)
    if weight_match:
        value = float(weight_match.group(1))
        unit = weight_match.group(2).lower()
        if unit == "kg":
            result["weight_g"] = round(value * 1000, 2)
        elif unit in {"lb", "lbs"}:
            result["weight_g"] = round(value * 453.592, 2)
        elif unit == "oz":
            result["weight_g"] = round(value * 28.3495, 2)
        else:
            result["weight_g"] = round(value, 2)

    dimension_match = re.search(
        r"(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*(cm|mm|in|inch|inches)?",
        text,
        flags=re.IGNORECASE,
    )
    if dimension_match:
        values = [float(dimension_match.group(i)) for i in range(1, 4)]
        unit = (dimension_match.group(4) or "cm").lower()
        if unit == "mm":
            values = [value / 10 for value in values]
        elif unit in {"in", "inch", "inches"}:
            values = [value * 2.54 for value in values]
        result["length_cm"], result["width_cm"], result["height_cm"] = [round(value, 2) for value in values]
    return result


def _extract_product_id(url: str, scripts: list[str], text: str) -> str:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    for key in ("goods_id", "goodsId", "product_id", "productId", "itemId", "id"):
        if query.get(key):
            return query[key][0]
    combined = "\n".join(scripts[:10]) + "\n" + text[:5000]
    match = re.search(r"(?:goodsId|productId|product_id|itemId|skuId)[\"'\s:=]+([A-Za-z0-9_-]{6,})", combined)
    return match.group(1) if match else ""


def _extract_sku(scripts: list[str], text: str) -> str:
    combined = "\n".join(scripts[:10]) + "\n" + text[:5000]
    match = re.search(r"(?:sku|skuId|sku_id)[\"'\s:=]+([A-Za-z0-9_-]{4,})", combined, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def _extract_description_from_text(text: str) -> str:
    lines = [_clean_text(line) for line in text.splitlines()]
    useful = [
        line
        for line in lines
        if 35 <= len(line) <= 260 and not re.search(r"(sign in|add to cart|shipping|review|coupon)", line, re.I)
    ]
    return "\n".join(useful[:6])


def _clean_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def specs_to_json(specs: dict[str, Any]) -> str:
    return json.dumps(specs or {}, ensure_ascii=False, sort_keys=True)
