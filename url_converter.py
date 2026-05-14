from __future__ import annotations

import asyncio
import random
import re
from dataclasses import dataclass
from typing import Any, Callable, Optional
from urllib.parse import parse_qs, urlparse


LogCallback = Callable[[str], None]


@dataclass(frozen=True)
class UrlConverterConfig:
    max_concurrent: int = 3
    headless: bool = True
    proxy: str = ""
    timeout_ms: int = 30000
    retry_count: int = 2
    min_delay_seconds: float = 1.5
    max_delay_seconds: float = 4.0


class TemuUrlConverter:
    """Convert Yunqi/search/id inputs into real Temu product detail URLs."""

    def __init__(self, config: UrlConverterConfig, log: Optional[LogCallback] = None) -> None:
        self.config = config
        self.log = log or (lambda message: None)
        self.semaphore = asyncio.Semaphore(max(1, int(config.max_concurrent or 1)))
        self.playwright = None
        self.browser = None

    async def __aenter__(self) -> "TemuUrlConverter":
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

    async def resolve(self, value: str) -> dict[str, Any]:
        async with self.semaphore:
            last_error = ""
            for attempt in range(self.config.retry_count + 1):
                context = None
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
                    await apply_stealth(page)
                    detail_url = await self._resolve_with_page(page, value)
                    await context.close()
                    await asyncio.sleep(random.uniform(self.config.min_delay_seconds, self.config.max_delay_seconds))
                    return {"ok": True, "input": value, "url": detail_url, "search_key": search_key_from_input(value)}
                except Exception as exc:
                    last_error = str(exc)
                    self.log(f"Temu URL resolve failed attempt {attempt + 1}: {value} | {last_error}")
                    try:
                        if context:
                            await context.close()
                    except Exception:
                        pass
                    await asyncio.sleep(min(8, 1.5 * (attempt + 1)))
            return {"ok": False, "input": value, "url": "", "search_key": search_key_from_input(value), "error": last_error}

    async def _resolve_with_page(self, page, value: str) -> str:
        cleaned = str(value or "").strip()
        search_key = search_key_from_input(cleaned)

        if cleaned.startswith(("http://", "https://")) and not is_search_like_url(cleaned):
            await page.goto(cleaned, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
            if not is_search_like_url(page.url):
                return canonical_product_url(page.url)

        if not search_key:
            raise RuntimeError(f"No product id/search_key found: {value}")

        return await self._search_in_site(page, search_key)

    async def _search_in_site(self, page, keyword: str) -> str:
        await page.goto("https://www.temu.com/", wait_until="domcontentloaded", timeout=self.config.timeout_ms)
        await page.wait_for_selector("body", timeout=self.config.timeout_ms)
        await page.wait_for_timeout(random.randint(1200, 2500))

        search_input = page.locator(
            "input[type='search'], input[placeholder*='Search'], input[aria-label*='Search'], input[type='text']"
        ).first()
        try:
            await search_input.fill(keyword, timeout=6000)
            await page.wait_for_timeout(random.randint(300, 800))
            await search_input.press("Enter")
        except Exception:
            await page.evaluate(
                """(keyword) => {
                    const inputs = Array.from(document.querySelectorAll('input'));
                    const input = inputs.find((el) => {
                        const text = [el.type, el.placeholder, el.getAttribute('aria-label'), el.name]
                            .join(' ')
                            .toLowerCase();
                        return !el.disabled && (text.includes('search') || el.type === 'text' || el.type === 'search');
                    }) || inputs.find((el) => !el.disabled);
                    if (!input) return false;
                    input.focus();
                    input.value = keyword;
                    input.dispatchEvent(new Event('input', {bubbles: true}));
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                    return true;
                }""",
                keyword,
            )
            await page.keyboard.press("Enter")

        try:
            await page.wait_for_load_state("domcontentloaded", timeout=self.config.timeout_ms)
        except Exception:
            pass
        await page.wait_for_timeout(random.randint(3000, 5200))
        await soft_scroll(page)

        detail_url = await first_product_detail_url(page)
        if detail_url:
            return canonical_product_url(detail_url)

        clicked = await click_first_product_card(page)
        if clicked:
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=self.config.timeout_ms)
            except Exception:
                pass
            await page.wait_for_timeout(random.randint(1500, 2600))
            if not is_search_like_url(page.url):
                return canonical_product_url(page.url)

        raise RuntimeError(f"No product card found after Temu in-site search: {keyword}")


async def apply_stealth(page) -> None:
    try:
        from playwright_stealth import stealth_async

        await stealth_async(page)
    except Exception:
        try:
            await page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
        except Exception:
            pass


async def soft_scroll(page) -> None:
    for _ in range(2):
        await page.mouse.wheel(0, random.randint(450, 850))
        await page.wait_for_timeout(random.randint(450, 900))


async def first_product_detail_url(page) -> str:
    return await page.evaluate(
        """() => {
            const detailPattern = /(-g-\\d+|goods[_-]?id=\\d+|product[_-]?id=\\d+)/i;
            const badPattern = /(search_result|cart|login|support|orders|category|about|privacy)/i;
            const anchors = Array.from(document.querySelectorAll('a[href]'));
            for (const anchor of anchors) {
                const href = new URL(anchor.getAttribute('href'), location.href).href;
                const text = (anchor.innerText || anchor.getAttribute('aria-label') || '').trim();
                const imageCount = anchor.querySelectorAll('img').length;
                if (detailPattern.test(href) && !badPattern.test(href) && (imageCount || text.length > 5)) {
                    return href;
                }
            }
            return '';
        }"""
    )


async def click_first_product_card(page) -> bool:
    selectors = [
        "a[href*='-g-']",
        "a[href*='goods_id']",
        "a[href*='goodsId']",
        "a[href*='product_id']",
        "a:has(img)",
        "[role='listitem']",
        "[class*='goods']",
        "[class*='product']",
    ]
    for selector in selectors:
        try:
            locator = page.locator(selector).first()
            if await locator.count():
                await locator.click(timeout=5000)
                return True
        except Exception:
            continue
    return False


def is_search_like_url(url: str) -> bool:
    parsed = urlparse(str(url or ""))
    text = f"{parsed.path}?{parsed.query}".lower()
    return "search" in text or "search_key" in text or "keyword" in text


def search_key_from_input(value: str) -> str:
    text = str(value or "").strip()
    parsed = urlparse(text)
    query = parse_qs(parsed.query)
    for key in ("search_key", "searchKey", "keyword", "q", "query", "goods_id", "goodsId", "product_id", "productId"):
        values = query.get(key)
        if values and str(values[0]).strip():
            return str(values[0]).strip()
    if re.fullmatch(r"\d{8,}", text):
        return text
    match = re.search(r"(\d{8,})", text)
    return match.group(1) if match else ""


def canonical_product_url(url: str) -> str:
    goods_id = search_key_from_input(url)
    if goods_id:
        return f"https://www.temu.com/goods.html?_bg_fs=1&goods_id={goods_id}"
    return str(url or "").strip()
