from __future__ import annotations

import asyncio
import json
from dataclasses import dataclass
from typing import Any, Callable, Optional

from scraper_integration import ExternalScraperConfig, ExternalTemuListingsScraper
from url_converter import TemuUrlConverter, UrlConverterConfig


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
    external_scraper_dir: str = ""
    external_scraper_command: str = "npm run scrape -- --input {input} --output {output}"
    external_scraper_timeout_seconds: int = 180


class TemuScraper:
    """Resolve Yunqi links, then delegate product extraction to external TS scraper."""

    def __init__(self, config: ScraperConfig, log: Optional[LogCallback] = None) -> None:
        self.config = config
        self.log = log or (lambda message: None)
        self.url_converter: TemuUrlConverter | None = None
        self.external_scraper = ExternalTemuListingsScraper(
            ExternalScraperConfig(
                project_dir=config.external_scraper_dir,
                command=config.external_scraper_command,
                timeout_seconds=config.external_scraper_timeout_seconds,
            ),
            log=self.log,
        )

    async def __aenter__(self) -> "TemuScraper":
        await self.start()
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        await self.close()

    async def start(self) -> None:
        self.url_converter = TemuUrlConverter(
            UrlConverterConfig(
                max_concurrent=self.config.max_concurrent,
                headless=self.config.headless,
                proxy=self.config.proxy,
                timeout_ms=self.config.timeout_ms,
                retry_count=self.config.retry_count,
                min_delay_seconds=self.config.min_delay_seconds,
                max_delay_seconds=self.config.max_delay_seconds,
            ),
            log=self.log,
        )
        await self.url_converter.start()

    async def close(self) -> None:
        if self.url_converter:
            await self.url_converter.close()
            self.url_converter = None

    async def scrape_product(self, url: str) -> dict[str, Any]:
        results = await self.scrape_products([url])
        return results[0] if results else {"ok": False, "url": url, "error": "No result"}

    async def scrape_products(self, urls: list[str]) -> list[dict[str, Any]]:
        if not self.url_converter:
            await self.start()

        assert self.url_converter is not None
        resolved_results = await asyncio.gather(*(self.url_converter.resolve(url) for url in urls))
        real_urls = [item["url"] for item in resolved_results if item.get("ok") and item.get("url")]
        failed_by_input = {item["input"]: item for item in resolved_results if not item.get("ok")}

        scraped_by_url: dict[str, dict[str, Any]] = {}
        if real_urls:
            external_results = await self.external_scraper.scrape_urls(real_urls)
            scraped_by_url = {str(item.get("url") or item.get("real_detail_url") or ""): item for item in external_results}

        final_results: list[dict[str, Any]] = []
        for resolved in resolved_results:
            original = str(resolved.get("input") or "")
            real_url = str(resolved.get("url") or "")
            if original in failed_by_input:
                final_results.append({"ok": False, "url": original, "real_detail_url": "", "error": resolved.get("error", "")})
                continue

            item = scraped_by_url.get(real_url) or _find_result_by_goods_id(scraped_by_url.values(), real_url)
            if not item:
                final_results.append({"ok": False, "url": original, "real_detail_url": real_url, "error": "External scraper returned no data"})
                continue

            item = dict(item)
            item["url"] = real_url
            item["real_detail_url"] = real_url
            item["source_input"] = original
            final_results.append(item)
        return final_results


def _find_result_by_goods_id(items, url: str) -> dict[str, Any] | None:
    import re

    match = re.search(r"(\d{8,})", url)
    goods_id = match.group(1) if match else ""
    if not goods_id:
        return None
    for item in items:
        if goods_id in json.dumps(item, ensure_ascii=False):
            return item
    return None


def specs_to_json(specs: dict[str, Any]) -> str:
    return json.dumps(specs or {}, ensure_ascii=False, sort_keys=True)
