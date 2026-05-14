from __future__ import annotations

import asyncio
import json
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

from attribute_filter import detect_no_attribute_product


LogCallback = Callable[[str], None]


@dataclass(frozen=True)
class ExternalScraperConfig:
    project_dir: str = ""
    command: str = "npm run scrape -- --input {input} --output {output}"
    timeout_seconds: int = 180


class ExternalTemuListingsScraper:
    """Subprocess adapter for aston-llrich/temu-listings-scraper.

    The upstream public README describes structured JSON output but does not
    document a stable importable API. A subprocess adapter keeps this project
    decoupled: install/update the TypeScript scraper separately, then configure
    the command template here.
    """

    def __init__(self, config: ExternalScraperConfig, log: Optional[LogCallback] = None) -> None:
        self.config = config
        self.log = log or (lambda message: None)

    async def scrape_urls(self, urls: list[str]) -> list[dict[str, Any]]:
        clean_urls = [str(url).strip() for url in urls if str(url or "").strip()]
        if not clean_urls:
            return []

        project_dir = Path(self.config.project_dir).expanduser() if self.config.project_dir else Path.cwd()
        if self.config.project_dir and not project_dir.exists():
            raise FileNotFoundError(f"External scraper directory not found: {project_dir}")

        with tempfile.TemporaryDirectory(prefix="temu_external_scraper_") as temp_dir:
            temp_path = Path(temp_dir)
            input_json = temp_path / "urls.json"
            input_txt = temp_path / "urls.txt"
            output_json = temp_path / "output.json"
            input_json.write_text(json.dumps(clean_urls, ensure_ascii=False, indent=2), encoding="utf-8")
            input_txt.write_text("\n".join(clean_urls), encoding="utf-8")

            command = self.config.command.format(
                input=str(input_json),
                input_txt=str(input_txt),
                output=str(output_json),
            )
            self.log(f"调用 temu-listings-scraper：{command}")
            process = await asyncio.create_subprocess_shell(
                command,
                cwd=str(project_dir),
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )
            try:
                stdout, stderr = await asyncio.wait_for(process.communicate(), timeout=self.config.timeout_seconds)
            except asyncio.TimeoutError as exc:
                process.kill()
                raise TimeoutError(f"External Temu scraper timed out after {self.config.timeout_seconds}s") from exc

            stdout_text = stdout.decode("utf-8", errors="replace").strip()
            stderr_text = stderr.decode("utf-8", errors="replace").strip()
            if process.returncode != 0:
                raise RuntimeError(f"External Temu scraper failed: {stderr_text or stdout_text}")

            raw_data = _read_scraper_json(output_json, stdout_text)
            by_url = {_normalise_product_url(item): item for item in raw_data if isinstance(item, dict)}

            results: list[dict[str, Any]] = []
            for url in clean_urls:
                item = by_url.get(url) or by_url.get(_strip_url(url)) or _match_by_goods_id(raw_data, url)
                if isinstance(item, dict):
                    results.append(normalise_external_item(item, fallback_url=url))
                else:
                    results.append({"ok": False, "url": url, "error": "No scraper result for URL"})
            return results


def normalise_external_item(item: dict[str, Any], fallback_url: str = "") -> dict[str, Any]:
    attribute_info = detect_no_attribute_product(item)
    images = _list_value(item.get("additionalImages"))
    if item.get("imageUrl"):
        images = [str(item["imageUrl"]), *[image for image in images if image != str(item["imageUrl"])]]

    specs = _dict_value(item.get("specs") or item.get("specifications") or item.get("attributes") or item.get("properties"))
    product_url = str(item.get("productUrl") or item.get("url") or fallback_url or "").strip()
    product_id = str(item.get("id") or item.get("goods_id") or item.get("goodsId") or "").strip()

    return {
        "ok": True,
        "url": product_url,
        "real_detail_url": product_url,
        "title": str(item.get("title") or "").strip(),
        "description": str(item.get("description") or "").strip(),
        "images": images,
        "specs": specs,
        "price": item.get("price"),
        "original_price": item.get("originalPrice"),
        "currency": item.get("currency"),
        "rating": item.get("rating"),
        "sales_count": item.get("salesCount"),
        "product_id": product_id,
        "sku": str(item.get("sku") or item.get("skuId") or "").strip(),
        **attribute_info,
    }


def _read_scraper_json(output_file: Path, stdout_text: str) -> list[Any]:
    if output_file.exists() and output_file.stat().st_size > 0:
        data = json.loads(output_file.read_text(encoding="utf-8"))
    else:
        data = json.loads(stdout_text)
    if isinstance(data, dict):
        for key in ("items", "products", "data", "results"):
            if isinstance(data.get(key), list):
                return data[key]
        return [data]
    if isinstance(data, list):
        return data
    return []


def _normalise_product_url(item: dict[str, Any]) -> str:
    return _strip_url(str(item.get("productUrl") or item.get("url") or ""))


def _strip_url(url: str) -> str:
    return str(url or "").strip().rstrip("/")


def _match_by_goods_id(items: list[Any], url: str) -> dict[str, Any] | None:
    goods_id = _goods_id(url)
    if not goods_id:
        return None
    for item in items:
        if not isinstance(item, dict):
            continue
        values = [item.get("id"), item.get("goods_id"), item.get("goodsId"), item.get("productUrl"), item.get("url")]
        if any(goods_id in str(value or "") for value in values):
            return item
    return None


def _goods_id(url: str) -> str:
    import re

    match = re.search(r"(\d{8,})", str(url or ""))
    return match.group(1) if match else ""


def _list_value(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item or "").strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []


def _dict_value(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}
