from __future__ import annotations

import base64
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse

import requests
from openai import OpenAI
from PIL import Image, ImageOps

from processor import ProcessingConfig


class ImageTuneError(RuntimeError):
    pass


def tune_image(source: str, output_dir: Path, config: ProcessingConfig, stem: str) -> Path:
    """Edit one product image using an OpenAI-compatible image edit API."""

    if not config.image_llm.api_key:
        raise ImageTuneError("请先在设置页填写图片模型 API Key")

    output_dir.mkdir(parents=True, exist_ok=True)
    source_path = _materialize_source_image(source)
    target = output_dir / f"{_safe_stem(stem)}_{int(time.time() * 1000)}.jpg"

    client_kwargs = {"api_key": config.image_llm.api_key}
    if config.image_llm.base_url:
        client_kwargs["base_url"] = config.image_llm.base_url
    client = OpenAI(**client_kwargs)

    prompt = config.image_tune_prompt.strip()
    if not prompt:
        raise ImageTuneError("请先在设置页填写图片微调提示词模板")

    prompt = _with_output_requirements(prompt)

    try:
        with source_path.open("rb") as image_file:
            response = client.images.edit(
                model=config.image_llm.model or "gpt-image-1",
                image=image_file,
                prompt=prompt,
                n=1,
                size="1024x1024",
                output_format="jpeg",
            )
        data = response.data[0]
        if getattr(data, "b64_json", None):
            target.write_bytes(base64.b64decode(data.b64_json))
        elif getattr(data, "url", None):
            image_response = requests.get(data.url, timeout=60)
            image_response.raise_for_status()
            target.write_bytes(image_response.content)
        else:
            raise ImageTuneError("图片模型没有返回可保存的图片")
        _normalize_square_800(target)
    except Exception as exc:
        if isinstance(exc, ImageTuneError):
            raise
        raise ImageTuneError(f"图片微调失败：{exc}") from exc
    finally:
        if source_path.parent == Path(tempfile.gettempdir()) and source_path.exists():
            source_path.unlink(missing_ok=True)

    return target


def _with_output_requirements(prompt: str) -> str:
    return (
        prompt.rstrip()
        + "\n\nOutput requirements:\n"
        + "- Create a square 1:1 image.\n"
        + "- Final image size must be 800 x 800 pixels.\n"
        + "- Keep the full product visible within the square canvas.\n"
        + "- Do not crop out any part of the product.\n"
        + "- Use clean ecommerce-style framing suitable for a Temu carousel image.\n"
    )


def _normalize_square_800(path: Path) -> None:
    with Image.open(path) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        image.thumbnail((800, 800), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (800, 800), "white")
        left = (800 - image.width) // 2
        top = (800 - image.height) // 2
        canvas.paste(image, (left, top))
        canvas.save(path, "JPEG", quality=92, optimize=True)


def _materialize_source_image(source: str) -> Path:
    parsed = urlparse(source)
    if parsed.scheme in {"http", "https"}:
        suffix = Path(parsed.path).suffix.lower()
        suffix = suffix if suffix in {".jpg", ".jpeg", ".png", ".webp"} else ".jpg"
        temp_path = Path(tempfile.gettempdir()) / f"temu_source_{int(time.time() * 1000)}{suffix}"
        response = requests.get(source, timeout=30)
        response.raise_for_status()
        temp_path.write_bytes(response.content)
        return temp_path

    path = Path(source)
    if not path.exists():
        raise ImageTuneError(f"源图片不存在：{source}")
    return path


def _safe_stem(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in value)
    return cleaned[:80] or "tuned_image"
