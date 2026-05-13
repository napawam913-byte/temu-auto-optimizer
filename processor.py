from __future__ import annotations

import json
import math
import re
import shutil
import time
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from html import escape
from io import BytesIO
from pathlib import Path
from typing import Callable, Iterable, Optional
from urllib.parse import urlparse

import pandas as pd
import requests
from openpyxl import load_workbook
from PIL import Image, ImageEnhance

from llm_optimizer import LLMConfig, LLMOptimizer


DXM_COLUMNS = [
    "*产品标题",
    "*英文标题",
    "产品描述",
    "产品货号",
    "*变种属性名称一",
    "*变种属性值一",
    "变种属性名称二",
    "变种属性值二",
    "预览图",
    "*申报价格\n(店铺币种)",
    "SKU货号",
    "*长（cm）",
    "*宽（cm）",
    "*高（cm）",
    "*重量（g）",
    "识别码类型",
    "识别码",
    "站外产品链接",
    "*轮播图",
    "*产品素材图",
    "外包装形状",
    "外包装类型",
    "外包装图片",
    "建议售价（USD）",
    "库存",
    "发货时效（天）",
]

CLOTHING_KEYWORDS = ("服装", "女装", "男装", "童装", "鞋", "靴", "尺码", "clothing", "shoes", "boots")

DEFAULT_IMAGE_TUNE_PROMPT = """Keep the original product exactly the same. Do not change the product shape, color, material, pattern, logo, text, quantity, packaging, or visible details.

Only make subtle ecommerce listing improvements:
1. Clean and brighten the background to pure white or near-white.
2. Improve lighting and contrast naturally.
3. Remove minor shadows, noise, compression artifacts, and background clutter.
4. Keep the product centered and fully visible.
5. Preserve realistic texture and edges.
6. Do not add props, text, labels, watermarks, badges, hands, models, or extra objects.

Output a clean marketplace-ready product image suitable for Temu listing."""


@dataclass
class ProcessingConfig:
    llm: LLMConfig
    image_llm: LLMConfig = field(default_factory=lambda: LLMConfig(provider="openai", model="gpt-image-1"))
    template_file: str = ""
    split_output_directory: str = ""
    output_directory: str = ""
    output_filename: str = ""
    title_max_length: int = 100
    filter_clothing: bool = True
    deduplicate: bool = True
    dedupe_field: str = "title"
    download_images: bool = False
    image_tune_count: int = 2
    image_tune_prompt: str = DEFAULT_IMAGE_TUNE_PROMPT
    description_image_count: int = 4
    default_price: float = 9.99
    default_stock: int = 100
    default_ship_days: str = "2"
    default_length_cm: float = 10
    default_width_cm: float = 10
    default_height_cm: float = 5
    default_weight_g: float = 200
    image_max_px: int = 1000
    enhance_white_bg: bool = False
    request_timeout: int = 20
    retry_count: int = 2
    default_variant_name: str = "颜色"
    default_variant_value: str = "如图"
    default_package_shape: str = "不规则"
    default_package_type: str = "硬包装"


@dataclass
class ProcessingResult:
    output_file: Path
    output_dir: Path
    success_count: int
    failure_count: int
    skipped_count: int
    failures: list[str]


ProgressCallback = Callable[[int, int, str], None]
LogCallback = Callable[[str], None]


def load_config(path: Path) -> ProcessingConfig:
    if not path.exists():
        return ProcessingConfig(llm=LLMConfig())
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    llm_data = data.pop("llm", {})
    image_llm_data = data.pop("image_llm", {})
    valid_fields = set(ProcessingConfig.__dataclass_fields__)
    clean_data = {key: value for key, value in data.items() if key in valid_fields}
    return ProcessingConfig(
        llm=LLMConfig(**llm_data),
        image_llm=LLMConfig(**image_llm_data) if image_llm_data else LLMConfig(provider="openai", model="gpt-image-1"),
        **clean_data,
    )


def save_config(config: ProcessingConfig, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(config), ensure_ascii=False, indent=2), encoding="utf-8")


def split_excel(
    input_file: Path,
    rows_per_file: int,
    output_dir: Path,
    progress: Optional[ProgressCallback] = None,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    split_dir = output_dir / f"split_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    split_dir.mkdir(parents=True, exist_ok=True)

    df = read_source_table(input_file)
    total = len(df)
    if total == 0:
        raise ValueError("Excel 文件没有可拆分的数据行")
    if rows_per_file <= 0:
        raise ValueError("每个文件行数必须大于 0")

    batch_count = math.ceil(total / rows_per_file)
    generated: list[Path] = []
    for index in range(batch_count):
        start = index * rows_per_file
        end = min(start + rows_per_file, total)
        batch = df.iloc[start:end]
        file_name = f"第{index + 1:02d}批_{start + 1}-{end}行.xlsx"
        target = split_dir / file_name
        batch.to_excel(target, index=False)
        generated.append(target)
        if progress:
            progress(index + 1, batch_count, f"已生成 {file_name}")

    zip_path = split_dir / "split_files.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in generated:
            archive.write(file_path, arcname=file_path.name)
    if progress:
        progress(batch_count, batch_count, f"拆分完成：{zip_path}")
    return zip_path


def read_source_table(file_path: Path) -> pd.DataFrame:
    """Read Excel, real CSV, and cloud exports whose bytes are xlsx but extension is .csv."""

    suffix = file_path.suffix.lower()
    if _is_zip_excel(file_path):
        frame = pd.read_excel(BytesIO(file_path.read_bytes()), header=None, engine="openpyxl")
        return _promote_detected_header(frame)

    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}:
        frame = pd.read_excel(file_path, header=None)
        return _promote_detected_header(frame)

    for encoding in ("utf-8-sig", "gb18030", "utf-16", "latin1"):
        try:
            frame = pd.read_csv(file_path, encoding=encoding, header=None)
            return _promote_detected_header(frame)
        except Exception:
            continue
    raise ValueError(f"无法读取文件：{file_path}")


class TemuProcessor:
    def __init__(
        self,
        config: ProcessingConfig,
        progress: Optional[ProgressCallback] = None,
        log: Optional[LogCallback] = None,
    ) -> None:
        self.config = config
        self.optimizer = LLMOptimizer(config.llm)
        self.progress = progress
        self.log = log or (lambda message: None)
        self.seen: set[str] = set()

    def process_files(self, files: Iterable[Path], output_root: Path) -> ProcessingResult:
        output_dir = output_root / f"temu_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        image_dir = output_dir / "images"
        output_dir.mkdir(parents=True, exist_ok=True)
        image_dir.mkdir(parents=True, exist_ok=True)

        rows = self._load_rows(files)
        total = len(rows)
        output_rows: list[dict[str, object]] = []
        failures: list[str] = []
        skipped = 0

        for index, row in enumerate(rows, start=1):
            source_title = self._first(row, ["商品标题（中文）", "商品标题", "产品标题", "标题", "*产品标题", "title", "Title"])
            sku = self._first(row, ["SKU货号", "SKU", "sku", "货号"])
            product_id = self._first(row, ["商品ID", "产品货号", "product_id"])
            category = self._first(row, ["前台分类（英文）", "前台分类（中文）", "后台分类", "类目", "分类", "category", "Category"])
            self._emit(index, total, f"处理中 {index}/{total}: {source_title or sku or '未命名商品'}")

            try:
                if self.config.filter_clothing and self._is_clothing(source_title, category):
                    skipped += 1
                    self.log(f"跳过需尺码表类目：{source_title}")
                    continue

                dedupe_key = self._dedupe_key(source_title, sku)
                if self.config.deduplicate and dedupe_key in self.seen:
                    skipped += 1
                    self.log(f"跳过去重商品：{dedupe_key}")
                    continue
                self.seen.add(dedupe_key)

                existing_english_title = self._first(row, ["商品标题（英文）", "*英文标题", "英文标题"])
                title_seed = existing_english_title or source_title
                english_title = self.optimizer.optimize_title(
                    title_seed,
                    category,
                    self.config.title_max_length,
                    source_title=source_title,
                    sensitive_words_file=self._sensitive_words_file(),
                )
                chinese_title = self.optimizer.translate_title_to_chinese(english_title, source_title, category)
                description = self.optimizer.generate_description(source_title, english_title, category)
                output_sku = self._listing_sku(sku, product_id, english_title, index)
                image_values = self._process_images(row, image_dir, output_sku or f"item_{index}")
                description = self._append_description_images(description, image_values)

                output_rows.append(self._build_output_row(row, chinese_title, english_title, description, output_sku, image_values))
            except Exception as exc:
                failures.append(f"第 {index} 行失败：{exc}")
                self.log(f"失败：第 {index} 行，{exc}")

        output_file = output_dir / self._output_filename()
        self._write_output(output_rows, output_file)
        return ProcessingResult(output_file, output_dir, len(output_rows), len(failures), skipped, failures)

    def _load_rows(self, files: Iterable[Path]) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for file_path in files:
            self.log(f"读取文件：{file_path}")
            frame = read_source_table(file_path)
            rows.extend(frame.fillna("").to_dict("records"))
        return rows

    def _write_output(self, rows: list[dict[str, object]], output_file: Path) -> None:
        template = Path(self.config.template_file) if self.config.template_file else None
        if template and template.exists():
            workbook = load_workbook(template)
            worksheet = workbook["导入模板"] if "导入模板" in workbook.sheetnames else workbook.active
            headers = [worksheet.cell(1, column).value for column in range(1, worksheet.max_column + 1)]
            for row_index in range(2, worksheet.max_row + 1):
                for column_index in range(1, worksheet.max_column + 1):
                    worksheet.cell(row_index, column_index).value = None
            for excel_row, item in enumerate(rows, start=2):
                for column_index, header in enumerate(headers, start=1):
                    if header:
                        worksheet.cell(excel_row, column_index).value = item.get(str(header), "")
            workbook.save(output_file)
            return

        pd.DataFrame(rows, columns=DXM_COLUMNS).to_excel(output_file, index=False)

    def _build_output_row(
        self,
        row: dict[str, object],
        source_title: str,
        english_title: str,
        description: str,
        sku: str,
        image_values: dict[str, str],
    ) -> dict[str, object]:
        product_id = self._first(row, ["商品ID", "产品货号", "product_id"])
        product_link = self._first(row, ["商品链接", "站外产品链接", "product_url"])
        price = self._first(row, ["*申报价格\n(店铺币种)", "美元价格($)", "价格", "售价", "price"]) or self.config.default_price
        suggestion_price = self._first(row, ["建议售价（USD）", "美元价格($)", "price"]) or price
        preview = image_values.get("preview", self._first(row, ["预览图", "商品主图", "preview_image"]))
        carousel = image_values.get("carousel", self._first(row, ["*轮播图", "商品轮播图", "轮播图"]))
        material = image_values.get("materials", self._first(row, ["*产品素材图", "商品主图", "素材图", "主图"]))
        package_image = image_values.get("package", self._first(row, ["外包装图片", "包装图"]))

        return {
            "*产品标题": source_title,
            "*英文标题": english_title,
            "产品描述": description,
            "产品货号": product_id,
            "*变种属性名称一": self._variant_name_one(row),
            "*变种属性值一": self._variant_value_one(row),
            "变种属性名称二": self._first(row, ["变种属性名称二"]),
            "变种属性值二": self._first(row, ["变种属性值二"]),
            "预览图": preview,
            "*申报价格\n(店铺币种)": price,
            "SKU货号": sku or product_id or self._make_sku(english_title),
            "*长（cm）": self._first(row, ["*长（cm）", "长(cm)", "长", "length"]) or self.config.default_length_cm,
            "*宽（cm）": self._first(row, ["*宽（cm）", "宽(cm)", "宽", "width"]) or self.config.default_width_cm,
            "*高（cm）": self._first(row, ["*高（cm）", "高(cm)", "高", "height"]) or self.config.default_height_cm,
            "*重量（g）": self._first(row, ["*重量（g）", "重量(g)", "重量", "weight"]) or self.config.default_weight_g,
            "识别码类型": self._first(row, ["识别码类型"]),
            "识别码": self._first(row, ["识别码"]),
            "站外产品链接": product_link,
            "*轮播图": carousel or preview,
            "*产品素材图": material or preview,
            "外包装形状": self._first(row, ["外包装形状"]) or self.config.default_package_shape,
            "外包装类型": self._first(row, ["外包装类型"]) or self.config.default_package_type,
            "外包装图片": package_image,
            "建议售价（USD）": suggestion_price,
            "库存": self._first(row, ["库存", "stock"]) or self.config.default_stock,
            "发货时效（天）": self._first(row, ["发货时效（天）", "发货时效"]) or self.config.default_ship_days,
        }

    def _append_description_images(self, description: str, image_values: dict[str, str]) -> str:
        urls = self._description_image_urls(image_values)
        if not urls:
            return description

        image_html = "\n".join(f'<p><img src="{escape(url, quote=True)}" /></p>' for url in urls)
        clean_description = str(description or "").strip()
        if not clean_description:
            return image_html
        return f"{clean_description}\n\n{image_html}"

    def _description_image_urls(self, image_values: dict[str, str]) -> list[str]:
        count = max(0, int(self.config.description_image_count or 0))
        if count == 0:
            return []

        selected: list[str] = []
        seen: set[str] = set()
        for key in ("preview", "carousel", "materials", "package"):
            for value in self._split_image_values(image_values.get(key, "")):
                if value in seen:
                    continue
                selected.append(value)
                seen.add(value)
                if len(selected) >= count:
                    return selected
        return selected

    def _process_images(self, row: dict[str, object], image_dir: Path, sku: str) -> dict[str, str]:
        image_sources = {
            "preview": self._first(row, ["预览图", "商品主图", "preview_image", "图片", "image"]),
            "carousel": self._first(row, ["*轮播图", "商品轮播图", "轮播图"]),
            "materials": self._first(row, ["*产品素材图", "商品主图", "素材图", "主图", "main_image"]),
            "package": self._first(row, ["外包装图片", "包装图", "package_image"]),
        }
        result: dict[str, str] = {}
        for key, value in image_sources.items():
            urls = self._split_image_values(value)
            if not self.config.download_images:
                if urls:
                    result[key] = "\n".join(urls)
                continue
            local_paths = []
            for image_index, url in enumerate(urls, start=1):
                path = self._download_and_optimize_image(url, image_dir, f"{sku}_{key}_{image_index}")
                if path:
                    local_paths.append(str(path))
            if local_paths:
                result[key] = "\n".join(local_paths)
        return result

    def _download_and_optimize_image(self, value: str, image_dir: Path, stem: str) -> str:
        if not value:
            return ""
        parsed = urlparse(value)
        suffix = Path(parsed.path).suffix.lower() or ".jpg"
        suffix = suffix if suffix in {".jpg", ".jpeg", ".png", ".webp"} else ".jpg"
        raw_path = image_dir / f"{self._safe_name(stem)}{suffix}"

        for attempt in range(self.config.retry_count + 1):
            try:
                if parsed.scheme in {"http", "https"}:
                    response = requests.get(value, timeout=self.config.request_timeout)
                    response.raise_for_status()
                    raw_path.write_bytes(response.content)
                else:
                    source = Path(value)
                    if source.exists():
                        shutil.copyfile(source, raw_path)
                    else:
                        return value
                return self._optimize_image(raw_path)
            except Exception as exc:
                if attempt >= self.config.retry_count:
                    self.log(f"图片处理失败：{value}，{exc}")
                else:
                    time.sleep(0.5 * (attempt + 1))
        return value

    def _optimize_image(self, path: Path) -> str:
        with Image.open(path) as image:
            image = image.convert("RGB")
            image.thumbnail((self.config.image_max_px, self.config.image_max_px))
            if self.config.enhance_white_bg:
                image = ImageEnhance.Brightness(image).enhance(1.04)
                image = ImageEnhance.Contrast(image).enhance(1.06)
            target = path.with_suffix(".jpg")
            image.save(target, "JPEG", quality=88, optimize=True)
        if target != path and path.exists():
            path.unlink(missing_ok=True)
        return str(target)

    def _emit(self, current: int, total: int, message: str) -> None:
        self.log(message)
        if self.progress:
            self.progress(current, total, message)

    @staticmethod
    def _sensitive_words_file() -> Path:
        return Path(__file__).resolve().parent / "config" / "sensitive_words.json"

    def _dedupe_key(self, title: str, sku: str) -> str:
        value = sku if self.config.dedupe_field.lower() == "sku" else title
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    def _variant_name_one(self, row: dict[str, object]) -> str:
        return (
            self._first(row, ["*变种属性名称一", "变种属性名称一", "属性名称一", "规格名称", "规格名", "变体名称", "变种名称"])
            or self.config.default_variant_name
            or "颜色"
        )

    def _variant_value_one(self, row: dict[str, object]) -> str:
        value = self._first(
            row,
            [
                "*变种属性值一",
                "变种属性值一",
                "属性值一",
                "颜色",
                "颜色分类",
                "规格",
                "规格值",
                "款式",
                "型号",
                "变体值",
                "变种值",
            ],
        )
        if value:
            return value
        configured = str(self.config.default_variant_value or "").strip()
        if not configured or configured.lower() == "default":
            return "如图"
        return configured

    def _listing_sku(self, source_sku: str, product_id: str, english_title: str, row_index: int) -> str:
        if source_sku:
            return source_sku
        digits = re.sub(r"\D+", "", str(product_id or ""))
        if digits:
            return f"TAO-{digits[-6:]}-{row_index:03d}"
        return f"{self._make_sku(english_title)}-{row_index:03d}"

    @staticmethod
    def _first(row: dict[str, object], names: list[str]) -> str:
        for name in names:
            value = row.get(name)
            if value is not None and str(value).strip():
                return str(value).strip()
        return ""

    @staticmethod
    def _split_image_values(value: str) -> list[str]:
        if not value:
            return []
        text = str(value).strip().strip("[]")
        return [part.strip().strip("'\"") for part in re.split(r"[;,\n]", text) if part.strip()]

    @staticmethod
    def _is_clothing(title: str, category: str) -> bool:
        text = f"{title} {category}".lower()
        return any(keyword.lower() in text for keyword in CLOTHING_KEYWORDS)

    @staticmethod
    def _make_sku(title: str) -> str:
        base = re.sub(r"[^A-Za-z0-9]+", "-", title.upper()).strip("-")
        return base[:32] or f"TEMU-{int(time.time())}"

    @staticmethod
    def _safe_name(value: str) -> str:
        return re.sub(r"[^A-Za-z0-9_-]+", "_", value)[:80] or "image"

    def _output_filename(self) -> str:
        raw = (self.config.output_filename or "").strip()
        if not raw:
            return "店小秘_TEMU半托管导入模板.xlsx"
        safe = re.sub(r'[<>:"/\\|?*]+', "_", raw).strip().strip(".")
        if not safe:
            return "店小秘_TEMU半托管导入模板.xlsx"
        if not safe.lower().endswith(".xlsx"):
            safe += ".xlsx"
        return safe


def _is_zip_excel(file_path: Path) -> bool:
    with file_path.open("rb") as handle:
        return handle.read(2) == b"PK"


def _promote_detected_header(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.dropna(how="all").reset_index(drop=True)
    header_index = 0
    markers = {"商品标题（中文）", "商品标题（英文）", "商品ID", "*产品标题", "*英文标题"}
    for index, row in frame.head(10).iterrows():
        values = {str(value).strip() for value in row.tolist() if pd.notna(value) and str(value).strip()}
        if values & markers:
            header_index = index
            break
    header = [
        str(value).strip() if pd.notna(value) and str(value).strip() else f"未命名列{idx + 1}"
        for idx, value in enumerate(frame.iloc[header_index].tolist())
    ]
    data = frame.iloc[header_index + 1 :].copy()
    data.columns = header
    return data.dropna(how="all").reset_index(drop=True)
