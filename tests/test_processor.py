from pathlib import Path

import pandas as pd
from pandas import isna

from llm_optimizer import LLMConfig
from processor import ProcessingConfig, TemuProcessor, read_source_table, split_excel


def test_split_excel_creates_batches_and_zip(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    pd.DataFrame({"标题": [f"商品{i}" for i in range(5)]}).to_excel(source, index=False)

    zip_path = split_excel(source, rows_per_file=2, output_dir=tmp_path)

    assert zip_path.exists()
    assert zip_path.name == "split_files.zip"
    assert len(list(zip_path.parent.glob("第*.xlsx"))) == 3


def test_processor_outputs_dxm_columns_without_llm(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {"商品标题": "Stainless steel cup", "SKU": "SKU001", "价格": 12.5, "类目": "Kitchen"},
            {"商品标题": "Stainless steel cup", "SKU": "SKU001", "价格": 12.5, "类目": "Kitchen"},
        ]
    ).to_excel(source, index=False)
    config = ProcessingConfig(llm=LLMConfig(api_key=""), deduplicate=True, dedupe_field="sku")

    result = TemuProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert result.success_count == 1
    assert result.skipped_count == 1
    assert "*英文标题" in frame.columns
    assert "产品描述" in frame.columns


def test_processor_writes_chinese_title_translated_from_generated_english_title(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {"商品标题": "复古恐龙绒布毯", "SKU": "SKU002", "价格": 8.8, "类目": "Home Textiles"},
        ]
    ).to_excel(source, index=False)

    class FakeOptimizer:
        def optimize_title(self, title, category="", max_length=100, source_title="", sensitive_words_file=None):
            return "Dinosaur Fleece Blanket Cozy Plush Throw For Sofa And Bed"

        def translate_title_to_chinese(self, english_title, source_title="", category=""):
            return "恐龙法兰绒毯柔软沙发床盖毯"

        def generate_description(self, title, english_title, category=""):
            return "description"

    config = ProcessingConfig(llm=LLMConfig(api_key=""), deduplicate=False)
    processor = TemuProcessor(config)
    processor.optimizer = FakeOptimizer()

    result = processor.process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert frame.loc[0, "*产品标题"] == "恐龙法兰绒毯柔软沙发床盖毯"
    assert frame.loc[0, "*英文标题"] == "Dinosaur Fleece Blanket Cozy Plush Throw For Sofa And Bed"


def test_processor_uses_listing_friendly_variant_value_when_source_has_no_variant(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {"商品标题": "Pink Bed Sheet", "商品ID": "605847883203024", "价格": 4.25},
        ]
    ).to_excel(source, index=False)
    config = ProcessingConfig(llm=LLMConfig(api_key=""), default_variant_value="Default", deduplicate=False)

    result = TemuProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert frame.loc[0, "*变种属性名称一"] == "颜色"
    assert frame.loc[0, "*变种属性值一"] == "如图"
    assert isna(frame.loc[0, "SKU货号"])
    assert isna(frame.loc[0, "产品货号"])


def test_processor_prefers_source_variant_value_when_present(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {"商品标题": "Pink Bed Sheet", "颜色": "粉色", "商品ID": "605847883203024", "价格": 4.25},
        ]
    ).to_excel(source, index=False)
    config = ProcessingConfig(llm=LLMConfig(api_key=""), deduplicate=False)

    result = TemuProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert frame.loc[0, "*变种属性值一"] == "粉色"


def test_processor_leaves_product_and_sku_codes_empty_when_source_sku_is_present(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {"商品标题": "Pink Bed Sheet", "SKU": "MY-SKU-001", "商品ID": "605847883203024", "价格": 4.25},
        ]
    ).to_excel(source, index=False)
    config = ProcessingConfig(llm=LLMConfig(api_key=""), deduplicate=False)

    result = TemuProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert isna(frame.loc[0, "SKU货号"])
    assert isna(frame.loc[0, "产品货号"])


def test_processor_merges_scraped_product_data_before_output(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {
                "商品标题": "Fallback title",
                "商品链接": "https://www.temu.com/product.html?goods_id=123456789",
                "商品轮播图": "https://source.example.com/a.jpg\nhttps://source.example.com/b.jpg",
            },
        ]
    ).to_excel(source, index=False)

    class FakeProcessor(TemuProcessor):
        def _run_scraper_enrichment(self, frame):
            enriched = frame.copy()
            self._merge_scraped_data(
                enriched,
                0,
                {
                    "title": "Scraped English Product Title",
                    "description": "Scraped detailed product description.",
                    "images": ["https://img.example.com/1.jpg", "https://img.example.com/2.jpg"],
                    "specs": {"Color": "Blue", "Size": "Queen", "Material": "Polyester"},
                    "weight_g": 250,
                    "length_cm": 20,
                    "width_cm": 10,
                    "height_cm": 5,
                    "product_id": "123456789",
                    "sku": "SCRAPED-SKU",
                },
            )
            return enriched

    config = ProcessingConfig(llm=LLMConfig(api_key=""), enable_scraper=True, deduplicate=False)
    result = FakeProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    assert frame.loc[0, "*英文标题"] == "Scraped English Product Title"
    assert "Scraped detailed product description." not in frame.loc[0, "产品描述"]
    assert frame.loc[0, "产品描述"].count("<img src=") == 2
    assert "https://source.example.com/a.jpg" in frame.loc[0, "*轮播图"]
    assert "https://img.example.com/1.jpg" not in frame.loc[0, "*轮播图"]
    assert frame.loc[0, "*变种属性名称一"] == "颜色"
    assert frame.loc[0, "*变种属性值一"] == "Blue"
    assert frame.loc[0, "变种属性名称二"] == "尺寸"
    assert frame.loc[0, "变种属性值二"] == "Queen"
    assert isna(frame.loc[0, "SKU货号"])
    assert isna(frame.loc[0, "产品货号"])
    assert frame.loc[0, "*重量（g）"] == 250


def test_processor_appends_four_unique_images_to_description(tmp_path: Path) -> None:
    source = tmp_path / "batch.xlsx"
    pd.DataFrame(
        [
            {
                "商品标题": "Pink Bed Sheet",
                "商品ID": "605847883203024",
                "商品主图": "https://img.example.com/main.jpg",
                "商品轮播图": "\n".join(
                    [
                        "https://img.example.com/main.jpg",
                        "https://img.example.com/1.jpg",
                        "https://img.example.com/2.jpg",
                        "https://img.example.com/3.jpg",
                        "https://img.example.com/4.jpg",
                    ]
                ),
            },
        ]
    ).to_excel(source, index=False)
    config = ProcessingConfig(llm=LLMConfig(api_key=""), deduplicate=False, description_image_count=4)

    result = TemuProcessor(config).process_files([source], tmp_path)

    frame = pd.read_excel(result.output_file)
    description = frame.loc[0, "产品描述"]
    assert description.count("<img src=") == 4
    assert "https://img.example.com/main.jpg" in description
    assert "https://img.example.com/3.jpg" in description
    assert "https://img.example.com/4.jpg" not in description


def test_read_source_table_handles_cloud_export_with_csv_extension() -> None:
    source = Path(r"D:\Downloads\2054441176885702657.csv")
    if not source.exists():
        return

    frame = read_source_table(source)

    assert "商品标题（中文）" in frame.columns
    assert "商品标题（英文）" in frame.columns
    assert "商品轮播图" in frame.columns


def test_processor_writes_user_template_columns(tmp_path: Path) -> None:
    source = Path(r"D:\Downloads\2054441176885702657.csv")
    template = Path(r"D:\Downloads\import_created_product_popTemu.xlsx")
    if not source.exists() or not template.exists():
        return
    config = ProcessingConfig(
        llm=LLMConfig(api_key=""),
        template_file=str(template),
        deduplicate=True,
        dedupe_field="sku",
    )

    result = TemuProcessor(config).process_files([source], tmp_path)
    frame = pd.read_excel(result.output_file, sheet_name="导入模板")

    assert "*产品标题" in frame.columns
    assert "*申报价格\n(店铺币种)" in frame.columns
    assert "*轮播图" in frame.columns
    assert result.success_count > 0
