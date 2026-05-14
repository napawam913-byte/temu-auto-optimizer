from __future__ import annotations

from typing import Any


ATTRIBUTE_KEYWORDS = (
    "color",
    "colour",
    "size",
    "style",
    "type",
    "material",
    "pattern",
    "capacity",
    "quantity",
    "颜色",
    "色",
    "尺寸",
    "尺码",
    "规格",
    "款式",
    "型号",
    "材质",
    "图案",
    "容量",
    "数量",
)

VARIANT_KEYS = (
    "variants",
    "variant",
    "variantOptions",
    "options",
    "skuList",
    "skcList",
    "skuOptions",
    "specs",
    "specifications",
    "attributes",
    "properties",
)


def detect_no_attribute_product(item: dict[str, Any]) -> dict[str, Any]:
    """Infer whether a scraped Temu product is no-attribute/single-spec.

    The external scraper can change field names, so this scans common variant/spec
    containers instead of relying on one exact schema.
    """

    detected_attributes: set[str] = set()
    variant_count = 0

    for key, value in _walk(item):
        key_text = str(key or "").strip()
        if _looks_like_attribute_name(key_text):
            detected_attributes.add(_normalise_attribute_name(key_text))

        if _looks_like_variant_container(key_text):
            count = _variant_count_from_value(value)
            if count > variant_count:
                variant_count = count
            for attribute in _attribute_names_from_value(value):
                detected_attributes.add(attribute)

    variant_count = max(variant_count, 1)
    has_multiple_variants = variant_count > 1
    has_multiple_attributes = len(detected_attributes) > 0
    return {
        "is_no_attribute": not has_multiple_variants and not has_multiple_attributes,
        "variant_count": variant_count,
        "detected_attributes": sorted(value for value in detected_attributes if value),
    }


def _walk(value: Any, key: str = ""):
    yield key, value
    if isinstance(value, dict):
        for child_key, child_value in value.items():
            yield from _walk(child_value, str(child_key))
    elif isinstance(value, list):
        for child_value in value:
            yield from _walk(child_value, key)


def _looks_like_variant_container(key: str) -> bool:
    lowered = key.lower()
    return any(candidate.lower() in lowered for candidate in VARIANT_KEYS)


def _looks_like_attribute_name(value: str) -> bool:
    lowered = value.lower()
    return any(keyword.lower() == lowered or keyword.lower() in lowered for keyword in ATTRIBUTE_KEYWORDS)


def _normalise_attribute_name(value: str) -> str:
    lowered = value.lower()
    if "color" in lowered or "colour" in lowered or "颜色" in value or value == "色":
        return "颜色"
    if "size" in lowered or "尺寸" in value or "尺码" in value or "规格" in value:
        return "尺寸"
    if "material" in lowered or "fabric" in lowered or "材质" in value:
        return "材质"
    if "style" in lowered or "款式" in value or "型号" in value:
        return "款式"
    return value.strip()


def _variant_count_from_value(value: Any) -> int:
    if isinstance(value, list):
        return len(value)
    if isinstance(value, dict):
        for nested_key in ("items", "list", "values", "options", "skus", "variants"):
            nested = value.get(nested_key)
            if isinstance(nested, list):
                return len(nested)
        if any(_looks_like_attribute_name(str(key)) for key in value):
            return max(1, len([key for key in value if value.get(key) not in (None, "", [])]))
    return 0


def _attribute_names_from_value(value: Any) -> set[str]:
    names: set[str] = set()
    if isinstance(value, dict):
        for key, child in value.items():
            if _looks_like_attribute_name(str(key)):
                names.add(_normalise_attribute_name(str(key)))
            if isinstance(child, (dict, list)):
                names.update(_attribute_names_from_value(child))
            elif _looks_like_attribute_name(str(child)):
                names.add(_normalise_attribute_name(str(child)))
    elif isinstance(value, list):
        for child in value:
            names.update(_attribute_names_from_value(child))
    return names
