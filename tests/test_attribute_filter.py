from attribute_filter import detect_no_attribute_product


def test_detect_no_attribute_product_without_variant_fields() -> None:
    result = detect_no_attribute_product({"title": "Simple item", "price": 9.99})

    assert result == {"is_no_attribute": True, "variant_count": 1, "detected_attributes": []}


def test_detect_no_attribute_product_with_color_and_size_variants() -> None:
    result = detect_no_attribute_product(
        {
            "title": "Variant item",
            "variantOptions": [
                {"name": "Color", "values": ["Blue", "Pink"]},
                {"name": "Size", "values": ["Queen", "King"]},
            ],
        }
    )

    assert result["is_no_attribute"] is False
    assert result["variant_count"] == 2
    assert result["detected_attributes"] == ["尺寸", "颜色"]
