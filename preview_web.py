from __future__ import annotations

import html
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


def create_preview_html(excel_file: Path, output_dir: Path | None = None) -> Path:
    frame = _read_preview_frame(excel_file).fillna("")
    output_dir = output_dir or excel_file.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    html_file = output_dir / f"preview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    html_file.write_text(_render_preview_html(excel_file, frame), encoding="utf-8")
    return html_file


def _read_preview_frame(excel_file: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(excel_file, sheet_name="导入模板")
    except Exception:
        return pd.read_excel(excel_file)


def _render_preview_html(excel_file: Path, frame: pd.DataFrame) -> str:
    rows = []
    for index, row in frame.iterrows():
        title = _first(row, ["*产品标题", "产品标题", "商品标题（中文）"])
        english_title = _first(row, ["*英文标题", "英文标题", "商品标题（英文）"])
        price = _first(row, ["*申报价格\n(店铺币种)", "美元价格($)", "价格"])
        preview_images = _image_values(_first(row, ["预览图", "商品主图"]))
        carousel_images = _image_values(_first(row, ["*轮播图", "商品轮播图", "*产品素材图"]))
        rows.append(
            f"""
            <article class="product-card">
              <header>
                <span class="row-index">#{index + 2}</span>
                <div>
                  <h2>{_esc(english_title or title)}</h2>
                  <p class="source-title">{_esc(title)}</p>
                </div>
                <strong class="price">{_esc(price)}</strong>
              </header>
              <section class="image-grid">
                <div>
                  <h3>预览图</h3>
                  {_render_images(preview_images[:1], "preview")}
                </div>
                <div>
                  <h3>轮播图</h3>
                  {_render_images(carousel_images[:10], "carousel")}
                </div>
              </section>
            </article>
            """
        )

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Temu 上架预览</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #1f2933;
      --muted: #667085;
      --line: #d9dee7;
      --accent: #2563eb;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, "Microsoft YaHei", sans-serif;
      background: var(--bg);
      color: var(--text);
    }}
    .app-header {{
      position: sticky;
      top: 0;
      z-index: 10;
      background: rgba(255,255,255,0.96);
      border-bottom: 1px solid var(--line);
      padding: 14px 24px;
      display: flex;
      gap: 18px;
      align-items: center;
      justify-content: space-between;
    }}
    .app-header h1 {{
      font-size: 18px;
      margin: 0;
    }}
    .app-header p {{
      margin: 4px 0 0;
      color: var(--muted);
      font-size: 13px;
      word-break: break-all;
    }}
    .search {{
      width: min(420px, 34vw);
      padding: 9px 11px;
      border: 1px solid var(--line);
      border-radius: 6px;
      font-size: 14px;
    }}
    main {{
      padding: 18px 24px 40px;
      display: grid;
      gap: 14px;
    }}
    .product-card {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
    }}
    .product-card header {{
      display: grid;
      grid-template-columns: auto 1fr auto;
      gap: 12px;
      align-items: start;
      border-bottom: 1px solid var(--line);
      padding-bottom: 12px;
      margin-bottom: 12px;
    }}
    .row-index {{
      background: #eef4ff;
      color: var(--accent);
      border-radius: 4px;
      padding: 5px 8px;
      font-weight: 700;
      font-size: 13px;
    }}
    h2 {{
      margin: 0;
      font-size: 16px;
      line-height: 1.35;
    }}
    .source-title {{
      margin: 5px 0 0;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }}
    .price {{
      color: #0f766e;
      white-space: nowrap;
      font-size: 15px;
    }}
    .image-grid {{
      display: grid;
      grid-template-columns: 220px 1fr;
      gap: 16px;
    }}
    h3 {{
      margin: 0 0 8px;
      font-size: 13px;
      color: var(--muted);
      font-weight: 700;
    }}
    .images {{
      display: flex;
      gap: 10px;
      overflow-x: auto;
      padding-bottom: 6px;
    }}
    .image-cell {{
      flex: 0 0 auto;
      width: 150px;
    }}
    .preview .image-cell {{
      width: 190px;
    }}
    img {{
      width: 100%;
      aspect-ratio: 1 / 1;
      object-fit: contain;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
    }}
    .url {{
      margin-top: 5px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.25;
      overflow-wrap: anywhere;
    }}
    .empty {{
      color: var(--muted);
      font-size: 13px;
      border: 1px dashed var(--line);
      border-radius: 6px;
      padding: 18px;
    }}
    @media (max-width: 860px) {{
      .app-header {{ align-items: stretch; flex-direction: column; }}
      .search {{ width: 100%; }}
      .product-card header {{ grid-template-columns: 1fr; }}
      .image-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header class="app-header">
    <div>
      <h1>Temu 上架预览</h1>
      <p>{_esc(str(excel_file))} · 共 {len(frame)} 行</p>
    </div>
    <input class="search" id="search" placeholder="搜索标题、英文标题或价格">
  </header>
  <main id="cards">
    {''.join(rows)}
  </main>
  <script>
    const search = document.getElementById('search');
    const cards = Array.from(document.querySelectorAll('.product-card'));
    search.addEventListener('input', () => {{
      const term = search.value.trim().toLowerCase();
      for (const card of cards) {{
        card.style.display = card.innerText.toLowerCase().includes(term) ? '' : 'none';
      }}
    }});
  </script>
</body>
</html>"""


def _first(row, names: list[str]) -> str:
    for name in names:
        if name in row and str(row[name]).strip():
            return str(row[name]).strip()
    return ""


def _image_values(value: str) -> list[str]:
    text = str(value or "").strip().strip("[]")
    if not text:
        return []
    return [part.strip().strip("'\"") for part in re.split(r"[;,\n]", text) if part.strip()]


def _render_images(images: list[str], class_name: str) -> str:
    if not images:
        return '<div class="empty">暂无图片</div>'
    cells = []
    for image in images:
        safe = _esc(image)
        cells.append(
            f"""
            <div class="image-cell">
              <img loading="lazy" src="{safe}" alt="product image">
              <div class="url">{safe}</div>
            </div>
            """
        )
    return f'<div class="images {class_name}">{"".join(cells)}</div>'


def _esc(value: object) -> str:
    return html.escape(str(value or ""), quote=True)
