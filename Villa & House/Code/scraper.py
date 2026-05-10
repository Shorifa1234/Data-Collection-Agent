"""
scraper.py — Villa & House
---------------------------
Platform: vandh.com (BigCommerce)

Site structure:
  Listing  : /category-slug/  — BigCommerce collection pages, JS-rendered
             Product cards in <article class="card"> with <a class="card-figure__link">
             Pagination: ?page=N (checked per listing)
  Product  : /product-slug   e.g. /aus-300-94

Product page fields (trade-only — no price shown without login):
  Product Name   : <h1> text
  SKU            : "SKU: {value}" text near product title
  Price          : empty (trade portal — LOG IN FOR PRICING)
  Dimensions     : inline text like "48W x 24D x 17.5H" in description/body
  Materials      : described in product description text
  Description    : full product description paragraph
  Specifications : Item Weight, Box Dimensions, etc.
  Image URL      : BigCommerce CDN image (high-res)
  Tearsheet      : constructed via "PRINT TEARSHEET" link if present

Run directly:
    python scraper.py
Or via orchestrator:
    python orchestrator.py "Villa & House"
    python orchestrator.py "Villa & House" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    sentence_case,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Villa & House")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://vandh.com"

# Map spec labels from SPECIFICATIONS tab → output field names
_SPEC_LABEL_MAP: dict[str, str] = {
    "item weight":         "Weight",
    "shipping weight":     "Shipping Weight",
    "box dimensions":      "Box Dimensions",
    "boxed weight":        "Boxed Weight",
    "distance between legs": "Distance Between Legs",
    "floor to lowest point height": "Floor to Lowest Height",
    "product has hand made qualities": "Handmade",
    "color":               "Color",
    "finish":              "Finish",
    "material":            "Materials",
    "style":               "Style",
    "collection":          "Collection",
    "designer":            "Designer",
    "lead time":           "Lead Time",
    "origin":              "Origin",
    "seat height":         "Seat Height",
    "seat depth":          "Seat Depth",
    "seat width":          "Seat Width",
    "arm height":          "Arm Height",
    "arm width":           "Arm Width",
    "arm length":          "Arm Length",
    "cushion":             "Cushion",
    "com":                 "COM",
    "wattage":             "Wattage",
    "socket":              "Socket",
    "shade":               "Shade Details",
    "shade details":       "Shade Details",
}


# ---------------------------------------------------------------------------
# Listing page
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect product URLs from a BigCommerce listing page.
    Cards use <article class="card"> > <a class="card-figure__link">.
    Pagination: ?page=N
    """
    links: list[str] = []
    seen: set[str] = set()
    page_num = 1

    while True:
        url = listing_url if page_num == 1 else f"{listing_url.rstrip('/')}?page={page_num}"
        try:
            await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
            await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"  [WARN] {url}: {e}")
            break

        hrefs: list[str] = await page.evaluate("""() => {
            const out = [];
            document.querySelectorAll("article.card a.card-figure__link").forEach(a => {
                if (a.href) out.push(a.href);
            });
            return out;
        }""")

        if not hrefs:
            break

        found_new = False
        for h in hrefs:
            if h not in seen:
                seen.add(h)
                links.append(h)
                found_new = True

        # Check for next-page link
        next_el = await page.query_selector("a[aria-label='Next page'], .pagination-item--next a")
        if not next_el or not found_new:
            break

        page_num += 1
        await async_polite_delay(0.5, 1.0)

    return links


# ---------------------------------------------------------------------------
# Product detail page
# ---------------------------------------------------------------------------

async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Villa & House product detail page. Returns a list with one dict."""
    row: dict = {"Source": url}

    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
    except Exception as e:
        print(f"    [WARN] {url}: {e}")
        return [row]

    # ── 1. Product Name ───────────────────────────────────────────────────
    for sel in ("h1.productView-title", "h1.product-name", "h1"):
        el = await page.query_selector(sel)
        if el:
            txt = clean_text(await el.inner_text())
            if txt:
                row["Product Name"] = sentence_case(txt)
                break

    # ── 2. SKU ───────────────────────────────────────────────────────────
    sku_el = await page.query_selector(".productView-info-value[data-product-sku], [itemprop='sku']")
    if sku_el:
        row["SKU"] = clean_text(await sku_el.inner_text())
    else:
        # Search body text for "SKU: XYZ"
        body_txt = await page.inner_text("body")
        m = re.search(r"SKU[:\s]+([A-Z0-9\-]+)", body_txt, re.I)
        if m:
            row["SKU"] = m.group(1).strip()

    # ── 3. Image URL ─────────────────────────────────────────────────────
    # BigCommerce CDN — prefer the highest-res stencil variant
    img_src = await page.evaluate("""() => {
        const candidates = [
            ...document.querySelectorAll(
                ".productView-image--default img, .productView-img-container img, img[itemprop='image']"
            )
        ];
        for (const img of candidates) {
            const src = img.getAttribute("data-zoom-image") || img.getAttribute("data-src") || img.src || "";
            if (src && src.includes("bigcommerce")) return src;
        }
        // Fallback: any BigCommerce CDN img
        const all = [...document.querySelectorAll("img")];
        for (const img of all) {
            const src = img.src || "";
            if (src.includes("bigcommerce.com/s-") && !src.includes("stencil/50")) return src;
        }
        return "";
    }""")
    if img_src:
        # Upgrade to max resolution (1280w or original)
        img_src = re.sub(r"stencil/\d+(?:x\d+)?/", "stencil/1280w/", img_src)
        img_src = img_src.split("?")[0]
        row["Image URL"] = img_src

    # ── 4. Description ───────────────────────────────────────────────────
    desc = ""
    for sel in (".productView-description", "[data-tab-content='description']",
                ".product-description", ".tab-content", ".productView-full-description"):
        el = await page.query_selector(sel)
        if el:
            txt = clean_text(await el.inner_text())
            if txt and len(txt) > 20:
                # Trim at known navigation / boilerplate sections
                for stop in ("SPECIFICATIONS", "CARE, SHIPPING", "VILLA & HOUSE CRAFTSMANSHIP",
                             "JOIN OUR MAILING LIST", "PRINT TEARSHEET"):
                    idx = txt.upper().find(stop)
                    if idx > 40:
                        txt = txt[:idx].strip()
                desc = txt
                break

    # Fallback: grab first descriptive paragraph from body text
    if not desc:
        for line in body_txt.split("\n"):
            line = line.strip()
            if len(line) > 80 and not any(kw in line.upper() for kw in
                    ("LOG IN", "SKU:", "SPECIFICATIONS", "FINISH:", "CART", "REGISTER")):
                desc = line
                break
    if desc:
        row["Description"] = desc[:1500]  # cap length

    # ── 5. Dimensions — from body text ───────────────────────────────────
    body_txt = await page.inner_text("body")
    # Pattern: "48W x 24D x 17.5H" or "48W x 24D" or "18.5Dia"
    dim_patterns = [
        r"(\d+(?:\.\d+)?[WLwl]\s*[xX×]\s*\d+(?:\.\d+)?[Dd]\s*[xX×]\s*\d+(?:\.\d+)?[Hh])",
        r"(\d+(?:\.\d+)?[WLwl]\s*[xX×]\s*\d+(?:\.\d+)?[Hh])",
        r"(\d+(?:\.\d+)?\s*[\"']?\s*(?:W|L)\s*[xX×]\s*\d+(?:\.\d+)?\s*[\"']?\s*D\s*[xX×]\s*\d+(?:\.\d+)?\s*[\"']?\s*H)",
    ]
    for pat in dim_patterns:
        m = re.search(pat, body_txt)
        if m:
            dim_str = m.group(1).strip()
            row.setdefault("Dimensions", dim_str)
            parsed = parse_dimensions(dim_str)
            for k, v in parsed.items():
                if k != "Dimensions":
                    row.setdefault(k, v)
            break

    # ── 6. Specifications ─────────────────────────────────────────────────
    # VandH BigCommerce — specs are in .productView-info-item elements (dt/dd style)
    spec_items = await page.query_selector_all(".productView-info-item")
    for item in spec_items:
        name_el = await item.query_selector(".productView-info-name")
        value_el = await item.query_selector(".productView-info-value")
        if name_el and value_el:
            label = clean_text(await name_el.inner_text()).lower().rstrip(":")
            value = clean_text(await value_el.inner_text())
            field = _SPEC_LABEL_MAP.get(label)
            if field and value:
                row.setdefault(field, value)

    # Also scan body text for key-value pairs from the SPECIFICATIONS tab section
    # VandH uses: "Item Weight:\n79 lbs" style
    spec_area_m = re.search(
        r"SPECIFICATIONS\s*\n(.*?)(?=CARE,|DESCRIPTION|VILLA|$)",
        body_txt, re.DOTALL | re.I
    )
    spec_area = spec_area_m.group(1) if spec_area_m else body_txt
    for label_lower, field in _SPEC_LABEL_MAP.items():
        if field in row:
            continue
        pattern = re.compile(
            re.escape(label_lower.title()) + r"[:\s]*\n?\s*(.+?)(?:\n|$)", re.I
        )
        m = pattern.search(spec_area)
        if m:
            value = clean_text(m.group(1))
            # Reject if it looks like navigation/button text
            if value and len(value) < 150 and value not in ("Yes", "No", "PRINT TEARSHEET", "SHARE"):
                row[field] = value

    # ── 7. Tearsheet ────────────────────────────────────────────────────
    tearsheet_el = await page.query_selector("a[href*='tearsheet'], a[href*='pdf']")
    if tearsheet_el:
        href = await tearsheet_el.get_attribute("href") or ""
        if href and "tearsheet" in href.lower():
            row["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)

    # ── 8. Product Family Id ─────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    row["Manufacturer"] = VENDOR_NAME
    return [row]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main() -> None:
    info       = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer     = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    categories = info["categories"]

    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            print(f"\n[{cat['group']}] {cat['name']} — collecting links…")

            seen_urls: set[str] = set()
            all_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_urls.append(u)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]
            print(f"  {len(all_urls)} products")

            global_idx = 1
            for url in all_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                    print(f"  [{global_idx - 1}] {url.rstrip('/').split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.8, 2.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
