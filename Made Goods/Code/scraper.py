"""
scraper.py  —  Made Goods
--------------------------
Platform: madegoods.com (Magento 2 + Varnish cache)

Site structure:
  Listing  : /furniture/{slug}.html?sortBy=...
             Product hrefs: li.product-item a.product-item-link
             Pagination   : ?p=N
  Product  : /furniture/{slug}.html  or  /lighting/{slug}.html  etc.

Product page fields:
  Product Name    : JSON-LD .name  or  h1.page-title span.base
  SKU             : JSON-LD .sku   or  [itemprop="sku"] .value
  Price           : trade-only (blank)
  Image URL       : .product.media img
  Description     : .product.attribute.description .value
  Dimensions      : .additional-attributes table  (W / D / H / Dia)
  Weight          : .additional-attributes table
  Finish          : .additional-attributes table
  COM/COL/COT     : .additional-attributes table (upholstery categories)
  Lighting fields : Socket, Wattage, Extension, etc.
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
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Made Goods")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://www.madegoods.com"
TIMEOUT_MS = 60_000


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a Made Goods category listing with pagination."""
    print(f"  [Listing] {listing_url}")
    all_links: list[str] = []
    seen: set[str] = set()

    base_url = listing_url.split("?")[0].rstrip("/")
    page_num = 1

    while True:
        url = listing_url if page_num == 1 else f"{base_url}?p={page_num}"
        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(3500)
        except Exception as e:
            print(f"  [WARN] page {page_num}: {e}")
            break

        links: list[str] = await page.evaluate(
            f"""() => {{
                const seen = new Set();
                const out  = [];
                // Strategy 1: standard product-item-link
                document.querySelectorAll("a.product-item-link, a.product-name-link").forEach(a => {{
                    const h = a.href || "";
                    if (h.includes("{BASE_URL}") && h.endsWith(".html")) {{
                        const canon = h.split("?")[0];
                        if (!seen.has(canon)) {{ seen.add(canon); out.push(canon); }}
                    }}
                }});
                // Strategy 2: li.product-item first anchor
                if (out.length === 0) {{
                    document.querySelectorAll("li.product-item").forEach(li => {{
                        const a = li.querySelector("a[href]");
                        if (a) {{
                            const h = a.href.split("?")[0];
                            if (h.includes("{BASE_URL}") && !h.includes("/category/") && h.endsWith(".html") && !seen.has(h)) {{
                                seen.add(h); out.push(h);
                            }}
                        }}
                    }});
                }}
                // Strategy 3: any product link deeper than category
                if (out.length === 0) {{
                    document.querySelectorAll("a[href]").forEach(a => {{
                        const h = a.href.split("?")[0];
                        if (h.startsWith("{BASE_URL}") && h.endsWith(".html") && h.split("/").length > 5 && !seen.has(h)) {{
                            seen.add(h); out.push(h);
                        }}
                    }});
                }}
                return out;
            }}"""
        )

        new_links = [l for l in links if l not in seen]
        for l in new_links:
            seen.add(l)
        all_links.extend(new_links)

        if not new_links:
            break

        next_el = await page.query_selector("a.action.next, li.pages-item-next a, a[aria-label='Next']")
        if not next_el:
            break

        page_num += 1
        await async_polite_delay(0.5, 1.0)

    print(f"  [Listing] {len(all_links)} products across {page_num} pages")
    return all_links


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Made Goods product detail page."""
    row: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    # ── 1. JSON-LD ─────────────────────────────────────────────────────────────
    ld_scripts = await page.evaluate("""
        () => Array.from(document.querySelectorAll('script[type="application/ld+json"]'))
                   .map(s => s.textContent)
    """)
    for raw in ld_scripts:
        try:
            d = json.loads(raw)
            candidates = d.get("@graph", [d]) if isinstance(d, dict) else (d if isinstance(d, list) else [d])
            for obj in candidates:
                if obj.get("@type") == "Product":
                    name = clean_text(obj.get("name", ""))
                    if name:
                        row["Product Name"] = name
                    if obj.get("sku"):
                        row["SKU"] = str(obj["sku"])
                    if obj.get("description"):
                        row.setdefault("Description", clean_text(obj["description"]))
                    img = obj.get("image")
                    if img and not row.get("Image URL"):
                        row["Image URL"] = img if isinstance(img, str) else (img[0] if img else "")
                    break
        except Exception:
            pass

    # ── 2. Product Name fallback ───────────────────────────────────────────────
    if not row.get("Product Name"):
        for sel in ["h1.page-title span.base", "h1 span[itemprop='name']", "h1.product-name", "h1"]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                if text and len(text) < 150:
                    row["Product Name"] = text
                    break

    # ── 3. SKU fallback ────────────────────────────────────────────────────────
    if not row.get("SKU"):
        for sel in ["[itemprop='sku']", ".product.sku .value", ".sku .value"]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                if text:
                    row["SKU"] = re.sub(r'^SKU[:\s]+', '', text, flags=re.IGNORECASE).strip()
                    break

    # ── 4. Image URL ───────────────────────────────────────────────────────────
    if not row.get("Image URL"):
        for sel in [
            ".product.media .fotorama__img",
            ".gallery-placeholder img",
            ".product-image-gallery img",
            ".product-image-photo",
            "[data-gallery-role='gallery-placeholder'] img",
        ]:
            img_el = await page.query_selector(sel)
            if img_el:
                src = (
                    await img_el.get_attribute("data-zoom-image")
                    or await img_el.get_attribute("data-src")
                    or await img_el.get_attribute("src")
                    or ""
                )
                if src and "placeholder" not in src and not src.startswith("data:"):
                    if src.startswith("//"):
                        src = "https:" + src
                    elif not src.startswith("http"):
                        src = urljoin(BASE_URL, src)
                    row["Image URL"] = src
                    break

    # ── 5. Description ─────────────────────────────────────────────────────────
    if not row.get("Description"):
        for sel in [
            ".product.attribute.description .value",
            "#description .value",
            ".product-description .value",
            ".product.attribute.overview .value",
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.evaluate("el => el.textContent"))
                if text:
                    row["Description"] = text
                    break

    # ── 6. Attributes Table ────────────────────────────────────────────────────
    attrs = await page.evaluate("""
        () => {
            const result = {};
            document.querySelectorAll(
                'table.additional-attributes tr, .product-attributes tr, .product.attributes tr'
            ).forEach(tr => {
                const th = tr.querySelector('th, td.col.label');
                const td = tr.querySelector('td.col.data, td:not(:first-child)');
                if (th && td) {
                    const label = th.textContent.trim().toLowerCase().replace(/[:\\s]+$/, '');
                    const value = td.textContent.trim().replace(/\\s+/g, ' ');
                    if (label && value) result[label] = value;
                }
            });
            return result;
        }
    """)

    if attrs:
        weight = attrs.get("weight", "")
        if weight:
            m = re.search(r"([\d.]+)", weight)
            if m:
                row["Weight"] = m.group(1)

        dims_raw = attrs.get("dimensions", attrs.get("overall dimensions", ""))
        if dims_raw and not row.get("Width"):
            parsed = parse_dimensions(dims_raw)
            row.update({k: v for k, v in parsed.items() if k not in row})

        for key, col in [("width", "Width"), ("depth", "Depth"), ("height", "Height"),
                         ("diameter", "Diameter"), ("length", "Length")]:
            if attrs.get(key) and not row.get(col):
                m = re.search(r"([\d.]+)", attrs[key])
                if m:
                    row[col] = m.group(1)

        for key in ["finish", "finishes", "finish options", "available finishes"]:
            if attrs.get(key) and not row.get("Finish"):
                row["Finish"] = attrs[key]
                break

        for key in ["material", "materials", "primary material"]:
            if attrs.get(key) and not row.get("Materials"):
                row["Materials"] = attrs[key]
                break

        # Upholstery / seating
        for key, col in [("com", "COM"), ("col", "COL"), ("cot", "COT"),
                         ("seat height", "Seat Height"), ("seat depth", "Seat Depth"),
                         ("arm height", "Arm Height"), ("cushion", "Cushion")]:
            if attrs.get(key):
                row[col] = attrs[key]

        # Lighting
        for key in ["socket", "socket type", "wattage", "bulb type", "extension",
                    "rating", "lightsource", "light source", "color temperature",
                    "canopy", "chain length", "shade details"]:
            if attrs.get(key):
                row[key.title()] = attrs[key]

        # Everything else
        known = {"weight", "dimensions", "overall dimensions", "width", "depth", "height",
                 "diameter", "length", "finish", "finishes", "finish options", "available finishes",
                 "material", "materials", "primary material", "com", "col", "cot",
                 "seat height", "seat depth", "arm height", "cushion",
                 "socket", "socket type", "wattage", "bulb type", "extension",
                 "rating", "lightsource", "light source", "color temperature",
                 "canopy", "chain length", "shade details"}
        for k, v in attrs.items():
            if k not in known and v:
                row[k.title()] = v

    # ── 7. Product Family Id ───────────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


async def main() -> None:
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: all {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor : {info['vendor_name']}")
    print(f"[Scraper] Mode   : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output : {OUTPUT_PATH}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        # Made Goods uses Varnish + bot protection — unblock all resources
        await page.unroute("**/*")

        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_urls:  list[str] = []

            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_urls.append(u)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]

            print(f"\n[Category] {cat['name']}: {len(all_urls)} products")

            for idx, url in enumerate(all_urls, 1):
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = info["vendor_name"]
                        writer.write_row(row, category_name=cat["name"])
                    print(f"  [{idx}] {url.rstrip('/').split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(1.0, 2.5)

            await async_polite_delay(1.0, 3.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
