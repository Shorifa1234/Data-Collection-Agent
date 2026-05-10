"""
scraper.py  —  Mr. Brown London
---------------------------------
Platform: mrbrownhome.com (WordPress + WooCommerce)

Site structure:
  Listing  : /products_by_category/{group}/{slug}/
             Product hrefs: /product/{slug}/
             Pagination   : /page/N/
  Product  : /product/{slug}/

Product page fields:
  Product Name    : h1.product_title (JSON-LD .name)
  SKU             : not provided (generate)
  Price           : .woocommerce-Price-amount bdi (variable product — base price)
  Image URL       : .woocommerce-product-gallery__image img
  Description     : .woocommerce-product-details__short-description
                    — "Finish: X\n\nDimensions:\n\nOverall: WxDxH"
  Weight          : woocommerce-product-attributes table
  Dimensions      : woocommerce-product-attributes table (WxDxH in)
  Finish          : woocommerce-product-attributes table + short description
  COM             : short description text
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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Mr. Brown London")
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

BASE_URL   = "https://mrbrownhome.com"
TIMEOUT_MS = 45_000


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a Mr. Brown London category listing with pagination."""
    print(f"  [Listing] {listing_url}")
    all_links: list[str] = []
    current_url = listing_url
    page_num = 1

    while True:
        try:
            await page.goto(current_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(1500)
        except Exception as e:
            print(f"  [WARN] page {page_num}: {e}")
            break

        links: list[str] = await page.evaluate(
            f"""() => {{
                const seen = new Set();
                const out  = [];
                document.querySelectorAll("a[href]").forEach(a => {{
                    const h = a.href || "";
                    if (h.includes("{BASE_URL}/product/")) {{
                        const canon = h.split("?")[0].replace(/\\/$/, "") + "/";
                        if (!seen.has(canon)) {{
                            seen.add(canon);
                            out.push(canon);
                        }}
                    }}
                }});
                return out;
            }}"""
        )

        if not links:
            break

        all_links.extend(links)

        # Check for next page link
        next_el = await page.query_selector("a.next.page-numbers, a[class*='next']")
        if not next_el:
            break

        next_href = await next_el.get_attribute("href")
        if not next_href or next_href == current_url:
            break

        current_url = next_href
        page_num += 1
        await async_polite_delay(0.5, 1.0)

    # Deduplicate
    seen: set[str] = set()
    unique: list[str] = []
    for u in all_links:
        if u not in seen:
            seen.add(u)
            unique.append(u)

    print(f"  [Listing] {len(unique)} products across {page_num} pages")
    return unique


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Mr. Brown London product detail page."""
    row: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    # ── 1. Product Name ────────────────────────────────────────────────────
    for sel in ["h1.product_title", "h1.entry-title", "h1"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            if text and len(text) < 150:
                row["Product Name"] = text
                break

    # ── 2. Price ──────────────────────────────────────────────────────────
    price_el = await page.query_selector(".woocommerce-Price-amount.amount bdi")
    if not price_el:
        price_el = await page.query_selector("p.price .woocommerce-Price-amount bdi")
    if price_el:
        raw = await price_el.evaluate("el => el.textContent")
        row["Price"] = clean_price(raw)

    # ── 3. Image URL ──────────────────────────────────────────────────────
    for sel in [".woocommerce-product-gallery__image img", ".wp-post-image"]:
        img_el = await page.query_selector(sel)
        if img_el:
            src = (
                await img_el.get_attribute("data-zoom-image")
                or await img_el.get_attribute("data-large_image")
                or await img_el.get_attribute("data-src")
                or await img_el.get_attribute("src")
                or ""
            )
            if src and "placeholder" not in src:
                if src.startswith("//"):
                    src = "https:" + src
                elif not src.startswith("http"):
                    src = urljoin(BASE_URL, src)
                row["Image URL"] = src
                break

    # ── 4. Short Description ──────────────────────────────────────────────
    desc_el = await page.query_selector(".woocommerce-product-details__short-description")
    if desc_el:
        # Use innerText to preserve newlines — needed so dimension regex stops
        # at the right line boundary before unrelated text like "3 drawers".
        desc_raw = await desc_el.evaluate("el => el.innerText")
        desc_text = clean_text(desc_raw)
        if desc_text:
            row["Description"] = desc_text

            # Parse Finish — runs on newline-preserved raw text so it stops at EOL
            finish_m = re.search(r"Finish[:\s]+([^\n\r]+)", desc_raw, re.IGNORECASE)
            if finish_m:
                row["Finish"] = clean_text(finish_m.group(1))

            # Parse Overall dimensions — [^\n\r]+ stops at the line break BEFORE
            # "3 drawers" etc., preventing Pattern A from matching stray numbers.
            dim_m = re.search(
                r"Overall[:\s]+([^\n\r]+)",
                desc_raw, re.IGNORECASE
            )
            if dim_m:
                raw_dim = dim_m.group(1).strip()
                # Normalize curly/smart double-quotes → straight " for parse_dimensions
                raw_dim = raw_dim.replace(chr(0x201c), chr(0x22)).replace(chr(0x201d), chr(0x22))
                parsed = parse_dimensions(raw_dim)
                row.update({k: v for k, v in parsed.items() if k not in row})

    # ── 5. Product Attributes Table ───────────────────────────────────────
    attrs = await page.evaluate("""
        () => {
            const result = {};
            const rows = document.querySelectorAll('table.woocommerce-product-attributes tr');
            rows.forEach(tr => {
                const th = tr.querySelector('th');
                const td = tr.querySelector('td');
                if (th && td) {
                    // strip trailing colon — some WooCommerce labels include it
                    const label = th.textContent.trim().toLowerCase().replace(/:+\\s*$/, '');
                    const value = td.textContent.trim().replace(/\\s+/g, ' ');
                    if (label) result[label] = value;
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

        dims = attrs.get("dimensions", attrs.get("dim", attrs.get("size", "")))
        if dims and not row.get("Width"):
            dims_clean = re.sub(r"\s*in\s*$", "", dims).strip()
            parsed = parse_dimensions(dims_clean)
            row.update({k: v for k, v in parsed.items() if k not in row})

        finish_attr = attrs.get("finishes", attrs.get("finish", attrs.get("finish shown", "")))
        if finish_attr and not row.get("Finish"):
            row["Finish"] = finish_attr

        # COM / upholstery info
        com = attrs.get("com", "")
        if com:
            row["COM"] = com

        # Any other attributes — skip dimension sub-fields; those come from parse_dimensions
        known = {
            "weight", "dimensions", "dim", "size",
            "finishes", "finish", "finish shown", "com",
            "height", "width", "depth", "length", "diameter",
            "h", "w", "d", "l",
        }
        for k, v in attrs.items():
            if k not in known and v:
                row[k.title()] = v

    # ── 6. Product Family Id ──────────────────────────────────────────────
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
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
