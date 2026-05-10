"""
scraper.py  —  Paul Ferrante
------------------------------
Platform: paulferrante.com (WordPress + WooCommerce + Elementor)

Site structure:
  Listing  : /product-category/{group}/{slug}/
             Product hrefs follow the pattern: /product/{slug}/
             No pagination (small catalog, all on one page)
  Product  : /product/{slug}/

Product page fields:
  Product Name : h1.product_title
  SKU          : numeric prefix from slug (e.g. "6204" from "6204-farm-table")
  Price        : .woocommerce-Price-amount bdi
  Image URL    : .woocommerce-product-gallery__image img (data-src or src)
  Description  : .woocommerce-product-details__short-description
  Dimensions   : woocommerce-product-attributes table — parse into individual fields
  Width        : extracted from Dimensions or individual row
  Depth        : extracted from Dimensions or individual row
  Height       : extracted from Dimensions or individual row
  Diameter     : extracted from Dimensions or individual row
  Length       : extracted from Dimensions or individual row
  Weight       : woocommerce-product-attributes table
  Finish       : woocommerce-product-attributes table (Finishes / Finish Shown row)
  Lighting fields: Wattage, Socket, Extension, etc.
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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Paul Ferrante")
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

BASE_URL   = "https://paulferrante.com"
TIMEOUT_MS = 90_000


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect product URLs from a Paul Ferrante category listing (no pagination)."""
    print(f"  [Listing] {listing_url}")
    for attempt in range(2):
        try:
            await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            # Check for Cloudflare challenge
            title = await page.title()
            if "just a moment" in title.lower() or "cloudflare" in title.lower():
                print(f"  [CF] Waiting for challenge...")
                await page.wait_for_timeout(8000)
            else:
                await page.wait_for_timeout(3000)
            break
        except Exception as e:
            if attempt == 0:
                print(f"  [WARN] Retrying {listing_url}: {e}")
                await page.wait_for_timeout(6000)
            else:
                print(f"  [WARN] {e}")
                return []

    links: list[str] = await page.evaluate(
        f"""() => {{
            const seen = new Set();
            const out  = [];
            document.querySelectorAll("a[href]").forEach(a => {{
                const h = a.href || "";
                if (h.includes("{BASE_URL}/product/") && !h.includes("/product-category/")) {{
                    const canon = h.replace(/\\/$/, "") + "/";
                    if (!seen.has(canon)) {{
                        seen.add(canon);
                        out.push(canon);
                    }}
                }}
            }});
            return out;
        }}"""
    )

    print(f"  [Listing] {len(links)} products")
    return links


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Paul Ferrante product detail page."""
    row: dict = {"Source": url}

    for attempt in range(2):
        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            title = await page.title()
            if "just a moment" in title.lower() or "cloudflare" in title.lower():
                await page.wait_for_timeout(8000)
            else:
                await page.wait_for_timeout(3000)
            break
        except Exception as e:
            if attempt == 0:
                print(f"    [WARN] Retrying: {e}")
                await page.wait_for_timeout(6000)
            else:
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

    # ── 2. SKU — numeric prefix from URL slug ─────────────────────────────
    slug = url.rstrip("/").split("/")[-1]
    sku_m = re.match(r"^(\d+)", slug)
    if sku_m:
        row["SKU"] = sku_m.group(1)

    # ── 3. Price ──────────────────────────────────────────────────────────
    price_el = await page.query_selector(".woocommerce-Price-amount.amount bdi")
    if not price_el:
        price_el = await page.query_selector(".woocommerce-Price-amount bdi")
    if price_el:
        raw_price = await price_el.evaluate("el => el.textContent")
        row["Price"] = clean_price(raw_price)

    # ── 4. Image URL ──────────────────────────────────────────────────────
    img_selectors = [
        ".woocommerce-product-gallery__image img",
        ".wp-post-image",
        "img.attachment-woocommerce_single",
    ]
    for sel in img_selectors:
        img_el = await page.query_selector(sel)
        if img_el:
            src = (
                await img_el.get_attribute("data-zoom-image")
                or await img_el.get_attribute("data-src")
                or await img_el.get_attribute("src")
                or ""
            )
            if src and "placeholder" not in src and not src.endswith(".gif"):
                if src.startswith("//"):
                    src = "https:" + src
                elif not src.startswith("http"):
                    src = urljoin(BASE_URL, src)
                row["Image URL"] = src
                break

    # ── 5. Short Description ──────────────────────────────────────────────
    desc_el = await page.query_selector(".woocommerce-product-details__short-description")
    if desc_el:
        desc = clean_text(await desc_el.evaluate("el => el.textContent"))
        if desc:
            row["Description"] = desc

    # ── 6. Product Attributes Table ───────────────────────────────────────
    # IMPORTANT: WooCommerce labels often include trailing colons (e.g. "Dimensions:")
    # Strip them in JS so lookups work correctly.
    attrs = await page.evaluate("""
        () => {
            const result = {};
            const rows = document.querySelectorAll('table.woocommerce-product-attributes tr');
            rows.forEach(tr => {
                const th = tr.querySelector('th');
                const td = tr.querySelector('td');
                if (th && td) {
                    const label = th.textContent.trim().toLowerCase().replace(/:+\\s*$/, '');
                    const value = td.textContent.trim().replace(/\\s+/g, ' ');
                    if (label && value) result[label] = value;
                }
            });
            return result;
        }
    """)

    if attrs:
        # Weight
        weight = attrs.get("weight", "")
        if weight:
            m = re.search(r"([\d.]+)", weight)
            if m:
                row["Weight"] = m.group(1)

        # Dimensions — stored as "W x D x H in", "119"l x 36"d x 30"h", etc.
        # parse_dimensions() returns a dict with Dimensions + individual W/D/H/Dia/L keys
        dims = attrs.get("dimensions", attrs.get("dim", attrs.get("size", "")))
        if dims:
            dims_clean = re.sub(r"\s*in\s*$", "", dims, flags=re.IGNORECASE).strip()
            parsed = parse_dimensions(dims_clean)
            for k, v in parsed.items():
                if k not in row:
                    row[k] = v

        # Individual dimension rows (sometimes listed separately)
        for attr_key, col_name in [
            ("width", "Width"), ("depth", "Depth"), ("height", "Height"),
            ("diameter", "Diameter"), ("length", "Length"),
        ]:
            if attrs.get(attr_key) and not row.get(col_name):
                m = re.search(r"([\d.]+)", attrs[attr_key])
                if m:
                    row[col_name] = m.group(1)

        # Paul Ferrante unlabeled format: "43 x 63 x 20h" (only H labeled, W/D by position)
        # Only run when Length is also unset — if Length was parsed, all labels were present
        # and the positional inference would wrongly assign Length → Width.
        if row.get("Height") and not row.get("Width") and not row.get("Length") and row.get("Dimensions"):
            numbers = re.findall(r'[\d.]+', row["Dimensions"])
            h_val = str(row["Height"])
            non_h = [n for n in numbers if n != h_val]
            if len(non_h) == 2:
                row.setdefault("Width", non_h[0])
                row.setdefault("Depth", non_h[1])
            elif len(non_h) == 1:
                row.setdefault("Width", non_h[0])

        # Rebuild Dimensions string if we have individual values but no Dimensions
        if not row.get("Dimensions"):
            dim_parts = []
            for col, label in [("Width", "W"), ("Depth", "D"), ("Height", "H"),
                                ("Diameter", "Dia"), ("Length", "L")]:
                if row.get(col):
                    dim_parts.append(f"{label} {row[col]}")
            if dim_parts:
                row["Dimensions"] = " x ".join(dim_parts)

        # Finishes — "finish shown:" is common on Paul Ferrante
        finish = attrs.get("finishes", attrs.get("finish", attrs.get("finish shown", "")))
        if finish:
            row.setdefault("Finish", finish)

        # Lighting specifics
        for key in ("wattage", "socket", "socket type", "voltage", "extension",
                    "rating", "shade details", "canopy", "chain length", "bulb type",
                    "bulb qty", "base"):
            val = attrs.get(key, "")
            if val:
                row[key.title()] = val

        # Any other attributes — skip item-number duplicates
        known = {
            "weight", "dimensions", "dim", "size",
            "width", "depth", "height", "diameter", "length",
            "finishes", "finish", "finish shown",
            "wattage", "socket", "socket type", "voltage", "extension",
            "rating", "shade details", "canopy", "chain length", "bulb type",
            "bulb qty", "base",
            "item #", "item#", "item", "item no", "item no.", "item number",
        }
        for k, v in attrs.items():
            if k not in known and v:
                col = k.title()
                if col.lower() not in ("sku", "item #", "item#"):
                    row.setdefault(col, v)

    # ── 7. Additional info from tabs (Lighting: wattage etc.) ─────────────
    tab_text = await page.evaluate("""
        () => {
            const panel = document.querySelector('.woocommerce-Tabs-panel--description, .woocommerce-Tabs-panel');
            return panel ? panel.textContent.trim() : '';
        }
    """)
    if tab_text:
        for label, col in [
            (r"wattage[:\\s]+([\w\\s.]+)", "Wattage"),
            (r"socket[:\\s]+([\w\\s/]+)", "Socket"),
            (r"bulb[:\\s]+([\w\\s]+)", "Bulb Type"),
            (r"cord[:\\s]+([\d.]+)", "Cord Length"),
        ]:
            if col not in row:
                m = re.search(label, tab_text, re.IGNORECASE)
                if m:
                    row[col] = clean_text(m.group(1))

    # ── 8. Product Family Id ──────────────────────────────────────────────
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
        # Unblock all resources so CF challenges can render
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
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
