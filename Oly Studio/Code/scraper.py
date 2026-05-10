"""
scraper.py  —  Oly Studio
--------------------------
Platform: olystudio.com (Magento 2 + Cloudflare)

Site structure:
  Listing  : /{category}.html?availability=262
             Product hrefs: li.product-item a.product-item-link
             Pagination   : ?p=N or a.action.next
  Product  : /{product-slug}.html

Product page fields:
  Product Name    : JSON-LD .name  or  h1.page-title span.base
  SKU             : JSON-LD .sku   or  [itemprop="sku"]
  Price           : JSON-LD offers.price
  Image URL       : .product.media img (data-zoom-image / data-src / src)
  Description     : .product.attribute.description .value
  Dimensions      : .additional-attributes table  (Width / Depth / Height / Dia)
  Weight          : .additional-attributes table
  Finish          : .additional-attributes table  (Finishes / Finish row)
  Materials       : .additional-attributes table
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlencode, parse_qs

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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Oly Studio")
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

# When running headed (HEADLESS=false), the browser opens and waits this many
# seconds before scraping starts — giving time to manually pass CF challenges.
# Override with MANUAL_WAIT=0 to skip the wait even in headed mode.
MANUAL_WAIT_SECS = int(os.environ.get("MANUAL_WAIT", "60" if not HEADLESS else "0"))

BASE_URL   = "https://www.olystudio.com"
TIMEOUT_MS = 60_000


def _paged_url(listing_url: str, page_num: int) -> str:
    """Build a paginated URL that preserves all existing query parameters."""
    if page_num == 1:
        return listing_url
    parsed = urlparse(listing_url)
    params = parse_qs(parsed.query, keep_blank_values=True)
    params["p"] = [str(page_num)]
    return parsed._replace(query=urlencode({k: v[0] for k, v in params.items()})).geturl()


async def _wait_for_cf(page, label: str = "") -> None:
    """Poll page title until Cloudflare challenge clears (up to 30 s)."""
    for i in range(30):
        title = await page.title()
        if "just a moment" not in title.lower() and "checking" not in title.lower():
            return
        if i == 0:
            print(f"  [CF] Waiting for challenge to clear{' — ' + label if label else ''}...")
        await page.wait_for_timeout(1000)
    print(f"  [CF] Challenge may not have cleared — proceeding anyway")


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from an Oly Studio category listing with pagination."""
    print(f"  [Listing] {listing_url}")
    all_links: list[str] = []
    seen: set[str] = set()

    page_num = 1

    while True:
        url = _paged_url(listing_url, page_num)
        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await _wait_for_cf(page, url)
            # Wait for Magento product grid to render (AJAX-loaded after DOM ready)
            try:
                await page.wait_for_selector(
                    "a.product-item-link, li.product-item, .product-item-info",
                    timeout=12_000,
                )
            except Exception:
                await page.wait_for_timeout(5000)
        except Exception as e:
            print(f"  [WARN] page {page_num}: {e}")
            break

        # Extract product links via JS
        links: list[str] = await page.evaluate(
            f"""() => {{
                const seen = new Set();
                const out  = [];
                // Strategy 1: standard Magento product-item-link
                document.querySelectorAll("a.product-item-link, a.product-name-link").forEach(a => {{
                    const h = a.href || "";
                    if (h.includes("{BASE_URL}") && h.endsWith(".html")) {{
                        const canon = h.split("?")[0];
                        if (!seen.has(canon)) {{ seen.add(canon); out.push(canon); }}
                    }}
                }});
                // Strategy 2: li.product-item first link
                if (out.length === 0) {{
                    document.querySelectorAll("li.product-item, div.product-item-info").forEach(li => {{
                        const a = li.querySelector("a[href]");
                        if (a) {{
                            const h = a.href.split("?")[0];
                            if (h.includes("{BASE_URL}") && h.endsWith(".html") && !seen.has(h)) {{
                                seen.add(h); out.push(h);
                            }}
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

        # Check for next page
        next_el = await page.query_selector("a.action.next, li.pages-item-next a, a[aria-label='Next']")
        if not next_el:
            break

        page_num += 1
        await async_polite_delay(0.5, 1.0)

    print(f"  [Listing] {len(all_links)} products across {page_num} pages")
    return all_links


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape an Oly Studio product detail page."""
    row: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await _wait_for_cf(page)
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    html = await page.content()

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
                    offers = obj.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0] if offers else {}
                    price = offers.get("price") or offers.get("lowPrice")
                    if price:
                        row["Price"] = clean_price(str(price))
                    img = obj.get("image")
                    if img and not row.get("Image URL"):
                        row["Image URL"] = img if isinstance(img, str) else (img[0] if img else "")
                    break
        except Exception:
            pass

    # ── 2. Product Name fallback ───────────────────────────────────────────────
    if not row.get("Product Name"):
        for sel in ["h1.page-title span.base", "h1.product-name span", "h1 span[itemprop='name']", "h1"]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                if text and len(text) < 150:
                    row["Product Name"] = text
                    break

    # ── 3. SKU fallback ────────────────────────────────────────────────────────
    if not row.get("SKU"):
        for sel in ["[itemprop='sku']", ".product.sku .value", ".sku .value", "div.sku"]:
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
            ".product-description",
            "[data-ui-id='page-title-wrapper'] + div",
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.evaluate("el => el.textContent"))
                if text and len(text) > 10:
                    row["Description"] = text
                    break

    # ── 6. Product Attributes Table ────────────────────────────────────────────
    attrs = await page.evaluate("""
        () => {
            const result = {};

            // Magento 2 additional-attributes table (most common on Oly Studio)
            document.querySelectorAll(
                'table.additional-attributes tr, ' +
                '.product-attributes tr, ' +
                '.product.info.detailed table tr, ' +
                '#product-attribute-specs-table tr'
            ).forEach(tr => {
                const th = tr.querySelector('th, td.col.label');
                const td = tr.querySelector('td.col.data, td:last-child');
                if (th && td && th !== td) {
                    const label = th.textContent.trim().toLowerCase().replace(/[:\\s]+$/, '');
                    const value = td.textContent.trim().replace(/\\s+/g, ' ');
                    if (label && value) result[label] = value;
                }
            });

            // dl/dt/dd pattern inside description or specs tabs
            document.querySelectorAll(
                '.product.attribute.description dl, .specs dl, .additional-attributes dl'
            ).forEach(dl => {
                const dts = dl.querySelectorAll('dt');
                const dds = dl.querySelectorAll('dd');
                dts.forEach((dt, i) => {
                    const label = dt.textContent.trim().toLowerCase();
                    const value = dds[i] ? dds[i].textContent.trim() : '';
                    if (label && value) result[label] = value;
                });
            });

            return result;
        }
    """)

    if attrs:
        weight = attrs.get("weight", attrs.get("ship weight", attrs.get("shipping weight", "")))
        if weight:
            m = re.search(r"([\d.]+)", weight)
            if m:
                row["Weight"] = m.group(1)

        # Dimensions — try combined field first, then individual fields
        dims_raw = attrs.get("dimensions", attrs.get("overall dimensions", attrs.get("size", "")))
        if dims_raw:
            parsed = parse_dimensions(dims_raw)
            for k, v in parsed.items():
                row.setdefault(k, v)

        # Individual dimension attributes (override combined only if not already set)
        for key, col in [("width", "Width"), ("depth", "Depth"), ("height", "Height"),
                         ("diameter", "Diameter"), ("length", "Length")]:
            if attrs.get(key) and not row.get(col):
                m = re.search(r"([\d.]+)", attrs[key])
                if m:
                    row[col] = m.group(1)

        # Rebuild Dimensions string if still missing
        if not row.get("Dimensions"):
            parts = []
            for lbl, field in [("W", "Width"), ("D", "Depth"), ("H", "Height"),
                                ("Dia", "Diameter"), ("L", "Length")]:
                if row.get(field):
                    parts.append(f"{lbl} {row[field]}")
            if parts:
                row["Dimensions"] = " x ".join(parts)

        for key in ["finish", "finishes", "finish options", "available finishes"]:
            if attrs.get(key) and not row.get("Finish"):
                row["Finish"] = attrs[key]
                break

        for key in ["material", "materials", "primary material", "frame material"]:
            if attrs.get(key) and not row.get("Materials"):
                row["Materials"] = attrs[key]
                break

        for key in ["collection"]:
            if attrs.get(key) and not row.get("Collection"):
                row["Collection"] = attrs[key]

        # Lighting-specific
        for key, col in [
            ("wattage", "Wattage"), ("socket", "Socket"), ("socket type", "Socket Type"),
            ("bulb type", "Bulb Type"), ("bulb qty", "Bulb Qty"), ("voltage", "Voltage"),
            ("lamping", "Lamping"), ("shade", "Shade Details"),
            ("canopy", "Canopy"), ("chain length", "Chain Length"),
            ("hanging length", "Hanging Length"), ("extension", "Extension"),
        ]:
            if attrs.get(key) and not row.get(col):
                row[col] = attrs[key]

        # Any remaining unknown attributes
        known = {
            "weight", "ship weight", "shipping weight",
            "dimensions", "overall dimensions", "size",
            "width", "depth", "height", "diameter", "length",
            "finish", "finishes", "finish options", "available finishes",
            "material", "materials", "primary material", "frame material",
            "collection",
            "wattage", "socket", "socket type", "bulb type", "bulb qty",
            "voltage", "lamping", "shade", "canopy", "chain length",
            "hanging length", "extension",
        }
        for k, v in attrs.items():
            if k not in known and v:
                row.setdefault(k.title(), v)

    # ── 7. Tearsheet link ──────────────────────────────────────────────────────
    tearsheet = await page.evaluate("""
        () => {
            const a = document.querySelector(
                'a[href*="tearsheet"], a[href*="spec-sheet"], ' +
                'a[href*="specsheet"], a[href*="cut-sheet"]'
            );
            return a ? a.href : '';
        }
    """)
    if tearsheet:
        row["Tearsheet Link"] = tearsheet

    # ── 8. Product Family Id ───────────────────────────────────────────────────
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
        # Unblock ALL resources — Cloudflare challenge requires JS + CSS to run
        await page.unroute("**/*")

        # Extra stealth: mask automation signals Cloudflare checks for
        await page.add_init_script("""
            window.chrome = { runtime: {}, loadTimes: function(){}, csi: function(){} };
            Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
            Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
            Object.defineProperty(navigator, 'platform', { get: () => 'Win32' });
        """)

        # Open the homepage so the browser is visible and CF can be handled.
        print("[Init] Opening homepage...")
        try:
            await page.goto(BASE_URL, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        except Exception as e:
            print(f"[Init] Homepage load warning: {e}")

        if MANUAL_WAIT_SECS > 0:
            print(f"[Init] Browser is open — you have {MANUAL_WAIT_SECS}s to solve any")
            print(f"[Init] Cloudflare challenge, log in, or browse manually.")
            print(f"[Init] Scraping will start automatically when the timer ends.")
            for remaining in range(MANUAL_WAIT_SECS, 0, -5):
                print(f"[Init]   {remaining}s remaining...", flush=True)
                await asyncio.sleep(5)
            print("[Init] Timer done — starting scrape now.")
        else:
            await _wait_for_cf(page, "homepage")
            await page.wait_for_timeout(3000)
            print("[Init] CF session ready")

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

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
