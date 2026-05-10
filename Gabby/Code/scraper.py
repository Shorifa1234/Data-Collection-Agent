"""
scraper.py  —  Gabby
-----------------------
Platform: gabriellawhite.com (Shopify) — accessed via gabby.com redirect

REDIRECT: gabby.com/products/{cat}/{subcat}  →  301  →  gabriellawhite.com/collections/{collection}
Playwright follows redirects automatically; the scraper uses the final URL for
product discovery and links.

Site structure:
  Listing  : /collections/{collection} — all products in one page (no pagination needed)
             (Playwright follows redirect from gabby.com links in vendor_info.json)
  Product  : /products/{slug}
  JSON-LD  : Present on product pages as Product schema

Fields collected:
  Product Name, SKU (SCH-XXXXX), Price (list price), Image URL (2048px),
  Description, Dimensions, Width, Depth, Height, Diameter, Weight,
  Materials, Finish, Collection, Designer, Seat Height, Seat Depth,
  Arm Height, COM, COL, Shade Details, Base, Canopy, Lamping, Wattage,
  Color, Pattern, Tearsheet Link

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Gabby"
    python orchestrator.py "Gabby" --test
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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Gabby")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://gabriellawhite.com"
TIMEOUT_MS = 45_000


# ---------------------------------------------------------------------------
# Listing page
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str, max_products: int = 0) -> list[str]:
    """
    Collect all product URLs from a Gabriella White / Gabby Shopify collection.
    gabby.com URLs redirect automatically to gabriellawhite.com.
    Most collections show all products on a single page.
    Handles /page/N pagination as fallback for large collections.
    """
    links: list[str] = []
    seen:  set[str]  = set()
    page_num = 1

    while True:
        url = listing_url if page_num == 1 else f"{listing_url}?page={page_num}"
        print(f"  [Listing p{page_num}] {url}")

        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
        except Exception as e:
            print(f"  [WARN] {e}")
            break

        # Wait for Shopify product grid
        try:
            await page.wait_for_selector(
                "a[href*='/products/'], .product-card a, .product-item a",
                timeout=8000,
            )
        except Exception:
            pass

        # Extract all /products/{slug} links
        hrefs: list[str] = await page.evaluate(
            """() => {
                const seen = new Set();
                const out  = [];
                document.querySelectorAll("a[href*='/products/']").forEach(a => {
                    const h = a.href.split('?')[0].split('#')[0];
                    if (h.includes('/products/') && !seen.has(h)) {
                        seen.add(h);
                        out.push(h);
                    }
                });
                return out;
            }"""
        )

        # Filter out collection-level links like /collections/.../products/...
        # and keep only canonical /products/{slug} paths
        new_count = 0
        for h in hrefs:
            # Normalise: strip collection prefix if present
            canonical = re.sub(r"/collections/[^/]+(/products/)", r"\1", h)
            if not canonical.startswith("http"):
                canonical = urljoin(BASE_URL, canonical)
            if canonical not in seen:
                seen.add(canonical)
                links.append(canonical)
                new_count += 1

        print(f"  [Listing p{page_num}] {new_count} new URLs ({len(links)} total)")

        if new_count == 0 or (max_products and len(links) >= max_products):
            break

        # Check for next page link
        next_el = await page.query_selector(
            'a[href*="?page="], .pagination__next, [aria-label*="Next"]'
        )
        if not next_el:
            break

        page_num += 1
        await async_polite_delay(0.8, 1.5)

    return links[:max_products] if max_products else links


# ---------------------------------------------------------------------------
# Product detail
# ---------------------------------------------------------------------------

def _best_shopify_image(url: str) -> str:
    """Upgrade Shopify CDN image URL to maximum available resolution."""
    if not url:
        return ""
    url = url.split("?")[0]   # strip existing params
    if url.startswith("//"):
        url = "https:" + url
    # Shopify supports ?width=2048 for high-res
    return url + "?width=2048"


def _extract_jsonld(scripts: list[str]) -> dict:
    for raw in scripts:
        try:
            obj = json.loads(raw)
            if isinstance(obj, list):
                for item in obj:
                    if isinstance(item, dict) and item.get("@type") in ("Product", "ProductGroup"):
                        return item
            elif isinstance(obj, dict) and obj.get("@type") in ("Product", "ProductGroup"):
                return obj
        except Exception:
            pass
    return {}


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape a Gabriella White product page.
    Returns one or more dicts — one per size/finish variant if detected.
    """
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [base]

    # ── 1. JSON-LD ─────────────────────────────────────────────────────────
    script_texts = await page.eval_on_selector_all(
        'script[type="application/ld+json"]', "els => els.map(e => e.innerText)"
    )
    ld = _extract_jsonld(script_texts)

    if ld:
        base["Product Name"] = clean_text(ld.get("name", ""))
        if ld.get("sku"):
            base["SKU"] = clean_text(ld["sku"])
        if ld.get("description"):
            base["Description"] = clean_text(re.sub(r"<[^>]+>", " ", ld["description"]))
        # Image
        img = ld.get("image", "")
        if isinstance(img, list):
            img = img[0] if img else ""
        if isinstance(img, dict):
            img = img.get("url", img.get("contentUrl", ""))
        if img:
            base["Image URL"] = _best_shopify_image(str(img))
        # Price — use list price (higher value) if available
        offers = ld.get("offers", {})
        if isinstance(offers, list):
            prices = [clean_price(str(o.get("price", ""))) for o in offers if o.get("price")]
            prices = [p for p in prices if p]
            if prices:
                base["Price"] = max(prices)  # list price is higher
        elif isinstance(offers, dict) and offers.get("price"):
            base["Price"] = clean_price(str(offers["price"]))

    # ── 2. Product Name fallback ───────────────────────────────────────────
    if not base.get("Product Name"):
        for sel in ["h1.product__title", "h1.product-single__title", "h1.product-title", "h1"]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                if text:
                    base["Product Name"] = text
                    break

    # ── 3. SKU ────────────────────────────────────────────────────────────
    if not base.get("SKU"):
        body_text = clean_text(await page.evaluate("() => document.body.innerText"))
        sku_m = re.search(r"SKU[:\s]+([A-Z0-9\-]+)", body_text, re.IGNORECASE)
        if sku_m:
            base["SKU"] = sku_m.group(1).strip()

    # ── 4. Price fallback ─────────────────────────────────────────────────
    if not base.get("Price"):
        body_text = clean_text(await page.evaluate("() => document.body.innerText"))
        # Prefer list price: "$1,939 List Price"
        lp_m = re.search(r"\$([\d,]+(?:\.\d{2})?)\s+List\s+Price", body_text, re.IGNORECASE)
        if lp_m:
            base["Price"] = clean_price(lp_m.group(1))
        else:
            for sel in [".product__price", ".product-single__price", ".price", "[class*='price']"]:
                el = await page.query_selector(sel)
                if el:
                    p = clean_price(clean_text(await el.inner_text()))
                    if p:
                        base["Price"] = p
                        break

    # ── 5. Image fallback ─────────────────────────────────────────────────
    if not base.get("Image URL"):
        for sel in [
            ".product__media img",
            ".product-single__photo img",
            ".product-featured-img",
            "img.product-image",
            "img[src*='cdn.shopify']",
        ]:
            img_el = await page.query_selector(sel)
            if img_el:
                src = (
                    await img_el.get_attribute("data-src")
                    or await img_el.get_attribute("data-zoom-image")
                    or await img_el.get_attribute("src")
                    or ""
                )
                if src:
                    base["Image URL"] = _best_shopify_image(src)
                    break

    # ── 6. Spec table ─────────────────────────────────────────────────────
    spec_dict: dict[str, str] = {}

    # Shopify product tables typically render as <table> with <td> pairs
    rows = await page.query_selector_all("table tr, .product-specs tr, .specifications tr")
    for tr in rows:
        tds = await tr.query_selector_all("td, th")
        if len(tds) >= 2:
            k = clean_text(await tds[0].inner_text()).rstrip(":")
            v = clean_text(await tds[1].inner_text())
            if k and v:
                spec_dict[k] = v

    # Fallback: scan page text for "Label Value" or "Label: Value" patterns
    if not spec_dict:
        body = clean_text(await page.evaluate("() => document.body.innerText"))
        for line in body.split("\n"):
            m = re.match(r'^([A-Za-z][A-Za-z0-9 /]+?)\s*[:\s]\s*(.+)$', line.strip())
            if m:
                k = m.group(1).strip()
                v = m.group(2).strip()
                if k and v and len(k) < 40:
                    spec_dict[k] = v

    # Map spec_dict to fields
    _SPEC_MAP = {
        "width":              "Width",
        "product width":      "Width",
        "depth":              "Depth",
        "product depth":      "Depth",
        "height":             "Height",
        "product height":     "Height",
        "diameter":           "Diameter",
        "length":             "Length",
        "weight":             "Weight",
        "product weight":     "Weight",
        "material":           "Materials",
        "materials":          "Materials",
        "finish":             "Finish",
        "finish family":      "Finish",
        "color":              "Color",
        "colour":             "Color",
        "collection":         "Collection",
        "designer":           "Designer",
        "origin":             "Origin",
        "country of origin":  "Origin",
        "lead time":          "Lead Time",
        "seat height":        "Seat Height",
        "seat depth":         "Seat Depth",
        "arm height":         "Arm Height",
        "com":                "COM",
        "col":                "COL",
        "fabric":             "Fabric",
        "shade":              "Shade Details",
        "shade details":      "Shade Details",
        "base":               "Base",
        "canopy":             "Canopy",
        "lamping":            "Lamping",
        "lamp type":          "Lamping",
        "wattage":            "Wattage",
        "pattern":            "Pattern",
        "number of drawers":  "Components",
        "assembly required":  "Assembly Required",
    }
    for raw_k, raw_v in spec_dict.items():
        k_norm = raw_k.lower().strip()
        canon = _SPEC_MAP.get(k_norm)
        if canon is None:
            for alias, c in _SPEC_MAP.items():
                if k_norm.startswith(alias):
                    canon = c
                    break
        if canon is None:
            canon = raw_k.strip().title()

        if canon == "Dimensions":
            parsed = parse_dimensions(raw_v)
            for dk, dv in parsed.items():
                base.setdefault(dk, dv)
        elif canon in ("Width", "Depth", "Height", "Diameter", "Length"):
            num = re.sub(r'["\']|in\b', '', raw_v.split()[0]).strip()
            base.setdefault(canon, num or raw_v)
        else:
            base.setdefault(canon, raw_v)

    # ── 7. Description fallback ───────────────────────────────────────────
    if not base.get("Description"):
        for sel in [
            ".product__description",
            ".product-single__description",
            ".product-description",
            "[class*='description']",
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(re.sub(r"<[^>]+>", " ", await el.inner_html()))
                if text and len(text) > 10:
                    base["Description"] = text
                    break

    # ── 8. Tearsheet ──────────────────────────────────────────────────────
    for sel in ["a[href*='tearsheet']", "a[href*='spec-sheet']", "a[href*='.pdf']"]:
        el = await page.query_selector(sel)
        if el:
            href = await el.get_attribute("href") or ""
            if href:
                base["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)
                break

    # ── 9. Product Family Id ──────────────────────────────────────────────
    if not base.get("Product Family Id") and base.get("SKU"):
        # SCH-175831 → SCH-175831 (no variant suffix on Gabby SKUs usually)
        base["Product Family Id"] = re.sub(r"-\d+$", "", base["SKU"])
    if not base.get("Product Family Id") and base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    return [base]


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
    print(f"[NOTE] gabby.com URLs redirect to gabriellawhite.com — handled automatically")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str] = set()
            all_urls:  list[str] = []

            for listing_url in cat["links"]:
                max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0
                for u in await get_product_links(page, listing_url, max_products=max_p):
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
                    print(f"  [{idx}] {url.split('/')[-1] or url.split('/')[-2]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
