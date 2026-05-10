"""
scraper.py  —  Curry & Company
---------------------------------
Platform: curreyandcompany.com (EPiServer CMS, fully JS-rendered)

APPROACH:
  Listing pages are 100% JS-rendered.
  We fetch the sitemap (https://www.curreyandcompany.com/sitemap.xml) once
  to get all product URLs, filter per category, then scrape each product
  page via Playwright with networkidle wait + content verification.

BROWSER MODE:
  Defaults to headless=FALSE so Cloudflare JS challenges complete fully
  and the EPiServer CMS can finish client-side rendering before extraction.
  Pass HEADLESS=true environment variable only on trusted IPs with CF bypass.

Run directly:
    python scraper.py
    HEADLESS=true python scraper.py

Or via orchestrator:
    python orchestrator.py "Curry"
    python orchestrator.py "Curry" --test --headless false
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urljoin, urlparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    async_polite_delay,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Curry")
# Default headless=False — Cloudflare + EPiServer requires full browser render
HEADLESS    = os.environ.get("HEADLESS", "false").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL    = "https://www.curreyandcompany.com"
SITEMAP_URL = "https://www.curreyandcompany.com/sitemap.xml"
TIMEOUT_MS  = 60_000   # 60s per page


# ---------------------------------------------------------------------------
# Sitemap — fetch all product URLs once
# ---------------------------------------------------------------------------

def fetch_all_product_urls() -> list[str]:
    """Download the sitemap and return every /c/{cat}/{subcat}/{sku}/ URL."""
    import requests

    print(f"[Sitemap] Fetching {SITEMAP_URL} …")
    try:
        resp = requests.get(SITEMAP_URL, timeout=30, headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            )
        })
        resp.raise_for_status()
    except Exception as e:
        print(f"[Sitemap] ERROR: {e}")
        return []

    xml_text = re.sub(r' xmlns="[^"]+"', '', resp.text)
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        print(f"[Sitemap] XML parse error: {e}")
        return []

    # Last segment must start with a digit — avoids slug URLs like /abu-gold-accent-table-4000-0118/
    product_re = re.compile(r"/c/[\w-]+/[\w-]+/\d[\d\w-]*/?$")
    urls = []
    for loc in root.iter("loc"):
        url = (loc.text or "").strip()
        path = urlparse(url).path
        if product_re.search(path):
            if not url.endswith("/"):
                url += "/"
            urls.append(url)

    print(f"[Sitemap] Found {len(urls)} product URLs total")
    return urls


def filter_urls_for_category(all_urls: list[str], listing_url: str) -> list[str]:
    """Return URLs whose path starts with the listing URL's path prefix."""
    parsed = urlparse(listing_url)
    prefix = parsed.path.rstrip("/") + "/"
    if prefix.count("/") < 3:
        return []
    return [u for u in all_urls if urlparse(u).path.startswith(prefix)]


# ---------------------------------------------------------------------------
# Browser — non-headless by default so CF + EPiServer JS renders fully
# ---------------------------------------------------------------------------

async def create_page(playwright, headless: bool):
    """
    Launch Chromium with stealth settings.
    headless=False (default) ensures Cloudflare passes and EPiServer renders.
    Media resources are blocked to keep load times fast.
    """
    args = [
        "--disable-blink-features=AutomationControlled",
        "--disable-dev-shm-usage",
        "--no-sandbox",
        "--disable-extensions",
        "--start-maximized",
    ]
    if headless:
        args += ["--disable-gpu", "--disable-background-networking"]

    browser = await playwright.chromium.launch(
        headless=headless,
        args=args,
        slow_mo=0,
    )
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1920, "height": 1080},
        locale="en-US",
        extra_http_headers={
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8"
            ),
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Cache-Control": "max-age=0",
            "Sec-Ch-Ua": '"Not A(Brand";v="99", "Google Chrome";v="124", "Chromium";v="124"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
        },
    )
    await context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'plugins', {
            get: () => [
                {name: 'Chrome PDF Plugin'},
                {name: 'Chrome PDF Viewer'},
                {name: 'Native Client'},
            ]
        });
        window.chrome = {runtime: {}, loadTimes: function(){}, csi: function(){}, app: {}};
        Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
        Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
        Object.defineProperty(screen, 'width', {get: () => 1920});
        Object.defineProperty(screen, 'height', {get: () => 1080});
    """)
    page = await context.new_page()

    # Block only media (video/audio) — keep images, fonts, stylesheets so page renders fully
    await page.route(
        "**/*",
        lambda route: (
            route.abort()
            if route.request.resource_type == "media"
            else route.continue_()
        ),
    )
    return browser, context, page


async def _is_cf_blocked(page) -> bool:
    """Return True if Cloudflare is serving a challenge/block page."""
    try:
        title = (await page.title()).lower()
        if any(k in title for k in ("just a moment", "attention required", "blocked", "cloudflare")):
            return True
        cf_el = await page.query_selector(
            "#cf-challenge-running, #challenge-running, .cf-browser-verification, "
            "#challenge-form, .ray-id"
        )
        return cf_el is not None
    except Exception:
        return False


async def safe_goto(page, url: str, retries: int = 3) -> bool:
    """
    Navigate with networkidle wait + Cloudflare detection + content verification.
    Returns True when product content is confirmed loaded.
    """
    from playwright.async_api import TimeoutError as PlaywrightTimeoutError

    for attempt in range(retries + 1):
        try:
            # Use networkidle so EPiServer finishes all client-side rendering
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="networkidle")
        except PlaywrightTimeoutError:
            # networkidle timed out — page may still have content; continue checking
            print(f"    [INFO] networkidle timeout (attempt {attempt + 1}), checking content …")
        except Exception as e:
            print(f"    [WARN] goto error (attempt {attempt + 1}): {e}")
            if attempt < retries:
                await asyncio.sleep(4)
            continue

        # Extra wait for any remaining lazy JS
        await page.wait_for_timeout(2500)

        if await _is_cf_blocked(page):
            print(f"    [CF] Challenge detected, waiting 12s …")
            await page.wait_for_timeout(12000)
            if await _is_cf_blocked(page):
                print(f"    [CF] Still blocked after wait: {url}")
                if attempt < retries:
                    await asyncio.sleep(6)
                    continue
                return False

        # Try to wait for actual product content
        try:
            await page.wait_for_selector(
                "h1, [class*='product'], [class*='catalog-entry'], [class*='pdp'], main",
                timeout=10_000,
            )
        except Exception:
            pass

        # Final extra wait for any remaining async renders
        await page.wait_for_timeout(1500)
        return True

    return False


# ---------------------------------------------------------------------------
# Spec field mapping
# ---------------------------------------------------------------------------

_SPEC_MAP: dict[str, str] = {
    "width":                     "Width",
    "depth":                     "Depth",
    "height":                    "Height",
    "diameter":                  "Diameter",
    "length":                    "Length",
    "dimensions":                "Dimensions",
    "overall dimensions":        "Dimensions",
    "overall dim":               "Dimensions",
    "overall":                   "Dimensions",   # accordion label for the dim row
    "item weight":               "Weight",
    "weight":                    "Weight",
    "material":                  "Materials",
    "materials":                 "Materials",
    "finish":                    "Finish",
    "finishes":                  "Finish",
    "collection":                "Collection",
    "product collection":        "Collection",
    "designer":                  "Designer",
    "origin":                    "Origin",
    "country of origin":         "Origin",
    "lead time":                 "Lead Time",
    "seat height":               "Seat Height",
    "seat depth":                "Seat Depth",
    "arm height":                "Arm Height",
    "arm length":                "Arm Length",
    "com":                       "COM",
    "col":                       "COL",
    "fabric":                    "Fabric",
    "wattage":                   "Wattage",
    "socket type":               "Socket",
    "socket":                    "Socket",
    "color temperature":         "Color Temperature",
    "colour temperature":        "Color Temperature",
    "shade details":             "Shade Details",
    "shade":                     "Shade Details",
    "canopy":                    "Canopy",
    "chain length":              "Chain Length",
    "cord length":               "Chain Length",
    "hanging length":            "Chain Length",
    "mounting":                  "Mounting",
    "number of lights":          "Lamp Quantity",
    "number of bulbs":           "Bulb Qty",
    "bulb type":                 "Bulb Type",
    "bulb wattage":              "Bulb Wattage",
    "lumens":                    "Lumens",
    "voltage":                   "Voltage",
    "style":                     "Style",
    "color":                     "Color",
    "shade color":               "Shade Color",
    "shade material":            "Shade Details",
    "max height":                "Height",
    "min height":                "Min Drop",
    "assembly":                  "Assembly Required",
    "assembly required":         "Assembly Required",
    "care instructions":         "Care Instructions",
    "maintenance & care":        "Care Instructions",
    "warranty":                  "Warranty",
    "hardware details":          "Hardware Details",
    "drawer details":            "Drawer Details",
    "floor protection":          "Floor Protection",
    "anti-tip safety hardware included": "Anti-Tip Hardware",
    "upholstery":                "Upholstery",
    "com yardage":               "COM Yardage",
    "com available":             "COM Available",
    "number of sockets":         "Socket Qty",
    "socket quantity":           "Socket Qty",
    "max wattage per socket":    "Wattage",
    "cord/chain":                "Chain Length",
    "hanging weight":            "Weight",
    "shade height":              "Shade Details",
    "shade diameter":            "Shade Details",
    "shade width":               "Shade Details",
}


# ---------------------------------------------------------------------------
# Product detail page
# ---------------------------------------------------------------------------

async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Currey & Company product detail page."""
    base: dict = {"Source": url}

    ok = await safe_goto(page, url)
    if not ok:
        print(f"    [WARN] Skipping — navigation failed: {url}")
        return [base]

    # ── 1. Product Name — breadcrumb JSON (most reliable, no h1 on this site)
    try:
        bc_el = await page.query_selector("#breadcrumbs-data")
        if bc_el:
            bc_json = await bc_el.get_attribute("data-breadcrumbs")
            if bc_json:
                crumbs = json.loads(bc_json)
                if crumbs:
                    base["Product Name"] = clean_text(crumbs[-1].get("linkText", ""))
    except Exception:
        pass

    # ── 2. SKU — last URL path segment (e.g. /9000-0022/ → 9000-0022)
    if not base.get("SKU"):
        url_parts = [p for p in url.rstrip("/").split("/") if p]
        if url_parts:
            candidate = url_parts[-1]
            # Match Currey SKU patterns: 4204, 9000-0022, 3000-0184, etc.
            if re.match(r"^\d{3,6}(-[\w]+)*$", candidate):
                base["SKU"] = candidate

    # ── 3. Image URL — first main gallery image (alt='product image')
    #    The gallery images in #overviewSection have alt='product image'
    #    Thumbnails have alt='product image thumbnail' — exclude those
    if not base.get("Image URL"):
        try:
            img_el = await page.query_selector(
                "#overviewSection img[alt='product image'], "
                ".image-gallery img[alt='product image']"
            )
            if not img_el:
                img_el = await page.query_selector("img[alt='product image']")
            if img_el:
                src = (await img_el.get_attribute("src") or "").strip()
                if src and "placeholder" not in src.lower():
                    base["Image URL"] = src if src.startswith("http") else urljoin(BASE_URL, src)
        except Exception:
            pass

    # ── 4. Price — span.paragraph-2a-reg containing '$'
    if not base.get("Price"):
        try:
            price_text = await page.evaluate("""() => {
                let spans = document.querySelectorAll('span.paragraph-2a-reg');
                for (let s of spans) {
                    let t = s.textContent.trim();
                    if (t.indexOf('$') !== -1) return t;
                }
                // fallback: any element with just a $ amount
                let all = document.querySelectorAll('*');
                for (let el of all) {
                    if (el.children.length === 0) {
                        let t = el.textContent.trim();
                        if (/^\\$[\\d,]+(\\.\\d{2})?$/.test(t)) return t;
                    }
                }
                return '';
            }""")
            if price_text:
                base["Price"] = clean_price(price_text)
        except Exception:
            pass

    # ── 5. Specs — MUI Accordion panels inside #specificationsSection
    #    Each accordion collapse panel has alternating label/value lines
    spec_dict: dict[str, str] = {}
    try:
        spec_text = await page.evaluate("""() => {
            const section = document.getElementById('specificationsSection');
            if (!section) return '';
            // Get all MUI collapse panels — each is one spec accordion section
            const collapses = section.querySelectorAll('[class*="MuiCollapse-root"]');
            let parts = [];
            for (let c of collapses) {
                parts.push(c.innerText.trim());
            }
            return parts.join('\\n---\\n');
        }""")
        if spec_text:
            for block in spec_text.split("---"):
                lines = [l.strip() for l in block.split("\n") if l.strip()]
                # alternating label / value pairs
                i = 0
                while i + 1 < len(lines):
                    k = lines[i]
                    v = lines[i + 1]
                    # skip if value looks like another label (all caps or too short)
                    if k and v and len(k) < 60 and len(v) < 200:
                        spec_dict[k] = v
                    i += 2
    except Exception:
        pass

    # Fallback: parse body text within #specificationsSection
    if not spec_dict:
        try:
            spec_body = await page.evaluate("""() => {
                const s = document.getElementById('specificationsSection');
                return s ? s.innerText : '';
            }""")
            if spec_body:
                for line in spec_body.split("\n"):
                    line = line.strip()
                    m = re.match(r'^([A-Za-z][A-Za-z0-9 /&-]+?)\s*:\s*(.+)$', line)
                    if m and len(m.group(1)) < 50:
                        spec_dict[m.group(1).strip()] = m.group(2).strip()
        except Exception:
            pass

    # Map spec keys → canonical field names
    _SECTION_HEADERS = {
        "dimensions", "furniture specifications", "lighting specifications",
        "seating specifications", "additional details", "certifications & ratings",
        "maintenance & care", "shipping specifications", "resources",
    }
    for raw_k, raw_v in spec_dict.items():
        k_norm = raw_k.lower().strip()
        if k_norm in _SECTION_HEADERS:
            continue

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
        elif canon == "Weight":
            num = re.sub(r'[^\d.]', '', raw_v.split()[0]).strip()
            base.setdefault(canon, num or raw_v)
        else:
            base.setdefault(canon, raw_v)

    # ── 6. Description — .account-paragraph-s in the spec section
    if not base.get("Description"):
        try:
            desc_el = await page.query_selector(
                "#specificationsSection .account-paragraph-s, "
                "#overviewSection .account-paragraph-s"
            )
            if desc_el:
                text = clean_text(await desc_el.inner_text())
                if text and len(text) > 10:
                    base["Description"] = text
        except Exception:
            pass

    # ── 7. Tearsheet — look in #resourcesSection or any PDF link
    try:
        tearsheet = await page.evaluate("""() => {
            const res = document.getElementById('resourcesSection');
            const container = res || document;
            const links = container.querySelectorAll('a[href]');
            for (let a of links) {
                let href = a.href || '';
                let text = a.textContent.toLowerCase();
                if (text.includes('tearsheet') || href.includes('tearsheet')) return href;
            }
            for (let a of links) {
                let href = a.href || '';
                if (href.endsWith('.pdf') || href.includes('.pdf')) return href;
            }
            return '';
        }""")
        if tearsheet:
            base["Tearsheet Link"] = tearsheet if tearsheet.startswith("http") else urljoin(BASE_URL, tearsheet)
    except Exception:
        pass

    # ── 8. Product Family Id — derive from product name, not SKU
    if not base.get("Product Family Id"):
        if base.get("Product Name"):
            base["Product Family Id"] = extract_family_id(base["Product Name"])
        elif base.get("SKU"):
            base["Product Family Id"] = base["SKU"]

    # Debug: warn if core fields still missing
    missing = [f for f in ("Product Name", "Image URL", "Price") if not base.get(f)]
    if missing:
        print(f"    [DBG] Missing: {missing} — {url}")

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
    print(f"[Scraper] Headless: {HEADLESS}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")

    all_sitemap_urls = fetch_all_product_urls()

    from playwright.async_api import async_playwright
    async with async_playwright() as pw:
        browser, context, page = await create_page(pw, headless=HEADLESS)
        try:
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
                    for u in filter_urls_for_category(all_sitemap_urls, listing_url):
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
                        short = url.rstrip("/").split("/")[-1]
                        print(f"  [{idx}] {short}")
                    except Exception as e:
                        print(f"  [ERROR] {url}: {e}")
                    await async_polite_delay(1.0, 2.0)

                await async_polite_delay(1.5, 3.0)
        finally:
            await context.close()
            await browser.close()

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
