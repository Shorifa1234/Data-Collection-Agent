"""
scraper.py — Alfonso Marina
-----------------------------
Platform: alfonsomarina.com (WooCommerce / WordPress — JS-rendered listing pages)

Site structure:
  Listing  : /product-category/{path}/  — JS-rendered; Playwright required
             Product cards with <a href="/product/{slug}/"> links
             Pagination: /page/N/ (WordPress style)
  Product  : /product/{slug}/

Product page fields (trade-only — no price shown without login):
  Product Name   : <h1 class="product_title">
  SKU            : extracted from product image filename (e.g. 509-318-02_NAME.webp → 509-318-02)
  Dimensions CM  : W: xx.x | D: xx.x | H: xx.x in DIMENSIONS CM section
  Dimensions IN  : W: xx | D: xx | H: xx in DIMENSIONS IN section (stored as primary)
  Finish Options : from dropdown select (04 DARK WALNUT, etc.)
  Materials      : from DETAILS / MATERIALS section
  Lead Time      : from specs (Regular Delivery / Made to Order)
  Character      : finish character label
  Description    : product description text
  Image URL      : wp-content/uploads high-res image (webp preferred)
  Tearsheet      : TEAR SHEET link if present

Run directly:
    python scraper.py
Or via orchestrator:
    python orchestrator.py "Alfonso Marina"
    python orchestrator.py "Alfonso Marina" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    async_polite_delay,
    clean_text,
    sentence_case,
    generate_sku,
    extract_family_id,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Alfonso Marina")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://alfonsomarina.com"

_REQ_SESSION = requests.Session()
_REQ_SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
})


# ---------------------------------------------------------------------------
# Browser helpers — custom context so fonts load (needed for CF challenge)
# ---------------------------------------------------------------------------

async def create_page(playwright, headless: bool):
    """
    Launch Chromium with stealth settings.
    Only media is blocked — fonts/images/stylesheets are allowed so that
    Cloudflare JS challenges can complete on VPS/datacenter IPs.
    """
    browser = await playwright.chromium.launch(
        headless=headless,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-extensions",
            "--disable-background-networking",
        ],
    )
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
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
            "Sec-Ch-Ua": '"Not A(Brand";v="99", "Google Chrome";v="122", "Chromium";v="122"',
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
    """)
    page = await context.new_page()
    await page.route(
        "**/*",
        lambda route: (
            route.abort()
            if route.request.resource_type in {"media"}
            else route.continue_()
        ),
    )
    return browser, context, page


async def _is_cf_blocked(page) -> bool:
    """Return True if Cloudflare is serving a challenge/block page."""
    title = (await page.title()).lower()
    if "just a moment" in title or "attention required" in title or "blocked" in title:
        return True
    cf_el = await page.query_selector(
        "#cf-challenge-running, #challenge-running, .cf-browser-verification"
    )
    return cf_el is not None


async def safe_goto(page, url: str, timeout: int = 90_000, retries: int = 2) -> bool:
    """
    Navigate to a URL with Cloudflare detection and retries.
    Returns True on success, False if blocked/failed.
    """
    for attempt in range(retries + 1):
        try:
            await page.goto(url, timeout=timeout, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
        except Exception as e:
            print(f"  [WARN] goto failed (attempt {attempt + 1}): {e}")
            if attempt < retries:
                await asyncio.sleep(3)
            continue

        if await _is_cf_blocked(page):
            print(f"  [INFO] Cloudflare challenge detected, waiting 8s …")
            await page.wait_for_timeout(8000)
            if await _is_cf_blocked(page):
                print(f"  [WARN] Still blocked after wait: {url}")
                if attempt < retries:
                    await asyncio.sleep(5)
                    continue
                return False

        return True

    return False


# ---------------------------------------------------------------------------
# Listing page — Playwright (JS-rendered WooCommerce)
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect all product URLs from an Alfonso Marina category listing page.
    WooCommerce — JS-rendered; /product/{slug}/ links.
    Pagination uses WordPress /page/N/ suffix.
    """
    links: list[str] = []
    seen: set[str] = set()
    # Build base URL without trailing slash for page construction
    base_listing = listing_url.rstrip("/")
    page_num = 1

    while True:
        url = base_listing + "/" if page_num == 1 else f"{base_listing}/page/{page_num}/"
        # Alfonso Marina server is slow — retry up to 3 times with 90 second timeout
        if not await safe_goto(page, url, timeout=90_000):
            print(f"  [WARN] Skipping listing page: {url}")
            break

        hrefs: list[str] = await page.evaluate("""() => {
            const out = [];
            document.querySelectorAll("a[href]").forEach(a => {
                const h = a.href;
                if (h.includes("/product/") && !h.includes("/product-category/")) {
                    out.push(h.split("?")[0]);
                }
            });
            return out;
        }""")

        if not hrefs:
            break

        found_new = False
        for h in hrefs:
            # normalise trailing slash
            canonical = h.rstrip("/") + "/"
            if canonical not in seen:
                seen.add(canonical)
                links.append(canonical)
                found_new = True

        # Check for next-page link
        next_el = await page.query_selector("a.next.page-numbers")
        if not next_el or not found_new:
            break

        page_num += 1
        await async_polite_delay(0.5, 1.0)

    return links


# ---------------------------------------------------------------------------
# Product detail — requests + BeautifulSoup (Alfonso Marina pages work without JS)
# ---------------------------------------------------------------------------

def _extract_sku_from_image(soup: BeautifulSoup) -> str:
    """
    Alfonso Marina does not show an explicit SKU element.
    The product reference is embedded in the image filename.
    Pattern: /uploads/.../509-318-02_PRODUCT_NAME.webp → "509-318-02"
    """
    for img in soup.find_all("img"):
        src = img.get("src", img.get("data-src", ""))
        m = re.search(r"/(\d{3}-\d{3}-\d{2,3})[-_]", src)
        if m:
            return m.group(1)
    return ""


def _parse_dimensions_block(text: str) -> dict:
    """
    Parse Alfonso Marina dimension blocks of the form:
      W: 55.0  D: 45.0  H: 70.0  or  W: 21-3/4  D: 17-3/4  H: 27-1/2
    Returns dict with Width, Height, Depth and a composed Dimensions string.
    """
    result: dict = {}
    mapping = {"W": "Width", "D": "Depth", "H": "Height", "L": "Length", "DIA": "Diameter"}
    # Pattern matches: "21-3/4" (mixed fraction), "55.0" (decimal), "21" (integer)
    dim_pat = r"(\d+(?:-\d+/\d+|\.\d+)?)"
    for abbr, field in mapping.items():
        m = re.search(rf"\b{abbr}\s*:\s*{dim_pat}", text, re.I)
        if m:
            val = m.group(1).strip()
            # Convert mixed fractions like "21-3/4" → "21.75"
            frac_m = re.match(r"^(\d+)-(\d+)/(\d+)$", val)
            if frac_m:
                whole = float(frac_m.group(1))
                num   = float(frac_m.group(2))
                den   = float(frac_m.group(3))
                val = str(round(whole + num / den, 2))
            result[field] = val

    # Compose Dimensions string
    parts = []
    for abbr, field in [("W", "Width"), ("D", "Depth"), ("H", "Height"), ("L", "Length"), ("DIA", "Diameter")]:
        if field in result:
            parts.append(f"{result[field]}{abbr}")
    if parts:
        result["Dimensions"] = " x ".join(parts)

    return result


def scrape_product(url: str) -> list[dict]:
    """
    Scrape an Alfonso Marina product page via requests.
    Returns one dict (no purchasable variants — all finishes are on the same product page).
    """
    row: dict = {"Source": url}

    try:
        resp = _REQ_SESSION.get(url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"    [WARN] {url}: {e}")
        return [row]

    soup = BeautifulSoup(resp.text, "html.parser")
    page_text = soup.get_text(separator="\n")

    # ── 1. Product Name ───────────────────────────────────────────────────
    h1 = soup.find("h1", class_="product_title") or soup.find("h1")
    if h1:
        row["Product Name"] = sentence_case(clean_text(h1.get_text()))

    # ── 2. SKU (from image filename) ─────────────────────────────────────
    sku = _extract_sku_from_image(soup)
    if sku:
        row["SKU"] = sku

    # ── 3. Image URL ─────────────────────────────────────────────────────
    # Alfonso Marina uses WordPress media uploads; prefer the highest-res image
    for img in soup.find_all("img"):
        src = img.get("src", img.get("data-src", ""))
        if "wp-content/uploads" in src and ("_2560" in src or "_1280" in src or ".webp" in src):
            row["Image URL"] = src
            break
    if not row.get("Image URL"):
        for img in soup.find_all("img"):
            src = img.get("src", img.get("data-src", ""))
            if "wp-content/uploads" in src and src.lower().endswith((".jpg", ".jpeg", ".webp", ".png")):
                row["Image URL"] = src
                break

    # ── 4. Dimensions ────────────────────────────────────────────────────
    # Alfonso Marina shows two sections: "DIMENSIONS IN :" and "DIMENSIONS CM :"
    # Page text has tabs/spaces around values — use simple string find + substring search.
    def _extract_dim_area(text: str, label: str) -> str:
        idx = text.upper().find(label.upper())
        return text[idx:idx + 600] if idx >= 0 else ""

    dim_in_area = _extract_dim_area(page_text, "DIMENSIONS IN")
    dim_cm_area = _extract_dim_area(page_text, "DIMENSIONS CM")

    # Prefer inches
    if dim_in_area:
        parsed = _parse_dimensions_block(dim_in_area)
        row.update(parsed)
    elif dim_cm_area:
        parsed = _parse_dimensions_block(dim_cm_area)
        row.update(parsed)

    # Store CM dimensions as extra column if inches were used
    if dim_in_area and dim_cm_area:
        parsed_cm = _parse_dimensions_block(dim_cm_area)
        if "Dimensions" in parsed_cm:
            row["Dimensions CM"] = parsed_cm["Dimensions"]

    # ── 5. Finish options ─────────────────────────────────────────────────
    # Finishes are in a <select> dropdown
    finish_select = soup.find("select", id=re.compile(r"finish|color", re.I))
    if not finish_select:
        finish_select = soup.find("select")  # fallback to any select
    if finish_select:
        options = [
            clean_text(opt.get_text())
            for opt in finish_select.find_all("option")
            if opt.get_text(strip=True) and "choose" not in opt.get_text(strip=True).lower()
        ]
        if options:
            row["Finish"] = " | ".join(options)

    # Also check for visible finish labels (span/div with finish names)
    if not row.get("Finish"):
        finish_area = soup.find(class_=re.compile(r"finish|swatch", re.I))
        if finish_area:
            labels = [clean_text(s.get_text()) for s in finish_area.find_all(["span", "label", "li"])]
            labels = [l for l in labels if l and "choose" not in l.lower()]
            if labels:
                row["Finish"] = " | ".join(labels)

    # ── 6. Materials ─────────────────────────────────────────────────────
    # In DETAILS section: "MATERIALS : Wood" (colon always present in product data)
    # Nav menu has "MATERIALS\n" with no colon — require colon to avoid false matches
    mat_m = re.search(r"MATERIALS\s*:\s*([^\n\t]+)", page_text, re.I)
    if mat_m:
        mat_val = clean_text(mat_m.group(1))
        if mat_val and len(mat_val) < 150:
            row["Materials"] = mat_val

    # ── 7. Lead Time ─────────────────────────────────────────────────────
    lead_m = re.search(r"Lead\s+Time\s*[:\s]+([^\n]+)", page_text, re.I)
    if lead_m:
        row["Lead Time"] = clean_text(lead_m.group(1))

    # ── 8. Character / style ─────────────────────────────────────────────
    char_m = re.search(r"Character\s*[:\s]+([^\n]+)", page_text, re.I)
    if char_m:
        row["Character"] = clean_text(char_m.group(1))

    # ── 9. Description ───────────────────────────────────────────────────
    # Alfonso Marina: description is in the DETAILS section after MATERIALS.
    # The DETAILS section ends at "RELATED PRODUCTS" or "NOTES:" or the footer nav.
    details_m = re.search(
        r"DETAILS\s*\n(.*?)(?=\nRELATED PRODUCTS|\nNOTES:|\nBE THE FIRST|\nPRODUCTS\s*\n|$)",
        page_text, re.DOTALL | re.I
    )
    if details_m:
        desc_raw = details_m.group(1).strip()
        # Remove MATERIALS line from start
        desc_raw = re.sub(r"^MATERIALS\s*[:\s]+[^\n]+\n?", "", desc_raw, flags=re.I | re.M)
        # Collapse excess whitespace
        desc_raw = re.sub(r"\n{3,}", "\n\n", desc_raw)
        desc = clean_text(desc_raw)
        if desc and len(desc) > 20:
            row["Description"] = desc[:1500]

    if not row.get("Description"):
        # Fallback: woocommerce short description element
        desc_el = soup.find("div", class_=re.compile(r"short.description|entry.summary", re.I))
        if desc_el:
            txt = clean_text(desc_el.get_text())
            if txt and len(txt) > 20:
                row["Description"] = txt[:1500]

    # ── 10. Tearsheet ────────────────────────────────────────────────────
    ts_el = soup.find("a", string=re.compile(r"tear\s*sheet|tearsheet", re.I))
    if not ts_el:
        ts_el = soup.find("a", href=re.compile(r"tearsheet|tear.sheet", re.I))
    if ts_el and ts_el.get("href"):
        href = ts_el["href"]
        row["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)

    # ── 11. Product Family Id ─────────────────────────────────────────────
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

    from playwright.async_api import async_playwright
    async with async_playwright() as pw:
        browser, context, playwright_page = await create_page(pw, headless=HEADLESS)
        try:
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
                    for u in await get_product_links(playwright_page, listing_url):
                        if u not in seen_urls:
                            seen_urls.add(u)
                            all_urls.append(u)

                if TEST_MODE:
                    all_urls = all_urls[:TEST_MAX_PRODUCTS]
                print(f"  {len(all_urls)} products")

                global_idx = 1
                for url in all_urls:
                    try:
                        rows = scrape_product(url)
                        for row in rows:
                            if not row.get("SKU"):
                                row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                            if not row.get("Product Family Id") and row.get("Product Name"):
                                row["Product Family Id"] = extract_family_id(row["Product Name"])
                            writer.write_row(row, category_name=cat["name"])
                            global_idx += 1
                        slug = url.rstrip("/").split("/")[-1]
                        print(f"  [{global_idx - 1}] {slug}")
                    except Exception as e:
                        print(f"  [ERROR] {url}: {e}")
                    await async_polite_delay(0.6, 1.5)
        finally:
            await context.close()
            await browser.close()

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
