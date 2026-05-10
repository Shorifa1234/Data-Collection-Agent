"""
scraper.py  —  Four Hands
---------------------------
Platform: fourhands.com (React SPA, trade-only login required)

AUTHENTICATION:
  The Four Hands wholesale portal requires a trade account.
  Set credentials in environment variables before running:

      set FH_EMAIL=your@email.com
      set FH_PASSWORD=yourpassword

  Or pass them inline:
      FH_EMAIL=you@email.com FH_PASSWORD=pass python scraper.py

STRATEGY:
  Phase 1 — Login via Playwright form submission.
  Phase 2 — Listing pages: wait for JS-rendered product cards, collect URLs.
             Pagination: "Load More" button or ?page=N query param.
  Phase 3 — Product detail pages: extract all fields via JSON-LD + CSS selectors.
             Tearsheet constructed from /product/{slug}/tearsheet/pdf.

FIELDS COLLECTED (all available):
  Product Name, SKU, Price, Image URL, Source,
  Description, Dimensions, Width, Depth, Height, Diameter, Weight,
  Materials, Finish, Collection, Designer, Origin, Lead Time,
  Seat Height, Seat Depth, Arm Height, COM, COL,
  Wattage, Socket, Lamping, Tearsheet Link

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Four Hands"
    python orchestrator.py "Four Hands" --test
    python orchestrator.py "Four Hands" --headless false
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
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
    parse_spec_block,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Four Hands")
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

# ── Credentials (required) ──────────────────────────────────────────────────
FH_EMAIL    = os.environ.get("FH_EMAIL", "")
FH_PASSWORD = os.environ.get("FH_PASSWORD", "")

BASE_URL   = "https://fourhands.com"
LOGIN_URL  = f"{BASE_URL}/login"
TIMEOUT_MS = 45_000
MAX_PAGES  = 200   # safety cap for pagination


# ---------------------------------------------------------------------------
# Browser helpers
# ---------------------------------------------------------------------------

async def create_page(playwright, headless: bool):
    """
    Launch Chromium with anti-detection headers.
    Does NOT block fonts or images so login forms and JS render correctly.
    """
    browser = await playwright.chromium.launch(
        headless=headless,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-extensions",
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
                "q=0.9,image/avif,image/webp,*/*;q=0.8"
            ),
            "Accept-Language": "en-US,en;q=0.9",
            "Sec-Ch-Ua": '"Not A(Brand";v="99", "Google Chrome";v="122", "Chromium";v="122"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Upgrade-Insecure-Requests": "1",
        },
    )
    await context.add_init_script(
        "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
        "window.chrome={runtime:{}};"
        "Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']});"
    )
    page = await context.new_page()
    # Only block media (video/audio) — allow fonts & images for proper rendering
    await page.route(
        "**/*",
        lambda route: (
            route.abort()
            if route.request.resource_type in {"media"}
            else route.continue_()
        ),
    )
    return browser, context, page


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

async def login(page, email: str, password: str) -> bool:
    """
    Log in to the Four Hands trade portal.
    Returns True on success, False on failure.
    """
    if not email or not password:
        print(
            "\n[ERROR] FH_EMAIL and FH_PASSWORD environment variables are required.\n"
            "  Set them before running:\n"
            "      set FH_EMAIL=your@email.com\n"
            "      set FH_PASSWORD=yourpassword\n"
        )
        return False

    print(f"[Login] Navigating to {LOGIN_URL} …")
    try:
        await page.goto(LOGIN_URL, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
    except Exception as e:
        print(f"[Login] Failed to load login page: {e}")
        return False

    # ── Fill email ────────────────────────────────────────────────────────
    email_filled = False
    for sel in [
        'input[name="email"]',
        'input[type="email"]',
        '#email',
        '#customer_email',
        '[placeholder*="email" i]',
        '[data-testid*="email" i]',
    ]:
        el = await page.query_selector(sel)
        if el:
            await el.fill(email)
            email_filled = True
            print(f"[Login] Email filled via: {sel}")
            break

    if not email_filled:
        print("[Login][WARN] Could not find email field — check login page selectors")

    # ── Fill password ─────────────────────────────────────────────────────
    pw_filled = False
    for sel in [
        'input[name="password"]',
        'input[type="password"]',
        '#password',
        '#customer_password',
        '[data-testid*="password" i]',
    ]:
        el = await page.query_selector(sel)
        if el:
            await el.fill(password)
            pw_filled = True
            print(f"[Login] Password filled via: {sel}")
            break

    if not pw_filled:
        print("[Login][WARN] Could not find password field — check login page selectors")

    # ── Submit form ───────────────────────────────────────────────────────
    submitted = False
    for sel in [
        'button[type="submit"]',
        'input[type="submit"]',
        'button:has-text("Sign In")',
        'button:has-text("Log In")',
        'button:has-text("Login")',
        '.login-btn',
        '[data-testid*="submit" i]',
        '[data-testid*="login" i]',
    ]:
        el = await page.query_selector(sel)
        if el:
            await el.click()
            submitted = True
            print(f"[Login] Submitted via: {sel}")
            break

    if not submitted:
        # Try pressing Enter on the password field as fallback
        pw_el = await page.query_selector('input[type="password"]')
        if pw_el:
            await pw_el.press("Enter")
            submitted = True
            print("[Login] Submitted via Enter key")

    # ── Wait for redirect after login ─────────────────────────────────────
    try:
        await page.wait_for_timeout(3000)
        await page.wait_for_load_state("networkidle", timeout=15_000)
    except Exception:
        pass

    current_url = page.url
    print(f"[Login] Post-login URL: {current_url}")

    # Login failed if still on login page or error visible
    if "/login" in current_url.lower() or "/account/login" in current_url.lower():
        # Check for error message
        err_text = ""
        for sel in [".error", ".alert-danger", "[class*='error']", "[class*='alert']"]:
            el = await page.query_selector(sel)
            if el:
                err_text = clean_text(await el.inner_text())
                break
        print(f"[Login] FAILED — still on login page. Error: {err_text or 'none visible'}")
        return False

    print("[Login] SUCCESS")
    return True


# ---------------------------------------------------------------------------
# Listing page — collect product URLs with pagination
# ---------------------------------------------------------------------------

async def get_product_links(page, base_listing_url: str, max_products: int = 0) -> list[str]:
    """
    Collect all product URLs from a Four Hands listing page.

    Pagination strategies attempted in order:
      1. "Load More" button — click repeatedly until gone
      2. ?page=N query parameter — increment until empty page
      3. /page/N/ path — increment until empty page
    """
    links: list[str] = []
    seen: set[str] = set()

    print(f"  [Listing] {base_listing_url}")

    try:
        await page.goto(base_listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
    except Exception as e:
        print(f"  [WARN] Failed to load listing: {e}")
        return links

    # ── Wait for product cards to render ─────────────────────────────────
    card_sel = await _detect_card_selector(page)
    print(f"  [Listing] Card selector: {card_sel}")

    # ── Strategy 1: "Load More" button ───────────────────────────────────
    load_more_sel = None
    for sel in [
        'button:has-text("Load More")',
        'button:has-text("Show More")',
        'a:has-text("Load More")',
        '[data-testid*="load-more" i]',
        '.load-more',
        '#load-more',
    ]:
        el = await page.query_selector(sel)
        if el and await el.is_visible():
            load_more_sel = sel
            break

    if load_more_sel:
        click_count = 0
        while click_count < MAX_PAGES:
            el = await page.query_selector(load_more_sel)
            if not el or not await el.is_visible():
                break
            try:
                await el.scroll_into_view_if_needed()
                await el.click()
                await page.wait_for_timeout(2000)
                click_count += 1
                # Check if we've hit the product limit
                hrefs = await _extract_card_links(page, card_sel)
                if max_products and len(hrefs) >= max_products:
                    break
            except Exception:
                break
        print(f"  [Listing] Clicked 'Load More' {click_count} times")
        hrefs = await _extract_card_links(page, card_sel)
        for h in hrefs:
            h = _normalise_url(h)
            if h and h not in seen:
                seen.add(h)
                links.append(h)
    else:
        # ── Strategy 2: Page-by-page (query param or path) ───────────────
        page_num = 1
        while page_num <= MAX_PAGES:
            if page_num > 1:
                # Try query param first: ?page=N
                url = _build_page_url(base_listing_url, page_num)
                try:
                    await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
                    await page.wait_for_timeout(2500)
                except Exception:
                    break

            hrefs = await _extract_card_links(page, card_sel)
            new_count = 0
            for h in hrefs:
                h = _normalise_url(h)
                if h and h not in seen:
                    seen.add(h)
                    links.append(h)
                    new_count += 1

            print(f"  [Listing p{page_num}] {new_count} new URLs ({len(links)} total)")

            if new_count == 0:
                break
            if max_products and len(links) >= max_products:
                break
            page_num += 1
            await async_polite_delay(0.8, 1.5)

    if max_products:
        links = links[:max_products]

    print(f"  [Listing] Total: {len(links)} product URLs")
    return links


async def _detect_card_selector(page) -> str:
    """
    Detect which CSS selector identifies product cards on the listing page.
    Waits up to 8s for cards to appear (JS rendering).
    """
    candidates = [
        # Four Hands specific (guesses based on React SPA patterns)
        '[data-testid="product-card"]',
        '[data-testid="product-item"]',
        '.product-card',
        '.product-item',
        '.product-tile',
        '.ProductCard',
        '.ProductItem',
        # Generic fallbacks
        'article.product',
        'li.product',
        '.grid-item a[href*="/product"]',
        'a[href*="/product/"]',
    ]
    for sel in candidates:
        try:
            await page.wait_for_selector(sel, timeout=5000)
            count = await page.eval_on_selector_all(sel, "els => els.length")
            if count > 0:
                return sel
        except Exception:
            continue
    return "a[href*='/product/']"  # last resort


async def _extract_card_links(page, card_sel: str) -> list[str]:
    """Extract href values from product cards."""
    try:
        # If selector is already an <a> tag
        if card_sel.startswith("a[") or card_sel == "a[href*='/product/']":
            hrefs = await page.eval_on_selector_all(
                card_sel, "els => els.map(e => e.href)"
            )
        else:
            # Get the <a> inside each card
            hrefs = await page.eval_on_selector_all(
                card_sel,
                """els => {
                    const out = [];
                    els.forEach(el => {
                        const a = el.tagName === 'A' ? el : el.querySelector('a[href*="/product"]') || el.querySelector('a');
                        if (a && a.href) out.push(a.href);
                    });
                    return out;
                }""",
            )
        return [h for h in hrefs if h and "/product/" in h]
    except Exception:
        # Broad fallback: all /product/ links on the page
        try:
            return await page.eval_on_selector_all(
                "a[href*='/product/']", "els => els.map(e => e.href)"
            )
        except Exception:
            return []


def _normalise_url(href: str) -> str:
    """Strip query params and fragments, ensure absolute URL."""
    if not href:
        return ""
    href = href.split("?")[0].split("#")[0].rstrip("/") + "/"
    if not href.startswith("http"):
        href = urljoin(BASE_URL, href)
    return href


def _build_page_url(base_url: str, page_num: int) -> str:
    """Build paginated URL. Tries query param first."""
    # Try: ?page=N appended
    sep = "&" if "?" in base_url else "?"
    return f"{base_url}{sep}page={page_num}"


# ---------------------------------------------------------------------------
# Product detail page scraping
# ---------------------------------------------------------------------------

def _extract_jsonld(page_text: str) -> dict:
    """Extract first Product JSON-LD object from raw script text."""
    try:
        obj = json.loads(page_text)
        if isinstance(obj, list):
            for item in obj:
                if isinstance(item, dict) and item.get("@type") == "Product":
                    return item
            return obj[0] if obj else {}
        if isinstance(obj, dict):
            if obj.get("@type") == "Product":
                return obj
            for item in obj.get("@graph", []):
                if isinstance(item, dict) and item.get("@type") == "Product":
                    return item
    except Exception:
        pass
    return {}


def _best_image(ld_obj: dict) -> str:
    img = ld_obj.get("image", "")
    if isinstance(img, list):
        img = img[0] if img else ""
    if isinstance(img, dict):
        img = img.get("url", img.get("contentUrl", ""))
    return str(img).strip()


def _slug_from_url(url: str) -> str:
    """Extract the product slug from a /product/{slug}/ URL."""
    parts = [p for p in url.rstrip("/").split("/") if p]
    # Find 'product' and take the next part
    try:
        idx = parts.index("product")
        return parts[idx + 1]
    except (ValueError, IndexError):
        pass
    return parts[-1] if parts else ""


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape a Four Hands product detail page.
    Returns a list of dicts — one per variant (size/finish), or single-element.
    """
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
    except Exception as e:
        print(f"    [WARN] Could not load product page: {e}")
        return [base]

    # ── 1. JSON-LD ─────────────────────────────────────────────────────────
    ld_obj: dict = {}
    for script in await page.query_selector_all('script[type="application/ld+json"]'):
        raw = (await script.inner_text()).strip()
        obj = _extract_jsonld(raw)
        if obj.get("@type") == "Product":
            ld_obj = obj
            break

    if ld_obj:
        base["Product Name"] = clean_text(ld_obj.get("name", ""))
        base["SKU"]          = clean_text(ld_obj.get("sku", ""))
        desc = ld_obj.get("description", "")
        if desc:
            base["Description"] = clean_text(re.sub(r"<[^>]+>", " ", desc))
        img = _best_image(ld_obj)
        if img:
            base["Image URL"] = img
        # Price
        offers = ld_obj.get("offers", {})
        if isinstance(offers, list):
            offers = offers[0] if offers else {}
        price = clean_price(str(offers.get("price", "")))
        if price:
            base["Price"] = price

    # ── 2. Product Name fallback ───────────────────────────────────────────
    if not base.get("Product Name"):
        for sel in [
            "h1.product-title",
            "h1.product-name",
            '[data-testid="product-name"]',
            '[data-testid="product-title"]',
            ".product-detail-name",
            ".pdp-title",
            "h1",
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                if text:
                    base["Product Name"] = text
                    break

    # ── 3. SKU fallback ────────────────────────────────────────────────────
    if not base.get("SKU"):
        for sel in [
            '[data-testid="product-sku"]',
            ".product-sku",
            ".sku",
            '[class*="sku" i]',
            '[class*="item-number" i]',
            '[class*="model" i]',
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                # Strip "SKU:", "Item #:", "Model:" prefix
                text = re.sub(r"(?i)^(sku|item\s*#?|model|part)\s*[:#]?\s*", "", text).strip()
                if text:
                    base["SKU"] = text
                    break

    # ── 4. Price fallback ──────────────────────────────────────────────────
    if not base.get("Price"):
        for sel in [
            '[data-testid="product-price"]',
            ".product-price",
            ".price",
            '[class*="price" i]',
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(await el.inner_text())
                p = clean_price(text)
                if p:
                    base["Price"] = p
                    break

    # ── 5. Image URL fallback ──────────────────────────────────────────────
    if not base.get("Image URL"):
        for sel in [
            '[data-testid="product-image"] img',
            ".product-image img",
            ".pdp-gallery img",
            ".product-gallery img",
            "img.product-main-image",
            ".hero-image img",
            "img[alt][src*='fourhands']",
            "img[alt][src*='cloudinary']",
            "img[alt][src*='imgix']",
        ]:
            el = await page.query_selector(sel)
            if el:
                src = (
                    await el.get_attribute("data-zoom-image")
                    or await el.get_attribute("data-large")
                    or await el.get_attribute("data-src")
                    or await el.get_attribute("src")
                    or ""
                ).strip()
                if src and "placeholder" not in src and src.startswith("http"):
                    base["Image URL"] = src
                    break

    # ── 6. Specification table / accordion ─────────────────────────────────
    spec_dict: dict[str, str] = {}

    # Try HTML tables first (th/td pairs)
    table_rows = await page.query_selector_all(
        "table.product-specs tr, table.specs tr, "
        ".specifications table tr, .product-details table tr, "
        "[class*='spec'] table tr, [data-testid*='spec'] tr"
    )
    for tr in table_rows:
        th = await tr.query_selector("th")
        td = await tr.query_selector("td")
        if th and td:
            k = clean_text(await th.inner_text())
            v = clean_text(await td.inner_text())
            if k and v:
                spec_dict[k] = v

    # Try definition list (dt/dd)
    if not spec_dict:
        dts = await page.query_selector_all(
            ".specs dt, .specifications dt, .product-attrs dt, [class*='spec'] dt"
        )
        dds = await page.query_selector_all(
            ".specs dd, .specifications dd, .product-attrs dd, [class*='spec'] dd"
        )
        for dt, dd in zip(dts, dds):
            k = clean_text(await dt.inner_text())
            v = clean_text(await dd.inner_text())
            if k and v:
                spec_dict[k] = v

    # Try key-value rows (div pairs, common in React SPAs)
    if not spec_dict:
        spec_rows = await page.query_selector_all(
            "[class*='spec-row'], [class*='spec-item'], [class*='detail-row'], "
            "[class*='attribute-row'], [data-testid*='spec']"
        )
        for row in spec_rows:
            children = await row.query_selector_all("span, div, p")
            if len(children) >= 2:
                k = clean_text(await children[0].inner_text())
                v = clean_text(await children[1].inner_text())
                if k and v:
                    spec_dict[k] = v

    # Broad fallback: scrape the whole spec section as text and parse it
    if not spec_dict:
        for container_sel in [
            '[data-testid*="specifications"]',
            '[data-testid*="details"]',
            ".product-specs",
            ".specifications",
            ".spec-section",
            "#specifications",
            "#product-details",
            "#details",
            ".product-attributes",
            "[class*='Specifications']",
        ]:
            el = await page.query_selector(container_sel)
            if el:
                raw_text = clean_text(await el.inner_text())
                if raw_text:
                    parsed = parse_spec_block(raw_text.replace("\n", " | "))
                    spec_dict.update(parsed)
                    break

    # ── 7. Map spec_dict → product fields ─────────────────────────────────
    _apply_specs(base, spec_dict)

    # ── 8. Description fallback ────────────────────────────────────────────
    if not base.get("Description"):
        for sel in [
            '[data-testid="product-description"]',
            ".product-description",
            ".pdp-description",
            ".product-short-desc",
            ".description",
        ]:
            el = await page.query_selector(sel)
            if el:
                text = clean_text(re.sub(r"<[^>]+>", " ", await el.inner_html()))
                if text and len(text) > 10:
                    base["Description"] = text
                    break

    # ── 9. Tearsheet link ──────────────────────────────────────────────────
    # Construct from URL slug: /product/{slug}/tearsheet/pdf
    slug = _slug_from_url(url)
    if slug:
        base["Tearsheet Link"] = f"{BASE_URL}/product/{slug}/tearsheet/pdf"

    # Also look for an explicit tearsheet/spec-sheet link on the page
    if not base.get("Tearsheet Link"):
        for sel in ["a[href*='tearsheet']", "a[href*='spec-sheet']", "a[href*='.pdf']"]:
            el = await page.query_selector(sel)
            if el:
                href = await el.get_attribute("href") or ""
                if href:
                    base["Tearsheet Link"] = urljoin(BASE_URL, href)
                    break

    # ── 10. Product Family Id ─────────────────────────────────────────────
    if not base.get("Product Family Id") and base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    return [base]


_SPEC_KEY_ALIASES: dict[str, str] = {
    # Dimensions
    "width":                "Width",
    "w":                    "Width",
    "depth":                "Depth",
    "d":                    "Depth",
    "height":               "Height",
    "h":                    "Height",
    "diameter":             "Diameter",
    "dia":                  "Diameter",
    "dia.":                 "Diameter",
    "length":               "Length",
    "dimensions":           "Dimensions",
    "overall dimensions":   "Dimensions",
    "overall size":         "Dimensions",
    "size":                 "Dimensions",
    # Physical
    "weight":               "Weight",
    "item weight":          "Weight",
    "shipping weight":      "Weight",
    # Materials / finish
    "material":             "Materials",
    "materials":            "Materials",
    "finish":               "Finish",
    "finishes":             "Finish",
    "color":                "Finish",
    "colour":               "Finish",
    "wood finish":          "Finish",
    # Provenance
    "collection":           "Collection",
    "designer":             "Designer",
    "origin":               "Origin",
    "country of origin":    "Origin",
    "made in":              "Origin",
    "lead time":            "Lead Time",
    # Seating
    "seat height":          "Seat Height",
    "seat depth":           "Seat Depth",
    "seat width":           "Seat Width",
    "arm height":           "Arm Height",
    "back height":          "Back",
    "com":                  "COM",
    "col":                  "COL",
    "fabric":               "Fabric",
    "upholstery":           "Fabric",
    "seat construction":    "Seat Construction",
    "frame":                "Frame",
    "base":                 "Base",
    # Lighting
    "wattage":              "Wattage",
    "socket":               "Socket",
    "socket type":          "Socket",
    "lamping":              "Lamping",
    "lamp type":            "Lamping",
    "light source":         "Lamping",
    "bulb type":            "Lamping",
    "voltage":              "Voltage",
    "lumens":               "Lumens",
    "color temperature":    "Color Temperature",
    "colour temperature":   "Color Temperature",
    "cri":                  "CRI",
    "canopy":               "Canopy",
    "canopy size":          "Canopy",
    "chain length":         "Chain Length",
    "hanging length":       "Chain Length",
    "min drop":             "Chain Length",
    "shade":                "Shade Details",
    "shade details":        "Shade Details",
    "mounting":             "Mounting",
    "bulb qty":             "Bulb Qty",
    "number of bulbs":      "Bulb Qty",
    "extension":            "Extension",
    # Other
    "assembly required":    "Assembly Required",
    "hardware details":     "Hardware Details",
    "carton size":          "Carton Size",
    "box size":             "Carton Size",
}


def _apply_specs(row: dict, spec_dict: dict[str, str]) -> None:
    """Map raw spec key/value pairs onto the product row dict."""
    for raw_key, raw_val in spec_dict.items():
        if not raw_val:
            continue
        k_norm = raw_key.lower().strip().rstrip(":")
        canonical = _SPEC_KEY_ALIASES.get(k_norm)

        # Fuzzy: try startswith match
        if canonical is None:
            for alias, canon in _SPEC_KEY_ALIASES.items():
                if k_norm.startswith(alias):
                    canonical = canon
                    break

        if canonical is None:
            # Unknown field — keep as title-cased
            canonical = raw_key.strip().title()

        # Special handling: parse Dimensions into sub-fields
        if canonical == "Dimensions":
            parsed = parse_dimensions(raw_val)
            for dim_key, dim_val in parsed.items():
                row.setdefault(dim_key, dim_val)
        elif canonical in ("Width", "Height", "Depth", "Diameter", "Length"):
            # Store numeric value only
            numeric = re.sub(r"[^\d.]", "", raw_val.split()[0])
            row.setdefault(canonical, numeric or raw_val)
        else:
            row.setdefault(canonical, raw_val)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main() -> None:
    from playwright.async_api import async_playwright

    if not FH_EMAIL or not FH_PASSWORD:
        print(
            "\n[ERROR] Credentials not set.\n"
            "  export FH_EMAIL='your@email.com'\n"
            "  export FH_PASSWORD='yourpassword'\n"
            "  Then re-run the scraper.\n"
        )
        sys.exit(1)

    info       = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer     = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    categories = info["categories"]

    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")
    print(f"[Scraper] Account : {FH_EMAIL}")

    async with async_playwright() as pw:
        browser, context, page = await create_page(pw, HEADLESS)
        try:
            # ── Login ─────────────────────────────────────────────────────
            ok = await login(page, FH_EMAIL, FH_PASSWORD)
            if not ok:
                print("[ERROR] Login failed — cannot proceed without authentication.")
                sys.exit(2)

            for cat in categories:
                if not cat["links"]:
                    continue

                writer.add_sheet(
                    cat["name"],
                    cat["links"][0],
                    studio_columns=cat["studio_columns"],
                )

                # ── Phase 2: collect all product URLs ─────────────────────
                seen_urls: set[str] = set()
                all_urls: list[str] = []

                for listing_url in cat["links"]:
                    max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0
                    new_links = await get_product_links(page, listing_url, max_products=max_p)
                    for u in new_links:
                        if u not in seen_urls:
                            seen_urls.add(u)
                            all_urls.append(u)
                    if TEST_MODE and len(all_urls) >= TEST_MAX_PRODUCTS:
                        break

                if TEST_MODE:
                    all_urls = all_urls[:TEST_MAX_PRODUCTS]

                print(f"\n[Category] {cat['name']}: {len(all_urls)} product URLs found")

                # ── Phase 3: scrape each product page ─────────────────────
                global_idx = 1
                for url in all_urls:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = info["vendor_name"]
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                    await async_polite_delay(0.8, 2.0)

                await async_polite_delay(1.0, 2.5)

        finally:
            await context.close()
            await browser.close()

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
