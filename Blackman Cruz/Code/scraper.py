import asyncio
import json
import os
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Blackman Cruz")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://blackmancruz.com"

# h5 labels that are navigation/UI elements, not product specs
_H5_SKIP = {
    "print tearsheet", "recommendations", "contact us",
    "sign up for our newsletter", "shopping cart", "menu",
    "search", "navigation", "footer",
}

# Map raw h5 label → canonical column name
_SPEC_MAP = {
    "origin":        "Origin",
    "materials":     "Materials",
    "material":      "Materials",
    "weight":        "Weight",
    "finish":        "Finish",
    "finishes":      "Finish",
    "collection":    "Collection",
    "designer":      "Designer",
    "maker":         "Maker",
    "designed by":   "Designer",
    "made by":       "Maker",
    "date":          "Date",
    "wattage":       "Wattage",
    "voltage":       "Voltage",
    "socket":        "Socket",
    "bulb type":     "Bulb Type",
    "color":         "Color",
    "colour":        "Color",
    "fabric":        "Fabric",
    "upholstery":    "Upholstery",
    "seat height":   "Seat Height",
    "seat depth":    "Seat Depth",
    "arm height":    "Arm Height",
    "back":          "Back",
    "lead time":     "Lead Time",
}

# Dimension label aliases
_DIM_MAP = {
    "h": "Height", "height": "Height",
    "w": "Width",  "width":  "Width",
    "d": "Depth",  "depth":  "Depth",
    "l": "Length", "length": "Length",
    "dia": "Diameter", "diam": "Diameter", "diameter": "Diameter",
}


def _normalize_inch_marks(s: str) -> str:
    """Normalise all inch-mark variants to a plain ASCII double-quote."""
    # Curly/smart double quotes
    s = s.replace('“', '"').replace('”', '"')
    # Double prime ″ and modifier letter double prime
    s = s.replace('″', '"').replace('ʺ', '"')
    # Two consecutive apostrophes / single quotes used as inch marks
    s = re.sub(r"['‘’ʹ]{2}", '"', s)
    return s


def _preprocess_mixed_fractions(s: str) -> str:
    """Convert mixed fractions like 1-1/4 → 1.25 before parsing."""
    def _replace(m):
        whole = int(m.group(1))
        num   = int(m.group(2))
        den   = int(m.group(3))
        return str(round(whole + num / den, 4))
    return re.sub(r'(\d+)-(\d+)/(\d+)', _replace, s)


def _parse_single_dim(s: str) -> str | None:
    """Extract numeric value from a single-dim h5 value like '33.25"' or '17 in'."""
    s = _preprocess_mixed_fractions(s)
    s = re.sub(r'["″]|\bin\.?\b', '', s, flags=re.IGNORECASE).strip()
    # Take the first numeric token
    m = re.match(r'^([\d]+(?:[./][\d]+)?(?:\.\d+)?)', s)
    if m:
        raw = m.group(1)
        if '/' in raw:
            parts = raw.split('/')
            try:
                val = int(parts[0]) / int(parts[1])
                return str(round(val, 4))
            except (ValueError, ZeroDivisionError):
                pass
        return raw
    return None


def _extract_nb_dims(text: str) -> dict:
    """
    Extract dimensions using ONLY number-before-label pattern.
    Avoids the parse_dimensions Pattern-A bug where 'H <next_num>' is
    misread as Height = next_num on strings like '17 H 28 W x 24 D'.
    """
    result: dict[str, str] = {}
    text = _preprocess_mixed_fractions(text)
    pat = re.compile(
        r'\b([\d]+(?:[./][\d]+)?(?:\.\d+)?)\s*(?:["″])?\s*'
        r'(dia(?:m(?:eter)?)?\.?|[wdhl])\b',
        re.IGNORECASE,
    )
    for m in pat.finditer(text):
        val_str = m.group(1)
        label   = m.group(2).lower().rstrip('.')
        col     = _DIM_MAP.get(label) or _DIM_MAP.get(label[:3] if len(label) >= 3 else label)
        if col and col not in result:
            if '/' in val_str:
                try:
                    a, b = val_str.split('/')
                    val_str = str(round(int(a) / int(b), 4))
                except (ValueError, ZeroDivisionError):
                    pass
            result[col] = val_str
    return result


def _parse_bc_dimensions(raw: str) -> dict:
    """
    Parse Blackman Cruz dimension strings. Returns a dict that may contain:
    Dimensions, Height, Width, Depth, Diameter, Length, Seat Height.

    Handles:
    - "18" H x 88" W x 32.5 "D"
    - "33.25" H x 52.75" W x 17.5" D"
    - "Back: 31.5" H, Seat: 17" H x 19" W x 20" D"  (seating complex format)
    - "Dia. 24" x H 36""
    - Mixed fractions: "1-1/4" → 1.25
    """
    if not raw:
        return {}

    result = {}

    # Normalise inch marks then mixed fractions
    processed = _preprocess_mixed_fractions(_normalize_inch_marks(raw))

    # Check for seating complex format: "Back: ... , Seat: ..."
    # Capture everything after the keyword until the OTHER keyword or end-of-string.
    # Using greedy .+ so commas / colons inside the value are included.
    back_match = re.search(r'\bback\s*[:\-/]?\s*(.+?)(?=\bseat\b|$)',
                           processed, re.IGNORECASE)
    seat_match = re.search(r'\bseat\s*[:\-/]?\s*(.+?)(?=\bback\b|$)',
                           processed, re.IGNORECASE)

    if back_match or seat_match:
        # Extract back height → Height (use nb extractor to avoid pattern-A ambiguity)
        if back_match:
            back_dims = _extract_nb_dims(back_match.group(1))
            if "Height" in back_dims:
                result["Height"] = back_dims["Height"]

        # Extract seat dimensions using nb extractor to correctly read "17 H 28 W"
        if seat_match:
            seat_str = seat_match.group(1)
            # Only take text up to the first occurrence of "back" to avoid re-reading back portion
            back_idx = re.search(r'\bback\b', seat_str, re.IGNORECASE)
            if back_idx:
                seat_str = seat_str[:back_idx.start()]
            seat_dims = _extract_nb_dims(seat_str)
            if "Height" in seat_dims:
                result["Seat Height"] = seat_dims["Height"]
            for k in ("Width", "Depth", "Length", "Diameter"):
                if k in seat_dims:
                    result[k] = seat_dims[k]
    else:
        # Standard combined dimensions string
        dims = parse_dimensions(processed)
        result.update(dims)

    # Always store a cleaned Dimensions string (no inch marks)
    clean_dim = re.sub(r'["″]', '', _normalize_inch_marks(raw))
    clean_dim = re.sub(r'\bin\.?\b', '', clean_dim, flags=re.IGNORECASE)
    clean_dim = re.sub(r'\s+', ' ', clean_dim).strip()
    result["Dimensions"] = clean_dim

    return result


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape one product page. Returns a list of dicts (always single-row for this vendor)."""
    handle = url.rstrip("/").split("/")[-1]
    canonical_url = f"{BASE_URL}/products/{handle}"

    await _goto_with_retry(page, canonical_url)
    await page.wait_for_timeout(1500)

    # Fetch JSON API (same origin) for: title, price, images, vendor
    api_data = await page.evaluate("""
        async () => {
            try {
                const r = await fetch(window.location.pathname + '.json');
                const j = await r.json();
                return j.product || {};
            } catch (e) { return {}; }
        }
    """)

    # --- Basic data from API ---
    product_name = clean_text(api_data.get("title", "") or "")
    vendor       = (api_data.get("vendor") or "").strip()

    images   = api_data.get("images") or []
    img_src  = images[0].get("src", "") if images else ""
    if img_src.startswith("//"):
        img_src = "https:" + img_src

    variants = api_data.get("variants") or []
    price    = None
    if variants:
        raw_p = variants[0].get("price", "")
        price = clean_price(str(raw_p)) if raw_p else None

    # --- Specs from h5 tags in HTML ---
    # Collect all h5 labels + values; take first occurrence only (sidebar repeats them)
    raw_specs = await page.evaluate("""
        () => {
            const result = [];
            const seen   = new Set();
            document.querySelectorAll('h5').forEach(h5 => {
                const label = h5.textContent.trim();
                const key   = label.toLowerCase();
                if (seen.has(key)) return;
                seen.add(key);

                // Try nextElementSibling first
                let value = '';
                const nextEl = h5.nextElementSibling;
                if (nextEl && !['H1','H2','H3','H4','H5','H6'].includes(nextEl.tagName)) {
                    value = nextEl.textContent.trim();
                }

                // Fallback: scan next sibling text nodes
                if (!value) {
                    let node = h5.nextSibling;
                    while (node) {
                        if (node.nodeType === 3) {
                            const t = node.textContent.trim();
                            if (t) { value = t; break; }
                        } else if (node.nodeType === 1) {
                            const tag = node.tagName || '';
                            if (['H1','H2','H3','H4','H5','H6'].includes(tag)) break;
                            const t = node.textContent.trim();
                            if (t) { value = t; break; }
                        }
                        node = node.nextSibling;
                    }
                }

                if (value) result.push([label, value]);
            });
            return result;
        }
    """)

    # --- Build the product row ---
    row = {
        "Source":           canonical_url,
        "Manufacturer":     VENDOR_NAME,
        "Product Name":     sentence_case(product_name) if product_name else "",
        "Product Family Id": extract_family_id(product_name) if product_name else "",
        "Image URL":        img_src,
        "Price":            price,
    }

    # Designer: use vendor field only when it's not the store itself
    if vendor and vendor.lower() not in ("blackman cruz", "blackmancruz"):
        row["Designer"] = vendor

    # Process spec h5 pairs
    for label, value in raw_specs:
        label_lower = label.lower().strip()

        if label_lower in _H5_SKIP or not value:
            continue

        # Dimension fields (combined or individual)
        if label_lower == "dimensions":
            dim_data = _parse_bc_dimensions(value)
            for k, v in dim_data.items():
                if k not in row:
                    row[k] = v
            continue

        if label_lower == "height":
            v = _parse_single_dim(value)
            if v and "Height" not in row:
                row["Height"] = v
            continue

        if label_lower == "width":
            v = _parse_single_dim(value)
            if v and "Width" not in row:
                row["Width"] = v
            continue

        if label_lower == "depth":
            v = _parse_single_dim(value)
            if v and "Depth" not in row:
                row["Depth"] = v
            continue

        if label_lower == "diameter":
            v = _parse_single_dim(value)
            if v and "Diameter" not in row:
                row["Diameter"] = v
            continue

        if label_lower == "length":
            v = _parse_single_dim(value)
            if v and "Length" not in row:
                row["Length"] = v
            continue

        if label_lower == "seat height":
            v = _parse_single_dim(value)
            if v:
                row["Seat Height"] = v
            continue

        if label_lower == "weight":
            wt = re.search(r'[\d.]+', value)
            if wt:
                row["Weight"] = safe_float(wt.group())
            continue

        # All other spec fields: map to canonical name
        canonical = _SPEC_MAP.get(label_lower)
        if canonical:
            if canonical not in row:
                row[canonical] = clean_text(value)
        else:
            # Unknown field — store as-is (title-cased label)
            col = label.strip().title()
            if col not in row:
                row[col] = clean_text(value)

    # Build a Dimensions summary if we have individual dims but no Dimensions
    if "Dimensions" not in row:
        parts = []
        for axis, abbr in [("Width", "W"), ("Depth", "D"), ("Height", "H"),
                           ("Diameter", "Dia"), ("Length", "L")]:
            if row.get(axis):
                parts.append(f"{row[axis]} {abbr}")
        if parts:
            row["Dimensions"] = " x ".join(parts)

    return [row]


async def _goto_with_retry(page, url: str, retries: int = 3) -> bool:
    """Navigate to url with up to `retries` attempts on network errors. Returns True on success."""
    for attempt in range(1, retries + 1):
        try:
            await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
            return True
        except Exception as e:
            err = str(e)
            if attempt < retries and ("ERR_INTERNET_DISCONNECTED" in err
                                      or "ERR_NETWORK_CHANGED" in err
                                      or "ERR_CONNECTION_RESET" in err
                                      or "net::" in err):
                print(f"  [Retry {attempt}/{retries}] Network error on {url}: {e}", file=sys.stderr)
                await asyncio.sleep(5)
            else:
                raise
    return False


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all unique canonical product URLs from a collection listing page.
    Handles Shopify ?page=N pagination."""
    seen:   set[str] = set()
    result: list[str] = []

    page_num = 1
    while True:
        url = f"{listing_url}?page={page_num}" if page_num > 1 else listing_url
        await _goto_with_retry(page, url)
        await page.wait_for_timeout(1500)

        hrefs = await page.evaluate("""
            () => {
                const seen = new Set();
                const out  = [];
                document.querySelectorAll('a[href*="/products/"]').forEach(a => {
                    const h = a.getAttribute('href');
                    if (h && !seen.has(h)) { seen.add(h); out.push(h); }
                });
                return out;
            }
        """)

        new_found = False
        for href in hrefs:
            # Normalize to canonical /products/{handle}
            m = re.search(r'/products/([^/?#]+)', href)
            if not m:
                continue
            handle = m.group(1)
            canonical = f"{BASE_URL}/products/{handle}"
            if canonical not in seen:
                seen.add(canonical)
                result.append(canonical)
                new_found = True

        if not new_found:
            break

        page_num += 1

    return result


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            print(f"\n[{cat['group']}] {cat['name']} ({len(cat['links'])} link(s))")

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            # Collect product URLs from ALL links for this category (deduplicated)
            seen_urls:        set[str]  = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            print(f"  Found {len(all_product_urls)} products")

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(
                                info["vendor_name"], cat["name"], global_idx
                            )
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(
                                variant["Product Name"]
                            )
                        writer.write_row(variant, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [Error] {url}: {e}", file=sys.stderr)

                await async_polite_delay()

    writer.save()
    print(f"\n[Done] Saved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
