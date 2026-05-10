"""
scraper.py  —  Julian Chichester
----------------------------------
Platform: julianchichester.com (Laravel + Inertia.js SPA)

Site structure:
  Listing  : /category/{slug}  — all products in page JSON (no pagination)
             Data embedded in <div id="app" data-page="..."> as JSON
  Product  : /product/{slug}   — basic info in page JSON; specs in popup

Spec popup:
  Clicking "VIEW SPECIFICATIONS" opens a modal with:
    - Key/value rows (BASE FINISH, TOP FINISH, DIMENSIONS, ...)
    - DIMENSIONS format: "W 86.61 in, D 55.1 in, H 29.9 in"
    - FINISHES USED: swatch names
  Dimensions are in inches — no mm conversion needed.
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import html as htmllib
from pathlib import Path
from urllib.parse import urljoin

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Julian Chichester")
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

BASE_URL   = "https://julianchichester.com"
TIMEOUT_MS = 45_000


def _strip_html(text: str) -> str:
    return clean_text(re.sub(r"<[^>]+>", " ", text or ""))


def _convert_mm_dims(dim_str: str) -> str:
    """Convert mm values to inches in a dimension string like 'W 1200 mm / D 30 mm / H 1200 mm'."""
    if not dim_str or "mm" not in dim_str.lower():
        return dim_str
    def _repl(m: re.Match) -> str:
        return str(round(float(m.group(1)) / 25.4, 2))
    return re.sub(r"([\d]+(?:\.\d+)?)\s*mm\b", _repl, dim_str, flags=re.IGNORECASE)


def _extract_page_data(html: str) -> dict:
    """Extract Inertia.js page data from <div id="app" data-page="...">."""
    m = re.search(r'id="app"\s+data-page="(.*?)"(?:\s|>)', html, re.DOTALL)
    if not m:
        return {}
    try:
        return json.loads(htmllib.unescape(m.group(1)))
    except Exception:
        return {}


async def _get_spec_popup(page) -> dict:
    """
    Click VIEW SPECIFICATIONS, wait for the popup, and return a dict of all
    key/value pairs found inside it, plus parsed dimension fields.
    """
    result = {}

    # Try to find and click the specifications trigger
    clicked = False
    for sel in [
        "text=VIEW SPECIFICATIONS",
        "a:has-text('SPECIFICATIONS')",
        "button:has-text('SPECIFICATIONS')",
        "[class*='specification' i]",
    ]:
        try:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                await el.click()
                clicked = True
                break
        except Exception:
            continue

    if not clicked:
        return result

    await page.wait_for_timeout(2000)

    # Extract all label+value pairs from the visible popup.
    # Julian Chichester uses a Tailwind CSS fixed-position right panel.
    # IMPORTANT: fixed-position elements always have offsetParent===null so we
    # must use getBoundingClientRect + getComputedStyle for visibility checks.
    popup_data = await page.evaluate("""
        () => {
            const POPUP_TITLES = [
                'SPECIFICATION & OPTIONS',
                'SPECIFICATIONS & OPTIONS',
                'SPECIFICATION AND OPTIONS',
            ];

            // ── Strategy A: exact match on the popup heading title ────────
            // Uses EXACT match (not includes) to avoid matching
            // the "VIEW SPECIFICATIONS" button which appears earlier in DOM.
            let modal = null;
            const heading = [...document.querySelectorAll('*')].find(el => {
                const t = el.textContent.trim().toUpperCase();
                return POPUP_TITLES.includes(t);
            });
            if (heading) {
                let node = heading.parentElement;
                for (let i = 0; i < 10 && node && node !== document.body; i++) {
                    const r = node.getBoundingClientRect();
                    if (r.width >= 250 && r.height >= 300) { modal = node; break; }
                    node = node.parentElement;
                }
            }

            // ── Strategy B: CSS class scan with proper visibility check ───
            // getBoundingClientRect works for fixed-position elements;
            // offsetParent does NOT (always null for position:fixed).
            if (!modal) {
                const isVisible = el => {
                    const r = el.getBoundingClientRect();
                    const s = window.getComputedStyle(el);
                    return r.width > 0 && r.height > 0 &&
                           s.display !== 'none' &&
                           s.visibility !== 'hidden' &&
                           parseFloat(s.opacity || '1') > 0;
                };
                const candidates = [
                    ...document.querySelectorAll('[role="dialog"]'),
                    ...document.querySelectorAll('[class*="modal"]'),
                    ...document.querySelectorAll('[class*="overlay"]'),
                    ...document.querySelectorAll('[class*="drawer"]'),
                    ...document.querySelectorAll('[class*="panel"]'),
                    ...document.querySelectorAll('.fixed'),
                    ...document.querySelectorAll('.absolute'),
                ];
                modal = candidates.find(el =>
                    isVisible(el) &&
                    el.getBoundingClientRect().width > 200 &&
                    POPUP_TITLES.some(t => el.textContent.toUpperCase().includes(t))
                ) || null;
            }

            const raw = modal ? modal.innerText : document.body.innerText;
            return { _raw: raw, _modal_found: !!modal };
        }
    """)

    if not popup_data:
        return result

    # ── Parse raw innerText ───────────────────────────────────────────────
    raw = popup_data.get("_raw", "")
    # Always narrow to start from the popup heading — this is a safety net
    # in case the wrong container was found or the full body was returned.
    if raw:
        raw_up = raw.upper()
        for title in ("SPECIFICATION & OPTIONS", "SPECIFICATIONS & OPTIONS", "SPECIFICATION"):
            start = raw_up.find(title)
            if start != -1:
                raw = raw[start:]
                break
        # Trim off anything after the product page resumes
        for end_marker in ("LOGIN TO SEE PRICE", "INQUIRE", "ADD TO CART"):
            end = raw.upper().find(end_marker)
            if end != -1:
                raw = raw[:end]
                break

    if raw:
        # Normalize tabs → newlines so "LABEL\tValue" becomes two lines
        raw_norm = raw.replace("\t", "\n")
        lines = [ln.strip() for ln in raw_norm.splitlines() if ln.strip()]

        SKIP_LABELS = {
            "SPECIFICATION & OPTIONS", "SPECIFICATIONS & OPTIONS",
            "SPECIFICATION", "SPECIFICATIONS", "FINISHES USED", "CLOSE", "X", "×",
        }
        KNOWN_LABELS = {
            "BASE FINISH", "TOP FINISH", "FINISH", "DIMENSIONS",
            "MATERIAL", "MATERIALS", "FABRIC", "UPHOLSTERY",
            "SEAT HEIGHT", "ARM HEIGHT", "SEAT DEPTH", "SEAT WIDTH",
            "SOCKET", "WATTAGE", "BULB TYPE", "VOLTAGE",
            "WEIGHT", "LEAD TIME", "ORIGIN", "COLLECTION",
            "SHADE DETAILS", "CANOPY", "CHAIN LENGTH", "EXTENSION",
        }

        i = 0
        while i < len(lines):
            line_up = lines[i].upper()

            if line_up in SKIP_LABELS:
                i += 1
                continue

            is_label = (
                line_up in KNOWN_LABELS or
                (lines[i].isupper() and 2 < len(lines[i]) <= 35 and
                 not re.match(r'^[\d\W]+$', lines[i]))
            )

            if is_label and i + 1 < len(lines):
                val = lines[i + 1]
                if not (val.isupper() and len(val) <= 35):
                    result.setdefault(lines[i].title(), val)
                    i += 2
                    continue
            i += 1

        # Swatch names after "FINISHES USED"
        if not result.get("Finishes Available"):
            fu_idx = next(
                (j for j, ln in enumerate(lines) if ln.upper() == "FINISHES USED"), None
            )
            if fu_idx is not None:
                sw = [ln for ln in lines[fu_idx + 1:]
                      if ln and not re.match(r'^[\d\W]+$', ln)
                      and not (ln.isupper() and len(ln) <= 35)]
                if sw:
                    result["Finishes Available"] = ", ".join(sw)

    # Parse DIMENSIONS into individual fields; convert mm → inches first
    dim_raw = result.pop("Dimensions", "")
    if dim_raw:
        dim_raw = _convert_mm_dims(dim_raw)
        parsed = parse_dimensions(dim_raw)
        result.update(parsed)   # Width, Depth, Height, Diameter, Length, Dimensions

    return result


async def get_product_slugs(page, listing_url: str) -> list[str]:
    """Return all product slugs from a category listing page."""
    print(f"  [Listing] {listing_url}")
    try:
        await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"  [WARN] {e}")
        return []

    html = await page.content()
    data = _extract_page_data(html)
    products = data.get("props", {}).get("category", {}).get("products", [])
    slugs = [p["slug"] for p in products if p.get("slug")]
    print(f"  [Listing] {len(slugs)} products")
    return slugs


async def scrape_product(page, slug: str) -> list[dict]:
    """
    Scrape a Julian Chichester product detail page.

    Strategy:
      1. Extract variant rows from the Inertia.js page JSON
         (name, sku, finish/color, description, image).
      2. Click VIEW SPECIFICATIONS → extract popup data
         (dimensions, base finish, top finish, any other spec rows).
      3. Apply shared popup specs to every variant row.
    """
    url = f"{BASE_URL}/product/{slug}"
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [base]

    html = await page.content()
    data = _extract_page_data(html)
    props = data.get("props", {})
    variants: list[dict] = props.get("variants", [])

    # ── 1. Build base row from page JSON (product-level fields) ───────────
    rows: list[dict] = []

    if variants:
        for v in variants:
            row: dict = {"Source": url}

            row["Product Name"]      = clean_text(v.get("name", ""))
            row["SKU"]               = v.get("sku", "") or ""
            row["Product Family Id"] = extract_family_id(row["Product Name"])

            # Price (null for trade-only site)
            price = v.get("price")
            if price is not None:
                try:
                    row["Price"] = float(price)
                except (TypeError, ValueError):
                    pass

            # Finish / Color / Size from JSON
            if v.get("color"):
                row["Finish"] = v["color"]
            if v.get("size"):
                row["Size"] = v["size"]

            # Description
            desc = _strip_html(v.get("description", ""))
            if desc:
                row["Description"] = desc

            # Image URL
            images = v.get("images", [])
            if images:
                img_url = images[0].get("url", "")
                if img_url:
                    row["Image URL"] = img_url

            # Dimensions from JSON (in mm — fallback if popup fails)
            specs_json = v.get("specifications", {})
            dims_json  = specs_json.get("dimensions", {})
            if dims_json:
                units = dims_json.get("units", "mm")
                def _to_in(val):
                    try:
                        f = float(val)
                        if f <= 0:
                            return None
                        return str(round(f / 25.4, 2)) if units == "mm" else str(round(f, 2))
                    except (TypeError, ValueError):
                        return None

                for field, key in [("width","Width"),("depth","Depth"),
                                   ("height","Height"),("diameter","Diameter")]:
                    v_in = _to_in(dims_json.get(field))
                    if v_in:
                        row[f"_json_{key}"] = v_in   # store with prefix; popup wins

            # Spec details from JSON
            details = specs_json.get("details", [])
            if details:
                row["_json_Specifications"] = " | ".join(
                    f"{item.get('label','')}: {_strip_html(item.get('value',''))}"
                    for item in details if item.get("value")
                )

            rows.append(row)
    else:
        rows = [base]

    # ── 2. Click VIEW SPECIFICATIONS popup ────────────────────────────────
    popup = await _get_spec_popup(page)

    # ── 3. Merge popup data into every row ────────────────────────────────
    for row in rows:
        # Popup wins over JSON for dimensions — apply popup fields
        for key, val in popup.items():
            row.setdefault(key, val)

        # If popup gave us no dimensions, promote the JSON fallback values
        for dim_field in ("Width", "Depth", "Height", "Diameter", "Length"):
            jkey = f"_json_{dim_field}"
            if not row.get(dim_field) and row.get(jkey):
                row[dim_field] = row[jkey]
            row.pop(jkey, None)   # clean up temp keys

        # Promote JSON specifications if popup had nothing
        if not row.get("Specifications") and row.get("_json_Specifications"):
            row["Specifications"] = row["_json_Specifications"]
        row.pop("_json_Specifications", None)

        # Rebuild Dimensions string from individual fields if missing
        if not row.get("Dimensions"):
            parts = []
            for lbl, field in [("W","Width"),("D","Depth"),("H","Height"),
                                ("Dia","Diameter"),("L","Length")]:
                if row.get(field):
                    parts.append(f"{lbl} {row[field]}")
            if parts:
                row["Dimensions"] = " x ".join(parts)

    return rows if rows else [base]


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

            seen_slugs: set[str] = set()
            all_slugs:  list[str] = []

            for listing_url in cat["links"]:
                for s in await get_product_slugs(page, listing_url):
                    if s not in seen_slugs:
                        seen_slugs.add(s)
                        all_slugs.append(s)

            if TEST_MODE:
                all_slugs = all_slugs[:TEST_MAX_PRODUCTS]

            print(f"\n[Category] {cat['name']}: {len(all_slugs)} products")

            global_idx = 1
            for slug in all_slugs:
                try:
                    rows = await scrape_product(page, slug)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = info["vendor_name"]
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                    print(f"  [{global_idx - len(rows)}] {slug}")
                except Exception as e:
                    print(f"  [ERROR] {slug}: {e}")
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
