"""
scraper.py — Kuzco Lighting
----------------------------
Platform : kuzcolighting.com (Shopify)

Site structure:
  Listing  : Shopify /collections/{handle}/products.json?limit=250&page=N
             (navigates once to set domain context, then fetches JSON API)
  Product  : /products/{handle}.js for variant data (SKU, price, finish, images)
             + product page visit for SPECIFICATIONS table and DOWNLOADS section

Variants : Finish dropdown — one row per variant
           option1 = Finish name, SKU includes finish suffix (e.g. CH18537-BK)

Spec table: <details> accordion with heading "SPECIFICATIONS"
            Two-column <table>: Label | Value
            Keys mapped to standard column names (see SPEC_KEY_MAP below)

Downloads : <details> accordion with heading "DOWNLOADS"
            PDF links captured as Spec Sheet and Install Guide
"""

import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# ---------------------------------------------------------------------------
# Checkpoint helpers — save progress after each product so crashes are resumable
# ---------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent.parent / "Data"

def _ckpt_path():  return DATA_DIR / "Kuzco Lighting_progress.json"
def _rows_path():  return DATA_DIR / "Kuzco Lighting_rows.jsonl"

def load_checkpoint() -> dict:
    p = _ckpt_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_checkpoint(state: dict):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    _ckpt_path().write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")

def append_row(row: dict):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with _rows_path().open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")

def load_saved_rows() -> list[dict]:
    p = _rows_path()
    if not p.exists():
        return []
    rows = []
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            try:
                rows.append(json.loads(line))
            except Exception:
                pass
    return rows

def clear_checkpoint():
    for p in [_ckpt_path(), _rows_path()]:
        if p.exists():
            p.unlink()

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Kuzco Lighting")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
SCRAPE_CATEGORIES   = [
    c.strip() for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",")
    if c.strip()
]

BASE_URL = "https://kuzcolighting.com"

# Map Kuzco spec table labels → our standard column names
SPEC_KEY_MAP = {
    "Fixture Dimensions":      "Dimensions",
    "Canopy Dimensions":       "Canopy",
    "Backplate Dimensions":    "Backplate Dimensions",
    "Shade Dimensions":        "Shade Dimensions",
    "Primary Dimension":       "Primary Dimension",
    "Color Temperature":       "Color Temperature",
    "Location Rating":         "Location Rating",
    "Wattage":                 "Wattage",
    "Dimming Range":           "Dimming Range",
    "Dimming Type":            "Dimming",
    "CRI":                     "CRI",
    "LED Total Lumens":        "LED Total Lumens",
    "Fixture Delivered Lumens": "Fixture Delivered Lumens",
    "LED Life Span":           "LED Life Span",
    "Lumens":                  "Lumens",
    "Brightness":              "Lumens",
    "External Wire Type":      "Wiring Type",
    "Wire Length":             "Wire Length",
    "Minimum Hang Height":     "Min Drop",
    "Hanging Length":          "Hanging Length",
    "Fixture Input Voltage":   "Voltage",
    "Voltage":                 "Voltage",
    "Illumination Direction":  "Illumination Direction",
    "Sloped Ceiling Adaptable": "Sloped Ceiling Adaptable",
    "Hanging Method":          "Hanging Method",
    "Mounting Style":          "Mounting",
    "Adjustable":              "Adjustable",
    "Material":                "Material",
    "Glass Details":           "Glass Details",
    "Bulbs Included":          "Bulbs Included",
    "Number Of Bulbs":         "Bulb Qty",
    "Bulb Shape":              "Bulb Type",
    "Bulb Base Size":          "Socket",
    "Bulb Wattage":            "Bulb Wattage",
    "ETL Compliant":           "ETL Compliant",
    "UL Certified":            "UL Certified",
    "ADA Compliant":           "ADA Compliant",
    "Country Of Origin":       "Origin",
    "Height from center":      "Height from center",
}

# JS: open a named <details> section and return its inner table as {key: value}
_GET_SPEC_TABLE_JS = """
(heading) => {
    const result = {};
    // Find <details> whose <summary> text starts with `heading`
    const allDetails = document.querySelectorAll('details');
    for (const d of allDetails) {
        const s = d.querySelector('summary');
        if (!s) continue;
        if (!s.textContent.trim().toUpperCase().startsWith(heading.toUpperCase())) continue;
        d.open = true;
        const rows = d.querySelectorAll('tr');
        for (const row of rows) {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 2) {
                const key = cells[0].textContent.trim();
                const val = cells[1].textContent.trim();
                if (key && val && key.length < 80) result[key] = val;
            }
        }
        return result;
    }
    // Fallback: look for any element with a heading child matching the label
    const headingEls = document.querySelectorAll('h2, h3, h4, [class*="accordion"], [class*="section-title"]');
    for (const h of headingEls) {
        if (!h.textContent.trim().toUpperCase().startsWith(heading.toUpperCase())) continue;
        const container = h.parentElement;
        if (!container) continue;
        const rows = container.querySelectorAll('tr');
        for (const row of rows) {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 2) {
                const key = cells[0].textContent.trim();
                const val = cells[1].textContent.trim();
                if (key && val && key.length < 80) result[key] = val;
            }
        }
        if (Object.keys(result).length) return result;
    }
    return result;
}
"""

# JS: open DOWNLOADS <details> and collect PDF links
_GET_DOWNLOADS_JS = """
() => {
    const result = {};
    const allDetails = document.querySelectorAll('details');
    for (const d of allDetails) {
        const s = d.querySelector('summary');
        if (!s) continue;
        if (!s.textContent.trim().toUpperCase().startsWith('DOWNLOAD')) continue;
        d.open = true;
        const links = d.querySelectorAll('a[href]');
        for (const a of links) {
            const href = a.href || '';
            const text = a.textContent.trim().toLowerCase();
            if (!href) continue;
            if (text.includes('spec') && !result.spec_sheet) {
                result.spec_sheet = href;
            } else if ((text.includes('install') || text.includes('instruction')) && !result.install_guide) {
                result.install_guide = href;
            } else if (href.toLowerCase().includes('.pdf') && !result.spec_sheet) {
                result.spec_sheet = href;
            }
        }
        break;
    }
    return result;
}
"""


def _clean_inches(val: str) -> str:
    """Strip inch marks from a value string so dimensions stay numeric-friendly."""
    return val.replace('"', '').replace('”', '').strip()


async def scrape_product(page, url: str) -> list[dict]:
    """Return one dict per finish variant for the given product URL."""
    base_url = url.split("?")[0]
    handle = base_url.rstrip("/").split("/products/")[-1]

    # --- Fetch Shopify variant data (price in cents, images, SKU, finish) ---
    try:
        pj = await page.evaluate("""
            async (handle) => {
                const res = await fetch('/products/' + handle + '.js');
                if (!res.ok) return {};
                return await res.json();
            }
        """, handle)
    except Exception:
        pj = {}

    variants  = pj.get("variants", [])
    images    = pj.get("images", [])
    options   = pj.get("options", [])
    body_html = pj.get("body_html") or ""

    option_name = (options[0].get("name") if options else None) or "Finish"

    # --- Shared fields ---
    shared = {"Manufacturer": VENDOR_NAME}

    title = clean_text(pj.get("title") or "")
    if title:
        shared["Product Name"] = title
        shared["Product Family Id"] = extract_family_id(title)

    # Description: strip HTML tags from body_html, keep first paragraph(s)
    if body_html:
        desc_raw = re.sub(r"<[^>]+>", " ", body_html)
        desc_raw = re.sub(r"\s+", " ", desc_raw).strip()
        if desc_raw:
            shared["Description"] = desc_raw[:800]

    # Fallback image (first product image)
    if images:
        img0 = images[0]
        img_src = img0.get("src") if isinstance(img0, dict) else str(img0)
        if img_src:
            shared["Image URL"] = img_src if img_src.startswith("http") else ("https:" + img_src)

    # --- Visit page for specs and downloads ---
    await page.goto(base_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    raw_specs = await page.evaluate(_GET_SPEC_TABLE_JS, "SPECIFICATIONS")
    raw_specs = raw_specs or {}

    # Map spec keys → standard names
    specs = {}
    for raw_key, raw_val in raw_specs.items():
        col = SPEC_KEY_MAP.get(raw_key, raw_key)
        specs[col] = clean_text(_clean_inches(raw_val))

    # Dimensions → parse sub-fields
    if specs.get("Dimensions"):
        parsed = parse_dimensions(specs["Dimensions"]) or {}
        for k, v in parsed.items():
            if v and not specs.get(k):
                specs[k] = v

    # Build Specifications summary string from key lighting fields
    spec_summary_keys = [
        "Color Temperature", "Primary Dimension", "Location Rating",
        "Wattage", "Voltage", "CRI", "Dimming", "Material",
    ]
    spec_parts = [f"{k}: {specs[k]}" for k in spec_summary_keys if specs.get(k)]
    if spec_parts:
        shared["Specifications"] = " | ".join(spec_parts)

    shared.update(specs)

    # Downloads
    downloads = await page.evaluate(_GET_DOWNLOADS_JS)
    downloads = downloads or {}
    if downloads.get("spec_sheet"):
        shared["Spec Sheet"] = downloads["spec_sheet"]
    if downloads.get("install_guide"):
        shared["Install Guide"] = downloads["install_guide"]

    # --- Build one row per variant ---
    rows = []

    for v in variants:
        row = dict(shared)
        sku_raw = (v.get("sku") or "").strip()
        if sku_raw and sku_raw.lower() != "default title":
            row["SKU"] = sku_raw

        price_cents = v.get("price") or 0
        if price_cents:
            row["Price"] = round(price_cents / 100, 2)

        row["Source URL"] = f"{base_url}?variant={v['id']}"

        opt_val = (v.get("option1") or "").strip()
        if opt_val and opt_val.lower() != "default title":
            row[option_name] = opt_val

        # Per-variant image (highest resolution via src_set or main src)
        vi = v.get("featured_image")
        if vi:
            src = vi.get("src") if isinstance(vi, dict) else str(vi)
            if src:
                row["Image URL"] = src if src.startswith("http") else ("https:" + src)

        rows.append(row)

    if not rows:
        row = dict(shared)
        row["Source URL"] = base_url
        rows.append(row)

    return rows


async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Use Shopify collection products JSON API for fast, complete listing retrieval.
    Navigate once to set cookie/domain context, then paginate via API.
    """
    # Extract collection handle from URL
    handle_part = listing_url.rstrip("/").split("/collections/")[-1].split("?")[0]

    # Navigate to set domain context (needed for same-origin fetch)
    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)

    links: list[str] = []
    seen: set[str] = set()
    page_num = 1

    while True:
        api_path = f"/collections/{handle_part}/products.json?limit=250&page={page_num}"
        try:
            handles = await page.evaluate("""
                async (path) => {
                    const res = await fetch(path);
                    if (!res.ok) return [];
                    const j = await res.json();
                    return (j.products || []).map(p => p.handle);
                }
            """, api_path)
        except Exception:
            handles = []

        if not handles:
            break

        for h in handles:
            if not h:
                continue
            full = f"{BASE_URL}/products/{h}"
            if full not in seen:
                seen.add(full)
                links.append(full)

        if len(handles) < 250:
            break

        page_num += 1
        await async_polite_delay()

    return links


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if SCRAPE_CATEGORIES:
        categories = [c for c in categories if c["name"] in SCRAPE_CATEGORIES]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

    # --- Checkpoint / resume ---
    ckpt = load_checkpoint() if not TEST_MODE else {}
    done_cats  = set(ckpt.get("done_categories", []))   # fully completed categories
    cur_cat    = ckpt.get("current_category")           # category in progress
    done_urls  = set(ckpt.get("done_urls", []))         # products already scraped

    # Replay already-scraped rows into the writer
    if ckpt and not TEST_MODE:
        saved = load_saved_rows()
        if saved:
            print(f"  [Resume] Replaying {len(saved)} saved rows from checkpoint…")
            for row in saved:
                cat_name = row.get("Category", "")
                writer.add_sheet(
                    cat_name,
                    row.get("Source URL", "").split("?")[0],
                    studio_columns=next(
                        (c["studio_columns"] for c in categories if c["name"] == cat_name), []
                    ),
                )
                writer.write_row(row, category_name=cat_name)

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            cat_name = cat["name"]

            # Skip fully-done categories when resuming
            if cat_name in done_cats:
                print(f"  [{cat_name}] already done — skipping")
                continue

            writer.add_sheet(
                cat_name,
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str] = set()
            queue: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        queue.append(u)

            if TEST_MODE:
                queue = queue[:TEST_MAX_PRODUCTS]

            # Skip URLs already scraped in a previous interrupted run
            remaining = [u for u in queue if u not in done_urls] if not TEST_MODE else queue
            skipped   = len(queue) - len(remaining)
            print(f"  [{cat_name}] {len(queue)} products found"
                  + (f" ({skipped} already done, resuming from #{skipped+1})" if skipped else ""))

            global_idx = skipped + 1
            for url in remaining:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat_name, global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat_name)
                        if not TEST_MODE:
                            append_row(row)
                        global_idx += 1
                except Exception as e:
                    print(f"  [!] {url}: {e}")

                # Save checkpoint after every product
                if not TEST_MODE:
                    done_urls.add(url)
                    save_checkpoint({
                        "done_categories": list(done_cats),
                        "current_category": cat_name,
                        "done_urls": list(done_urls),
                    })

                await async_polite_delay()

            # Mark this category as fully done
            done_cats.add(cat_name)
            if not TEST_MODE:
                save_checkpoint({
                    "done_categories": list(done_cats),
                    "current_category": None,
                    "done_urls": list(done_urls),
                })

    writer.save()
    print(f"[Done] Saved to {OUTPUT_PATH}")

    # Clean up checkpoint files on successful completion
    if not TEST_MODE:
        clear_checkpoint()
        print("[Checkpoint] Cleared — run completed successfully.")


if __name__ == "__main__":
    asyncio.run(main())
