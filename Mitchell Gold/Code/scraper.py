import asyncio
import json
import os
import re
import sys
from pathlib import Path

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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Mitchell Gold")
HEADLESS = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
SCRAPE_CATEGORIES = [
    c.strip()
    for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",")
    if c.strip()
]

BASE_URL = "https://mgbw.com"


async def _dismiss_overlays(page):
    await page.evaluate(
        """() => {
        document.querySelectorAll(
            '.ui-widget-overlay, .ui-dialog, [class*=cookie], .cookie-consent-dialog'
        ).forEach(el => el.remove());
    }"""
    )


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return deduplicated product URLs from a listing page, handling infinite scroll."""
    try:
        await page.goto(listing_url, timeout=60_000, wait_until="domcontentloaded")
    except Exception as e:
        print(f"  [WARN] Listing page load failed ({listing_url}): {e}")
        return []
    await page.wait_for_timeout(2500)
    await _dismiss_overlays(page)

    prev_count = 0
    stall_rounds = 0
    while True:
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(1800)
        tiles = await page.query_selector_all(".product-tile")
        count = len(tiles)
        if count == prev_count:
            stall_rounds += 1
            if stall_rounds >= 2:
                break
        else:
            stall_rounds = 0
        prev_count = count

    tiles = await page.query_selector_all(".product-tile")
    seen: set[str] = set()
    links: list[str] = []
    for tile in tiles:
        a = await tile.query_selector("a[href]")
        if not a:
            continue
        href = await a.get_attribute("href")
        if not href:
            continue
        # Build absolute URL and strip query string for dedup
        if href.startswith("/"):
            href = BASE_URL + href
        base = href.split("?")[0]
        if base not in seen:
            seen.add(base)
            links.append(base)
    return links


def _parse_details_bullets(bullets: list[str]) -> dict:
    """Extract structured fields from detail bullet text (handles Key: Value format)."""
    data: dict = {}
    specs: list[str] = []
    care_parts: list[str] = []

    for bullet in bullets:
        bullet = clean_text(bullet)
        if not bullet:
            continue
        bl = bullet.lower()

        # Key: Value structured bullets (e.g. "Materials: Metal, Glass")
        kv_match = re.match(r"^([A-Za-z][A-Za-z &/]+?):\s*(.+)$", bullet)
        if kv_match:
            key_raw = kv_match.group(1).strip()
            val = kv_match.group(2).strip()
            key_lc = key_raw.lower()

            if key_lc in ("materials", "material"):
                data.setdefault("Materials", val)
                continue
            if key_lc == "finish":
                # Available finishes list — only store if no selected finish yet
                data.setdefault("Finish", val)
                continue
            if key_lc in ("care", "care instructions"):
                care_parts.append(val)
                continue
            if key_lc == "dimensions":
                data.setdefault("Dimensions", val)
                continue
            if key_lc in ("com approximate requirements", "com requirement", "com"):
                data.setdefault("COM", bullet)
                continue
            if key_lc == "origin":
                data.setdefault("Origin", val)
                continue
            # Any other Key: Value → add as-is to specs
            specs.append(bullet)
            continue

        # COM yardage (non-structured)
        if "com approximate" in bl or "com order" in bl or "com requirement" in bl:
            data.setdefault("COM", bullet)
            continue

        # Care instructions (non-structured)
        if any(k in bl for k in ["wipe clean", "dust with", "coasters recommended", "spray glass"]):
            care_parts.append(bullet)
            continue

        # Origin
        if "north carolina" in bl or "made in" in bl or "handcrafted in" in bl or "proudly crafted" in bl:
            data.setdefault("Origin", re.sub(r"(?i)proudly\s+crafted\s+by\s+hand\s+in\s+", "", bullet).strip())
            continue

        # Material / construction clues (non-structured)
        if any(k in bl for k in ["veneer", "hardwood", "wood frame", "engineered wood"]):
            data.setdefault("Material", bullet)
            continue

        specs.append(bullet)

    if care_parts:
        data["Care Instructions"] = " | ".join(care_parts)
    if specs:
        data["Specifications"] = " | ".join(specs)
    return data


async def scrape_product(page, url: str) -> list[dict]:
    """Return a list containing one dict for this product/variant listing."""
    base: dict = {"Source URL": url}

    for attempt in range(2):
        try:
            await page.goto(url, timeout=60_000, wait_until="domcontentloaded")
            break
        except Exception:
            if attempt == 1:
                raise
            await page.wait_for_timeout(3000)
    await page.wait_for_timeout(2200)
    await _dismiss_overlays(page)

    # --- DataLayer: name, price, SKU, Product Family Id ---
    try:
        dl_raw = await page.evaluate("() => JSON.stringify(window.dataLayer)")
        dl = json.loads(dl_raw)
        for entry in dl:
            ec = entry.get("ecommerce", {})
            products = ec.get("detail", {}).get("products", [])
            if products:
                p = products[0]
                base["Product Name"] = clean_text(p.get("name", ""))
                raw_sku = p.get("variant", "")
                base["SKU"] = raw_sku
                raw_price = p.get("price")
                if raw_price is not None:
                    base["Price"] = float(raw_price)
                master_id = p.get("id", "")
                if master_id:
                    base["Product Family Id"] = master_id
                break
    except Exception:
        pass

    # --- Product name fallback ---
    if not base.get("Product Name"):
        el = await page.query_selector("h1")
        if el:
            base["Product Name"] = clean_text(await el.inner_text())

    # --- Image: og:image meta (highest res hero) ---
    img_meta = await page.query_selector('meta[property="og:image"]')
    if img_meta:
        base["Image URL"] = await img_meta.get_attribute("content") or ""

    # --- Description ---
    desc_el = await page.query_selector(".tab-short-description")
    if desc_el:
        base["Description"] = clean_text(await desc_el.inner_text())

    # --- Lead time (shipping estimate shown on page) ---
    ship_el = await page.query_selector(".estimated-shipping .value")
    if ship_el:
        lt = clean_text(await ship_el.inner_text())
        if lt:
            base["Lead Time"] = lt

    # --- Finish (furniture variants) ---
    finish_el = await page.query_selector(".fabric-name.selected .name")
    if finish_el:
        finish = clean_text(await finish_el.inner_text())
        if finish and finish.lower() not in ("please select a finish", ""):
            base["Finish"] = finish

    # --- Fabric / Upholstery (seating: "Shown In: Fabric, Fill") ---
    shown_el = await page.query_selector(".shownin-content span")
    if shown_el:
        shown = clean_text(await shown_el.inner_text())
        if shown:
            parts = [p.strip() for p in shown.split(",")]
            if parts:
                base["Fabric"] = parts[0]
            if len(parts) >= 2:
                base["Upholstery"] = ", ".join(parts[1:])

    # --- Dimensions list (structured) ---
    dim_items = await page.query_selector_all(".dimensions-list li")
    dim_raw: dict[str, str] = {}
    for item in dim_items:
        divs = await item.query_selector_all("div")
        if len(divs) >= 2:
            label = clean_text(await divs[0].inner_text())
            value = clean_text(await divs[1].inner_text())
            if label and value:
                dim_raw[label] = value

    # Map dimension labels to field names
    label_map = {
        "Dimensions": "Dimensions",
        "Overall Height": "Height",
        "Overall Width": "Width",
        "Overall Depth": "Depth",
        "Overall Diameter": "Diameter",
        "Seat Height": "Seat Height",
        "Seat Depth": "Seat Depth",
        "Seat Width": "Seat Width",
        "Seat Length": "Seat Length",
        "Arm Height": "Arm Height",
        "Frame Height": "Height",  # fallback
        "Overall Length": "Length",
    }
    for label, value in dim_raw.items():
        field = label_map.get(label)
        if field and field not in base:
            base[field] = value

    # Build Dimensions string if not already set
    if "Dimensions" not in base:
        parts = []
        for lbl, fld in [("Overall Width", "W"), ("Overall Depth", "D"), ("Overall Height", "H")]:
            if lbl in dim_raw:
                parts.append(f"{dim_raw[lbl]}{fld}")
        if parts:
            base["Dimensions"] = " x ".join(parts)

    # --- Available sizes / widths (seating: "Available widths: 75 85 95 105") ---
    size_el = await page.query_selector(".attribute.size .avilable-size")
    if size_el:
        size_txt = clean_text(await size_el.inner_text())
        if size_txt:
            base["Size"] = size_txt

    # --- Details bullets (second print-column in design tab) ---
    bullets_raw: list[str] = []
    design_cols = await page.query_selector_all(".tab-design .print-column")
    if len(design_cols) >= 2:
        bullet_els = await design_cols[1].query_selector_all("li")
        for li in bullet_els:
            bullets_raw.append(await li.inner_text())

    parsed = _parse_details_bullets(bullets_raw)
    for k, v in parsed.items():
        if k not in base:
            base[k] = v

    # --- Collection (from URL or product family name inference) ---
    # MGBW uses family names as collections (e.g., MALIBU, GISELLE)
    pf = base.get("Product Family Id", "")
    if pf and "Collection" not in base:
        # Extract first word group as collection name
        col = re.sub(r"[_\-]+", " ", pf).strip().title()
        # Remove trailing size/type words
        col = re.sub(r"\s+(Sofa|Chair|Table|Nightstand|Desk|Ottoman|Bench|Sectional|Stool|Bed|Dresser|Console|Bookcase|Cabinet|Pillow|Loveseat|Sleeper|Lounge|Swivel|Recliner|Accent).*$", "", col, flags=re.IGNORECASE).strip()
        if col:
            base["Collection"] = col

    base["Manufacturer"] = VENDOR_NAME

    return [base]


async def main():
    info = json.loads(
        (Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8")
    )
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if SCRAPE_CATEGORIES:
        categories = [c for c in categories if c["name"] in SCRAPE_CATEGORIES]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

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
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(
                f"[{VENDOR_NAME}] {cat['name']}: {len(all_product_urls)} products"
            )

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(
                                VENDOR_NAME, cat["name"], global_idx
                            )
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(
                                row["Product Name"]
                            )
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [WARN] Failed {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"[{VENDOR_NAME}] Done → {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
