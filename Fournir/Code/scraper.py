import asyncio, json, os, sys, re, html
from html.parser import HTMLParser
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    safe_float,
)

VENDOR_NAME       = os.environ.get("VENDOR_NAME", "Fournir")
HEADLESS          = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH       = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE         = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))

IMAGES_BASE = "https://imagescdn.dessinfournir.com/High"
SITE_BASE   = "https://www.dessinfournir.com"

COMPANY_NAMES = {
    "DF":  "Dessin Fournir",
    "GD":  "Gerard",
    "TH":  "Therien",
    "QU":  "Quatrain",
    "EB":  "Erika Brunson",
    "HPP": "HPP",
    "KM":  "KM",
    "DFL": "Dessin Fournir Lighting",
}


def strip_html(raw: str) -> str:
    """Convert <br> to space, strip all other HTML tags, decode entities."""
    if not raw:
        return ""
    # Convert <br> variants to space
    s = re.sub(r"<br\s*/?>", " ", raw, flags=re.IGNORECASE)
    # Strip remaining tags
    s = re.sub(r"<[^>]+>", "", s)
    # Decode entities
    s = html.unescape(s)
    # Collapse whitespace
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s


def clean_dim_value(raw) -> str:
    """Decode HTML entities, remove inch marks, convert mixed fractions."""
    if not raw:
        return ""
    v = html.unescape(str(raw)).replace('"', "").strip()
    # Convert mixed fractions: 22 1/2 → 22.5
    def _repl(m):
        whole = int(m.group(1))
        num   = int(m.group(2))
        den   = int(m.group(3))
        return str(round(whole + num / den, 4)).rstrip("0").rstrip(".")
    v = re.sub(r"(\d+)\s+(\d+)/(\d+)", _repl, v)
    # Standalone fraction: 1/2 → 0.5
    v = re.sub(r"^(\d+)/(\d+)$", lambda m: str(round(int(m.group(1)) / int(m.group(2)), 4)), v)
    return v.strip()


async def get_product_oids(page, listing_url: str, max_oids: int = 0) -> list[int]:
    """
    Load a listing URL, scroll-load all products, return unique OIDs.
    If max_oids > 0, stop scrolling once we have enough.
    """
    await page.goto(listing_url, timeout=60000, wait_until="domcontentloaded")
    await page.wait_for_timeout(5000)

    prev_count = 0
    stall = 0
    while True:
        oids_js = await page.eval_on_selector_all(
            'a[href*="/item/"]',
            "els => [...new Set(els.map(e => { const m = e.href.match(/item\\/(\\d+)/); return m ? parseInt(m[1]) : null; }).filter(Boolean))]"
        )
        cur = len(oids_js)
        if max_oids > 0 and cur >= max_oids:
            break
        if cur == prev_count:
            stall += 1
            if stall >= 3:
                break
        else:
            stall = 0
        prev_count = cur
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(1500)

    all_oids = await page.eval_on_selector_all(
        'a[href*="/item/"]',
        "els => [...new Set(els.map(e => { const m = e.href.match(/item\\/(\\d+)/); return m ? parseInt(m[1]) : null; }).filter(Boolean))]"
    )
    return all_oids


async def get_item_detail(page, oid: int) -> dict:
    return await page.evaluate(f"""async () => {{
        const r = await fetch('/api/v1/item/{oid}');
        if (!r.ok) return {{}};
        return await r.json();
    }}""")


def map_item_to_row(detail: dict, cat_name: str) -> dict:
    oid = detail.get("oid")
    row = {
        "Manufacturer": VENDOR_NAME,
        "Source URL":   f"{SITE_BASE}/#/item/{oid}",
        "Image URL":    f"{IMAGES_BASE}/{oid}.jpg",
        "Product Name": clean_text(detail.get("item", "")),
        "SKU":          detail.get("displayItemnum", ""),
    }

    # Sub-brand
    company = detail.get("company", "")
    brand = COMPANY_NAMES.get(company, company)
    if brand and brand.lower() != VENDOR_NAME.lower():
        row["Brand"] = brand

    row["Product Family Id"] = extract_family_id(row["Product Name"])

    if detail.get("collection"):
        row["Collection"] = clean_text(detail["collection"])

    if detail.get("material"):
        row["Material"] = strip_html(detail["material"])
    if detail.get("finish"):
        row["Finish"] = strip_html(detail["finish"])

    # Description — combine misc (notes about seat, upholstery) + notes (pricing notes)
    desc_parts = []
    if detail.get("misc"):
        desc_parts.append(strip_html(detail["misc"]))
    if detail.get("notes"):
        desc_parts.append(strip_html(detail["notes"]))
    if desc_parts:
        row["Description"] = "\n".join(desc_parts)

    # Dimensions
    w   = clean_dim_value(detail.get("width"))
    d   = clean_dim_value(detail.get("depth"))
    h   = clean_dim_value(detail.get("height"))
    dia = clean_dim_value(detail.get("diameter"))
    lng = clean_dim_value(detail.get("length"))

    if w:   row["Width"]    = w
    if d:   row["Depth"]    = d
    if h:   row["Height"]   = h
    if dia: row["Diameter"] = dia
    if lng: row["Length"]   = lng

    # Build Dimensions string
    dim_parts = []
    if w:   dim_parts.append(f'W: {w}"')
    if d:   dim_parts.append(f'D: {d}"')
    if h:   dim_parts.append(f'H: {h}"')
    if dia: dim_parts.append(f'Dia: {dia}"')
    if lng: dim_parts.append(f'L: {lng}"')
    if dim_parts:
        row["Dimensions"] = " x ".join(dim_parts)
    dim_misc = strip_html(detail.get("dim_misc", ""))
    if dim_misc:
        # Append dimension notes as a separate Specifications entry
        row["Specifications"] = dim_misc

    # Seating
    sh = clean_dim_value(detail.get("seatheight"))
    sd = clean_dim_value(detail.get("seatdepth"))
    sl = clean_dim_value(detail.get("seatlength"))
    ah = clean_dim_value(detail.get("armheight"))
    if sh: row["Seat Height"] = sh
    if sd: row["Seat Depth"]  = sd
    if sl: row["Seat Length"] = sl
    if ah: row["Arm Height"]  = ah

    # COM / COL
    if detail.get("com"):  row["COM"] = clean_text(detail["com"])
    if detail.get("col"):  row["COL"] = clean_text(detail["col"])

    # Lighting
    if detail.get("bulb_quantity"): row["Bulb Qty"] = str(detail["bulb_quantity"])
    if detail.get("wattage"):       row["Wattage"]  = clean_text(detail["wattage"])

    # Bed components
    if clean_dim_value(detail.get("headboard")):
        row["Headboard"] = clean_dim_value(detail["headboard"])
    if clean_dim_value(detail.get("footboard")):
        row["Footboard"] = clean_dim_value(detail["footboard"])
    if clean_dim_value(detail.get("footrail")):
        row["Footrail"]  = clean_dim_value(detail["footrail"])

    # Lead time
    if detail.get("lead_time"):
        row["Lead Time"] = clean_text(detail["lead_time"])

    return row


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    cats = info["categories"]
    if TEST_MODE:
        cats = cats[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        # Establish session on the main site
        await page.goto(SITE_BASE, timeout=60000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        for cat in cats:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_oids: set = set()
            all_oids: list = []

            max_per_link = TEST_MAX_PRODUCTS if TEST_MODE else 0
            for listing_url in cat["links"]:
                oids = await get_product_oids(page, listing_url, max_oids=max_per_link)
                for oid in oids:
                    if oid not in seen_oids:
                        seen_oids.add(oid)
                        all_oids.append(oid)

            if TEST_MODE:
                all_oids = all_oids[:TEST_MAX_PRODUCTS]

            print(f"[{cat['name']}] {len(all_oids)} products")

            global_idx = 1
            for oid in all_oids:
                detail = await get_item_detail(page, oid)
                if not detail or "oid" not in detail:
                    print(f"  WARNING: no detail for oid={oid}")
                    continue

                row = map_item_to_row(detail, cat["name"])

                if not row.get("SKU"):
                    row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                if not row.get("Product Family Id") and row.get("Product Name"):
                    row["Product Family Id"] = extract_family_id(row["Product Name"])

                writer.write_row(row, category_name=cat["name"])
                global_idx += 1
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
