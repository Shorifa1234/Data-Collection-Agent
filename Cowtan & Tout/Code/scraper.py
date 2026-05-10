import asyncio, json, os, re, sys
import requests
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text, clean_price, generate_sku, extract_family_id,
    async_polite_delay,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Cowtan & Tout")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://designs.cowtan.com"
IMAGE_BASE = "https://d2mq91o692rj7w.cloudfront.net/flatshots"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Accept": "application/json",
}


def _url_safe(stock_code: str) -> str:
    return stock_code.replace("/", "-")


def _clean_inches(val) -> str | None:
    """Strip inch marks and return cleaned string, or None if empty."""
    if val is None:
        return None
    cleaned = str(val).replace('"', "").replace("in", "").strip()
    return cleaned if cleaned else None


def _fetch_search_page(session: requests.Session, type_code: str, page_index: int) -> dict:
    r = session.post(
        f"{BASE_URL}/api/ProductSearch",
        json={
            "Types": [{"Value": type_code, "Selected": True}],
            "Brands": [],
            "Colours": [],
            "Categories": [],
            "Keywords": "",
            "NewOnly": {"Selected": False, "Value": "NewOnly"},
            "PageIndex": page_index,
        },
        headers=HEADERS,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def _fetch_product_detail(session: requests.Session, stock_code: str) -> list[dict]:
    r = session.get(
        f"{BASE_URL}/api/Product/Detail/{stock_code}",
        headers=HEADERS,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def _build_row(cw: dict) -> dict:
    sc = cw.get("StockCode") or ""
    safe = _url_safe(sc)

    row = {
        "Manufacturer": VENDOR_NAME,
        "Brand":        clean_text(cw.get("BrandName") or ""),
        "Source":       f"{BASE_URL}/Design/{safe}",
        "Image URL":    f"{IMAGE_BASE}/2017-main/{safe}_m.jpg",
        "Product Name": clean_text(cw.get("ProductName") or ""),
        "SKU":          sc,
        "Color":        clean_text(cw.get("ColourName") or ""),
        "Pattern Book": clean_text(cw.get("PatternBook") or ""),
        "Tearsheet Link": f"{BASE_URL}/Product/{safe}.pdf",
    }

    # Family ID — ProductName is already the base name (colorway is in ColourName)
    name = row.get("Product Name") or ""
    row["Product Family Id"] = extract_family_id(name) if name else None

    # Dimensions / textile specs
    if cw.get("Width"):
        row["Width"] = _clean_inches(cw["Width"])
    if cw.get("Weight"):
        row["Weight"] = _clean_inches(cw["Weight"])
    if cw.get("RepeatV"):
        row["Vertical Repeat"] = _clean_inches(cw["RepeatV"])
    if cw.get("RepeatH"):
        row["Horizontal Repeat"] = _clean_inches(cw["RepeatH"])

    # Composition/contents (list → joined string)
    comp = cw.get("Composition")
    if comp:
        joined = ", ".join(c.strip() for c in comp if c and c.strip())
        if joined:
            row["Contents"] = joined

    # Price (numeric in API — already a number or null)
    if cw.get("Price") is not None:
        row["Price"] = clean_price(str(cw["Price"]))

    # Conditionally present fields
    for field, col in [
        ("Testing",      "Testing"),
        ("Qualities",    "Qualities"),
        ("Availability", "Availability"),
        ("Finish",       "Finish"),
    ]:
        val = cw.get(field)
        if val:
            row[col] = clean_text(str(val))

    # In-stock status
    if cw.get("InStock") is not None:
        row["In Stock"] = "Yes" if cw["InStock"] else "No"

    # Coordinate products (related fabric/wallcovering companions)
    coord = cw.get("Coordinate")
    if coord:
        row["Coordinate"] = ", ".join(str(c) for c in coord)

    # Remove None / empty values
    return {k: v for k, v in row.items() if v is not None and v != ""}


async def _scrape_category(session: requests.Session, cat: dict, writer: ExcelWriter) -> None:
    url = cat["links"][0]
    m = re.search(r"/Search/([A-Z])/", url)
    type_code = m.group(1) if m else "F"

    if TEST_MODE:
        print(f"  [TEST: max {TEST_MAX_PRODUCTS} products per category]")

    # ── Collect unique ProductCodes from search API ──────────────────────────
    seen_product_codes: set[str] = set()
    stock_codes_to_scrape: list[str] = []
    page_index = 0
    total_pages = 1

    while page_index < total_pages:
        try:
            result = await asyncio.to_thread(_fetch_search_page, session, type_code, page_index)
        except Exception as e:
            print(f"  Search page {page_index} error: {e}")
            break

        total_pages = result["Result"]["TotalPages"]

        for entry in result["Result"]["Products"]:
            pc = entry.get("ProductCode") or entry.get("StockCode", "")
            if pc not in seen_product_codes:
                seen_product_codes.add(pc)
                stock_codes_to_scrape.append(entry["StockCode"])

            if TEST_MODE and len(stock_codes_to_scrape) >= TEST_MAX_PRODUCTS:
                break

        if TEST_MODE and len(stock_codes_to_scrape) >= TEST_MAX_PRODUCTS:
            break

        page_index += 1
        await asyncio.sleep(0.1)

    label = " [TEST]" if TEST_MODE else f" ({page_index} pages)"
    print(f"  {cat['name']}: {len(stock_codes_to_scrape)} unique products{label}")

    # ── Fetch full detail per product and write all colorways ─────────────────
    global_idx = 1
    for sc in stock_codes_to_scrape:
        try:
            colorways = await asyncio.to_thread(_fetch_product_detail, session, sc)
            for cw in colorways:
                row = _build_row(cw)
                if not row.get("SKU"):
                    row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                if not row.get("Product Family Id") and row.get("Product Name"):
                    row["Product Family Id"] = extract_family_id(row["Product Name"])
                writer.write_row(row, category_name=cat["name"])
                global_idx += 1
        except Exception as e:
            print(f"  Error on {sc}: {e}")

        await asyncio.sleep(0.3)

    print(f"  {cat['name']}: wrote {global_idx - 1} rows total")


async def main() -> None:
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    session = requests.Session()

    categories = info["categories"]
    if TEST_MODE and TEST_MAX_CATEGORIES < len(categories):
        categories = categories[:TEST_MAX_CATEGORIES]

    for cat in categories:
        if not cat["links"]:
            continue
        writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])
        await _scrape_category(session, cat, writer)

    session.close()
    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
