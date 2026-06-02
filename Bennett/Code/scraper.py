import asyncio, json, os, sys, re, requests
from pathlib import Path
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    async_polite_delay, clean_text, clean_price,
    generate_sku,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Bennett")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://www.bennetttothetrade.com"
SESSION  = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
})


def get_family_id(sku: str) -> str:
    """Strip finish suffix — base is 2-4 letters + digits, suffix must start with a letter.
    BIZ1002TR → BIZ1002, BIZ1002 → BIZ1002, EFM6001 → EFM6001."""
    m = re.match(r'^([A-Z]{2,4}\d+)([A-Z].*)$', sku)
    return m.group(1) if m else sku


def parse_body(html: str) -> dict:
    """Extract dimensions, COM yardage, seat height, finish, description from body_html."""
    if not html:
        return {}
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator="\n").strip()
    data = {}

    # Dimensions: "59"x 42"x 20"H" or "20"x 18.5"x 32" (19" to seat)"
    # Pattern: num"  x  num"  x  num"[H]
    dim_pat = re.compile(
        r'(\d[\d.]*)\s*["”]\s*[xX]\s*(\d[\d.]*)\s*["”]\s*[xX]\s*(\d[\d.]*)\s*["”]?H?',
        re.I,
    )
    m = dim_pat.search(text)
    if m:
        w, d, h = m.group(1), m.group(2), m.group(3)
        data["Width"]      = w
        data["Depth"]      = d
        data["Height"]     = h
        data["Dimensions"] = f'{w}" x {d}" x {h}"H'

        # Seat height: "(19" to seat)"
        after = text[m.end():]
        seat  = re.search(r'\((\d[\d.]*)["”]?\s*to\s*seat\)', after, re.I)
        if seat:
            data["Seat Height"] = seat.group(1)

    # COM yardage: "COM- 3 yards"
    com = re.search(r'(?:COM\s*[-–]\s*)?(\d+(?:\.\d+)?)\s+yards?', text, re.I)
    if com:
        data["COM"] = com.group(1)

    # Finish: "Available in the following finish:\n\nBianco with Charcoal trim"
    fin = re.search(r'Available in the following finish[:\s]*\n+\s*(.+)', text, re.I)
    if not fin:
        # "w/ Antique Glass"
        fin = re.search(r'\bw/\s*(.+)', text, re.I)
    if fin:
        val = fin.group(1).strip()
        if val:
            data["Finish"] = val

    # Description — full cleaned text from body
    full = clean_text(text)
    if full:
        data["Description"] = full

    return data


def fetch_collection(collection_handle: str) -> list[dict]:
    """Return all products from a Shopify collection via JSON API."""
    products, page = [], 1
    while True:
        url  = f"{BASE_URL}/collections/{collection_handle}/products.json?limit=250&page={page}"
        resp = SESSION.get(url, timeout=30)
        if resp.status_code != 200:
            break
        batch = resp.json().get("products", [])
        if not batch:
            break
        products.extend(batch)
        if len(batch) < 250:
            break
        page += 1
    return products


def collection_handle(url: str) -> str:
    return url.rstrip("/").split("/collections/")[-1]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    test_mode   = os.environ.get("TEST_MODE",  "false").lower() == "true"
    test_limit  = int(os.environ.get("TEST_LIMIT", "5"))
    scrape_cats = [c.strip() for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",") if c.strip()]

    for cat in info["categories"]:
        if not cat["links"]:
            continue
        if scrape_cats and cat["name"] not in scrape_cats:
            continue

        writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])

        handle   = collection_handle(cat["links"][0])
        products = await asyncio.to_thread(fetch_collection, handle)

        if test_mode:
            products = products[:test_limit]

        print(f"  [{cat['name']}] {len(products)} products")

        for idx, product in enumerate(products, start=1):
            variant  = (product.get("variants") or [{}])[0]
            sku      = variant.get("sku") or product["title"]
            images   = product.get("images", [])

            row = {
                "Source URL":        f"{BASE_URL}/products/{product['handle']}",
                "Image URL":         images[0]["src"] if images else "",
                "Product Name":      product["title"],
                "SKU":               sku,
                "Product Family Id": get_family_id(sku),
                "Price":             clean_price(str(variant.get("price", ""))),
                "Manufacturer":      VENDOR_NAME,
            }

            row.update(parse_body(product.get("body_html", "")))

            raw_tags = product.get("tags", [])
            if isinstance(raw_tags, str):
                raw_tags = [t.strip() for t in raw_tags.split(",")]
            tags = [t.strip() for t in raw_tags if t.strip() and t.strip().lower() != "axll_enriched"]
            if tags:
                row["Tags"] = ", ".join(tags)

            if not row.get("SKU"):
                row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], idx)

            writer.write_row(row, category_name=cat["name"])
            await async_polite_delay()

    writer.save()


if __name__ == "__main__":
    asyncio.run(main())
