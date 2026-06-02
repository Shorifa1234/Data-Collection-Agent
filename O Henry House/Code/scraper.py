import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    generate_sku, extract_family_id,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "O Henry House")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://ohenryhouseltd.com"

CATEGORY_FOLDERS = {
    "Dining Chairs":    "dining",
    "Sofas & Loveseats": "sofas",
    "Lounge Chairs":    "chairs",
    "Ottomans":         "ottomans",
}


def _parse_ohh_dims(text: str) -> dict:
    """Parse O Henry House dimension strings like '20"W 24.5"D 39"H' or '90"W × 36"D × 34"H'."""
    result = {}
    pat = re.compile(r'(\d+(?:\.\d+)?)\s*(?:"|in)?\s*([WwDdHhLl]|Dia)\b')
    for m in pat.finditer(text):
        val, label = m.group(1), m.group(2).upper()
        if label == "W":
            result.setdefault("Width", val)
        elif label == "D":
            result.setdefault("Depth", val)
        elif label == "H":
            result.setdefault("Height", val)
        elif label == "L":
            result.setdefault("Length", val)
        elif label == "DIA":
            result.setdefault("Diameter", val)
    return result


def _extract_sku(url: str) -> str:
    """Extract model number from URL slug: '/dining/detail/104-sterling' → '104'."""
    slug = url.rstrip("/").split("/")[-1]
    m = re.match(r"^([a-zA-Z]{0,4}\d+[a-zA-Z]?)", slug)
    return m.group(1) if m else slug.split("-")[0]


async def get_product_links(page, listing_url: str) -> list[str]:
    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2_000)

    hrefs = await page.eval_on_selector_all(
        "a[href*='/detail/']",
        "els => els.map(el => el.getAttribute('href'))"
    )

    seen: set[str] = set()
    links: list[str] = []
    for href in hrefs:
        if not href or "/detail/" not in href:
            continue
        full = href if href.startswith("http") else BASE_URL + href
        base = full.split("?")[0]
        if base not in seen:
            seen.add(base)
            links.append(base)
    return links


async def scrape_product(page, url: str, category_name: str) -> list[dict]:
    base: dict = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2_000)

    # --- Product Name ---
    try:
        raw_name = await page.eval_on_selector("h3", "el => el.textContent.trim()")
        # Names start with a number (e.g. "104 Sterling") — use title() so both
        # the word and the number prefix are preserved correctly.
        base["Product Name"] = clean_text(raw_name).title()
    except Exception:
        base["Product Name"] = ""

    # --- SKU ---
    base["SKU"] = _extract_sku(url)
    base["Product Family Id"] = extract_family_id(base["Product Name"])

    # --- Image URL (derived from URL pattern) ---
    folder = CATEGORY_FOLDERS.get(category_name, category_name.lower())
    sku = base["SKU"]
    base["Image URL"] = f"{BASE_URL}/img/products/{folder}/hero/{sku}.jpg"

    # --- Full page text for regex-based parsing ---
    body_text = await page.eval_on_selector("body", "el => el.innerText")

    # --- Overall dimensions ---
    overall_m = re.search(r'OVERALL[:\s]*([^\n]+)', body_text, re.IGNORECASE)
    if overall_m:
        raw_overall = overall_m.group(1).strip()
        # Clean for Dimensions field (remove inch marks)
        base["Dimensions"] = re.sub(r'"', '', raw_overall).strip()
        base.update(_parse_ohh_dims(raw_overall))

    # --- Seat-specific dimensions ---
    for label, key in [
        (r'SEAT\s+HEIGHT', "Seat Height"),
        (r'SEAT\s+DEPTH',  "Seat Depth"),
        (r'ARM\s+HEIGHT',  "Arm Height"),
        (r'ARM\s+WIDTH',   "Arm Width"),
        (r'COUNTER\s+HEIGHT', "Counter Height"),
    ]:
        m = re.search(rf'{label}[:\s]*([\d.]+)"?', body_text, re.IGNORECASE)
        if m:
            base[key] = m.group(1)

    # --- COM yardage ---
    com_m = re.search(r'\bCOM[:\s]*([\d.]+)\s*YARDS?', body_text, re.IGNORECASE)
    if com_m:
        base["COM Yardage"] = com_m.group(1)

    # --- Description ---
    # Design features text (e.g. "Tufted Back, Tight Seat, Tapered Legs") appears
    # before the dimension block. Filter out <p> tags that look like spec blocks.
    description = ""
    try:
        p_texts = await page.eval_on_selector_all(
            "p",
            "els => els.map(el => el.innerText.trim()).filter(t => t.length > 10)"
        )
        candidates = [
            t for t in p_texts
            if not re.search(
                r'cookie|privacy|copyright|\bsign in\b|@|OVERALL|COM:|SEAT HEIGHT|SEAT DEPTH|ARM HEIGHT',
                t, re.IGNORECASE
            )
            and len(t) > 10
        ]
        if candidates:
            description = candidates[0]
    except Exception:
        pass

    if description:
        base["Description"] = clean_text(description)

    # --- Finishes (wood/leg finish options) ---
    try:
        finish_alts = await page.eval_on_selector_all(
            "img[src*='finishes/thumb']",
            "els => els.map(el => el.getAttribute('alt')).filter(Boolean)"
        )
        if finish_alts:
            base["Finish"] = ", ".join(finish_alts)
    except Exception:
        pass

    # Tearsheet PDFs follow the pattern /pdf/{sku}.pdf for all products
    base["Tearsheet Link"] = f"{BASE_URL}/pdf/{sku}.pdf"

    return [base]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
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

            global_idx = 1
            for url in all_product_urls:
                variant_rows = await scrape_product(page, url, cat["name"])
                for variant in variant_rows:
                    if not variant.get("SKU"):
                        variant["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not variant.get("Product Family Id") and variant.get("Product Name"):
                        variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                    writer.write_row(variant, category_name=cat["name"])
                    global_idx += 1
                await async_polite_delay()

    writer.save()


if __name__ == "__main__":
    asyncio.run(main())
