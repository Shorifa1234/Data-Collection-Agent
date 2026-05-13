import asyncio, json, os, sys, re
from pathlib import Path
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    generate_sku, extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Hickory White")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://hickorywhite.com"

# Number token: handles integers, decimals, fractions like "93 1/2"
_NUM = r"\d+(?:\s+\d+/\d+)?(?:\.\d+)?"


def _frac(s: str) -> str:
    """Convert '93 1/2' to '93.5', '39 1/4' to '39.25', '64' unchanged."""
    s = s.strip()
    if " " in s:
        parts = s.split()
        for i, p in enumerate(parts):
            if "/" in p and i > 0:
                try:
                    n, d = p.split("/")
                    return str(round(float(parts[0]) + float(n) / float(d), 4))
                except (ValueError, ZeroDivisionError):
                    pass
    return s


def parse_description_block(div) -> dict:
    """Extract structured fields from div.product-description-copy."""
    data = {}
    if not div:
        return data

    text = div.get_text(separator="\n", strip=True)
    data["Description"] = text

    # Variant SKU from "As Shown:" e.g. "113W-ABK"
    m = re.search(r"As Shown[:\s]*\n?([A-Z0-9]+[-_][A-Z0-9]+)", text, re.I)
    if m:
        data["SKU"] = m.group(1).strip().upper()

    # Seating "Overall: W100 D38 H36 in." (checked first, highest priority)
    m = re.search(rf"Overall:\s*W({_NUM})\s+D({_NUM})\s+H({_NUM})", text, re.I)
    if m:
        data["Width"]      = _frac(m.group(1))
        data["Depth"]      = _frac(m.group(2))
        data["Height"]     = _frac(m.group(3))
        data["Dimensions"] = f"W{m.group(1)} D{m.group(2)} H{m.group(3)} in."

    # General WDH: "W64 D30 H30 in." or "W64 D30 H30" (no "in." required)
    if "Width" not in data:
        m = re.search(rf"\bW({_NUM})\s+D({_NUM})\s+H({_NUM})(?:\s*in\.?)?", text, re.I)
        if m:
            data["Width"]  = _frac(m.group(1))
            data["Depth"]  = _frac(m.group(2))
            data["Height"] = _frac(m.group(3))
            data.setdefault("Dimensions", f"W{m.group(1)} D{m.group(2)} H{m.group(3)} in.")

    # Labeled format: "Width/Dia: 23 \n Depth: 31 \n Height: 23" (some tables)
    if "Width" not in data:
        mw = re.search(r"Width(?:/Dia)?:\s*(\d+(?:\.\d+)?)", text, re.I)
        md = re.search(r"^Depth:\s*(\d+(?:\.\d+)?)", text, re.I | re.M)
        mh = re.search(r"^Height:\s*(\d+(?:\.\d+)?)", text, re.I | re.M)
        if mw:
            data["Width"] = mw.group(1)
            if re.search(r"Dia", mw.group(0), re.I):
                data["Diameter"] = mw.group(1)
        if md:
            data["Depth"] = md.group(1)
        if mh:
            data["Height"] = mh.group(1)

    # Diameter: "24 Dia" or "Dia 24"
    if "Diameter" not in data:
        m = re.search(rf"({_NUM})\s*[Dd]ia", text)
        if m:
            data["Diameter"] = _frac(m.group(1))

    # Seating specific dims (fraction-aware)
    m = re.search(rf"Inside Depth:\s*({_NUM})", text, re.I)
    if m:
        data["Seat Depth"] = _frac(m.group(1))

    m = re.search(rf"Arm Height:\s*({_NUM})", text, re.I)
    if m:
        data["Arm Height"] = _frac(m.group(1))

    m = re.search(rf"Seat Height:\s*({_NUM})", text, re.I)
    if m:
        data["Seat Height"] = _frac(m.group(1))

    # Finish: handles "Finish:", "Finish Shown:", "Finish As Shown:", "Selected Finish:"
    m = re.search(
        r"(?:Selected |Shown |As Shown )?Finish(?:\s+(?:Shown|As Shown))?:\s*(.+?)(?:\n|$)",
        text, re.I,
    )
    if m:
        finish_val = m.group(1).strip().lstrip("-").strip()
        # Strip any nested label prefix e.g. "Selected Finish: X" captured after "Standard Finish:"
        finish_val = re.sub(r"^(?:Selected|Standard|Shown|As Shown)\s+Finish:\s*", "", finish_val, flags=re.I)
        data["Finish"] = finish_val.lstrip("-").strip()

    # Discontinued finish: "SHOWN IN A DISCONTINUED AND UNAVAILABLE FINISH: K6 Cappuccino"
    if not data.get("Finish"):
        m = re.search(r"UNAVAILABLE FINISH[:\s\n]+([A-Z0-9][^\n]+)", text, re.I)
        if m:
            data["Finish"] = m.group(1).strip()

    # Fabric
    m = re.search(r"\bFabric:\s*(.+?)(?:\n|$)", text, re.I)
    if m:
        data["Fabric"] = m.group(1).strip()

    # Material: line after discontinued FINISH notice with explicit "Solids" or "Veneers"
    m = re.search(
        r"DISCONTINUED[^\n]*\n[^\n]*\n(.+?(?:Solid[s]?|Veneer[s]?)[^\n]*)",
        text, re.I,
    )
    if m:
        data["Material"] = m.group(1).strip()

    return data


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from listing pages (handles ?page=N pagination)."""
    links: list[str] = []
    seen: set[str] = set()
    current_url = listing_url

    while current_url:
        await page.goto(current_url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)

        html = await page.content()
        soup = BeautifulSoup(html, "html.parser")

        for a in soup.select("a.product-results-tile"):
            href = a.get("href", "")
            if href and "/catalog/" in href:
                full = BASE_URL + href if href.startswith("/") else href
                if full not in seen:
                    seen.add(full)
                    links.append(full)

        next_li = soup.select_one("ul.pager li.pager-next a")
        if next_li:
            next_href = next_li.get("href", "")
            current_url = BASE_URL + next_href if next_href.startswith("/") else next_href
        else:
            break

    return links


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a /catalog/{sku} product page. Returns a single-element list."""
    data = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)

    html = await page.content()
    soup = BeautifulSoup(html, "html.parser")

    # Product Name
    el = soup.select_one("p.product-title")
    if el:
        data["Product Name"] = clean_text(el.get_text())

    # Base SKU from h1
    el = soup.select_one("h1.heading-7")
    if el:
        data["SKU"] = clean_text(el.get_text()).upper()

    # Main product image
    img = soup.select_one("img.product-large-image")
    if img:
        src = img.get("src", "")
        if src:
            data["Image URL"] = src if src.startswith("http") else BASE_URL + src

    # Parse description block (may override SKU with full variant SKU)
    desc_div = soup.select_one("div.product-description-copy")
    parsed = parse_description_block(desc_div)

    variant_sku = parsed.pop("SKU", None)
    if variant_sku:
        data["SKU"] = variant_sku

    data.update(parsed)
    return [data]


async def scrape_leather_page(page) -> list[dict]:
    """Scrape leather gallery -- content is static HTML in hidden divs."""
    await page.goto(f"{BASE_URL}/leathers", timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    html = await page.content()
    soup = BeautifulSoup(html, "html.parser")

    items: list[dict] = []
    for container in soup.select("div.col-20-finishes"):
        row: dict = {"Source URL": f"{BASE_URL}/leathers", "Manufacturer": VENDOR_NAME}

        name_el = container.select_one("div.finish-name")
        if name_el:
            row["Product Name"] = clean_text(name_el.get_text())

        img = container.select_one("a.fancybox img.finish-image")
        if img:
            src = img.get("src", "")
            row["Image URL"] = src if src.startswith("http") else BASE_URL + src

        # Hidden div: <br/>-separated lines: name / grade / description
        hidden_div = container.select_one('div[style*="display:none"] div')
        if hidden_div:
            for br in hidden_div.find_all("br"):
                br.replace_with("\n")
            lines = [l.strip() for l in hidden_div.get_text().split("\n") if l.strip()]
            if len(lines) >= 2:
                row["Grade"] = lines[1]
            if len(lines) >= 3:
                row["Description"] = " ".join(lines[2:])

        if row.get("Product Name"):
            items.append(row)

    return items


async def scrape_fabric_page(page) -> list[dict]:
    """
    Hickory White fabrics are loaded via an iframe from hickorywhite.microdinc.com,
    which actively blocks headless Chromium (chrome-error on iframe load, 403 on direct fetch).
    This category is skipped until a bypass is available.
    """
    print("  [Fabric] SKIPPED: microdinc.com iframe blocks headless browsers (anti-bot detection)")
    return []


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category]")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        cats_done = 0
        for cat in info["categories"]:
            if not cat["links"]:
                continue
            if cats_done >= TEST_MAX_CATEGORIES:
                break

            cat_name = cat["name"]
            print(f"\n[{cat_name}] Starting...")

            writer.add_sheet(
                cat_name,
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            # Leather gallery
            if cat_name == "Leather":
                rows = await scrape_leather_page(page)
                if TEST_MODE:
                    rows = rows[:TEST_MAX_PRODUCTS]
                for idx, row in enumerate(rows, 1):
                    row.setdefault("Category", cat_name)
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(VENDOR_NAME, cat_name, idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat_name)
                print(f"  [{cat_name}] {len(rows)} items")
                cats_done += 1
                await async_polite_delay()
                continue

            # Fabric iframe
            if cat_name == "Fabric":
                rows = await scrape_fabric_page(page)
                if TEST_MODE:
                    rows = rows[:TEST_MAX_PRODUCTS]
                for idx, row in enumerate(rows, 1):
                    row.setdefault("Category", cat_name)
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(VENDOR_NAME, cat_name, idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat_name)
                print(f"  [{cat_name}] {len(rows)} items")
                cats_done += 1
                await async_polite_delay()
                continue

            # Regular catalog products
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"  [{cat_name}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        variant.setdefault("Category", cat_name)
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(VENDOR_NAME, cat_name, global_idx)
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                        writer.write_row(variant, category_name=cat_name)
                        global_idx += 1
                except Exception as e:
                    print(f"  ERROR scraping {url}: {e}")
                await async_polite_delay()

            print(f"  [{cat_name}] Done: {global_idx - 1} products")
            cats_done += 1

    writer.save()
    print(f"\n[Done] Saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
