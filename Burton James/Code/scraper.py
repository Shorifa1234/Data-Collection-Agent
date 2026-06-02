import asyncio, json, os, re, sys
from pathlib import Path
from urllib.parse import urlparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Burton James")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
SCRAPE_CATEGORIES = [
    c.strip() for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",") if c.strip()
]


def _parse_repeater_items(html_section: str) -> list[tuple[str, str]]:
    """Extract (label, value) pairs from a dce-acf-repeater-list HTML block."""
    return re.findall(
        r'<li[^>]*>\s*<span[^>]*>([^<]+)</span>\s*<span[^>]*>([^<]*)</span>',
        html_section
    )


def _clean_dim_value(val: str) -> str:
    """Strip inch/quote marks (straight and Unicode smart quotes) from dimension values."""
    # Remove all quote/prime variants: " " " ″ ' ' ′
    return re.sub(r'["“”″′’‘\']', '', val).strip()


async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Return product page URLs from a listing page.
    If the listing URL is itself a product page (e.g. Benches), returns [listing_url].
    """
    parsed = urlparse(listing_url)
    listing_path = parsed.path.rstrip("/")
    path_depth = len([p for p in listing_path.split("/") if p])

    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2500)

    all_hrefs = await page.eval_on_selector_all(
        "a[href]", "els => els.map(e => e.href)"
    )

    links = []
    seen = set()
    for href in all_hrefs:
        if "#" in href:
            continue
        p = urlparse(href)
        if p.netloc != parsed.netloc:
            continue
        p_path = p.path.rstrip("/")
        p_depth = len([seg for seg in p_path.split("/") if seg])
        # Product links are exactly one level deeper than the listing path
        if p_path.startswith(listing_path + "/") and p_depth == path_depth + 1:
            base = p_path
            if base not in seen:
                seen.add(base)
                links.append(f"{p.scheme}://{p.netloc}{p_path}/")

    # If no products found but the listing URL itself looks like a product
    # page (depth >= 3, e.g. /furniture/banquette/blake/), treat it as one
    if not links and path_depth >= 3:
        links = [listing_url]

    return links


async def scrape_product(page, url: str) -> list[dict]:
    """Return a list of row dicts for a product page (one row per product, no variants)."""
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2500)

    html = await page.content()

    # --- Product name from <title> ---
    # Some ottoman/product titles include dimensions: "Anastasia46" x 23" x 20" - Burton James"
    # Strip any trailing dimension patterns (e.g. 46" x 23" x 20") from the name
    title_match = re.search(r"<title>\s*(.+?)\s*[–\-|]\s*Burton James", html, re.IGNORECASE)
    raw_name = title_match.group(1).strip() if title_match else ""
    raw_name = re.sub(r'[\s\d\.]*[""″”“\']+\s*[xX×]\s*.*$', '', raw_name).strip()
    product_name = sentence_case(raw_name) if raw_name else ""

    # --- Description (from description-content widget) ---
    desc_match = re.search(
        r'class="[^"]*description-content[^"]*".*?<p[^>]*>(.*?)</p>',
        html, re.DOTALL
    )
    description = clean_text(re.sub(r"<[^>]+>", " ", desc_match.group(1))) if desc_match else ""

    # --- Image URL ---
    # Collect all upload images, skip Screenshots and sized thumbnails (-NNNxNNN)
    img_url = ""
    img_matches = re.findall(
        r'src="(https://(?:www\.)?burtonjames[^"]+/wp-content/uploads/[^"]+\.(?:jpg|jpeg|png|webp))"',
        html, re.IGNORECASE
    )
    for img in img_matches:
        if "Screenshot" in img or "screenshot" in img:
            continue
        if re.search(r"-\d+x\d+\.", img):
            continue
        img_url = img
        break
    if not img_url:
        # Fallback: first non-screenshot upload image
        for img in img_matches:
            if "Screenshot" not in img and "screenshot" not in img:
                img_url = img
                break

    # --- Dimensions and Standard Options from dce-acf-repeater sections ---
    repeater_sections = re.findall(
        r'<div[^>]+class="[^"]*dce-acf-repeater[^"]*"[^>]*>(.*?)</div>\s*</div>',
        html, re.DOTALL
    )

    dims: dict[str, str] = {}
    options: dict[str, str] = {}

    # Check surrounding container class to differentiate dimensions vs options
    container_pattern = re.compile(
        r'<div[^>]+class="([^"]*dce-acf-repeater[^"]*)"[^>]*>(.*?)</div>\s*</div>',
        re.DOTALL
    )
    for m in container_pattern.finditer(html):
        container_class = m.group(1)
        section_html = m.group(2)
        pairs = _parse_repeater_items(section_html)
        if "options-list" in container_class:
            for label, value in pairs:
                options[label.strip()] = value.strip()
        else:
            # dimension-list
            for label, value in pairs:
                label = label.strip()
                if label.lower() == "sizing":
                    continue
                dims[label] = _clean_dim_value(value)

    # --- Build row ---
    row: dict = {
        "Source URL": url,
        "Product Name": product_name,
        "Product Family Id": extract_family_id(product_name),
        "Manufacturer": VENDOR_NAME,
        "Description": description,
        "Image URL": img_url,
    }

    # Map dimension labels to standard fields
    dim_label_map = {
        "Width": "Width",
        "Depth": "Depth",
        "Height": "Height",
        "Diameter": "Diameter",
        "Length": "Length",
        "Frame Height": "Frame Height",
        "Inside Width": "Inside Width",
        "Seat Height": "Seat Height",
        "Seat Depth": "Seat Depth",
        "Seat Width": "Seat Width",
        "Arm Width": "Arm Width",
        "Arm Height": "Arm Height",
        "Base/Leg Height": "Base/Leg Height",
    }
    for label, value in dims.items():
        col = dim_label_map.get(label, label)
        if value:
            row[col] = value

    # Build Dimensions string — prefer W/D/H; fall back to L/D/H for ottomans
    def _is_num(v: str) -> bool:
        return bool(v and re.match(r"^\d+(\.\d+)?$", v.strip()))

    w = dims.get("Width", "")
    d = dims.get("Depth", "")
    h = dims.get("Height", "")
    length = dims.get("Length", "")
    if _is_num(w) and _is_num(d) and _is_num(h):
        row["Dimensions"] = f'W {w}" x D {d}" x H {h}"'
    elif _is_num(w) and _is_num(h):
        row["Dimensions"] = f'W {w}" x H {h}"'
    elif _is_num(length) and _is_num(d) and _is_num(h):
        row["Dimensions"] = f'L {length}" x D {d}" x H {h}"'
    elif _is_num(length) and _is_num(h):
        row["Dimensions"] = f'L {length}" x H {h}"'

    # Standard options → map to canonical column names
    options_map = {
        "Spring Construction": "Seat Construction",
        "Seat Cushion Construction": "Seat Construction",
        "Back Cushion Construction": "Back",
        "Finish": "Finish",
        "Welting": "Welting",
        "Nailhead Trim": "Nailhead Trim",
        "COM": "COM",
    }
    for label, value in options.items():
        if not value:
            continue
        col = options_map.get(label, label)
        if col not in row:
            row[col] = value

    # --- Tearsheet link ---
    # href and "Download Tearsheet/Brochure" text are in the same <a> tag
    tearsheet_match = re.search(
        r'<a\b[^>]*href="(https://[^"]+\.pdf)"[^>]*>((?!</a>).)*?Download\s+(?:Tearsheet|Brochure)((?!</a>).)*?</a>',
        html, re.IGNORECASE | re.DOTALL
    )
    if tearsheet_match:
        row["Tearsheet Link"] = tearsheet_match.group(1)

    return [row]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
            if not cat["links"]:
                continue
            if SCRAPE_CATEGORIES and cat["name"] not in SCRAPE_CATEGORIES:
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
                    base = u.split("?")[0]
                    if base not in seen_urls:
                        seen_urls.add(base)
                        all_product_urls.append(u)

            print(f"  [{cat['name']}] {len(all_product_urls)} products")

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
                            variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                        writer.write_row(variant, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"    ERROR scraping {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
