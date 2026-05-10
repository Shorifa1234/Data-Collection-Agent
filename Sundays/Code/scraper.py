import asyncio, json, os, sys, re
from pathlib import Path
import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Sundays")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.sundays-company.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# Words in a title segment that identify it as a size, not a color/material
_SIZE_WORDS = {
    "tall", "regular", "large", "small", "medium", "counter", "bar",
    "king", "queen", "twin", "full", "seats", "for", "square", "round",
    "oval", "people", "set", "pair",
}


def _sundays_family_id(name: str) -> str:
    """Strip trailing color/material/size segments from a Sundays product title."""
    parts = [p.strip() for p in name.split(",")]
    while len(parts) > 1:
        last = parts[-1].strip()
        words = last.lower().split()
        # Strip if last segment is 1-3 short words (color / material / size)
        if len(words) <= 3 and all(len(w) <= 15 for w in words):
            parts.pop()
        else:
            break
    result = ", ".join(parts)
    # Fall back to base_scraper helper if we couldn't strip anything useful
    return result if result != name else extract_family_id(name)


def _shopify_collection_products(collection_handle: str) -> list[str]:
    """Return product page URLs for a collection via the Shopify JSON API."""
    urls: list[str] = []
    page = 1
    while True:
        api_url = (
            f"{BASE_URL}/collections/{collection_handle}"
            f"/products.json?limit=250&page={page}"
        )
        try:
            r = requests.get(api_url, headers=HEADERS, timeout=20)
            if r.status_code != 200:
                break
            products = r.json().get("products", [])
            if not products:
                break
            for p in products:
                urls.append(f"{BASE_URL}/products/{p['handle']}")
            if len(products) < 250:
                break
            page += 1
        except Exception as e:
            print(f"  [WARN] API error for {collection_handle} page {page}: {e}")
            break
    return urls


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs for a collection listing page."""
    m = re.search(r"/collections/([^/?#]+)", listing_url)
    if not m:
        return []
    return _shopify_collection_products(m.group(1))


def _parse_secondary_dims(sec_div) -> dict:
    """Parse secondary dimension blocks like 'Seat Depth 16.0"' into a dict."""
    result: dict = {}
    if sec_div is None:
        return result
    text = sec_div.get_text(" ", strip=True)
    # Remove inch marks and normalise whitespace
    text = text.replace('"', "").replace("'", "").strip()
    # Each entry: one or more words (label) followed by a decimal number
    for m in re.finditer(r"([A-Za-z][A-Za-z\s]*?)\s+([\d.]+)(?:\s|$)", text):
        label = m.group(1).strip().title()
        val   = m.group(2).strip()
        if label:
            result[label] = val
    return result


def _first_item(soup: BeautifulSoup, class_suffix: str):
    """Return the first .product__additional-information__item with given suffix class."""
    return soup.find(
        "div",
        class_=lambda c: c and "product__additional-information__item" in c and class_suffix in c,
    )


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape one product page and return a list containing one row dict."""
    try:
        # Use requests directly: all data is in static HTML and Playwright JS
        # execution modifies <sundays-float-number> values (truncates decimals).
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        data: dict = {"Source": url, "Manufacturer": VENDOR_NAME}

        # ── JSON-LD ────────────────────────────────────────────────────────
        ld: dict = {}
        ld_tag = soup.find("script", type="application/ld+json")
        if ld_tag:
            try:
                ld = json.loads(ld_tag.string or "")
            except Exception:
                pass

        # Product Name
        name = ld.get("name", "")
        if not name:
            h1 = soup.find("h1")
            if h1:
                name = clean_text(h1.get_text())
                name = re.sub(r"\s*-\s*Sundays\s+Company\s*$", "", name, flags=re.I)
        data["Product Name"] = name
        data["Product Family Id"] = _sundays_family_id(name)

        # SKU (from JSON-LD; overridden later if design-details has one)
        data["SKU"] = ld.get("sku", "")

        # Price
        offers = ld.get("offers", [])
        if offers:
            offer = offers[0] if isinstance(offers, list) else offers
            data["Price"] = clean_price(str(offer.get("price", "")))

        # Image URL
        images = ld.get("image", [])
        if images:
            data["Image URL"] = images[0] if isinstance(images, list) else images
        else:
            og = soup.find("meta", property="og:image")
            if og:
                data["Image URL"] = og.get("content", "")

        # Description
        desc = clean_text(ld.get("description", ""))
        if not desc:
            desc_tag = soup.find("div", class_="product__description")
            if desc_tag:
                desc = clean_text(desc_tag.get_text(" ", strip=True))
        data["Description"] = desc

        # ── Dimensions ────────────────────────────────────────────────────
        dims_item = _first_item(soup, "dimensions")
        if dims_item:
            base_div = dims_item.find("div", class_="product__base-dimensions")
            if base_div:
                raw = base_div.get_text(" ", strip=True).replace('"', "")
                raw = re.sub(r"\s+", " ", raw).strip()
                parsed = parse_dimensions(raw)
                data.update(parsed)

            sec_div = dims_item.find("div", class_="product__secondary-dimensions")
            for k, v in _parse_secondary_dims(sec_div).items():
                data.setdefault(k, v)

        # ── Design Details ─────────────────────────────────────────────────
        details_item = _first_item(soup, "details")
        if details_item:
            from bs4 import NavigableString
            # Each <li> is nested inside the previous one; get only the direct
            # text of each <li> (exclude child tag text) to avoid duplication.
            def _direct_text(tag):
                return clean_text(
                    " ".join(s.strip() for s in tag.children
                             if isinstance(s, NavigableString) and s.strip())
                )
            bullets = [t for t in (_direct_text(li) for li in details_item.find_all("li")) if t]
            if bullets:
                data["Design Details"] = " | ".join(bullets)
                for bullet in bullets:
                    bl = bullet.lower()
                    # material
                    if "material" in bl and ":" in bullet and not data.get("Material"):
                        m = re.search(r"material[s]?:\s*(.+)", bullet, re.I)
                        if m:
                            data["Material"] = m.group(1).strip()
                    # assembly
                    if "fully assembled" in bl:
                        data["Assembly Required"] = "No"
                    # finish
                    if "finish" in bl and ":" in bullet and not data.get("Finish"):
                        m = re.search(r"finish:\s*(.+)", bullet, re.I)
                        if m:
                            data["Finish"] = m.group(1).strip()
            # SKU from design details (overrides JSON-LD if present)
            sku_div = details_item.find("div", class_="product__sku")
            if sku_div:
                sku_val = sku_div.get_text(strip=True)
                if sku_val:
                    data["SKU"] = sku_val

        # ── Features ──────────────────────────────────────────────────────
        features_item = _first_item(soup, "features")
        if features_item:
            feat = [li.get_text(" ", strip=True) for li in features_item.find_all("li")]
            if feat:
                data["Features"] = " | ".join(feat)

        # ── Materials & Care ───────────────────────────────────────────────
        materials_item = _first_item(soup, "materials")
        if materials_item:
            mat_text = materials_item.get_text(" ", strip=True)
            if "Care:" in mat_text:
                parts = mat_text.split("Care:", 1)
                mat_part  = parts[0].strip()
                care_part = "Care: " + parts[1].strip()
                if not data.get("Material"):
                    data["Material"] = mat_part
                data["Care Instructions"] = care_part
            else:
                if not data.get("Material"):
                    data["Material"] = mat_text

        # ── Delivery ───────────────────────────────────────────────────────
        delivery_item = _first_item(soup, "delivery")
        if delivery_item:
            del_text = delivery_item.get_text(" ", strip=True)
            box_m = re.search(r"Comes in\s+(\d+)\s+box", del_text, re.I)
            if box_m:
                data["Pack"] = box_m.group(1)

        # ── Color/Finish from title ────────────────────────────────────────
        if name and not data.get("Color"):
            parts = [p.strip() for p in name.split(",")]
            if len(parts) >= 2:
                last = parts[-1]
                size_words = {
                    "tall", "regular", "large", "small", "medium", "counter",
                    "bar", "king", "queen", "twin", "full", "seats",
                }
                if not any(w in last.lower() for w in size_words):
                    data["Color"] = last

        return [data]

    except Exception as e:
        print(f"  [ERROR] {url}: {e}")
        return []


async def main() -> None:
    info    = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer  = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category]")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        cats = info["categories"]
        if TEST_MODE:
            cats = cats[:TEST_MAX_CATEGORIES]

        for cat in cats:
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

            print(f"[{cat['name']}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = _sundays_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [SKIP] {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"[Done] Output: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
