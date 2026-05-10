import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urlencode, urljoin, urlparse, parse_qs, unquote

import requests

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
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Bernhardt")
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

BASE_URL = "https://www.bernhardt.com"
IMAGE_BASE = "https://s3.amazonaws.com/emuncloud-staticassets/productImages/bh074/medium"
LIST_API = f"{BASE_URL}/service/QueryBernhardtProducts.json"
DETAIL_API = f"{BASE_URL}/service/QueryBernhardtProducts.json?1=up&IncludeTagShards=*&Status=Active"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": BASE_URL,
}

# Params excluded from tag criteria (pagination / UI params)
_EXCLUDE_PARAMS = {"orderBy", "context", "page", "skip", "take"}


def _parse_tag_criteria(url: str) -> dict:
    """Parse URL hash fragment into tagCriteria dict for the listing API."""
    if "#?" in url:
        fragment = url.split("#?", 1)[1]
    elif "#" in url:
        fragment = url.split("#", 1)[1]
    else:
        return {}

    criteria = {}
    for item in fragment.split("&"):
        if "=" not in item:
            continue
        k, v = item.split("=", 1)
        k = unquote(k).strip()
        v = unquote(v).strip()
        if k and k not in _EXCLUDE_PARAMS:
            criteria.setdefault(k, []).append(v)
    return criteria


def _fetch_product_ids_for_url(listing_url: str, session: requests.Session) -> list[str]:
    """Call listing API with pagination, return all product IDs for a URL."""
    criteria = _parse_tag_criteria(listing_url)

    # Build base params
    base_params = {
        "op": "ProductQuery1.4",
        "Fields": "Id",
        "include": "Total",
        "retailerId": "*",
        "context": "shop",
        "take": "48",
    }
    if criteria:
        base_params["tagCriteria"] = json.dumps(criteria)

    all_ids: list[str] = []
    skip = 0
    total = None

    while True:
        params = {**base_params, "skip": str(skip)}
        try:
            r = session.get(LIST_API, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
        except Exception as exc:
            print(f"  [WARN] Listing API error (skip={skip}): {exc}")
            break

        results = data.get("results", [])
        for item in results:
            pid = str(item.get("id", "")).strip()
            if pid:
                all_ids.append(pid)

        if total is None:
            total = int(data.get("total") or 0)

        skip += 48
        if skip >= total:
            break

    return all_ids


def _dim_strip(val: str) -> str:
    """Remove 'in' and whitespace from a dimension string like '32 in'."""
    return re.sub(r"[^\d./]", "", val).strip() or val.strip()


def _tag_first(tags: dict, *keys: str) -> str | None:
    """Return first non-empty value for any of the given tag keys."""
    for key in keys:
        vals = tags.get(key, [])
        for v in vals:
            v = clean_text(str(v))
            if v:
                return v
    return None


def _build_row_from_api(product: dict, cat_name: str) -> dict:
    """Convert the raw API product dict into a flat row dict."""
    tags = product.get("tags", {})
    meta = product.get("meta") or {}
    pid = str(product.get("id", "")).strip()

    row: dict = {}

    # --- core ---
    row["Product Name"] = clean_text(product.get("shortDescription", ""))
    row["SKU"] = pid
    row["Source"] = f"{BASE_URL}/shop/{pid}"
    row["Manufacturer"] = VENDOR_NAME

    # Image – primary image at position 1
    images = product.get("productImages", [])
    primary_images = [img for img in images if img.get("position") == 1]
    if not primary_images and images:
        primary_images = [images[0]]
    if primary_images:
        row["Image URL"] = f"{IMAGE_BASE}/{primary_images[0]['productId']}.jpg"
    else:
        row["Image URL"] = f"{IMAGE_BASE}/{pid}.jpg"

    # --- price ---
    price_val = (
        product.get("retailListPrice")
        or product.get("listPrice")
        or product.get("price")
        or product.get("wholesalePrice")
    )
    if price_val:
        row["Price"] = float(price_val)

    # --- dimensions from tags (with 'in' stripped) ---
    w = _tag_first(tags, "Width")
    h = _tag_first(tags, "Height")
    d = _tag_first(tags, "Depth")

    if w:
        row["Width"] = _dim_strip(w)
    if h:
        row["Height"] = _dim_strip(h)
    if d:
        row["Depth"] = _dim_strip(d)

    # Diameter (some products are round)
    diam = _tag_first(tags, "Diameter")
    if diam:
        row["Diameter"] = _dim_strip(diam)

    # Seat dims (seating)
    seat_w = _tag_first(tags, "Seat Width")
    seat_d = _tag_first(tags, "Seat Depth", "Seat Depth 2")
    seat_h = _tag_first(tags, "Seat Height")
    arm_h = _tag_first(tags, "Arm Height")
    if seat_w:
        row["Seat Width"] = _dim_strip(seat_w)
    if seat_d:
        row["Seat Depth"] = _dim_strip(seat_d)
    if seat_h:
        row["Seat Height"] = _dim_strip(seat_h)
    if arm_h:
        row["Arm Height"] = _dim_strip(arm_h)

    # Weight (shipping weight is in lbs; stored without unit)
    wt = _tag_first(tags, "Shipping Weight", "Total Shipping Weight")
    if wt:
        wt_num = re.sub(r"[^\d.]", "", wt).strip()
        if wt_num:
            row["Weight"] = wt_num

    # Dimensions string
    dim_str = _tag_first(tags, "TearSheet_DimensionIN", "1up_Dimension")
    if dim_str:
        # Clean HTML entities
        dim_str = re.sub(r"&nbsp;", " ", dim_str).strip()
        dim_str = re.sub(r"\s{2,}", "  ", dim_str)
        row["Dimensions"] = dim_str

    # --- finish / material / collection ---
    finish = _tag_first(tags, "Website Finish Color") or meta.get("Finish") or ""
    if finish:
        row["Finish"] = finish

    material = _tag_first(tags, "Website Material")
    if material:
        row["Material"] = material

    collection = (
        _tag_first(tags, "Brand", "Brand Multiview")
        or meta.get("Collection")
        or product.get("productLineId")
        or ""
    )
    if collection:
        row["Collection"] = collection

    # --- dates ---
    intro = _tag_first(tags, "Date introduced")
    if intro:
        row["Date"] = intro

    # --- UPC ---
    upc = product.get("upcCode", "")
    if upc:
        row["UPC"] = upc

    # --- COM flag ---
    com_val = _tag_first(tags, "COM", "COM Yardage")
    if com_val:
        row["COM"] = com_val

    # Fabric grade
    fabric_grade = _tag_first(tags, "Fabric Grade")
    if fabric_grade:
        row["Fabric Grade"] = fabric_grade

    # Product type
    prod_type = _tag_first(tags, "Product Type")
    if prod_type:
        row["Product Type"] = prod_type

    # Division
    division = _tag_first(tags, "Division")
    if division:
        row["Division"] = division

    # Sub-category
    subcats = tags.get("Sub-Category", [])
    if subcats:
        row["Sub-Category"] = ", ".join(subcats)

    # Express Ship
    express = _tag_first(tags, "Express Ship")
    if express:
        row["Express Ship"] = express

    # Availability note
    avail_note = _tag_first(tags, "1upStock")
    if avail_note:
        row["Availability"] = avail_note

    # Cubes (cubic volume)
    cubes = _tag_first(tags, "Cubes")
    if cubes:
        row["CBM"] = cubes

    # Pieces per carton
    ppc = _tag_first(tags, "Pieces Per Carton")
    if ppc:
        row["Pieces Per Carton"] = ppc

    # Interior bed dims
    for label, key in [
        ("Interior Bed Width", "Interior Width"),
        ("Interior Bed Depth", "Interior Depth"),
        ("Interior Bed Height", "Interior Height"),
    ]:
        val = _tag_first(tags, label)
        if val:
            row[key] = _dim_strip(val)

    # Extra custom meta fields with meaningful data
    if meta.get("RomanceCopy"):
        row["Finish Note"] = str(meta["RomanceCopy"]).strip()

    # Tearsheet link (construct from ATearSheet_AProductID or id)
    ts_ids = tags.get("ATearSheet_AProductID", [])
    # The tag may contain a comma-separated string like "353772, 998054P" — use first only
    ts_id_raw = ts_ids[0] if ts_ids else pid
    ts_id = ts_id_raw.split(",")[0].strip() if ts_id_raw else pid
    if ts_id and _tag_first(tags, "TearSheet_Show") == "Yes":
        row["Tearsheet Link"] = f"{BASE_URL}/tearsheet/{ts_id}/"

    return row


async def scrape_product(page, product_id: str) -> dict:
    """
    Navigate to /shop/{product_id}, intercept the 1=up API, and collect all fields.
    Returns a single flat row dict.
    """
    url = f"{BASE_URL}/shop/{product_id}"
    captured_product: dict = {}

    async def capture_response(response):
        if "QueryBernhardtProducts.json" in response.url and "1=up" in response.url:
            try:
                body = await response.text()
                data = json.loads(body)
                results = data.get("results", [])
                if results:
                    captured_product.update(results[0])
            except Exception:
                pass

    page.on("response", capture_response)
    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(4000)

        # Build row from API data
        if captured_product:
            row = _build_row_from_api(captured_product, "")
        else:
            row = {
                "Source": url,
                "SKU": product_id,
                "Image URL": f"{IMAGE_BASE}/{product_id}.jpg",
                "Manufacturer": VENDOR_NAME,
            }

        row["Source"] = url

        # Get description from romance-copy section
        try:
            desc_el = await page.query_selector("#romance-copy p")
            if desc_el:
                desc_text = clean_text(await desc_el.inner_text())
                if desc_text:
                    row["Description"] = desc_text
        except Exception:
            pass

        # If no description, try other selectors
        if not row.get("Description"):
            try:
                desc_el2 = await page.query_selector(
                    ".one-up-description, .product-about p, [class*='description'] p"
                )
                if desc_el2:
                    desc_text2 = clean_text(await desc_el2.inner_text())
                    if desc_text2 and len(desc_text2) > 20:
                        row["Description"] = desc_text2
            except Exception:
                pass

    except Exception as exc:
        print(f"  [ERROR] Product {product_id}: {exc}")
        row = {
            "Source": url,
            "SKU": product_id,
            "Image URL": f"{IMAGE_BASE}/{product_id}.jpg",
            "Manufacturer": VENDOR_NAME,
        }
    finally:
        page.remove_listener("response", capture_response)

    return row


async def get_product_links(page, listing_url: str, session: requests.Session) -> list[str]:
    """Return all product IDs for a listing URL."""
    # Intercept first page load to capture the canonical API URL
    captured_url: list[str] = []
    captured_data: dict = {}

    async def cap_resp(response):
        if (
            "QueryBernhardtProducts.json" in response.url
            and "op=ProductQuery" in response.url
        ):
            captured_url.append(response.url)
            try:
                body = await response.text()
                captured_data["body"] = body
            except Exception:
                pass

    page.on("response", cap_resp)
    try:
        await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(5000)
    except Exception as exc:
        print(f"  [WARN] Listing page load issue: {exc}")
    finally:
        page.remove_listener("response", cap_resp)

    if not captured_data.get("body"):
        print("  [WARN] No listing API captured; falling back to URL parsing.")
        return _fetch_product_ids_for_url(listing_url, session)

    first_data = json.loads(captured_data["body"])
    total = int(first_data.get("total") or 0)
    all_ids = [str(r["id"]) for r in first_data.get("results", []) if r.get("id")]

    # Paginate remaining pages via requests (faster than Playwright)
    if captured_url and total > 48:
        captured_api = captured_url[0]
        # Remove skip/take from URL and add fresh ones
        base_api = re.sub(r"&?skip=\d+", "", captured_api)
        base_api = re.sub(r"&?take=\d+", "", base_api)
        base_api = re.sub(r"\?&", "?", base_api)
        skip = 48
        while skip < total:
            paginated_url = f"{base_api}&skip={skip}&take=48"
            try:
                r = session.get(paginated_url, headers=HEADERS, timeout=30)
                data = r.json()
                for item in data.get("results", []):
                    pid = str(item.get("id", "")).strip()
                    if pid:
                        all_ids.append(pid)
            except Exception as exc:
                print(f"  [WARN] Pagination error (skip={skip}): {exc}")
            skip += 48

    print(f"  Found {len(all_ids)} product IDs (total={total})")
    return all_ids


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    session = requests.Session()
    session.headers.update(HEADERS)
    # Prime session
    try:
        session.get(BASE_URL, timeout=15)
    except Exception:
        pass

    categories = info["categories"]
    if TEST_MODE:
        cats_with_links = [c for c in categories if c.get("links")]
        categories = cats_with_links[:TEST_MAX_CATEGORIES]
        print(f"[TEST MODE] Limiting to {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat.get("links"):
                continue

            cat_name = cat["name"]
            print(f"\n[Category] {cat_name}")

            writer.add_sheet(
                cat_name,
                cat["links"][0],
                studio_columns=cat.get("studio_columns", []),
            )

            # Collect product IDs across all links for this category (deduped)
            all_ids: list[str] = []
            seen_ids: set[str] = set()

            for link_url in cat["links"]:
                ids = await get_product_links(page, link_url, session)
                for pid in ids:
                    if pid not in seen_ids:
                        seen_ids.add(pid)
                        all_ids.append(pid)

            if TEST_MODE:
                all_ids = all_ids[:TEST_MAX_PRODUCTS]
                print(f"  [TEST: max {TEST_MAX_PRODUCTS} products]")

            print(f"  Scraping {len(all_ids)} products...")

            global_idx = 1
            for i, product_id in enumerate(all_ids):
                try:
                    row = await scrape_product(page, product_id)

                    # Ensure mandatory fields
                    if not row.get("Product Name"):
                        row["Product Name"] = f"Product {product_id}"
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(VENDOR_NAME, cat_name, global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])

                    row["Category"] = cat_name

                    writer.write_row(row, category_name=cat_name)
                    print(f"  [{global_idx}] {row.get('Product Name', product_id)}")
                except Exception as exc:
                    print(f"  [ERROR] Product {product_id}: {exc}")

                global_idx += 1
                await async_polite_delay()

    writer.save()
    print(f"\nDone -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
