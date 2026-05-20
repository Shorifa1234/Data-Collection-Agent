"""
scraper.py — Shiir Rugs
------------------------
Site: shiirrugs.com (WordPress)

Site structure:
  Listing : /all-rugs/ — JS-rendered grid, scroll to load all products
  Product : /rugs/{slug}/ — one page per rug design

Data model:
  Each rug has multiple colorways, each with its own image.
  We produce ONE ROW PER COLORWAY so Image URL is always populated.

Selectors discovered:
  Product links  : article.product-card a, or a[href*='/rugs/']
  Product name   : h1
  Collection     : div.collection-name, or text near "Collection" label
  Description    : div.product-description p, or .entry-summary p
  Material       : text after "Material:" in details section
  Construction   : text after "Construction:" in details section
  Origin         : text after "Made in" in details section
  Colorways      : .colorway-item, .swatch, or named anchors/labels
  Images         : gallery images with alt text containing colorway name
  Tearsheet link : a[href*='pdf'][href*='tearsheet'], a:has-text('TEARSHEET')
  Care           : section.care p, or accordion with "Care" header
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    generate_sku,
    extract_family_id,
)
from bs4 import BeautifulSoup

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Shiir")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get(
    "OUTPUT_PATH",
    str(PROJECT_ROOT / "Shiir" / "Data" / "Shiir.xlsx"),
))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL    = "https://www.shiirrugs.com"
TIMEOUT_MS  = 45_000


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all rug product URLs across all paginated listing pages.

    Pagination pattern: ?_paged=N  (6 pages as of May 2026)
    Stop when a page yields no new product links.
    """
    seen: set[str] = set()
    result: list[str] = []
    base_listing = listing_url.rstrip("/").split("?")[0]

    page_num = 1
    while True:
        paged_url = base_listing + "/" if page_num == 1 else f"{base_listing}/?_paged={page_num}"
        print(f"  [Listing] page {page_num}: {paged_url}")

        try:
            await page.goto(paged_url, timeout=TIMEOUT_MS, wait_until="networkidle")
            await page.wait_for_timeout(1500)
        except Exception as e:
            print(f"  [WARN] listing page {page_num}: {e}")
            break

        links = await page.eval_on_selector_all("a[href]", "els => els.map(e => e.href)")
        page_links = [
            l.split("?")[0].rstrip("/") + "/"
            for l in links
            if re.search(r"/rugs/[^/]+/$", l.split("?")[0].rstrip("/") + "/")
            and "/all-rugs/" not in l
        ]

        new_count = 0
        for url in page_links:
            if url not in seen:
                seen.add(url)
                result.append(url)
                new_count += 1

        print(f"  [Listing] page {page_num}: {new_count} new products (total {len(result)})")

        # Stop when a page yields no new products
        if new_count == 0:
            break

        page_num += 1

    print(f"  [Listing] {len(result)} total products found")
    return result


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape one Shiir product page.
    Returns a list of dicts — one per colorway (or one fallback row).

    Page line structure (discovered via Playwright):
      [nav lines ...]
      LEGACY COLLECTION      ← collection link text
      Acanthus               ← h1 product name
      Possessing a luster... ← description (long line after product name)
      Colorways:
      Cumulus                ← one line per colorway
      REQUEST A QUOTE / AVAILABLE IN STOCK
      Details
      Made in India          ← origin
      Material: Wool & Silk
      Construction: Hand-knotted, Persian
      [Pile Height: Medium]  ← optional extra detail fields
      Care
      Clean spills...        ← care instructions
      VIEW CATALOG
      DOWNLOAD TEARSHEET
      Similar Rugs           ← STOP — everything after is noise
    """
    base: dict = {
        "Source URL": url,
        "Manufacturer": VENDOR_NAME,
    }

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="networkidle")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] page load: {e}")
        return [base]

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # ── 1. Product Name ────────────────────────────────────────────────────────
    h1 = soup.find("h1")
    if h1:
        base["Product Name"] = clean_text(h1.get_text())
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    # ── 2. Collection — link to /collections/{slug}/ with no CSS class ─────────
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        if re.search(r"/collections/[a-z]", href) and not a.get("class"):
            text = clean_text(a.get_text())
            if text and len(text) < 80:
                base["Collection"] = text.title()  # "LEGACY COLLECTION" → "Legacy Collection"
                break

    # ── 3. Line-based content parsing ─────────────────────────────────────────
    # Use \n to preserve line boundaries; filter to non-empty lines only
    raw_lines = [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]

    # Truncate at "Similar Rugs" to avoid footer noise
    try:
        stop_idx = raw_lines.index("Similar Rugs")
        lines = raw_lines[:stop_idx]
    except ValueError:
        lines = raw_lines

    product_name = base.get("Product Name", "")

    # Find key landmark indices
    name_idx        = next((i for i, l in enumerate(lines) if l == product_name), None)
    colorways_idx   = next((i for i, l in enumerate(lines) if l.lower() == "colorways:"), None)
    details_idx     = next((i for i, l in enumerate(lines) if l == "Details"), None)
    care_idx        = next((i for i, l in enumerate(lines) if l == "Care"), None)
    view_catalog_idx = next((i for i, l in enumerate(lines) if l == "VIEW CATALOG"), None)

    # Description: the line immediately after the product name (if it's long enough)
    if name_idx is not None and name_idx + 1 < len(lines):
        candidate = lines[name_idx + 1]
        if len(candidate) > 40 and candidate.lower() != "colorways:":
            base["Description"] = candidate

    # Colorways: lines between "Colorways:" and the next status/section marker
    _STATUS_MARKERS = {"REQUEST A QUOTE", "AVAILABLE IN STOCK", "Details", "Care"}
    colorways: list[str] = []
    _SEPARATORS = {",", "&", "/", "|", "-", "and"}
    if colorways_idx is not None:
        for line in lines[colorways_idx + 1:]:
            if line in _STATUS_MARKERS:
                break
            if line and line not in _SEPARATORS and len(line) > 1:
                colorways.append(line)

    # Details section: lines between "Details" and "Care" (or end)
    extra_details: dict = {}
    if details_idx is not None:
        detail_end = care_idx if care_idx and care_idx > details_idx else len(lines)
        for line in lines[details_idx + 1:detail_end]:
            if line.startswith("Made in"):
                base["Origin"] = line
            elif ":" in line:
                key, _, val = line.partition(":")
                key = key.strip()
                val = val.strip()
                if key == "Material":
                    base["Material"] = val
                elif key == "Construction":
                    base["Construction"] = val
                elif key and val and len(key) < 50:
                    extra_details[key] = val

    # Care Instructions: lines between "Care" and "VIEW CATALOG" (or end)
    if care_idx is not None:
        care_end = view_catalog_idx if view_catalog_idx and view_catalog_idx > care_idx else len(lines)
        care_lines = [l for l in lines[care_idx + 1:care_end] if l and l not in _STATUS_MARKERS]
        if care_lines:
            base["Care Instructions"] = " ".join(care_lines)

    # Merge extra detail fields (Pile Height, etc.) into base
    for k, v in extra_details.items():
        base.setdefault(k, v)

    # ── 4. Tearsheet Link ─────────────────────────────────────────────────────
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = clean_text(a.get_text())
        if ("tearsheet" in href.lower() or "TEARSHEET" in text.upper()) and href.endswith(".pdf"):
            base["Tearsheet Link"] = urljoin(BASE_URL, href)
            break

    # ── 5. Images — keyed by colorway ─────────────────────────────────────────
    gallery_images: list[dict] = []  # [{colorway, url}]
    name_slug = re.sub(r"[^a-zA-Z0-9]", r"[-_]", product_name)

    for img in soup.find_all("img", src=True):
        src = img.get("src", "")
        if not src or src.startswith("data:"):
            continue
        src_lower = src.lower()
        if "wp-content/uploads" not in src_lower:
            continue
        if not any(ext in src_lower for ext in [".jpg", ".jpeg", ".png", ".webp"]):
            continue
        if any(skip in src_lower for skip in ["logo", "icon", "banner", "sprite"]):
            continue

        full_src = urljoin(BASE_URL, src)
        fname_no_ext = re.sub(r"\.[^.]+$", "", src.split("/")[-1])

        # Strategy 1: SHIIR-Rugs-{Name}-{Colorway}.jpg pattern
        colorway_hint = ""
        m = re.match(
            r"(?:SHIIR[-_]Rugs[-_])" + name_slug + r"[-_](.+)",
            fname_no_ext,
            re.I,
        )
        if m:
            colorway_hint = re.sub(r"[-_]", " ", m.group(1)).strip()
            colorway_hint = re.sub(r"\s*\d+$", "", colorway_hint).strip()

        # Strategy 2: {NAME}-{Colorway}.jpg (e.g. AIKO-Ever.jpg)
        if not colorway_hint:
            name_upper = re.sub(r"\s+", "-", product_name.upper())
            m2 = re.match(rf"{re.escape(name_upper)}-(.+)", fname_no_ext, re.I)
            if m2:
                colorway_hint = re.sub(r"[-_]", " ", m2.group(1)).strip()
                colorway_hint = re.sub(r"\s*\d+$", "", colorway_hint).strip()

        # Strategy 3: alt text — last word(s) after product name
        if not colorway_hint:
            alt = clean_text(img.get("alt", ""))
            if alt and product_name.lower() in alt.lower():
                after = re.sub(re.escape(product_name), "", alt, flags=re.I).strip(" -_")
                if after:
                    colorway_hint = after.strip()

        gallery_images.append({"colorway": colorway_hint, "url": full_src})

    # Deduplicate
    seen_img: set[str] = set()
    product_images: list[dict] = []
    for img_data in gallery_images:
        if img_data["url"] not in seen_img:
            seen_img.add(img_data["url"])
            product_images.append(img_data)

    # ── 6. Build rows — one per colorway ──────────────────────────────────────
    rows: list[dict] = []

    if colorways and product_images:
        matched_urls: set[str] = set()
        for cw in colorways:
            best_img = ""
            cw_key = cw.lower().replace(" ", "")
            # Exact/partial colorway match on filename-extracted hint
            for img_data in product_images:
                img_key = img_data["colorway"].lower().replace(" ", "")
                if img_key and (img_key == cw_key or cw_key in img_key or img_key in cw_key):
                    best_img = img_data["url"]
                    matched_urls.add(best_img)
                    break
            # Fallback: first unmatched image
            if not best_img:
                for img_data in product_images:
                    if img_data["url"] not in matched_urls:
                        best_img = img_data["url"]
                        matched_urls.add(best_img)
                        break
            row = dict(base)
            row["Color"] = cw
            if best_img:
                row["Image URL"] = best_img
            rows.append(row)

    elif colorways:
        for cw in colorways:
            row = dict(base)
            row["Color"] = cw
            rows.append(row)

    elif product_images:
        for img_data in product_images:
            row = dict(base)
            row["Image URL"] = img_data["url"]
            if img_data["colorway"]:
                row["Color"] = img_data["colorway"]
            rows.append(row)

    else:
        rows = [base]

    # Single-row: ensure image is set
    if len(rows) == 1 and not rows[0].get("Image URL") and product_images:
        rows[0]["Image URL"] = product_images[0]["url"]

    return rows if rows else [base]


async def main() -> None:
    info_path = Path(__file__).parent / "vendor_info.json"
    info      = json.loads(info_path.read_text(encoding="utf-8"))

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {VENDOR_NAME}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")

    writer = ExcelWriter(OUTPUT_PATH, VENDOR_NAME)

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat.get("links"):
                print(f"[Skip] {cat['name']} — no links")
                continue

            print(f"\n[Category] {cat['name']}")
            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat.get("studio_columns", []),
            )

            seen_urls: set[str] = set()
            all_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_urls.append(u)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]

            print(f"  {len(all_urls)} products to scrape")

            global_idx = 1
            for url in all_urls:
                slug = url.rstrip("/").split("/")[-1][:60]
                print(f"  [{global_idx}/{len(all_urls)}] {slug}")
                try:
                    variant_rows = await scrape_product(page, url)
                    for row in variant_rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                    global_idx += 1

                await async_polite_delay(0.8, 2.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
