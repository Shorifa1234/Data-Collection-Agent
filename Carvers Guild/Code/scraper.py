import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Carvers Guild")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://carversguild.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def get_all_mirror_links(listing_url: str) -> list[str]:
    """Paginate the browse listing to collect all mirror product URLs."""
    all_links = []
    seen = set()
    page = 0

    while True:
        url = listing_url if page == 0 else f"{listing_url}?page={page}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, "html.parser")
            links = [
                a["href"]
                for a in soup.find_all("a", href=True)
                if re.match(r"^/our-mirrors/number/\w+$", a["href"])
            ]
            new = [lnk for lnk in links if lnk not in seen]
            if not new:
                break
            for lnk in new:
                seen.add(lnk)
                all_links.append(BASE_URL + lnk)
            page += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"    Listing error page {page}: {e}")
            break

    return all_links


def scrape_product(url: str) -> list[dict]:
    """Scrape a single Carvers Guild mirror product page.

    Returns a list with one dict (no variant selection on this site).
    """
    r = requests.get(url, headers=HEADERS, timeout=25)
    if r.status_code != 200:
        return [{"Source": url}]

    soup = BeautifulSoup(r.text, "html.parser")
    data = {"Source": url}

    # --- product_caption block: SKU, Product Name, Dimensions ---
    cap = soup.find(class_="product_caption")
    if cap:
        spans = cap.find_all("span", class_="caption")
        h2 = cap.find("h2")

        if spans:
            # SKU: first caption span contains e.g. "#0012"
            sku_text = spans[0].get_text(strip=True).lstrip("#")
            if sku_text:
                data["SKU"] = sku_text

        if h2:
            data["Product Name"] = clean_text(h2.get_text())

        if len(spans) > 1:
            # Dimensions: second span e.g. "35&nbspx&nbsp47\"" (literal &nbsp; entities)
            # Also handles "42\"diam." (diameter-only mirrors)
            dims_raw = spans[1].get_text(strip=True)
            # Replace non-breaking spaces (both encoded and decoded forms)
            dims_raw = dims_raw.replace("\xa0", " ").replace("&nbsp;", " ").replace("&nbsp", " ")
            dims_raw = dims_raw.replace('"', "").strip()
            # Format: diameter only — e.g. "42diam." or "42 diam"
            m_diam = re.match(
                r"(\d+(?:\.\d+)?)\s*diam",
                dims_raw,
                re.IGNORECASE,
            )
            # Format: W x D x H
            m3 = re.match(
                r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)",
                dims_raw,
                re.IGNORECASE,
            )
            # Format: W x H
            m2 = re.match(
                r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)",
                dims_raw,
                re.IGNORECASE,
            )
            if m_diam:
                data["Diameter"] = m_diam.group(1)
                data["Dimensions"] = f'{m_diam.group(1)} diam.'
            elif m3:
                data["Width"] = m3.group(1)
                data["Depth"] = m3.group(2)
                data["Height"] = m3.group(3)
                data["Dimensions"] = f'{m3.group(1)} x {m3.group(2)} x {m3.group(3)}'
            elif m2:
                data["Width"] = m2.group(1)
                data["Height"] = m2.group(2)
                data["Dimensions"] = f'{m2.group(1)} x {m2.group(2)}'

    # --- Fallback product name from og:title ---
    if not data.get("Product Name"):
        og_title = soup.find("meta", property="og:title")
        if og_title:
            data["Product Name"] = clean_text(og_title.get("content", ""))

    # --- Fallback SKU from URL ---
    if not data.get("SKU"):
        m = re.search(r"/our-mirrors/number/(\w+)$", url)
        if m:
            data["SKU"] = m.group(1)

    # --- Description from product-body div ---
    body_div = soup.find(class_="product-body")
    if body_div:
        data["Description"] = clean_text(body_div.get_text(separator=" "))
    else:
        og_desc = soup.find("meta", property="og:description")
        if og_desc:
            raw_desc = BeautifulSoup(og_desc.get("content", ""), "html.parser").get_text()
            data["Description"] = clean_text(raw_desc)

    # --- Image URL: first ubercart product image ---
    ubercart_imgs = list(dict.fromkeys([
        img["src"]
        for img in soup.find_all("img", src=True)
        if "ubercart_images" in img["src"] and "carversn/images" not in img["src"]
    ]))
    if ubercart_imgs:
        img_url = ubercart_imgs[0]
        if not img_url.startswith("http"):
            img_url = BASE_URL + img_url
        data["Image URL"] = img_url

    # --- Price: Ubercart price (may require login — capture if visible) ---
    price_el = soup.find(class_=re.compile(r"uc-price|sell-price|price-amount"))
    if price_el:
        data["Price"] = clean_price(price_el.get_text())

    # --- Product Family Id ---
    if data.get("Product Name"):
        data["Product Family Id"] = extract_family_id(data["Product Name"])

    return [data]


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    for cat in categories:
        if not cat["links"]:
            continue

        cat_url = cat["links"][0]
        writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

        # Collect all product links across all listing URLs for this category
        seen_urls: set[str] = set()
        all_product_urls: list[str] = []
        for listing_url in cat["links"]:
            for u in get_all_mirror_links(listing_url):
                if u not in seen_urls:
                    seen_urls.add(u)
                    all_product_urls.append(u)

        if TEST_MODE:
            all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(all_product_urls)} mirrors found")

        global_idx = 1
        for url in all_product_urls:
            try:
                rows = scrape_product(url)
                for data in rows:
                    if not data.get("SKU"):
                        data["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not data.get("Product Family Id") and data.get("Product Name"):
                        data["Product Family Id"] = extract_family_id(data["Product Name"])
                    writer.write_row(data, category_name=cat["name"])
                    global_idx += 1
            except Exception as e:
                print(f"    ERROR on {url}: {e}")
            time.sleep(0.8)

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
