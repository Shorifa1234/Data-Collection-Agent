"""
scraper_chatgpt.py  -  Visual Comfort
-------------------------------------
Alternative Visual Comfort scraper implementation for comparison with
Claude-generated code.

Design goals:
  - Follow the project's dynamic-column contract exactly
  - Support orchestrator-style env vars and test mode
  - Use resilient, multi-strategy listing extraction because category pages
    may be JS-rendered and pagination patterns can vary
  - Collect all visible product fields plus known script-backed attributes

Run directly:
    python scraper_chatgpt.py

Env vars:
    HEADLESS             true | false
    OUTPUT_PATH          destination xlsx path
    VENDOR_NAME          vendor name string
    TEST_MODE            true | false
    TEST_MAX_CATEGORIES  max categories in test mode
    TEST_MAX_PRODUCTS    max products per category in test mode
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    PlaywrightBrowser,
    async_polite_delay,
    clean_price,
    clean_text,
    extract_family_id,
    generate_sku,
    parse_dimensions,
    parse_spec_block,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Visual Comfort")
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

BASE_URL = "https://www.visualcomfort.com"
TIMEOUT_MS = 45_000
MAX_PAGINATION_STEPS = 80

KNOWN_FIELD_ALIASES: dict[str, str] = {
    "finish": "Finish",
    "canopy": "Canopy",
    "socket": "Socket",
    "wattage": "Wattage",
    "shade details": "Shade Details",
    "shade": "Shade Details",
    "chain length": "Chain Length",
    "chain_length": "Chain Length",
    "lightsource": "Lightsource",
    "light source": "Lightsource",
    "rating": "Rating",
    "o/a height": "O/A Height",
    "oa height": "O/A Height",
    "oa_height": "O/A Height",
    "fixture height": "Fixture Height",
    "fixture_height": "Fixture Height",
    "min. custom height": "Min. Custom Height",
    "minimum custom height": "Min. Custom Height",
    "min_custom_height": "Min. Custom Height",
    "overall height": "Overall Height",
    "overall_height": "Overall Height",
    "extension": "Extension",
    "backplate": "Backplate",
    "base": "Base",
    "color temperature": "Color Temperature",
    "colour temperature": "Color Temperature",
    "color_temperature": "Color Temperature",
    "designer": "Designer",
    "collection": "Collection",
    "brand": "Collection",
    "bulb qty": "Bulb Qty",
    "bulb quantity": "Bulb Qty",
    "bulb_qty": "Bulb Qty",
    "ada compliant": "ADA Compliant",
    "ada": "ADA Compliant",
    "mounting": "Mounting",
}

SCRIPT_FIELD_ALIASES: dict[str, str] = {
    "finish": "Finish",
    "socket": "Socket",
    "wattage": "Wattage",
    "shade_details": "Shade Details",
    "shade": "Shade Details",
    "chain_length": "Chain Length",
    "lightsource": "Lightsource",
    "rating": "Rating",
    "oa_height": "O/A Height",
    "fixture_height": "Fixture Height",
    "min_custom_height": "Min. Custom Height",
    "overall_height": "Overall Height",
    "extension": "Extension",
    "backplate": "Backplate",
    "base": "Base",
    "color_temperature": "Color Temperature",
    "designer": "Designer",
    "collection": "Collection",
    "bulb_qty": "Bulb Qty",
    "canopy": "Canopy",
    "height": "Height",
    "width": "Width",
    "length": "Length",
    "diameter": "Diameter",
    "weight": "Weight",
}

DOWNLOAD_LABEL_MAP = {
    "tearsheet": "Tearsheet Link",
    "tear sheet": "Tearsheet Link",
    "spec sheet": "Spec Sheet Link",
    "cut sheet": "Spec Sheet Link",
    "installation guide": "Installation Guide Link",
    "instruction sheet": "Instruction Sheet Link",
    "warranty": "Warranty Link",
}


def _normalise_url(url: str) -> str:
    url = clean_text(url)
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/"):
        url = BASE_URL + url
    parsed = urlparse(url)
    clean_path = parsed.path.rstrip("/")
    query = urlencode(parse_qs(parsed.query), doseq=True)
    return urlunparse((parsed.scheme, parsed.netloc, clean_path, "", query, ""))


def _is_product_url(url: str) -> bool:
    parsed = urlparse(url)
    return parsed.netloc.endswith("visualcomfort.com") and "/us/p/" in parsed.path


def _clean_measurement(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ""
    value = value.replace("\u2033", '"').replace("\u2032", "'")
    value = re.sub(r"\b(inches|inch|in\.?|lbs?\.?|pounds?|kg)\b", "", value, flags=re.I)
    value = value.replace('"', "").replace("'", "")
    value = re.sub(r"\s+", " ", value).strip(" :,-")
    return value


def _canonical_field_name(label: str) -> str:
    label_clean = clean_text(label).rstrip(":")
    label_lower = label_clean.lower()
    if label_lower in KNOWN_FIELD_ALIASES:
        return KNOWN_FIELD_ALIASES[label_lower]
    for alias, canonical in KNOWN_FIELD_ALIASES.items():
        if label_lower == alias or label_lower.startswith(alias + " "):
            return canonical
    return label_clean.title()


def _maybe_store_field(data: dict, label: str, value: str) -> None:
    value = clean_text(value)
    if not value:
        return

    field = _canonical_field_name(label)

    if field == "Price":
        price = clean_price(value)
        if price is not None:
            data[field] = price
        return

    if field in {"Height", "Width", "Length", "Depth", "Diameter", "Weight"}:
        cleaned = _clean_measurement(value)
        if cleaned:
            data[field] = cleaned
        return

    if field == "Dimensions":
        dims = parse_dimensions(value)
        for k, v in dims.items():
            if v:
                data.setdefault(k, v)
        return

    data.setdefault(field, value)


def _extract_urls_from_html(html: str) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()

    for match in re.findall(r"https://www\.visualcomfort\.com/us/p/[^\"' <]+", html):
        url = _normalise_url(match)
        if _is_product_url(url) and url not in seen:
            seen.add(url)
            urls.append(url)

    soup = BeautifulSoup(html, "lxml")
    for tag in soup.select("a[href]"):
        href = _normalise_url(tag.get("href", ""))
        if _is_product_url(href) and href not in seen:
            seen.add(href)
            urls.append(href)

    return urls


async def _accept_cookies(page) -> None:
    selectors = [
        "button:has-text('Accept')",
        "button:has-text('Accept All')",
        "button:has-text('Allow All')",
        "#onetrust-accept-btn-handler",
    ]
    for selector in selectors:
        try:
            button = await page.query_selector(selector)
            if button:
                await button.click()
                await page.wait_for_timeout(750)
                return
        except Exception:
            continue


async def _scroll_listing_page(page) -> None:
    try:
        await page.evaluate(
            """
            async () => {
                const step = Math.max(500, Math.floor(window.innerHeight * 0.9));
                for (let y = 0; y < document.body.scrollHeight; y += step) {
                    window.scrollTo(0, y);
                    await new Promise(r => setTimeout(r, 180));
                }
                window.scrollTo(0, document.body.scrollHeight);
            }
            """
        )
        await page.wait_for_timeout(1200)
    except Exception:
        pass


async def _harvest_product_links(page) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()

    try:
        hrefs = await page.eval_on_selector_all(
            "a[href]",
            "els => els.map(el => el.href || el.getAttribute('href') || '').filter(Boolean)",
        )
        for href in hrefs:
            url = _normalise_url(href)
            if _is_product_url(url) and url not in seen:
                seen.add(url)
                urls.append(url)
    except Exception:
        pass

    try:
        html = await page.content()
        for url in _extract_urls_from_html(html):
            if url not in seen:
                seen.add(url)
                urls.append(url)
    except Exception:
        pass

    return urls


async def _goto_with_page_params(page, base_url: str, page_num: int) -> bool:
    candidates = [
        f"{base_url}?page={page_num}",
        f"{base_url}?p={page_num}",
        f"{base_url}?product_list_dir=asc&p={page_num}",
    ]
    for url in candidates:
        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            await _accept_cookies(page)
            return True
        except Exception:
            continue
    return False


async def get_product_links(page, listing_url: str, max_products: int | None = None) -> list[str]:
    """
    Collect product URLs from a category page.

    Strategy:
      1. Load the category page and harvest all /us/p/ links
      2. Scroll to trigger lazy-loaded product grids
      3. Click "next" / "load more" controls when present
      4. Fall back to common page query params if no explicit control exists
    """
    collected: list[str] = []
    seen_urls: set[str] = set()
    seen_states: set[tuple[str, int]] = set()

    try:
        await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
        await _accept_cookies(page)
    except Exception as exc:
        print(f"    [WARN] Failed to load listing: {listing_url} - {exc}")
        return []

    current_page_guess = 1

    for step in range(1, MAX_PAGINATION_STEPS + 1):
        await _scroll_listing_page(page)
        page_links = await _harvest_product_links(page)

        added = 0
        for url in page_links:
            if url not in seen_urls:
                seen_urls.add(url)
                collected.append(url)
                added += 1

        print(
            f"    [Listing step {step}] +{added} new products "
            f"(total {len(collected)})"
        )

        if max_products and len(collected) >= max_products:
            break

        state = (_normalise_url(page.url), len(seen_urls))
        if state in seen_states and added == 0:
            break
        seen_states.add(state)

        clicked_next = False
        next_selectors = [
            "a[rel='next']",
            "button[aria-label*='Next']",
            "button[aria-label*='next']",
            "a[aria-label*='Next']",
            "button:has-text('Next')",
            "a:has-text('Next')",
            "button:has-text('Load More')",
            "a:has-text('Load More')",
            "button:has-text('Show More')",
            ".ais-Pagination-link[aria-label*='Next']",
            ".pagination-next a",
        ]

        for selector in next_selectors:
            try:
                control = await page.query_selector(selector)
                if not control:
                    continue
                disabled = (
                    await control.get_attribute("disabled")
                    or await control.get_attribute("aria-disabled")
                )
                classes = (await control.get_attribute("class") or "").lower()
                if disabled or "disabled" in classes:
                    continue
                before = len(seen_urls)
                await control.click()
                await page.wait_for_timeout(2500)
                await _scroll_listing_page(page)
                after_links = await _harvest_product_links(page)
                if len({*seen_urls, *after_links}) > before:
                    clicked_next = True
                    break
            except Exception:
                continue

        if clicked_next:
            continue

        current_page_guess += 1
        navigated = await _goto_with_page_params(page, listing_url, current_page_guess)
        if not navigated:
            break

    return collected[:max_products] if max_products else collected


def _extract_jsonld_product(soup: BeautifulSoup, data: dict) -> None:
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or script.get_text() or ""
        raw = raw.strip()
        if not raw:
            continue
        try:
            payload = json.loads(raw)
        except Exception:
            continue

        if isinstance(payload, dict) and "@graph" in payload:
            candidates = payload["@graph"]
        elif isinstance(payload, list):
            candidates = payload
        else:
            candidates = [payload]

        for obj in candidates:
            if not isinstance(obj, dict):
                continue
            types = obj.get("@type")
            type_names = types if isinstance(types, list) else [types]
            if "Product" not in type_names:
                continue

            name = clean_text(obj.get("name", ""))
            if name:
                data["Product Name"] = name.upper()
                data.setdefault("Product Family Id", extract_family_id(data["Product Name"]))

            description = clean_text(obj.get("description", ""))
            if description:
                data.setdefault("Description", description)

            image = obj.get("image")
            if isinstance(image, list) and image:
                image = image[0]
            if image:
                data.setdefault("Image URL", str(image))

            sku = clean_text(obj.get("sku", ""))
            if sku:
                data.setdefault("SKU", sku)

            offers = obj.get("offers", {})
            if isinstance(offers, list) and offers:
                offers = offers[0]
            if isinstance(offers, dict):
                price = offers.get("price") or offers.get("lowPrice")
                if price is not None:
                    cleaned = clean_price(str(price))
                    if cleaned is not None:
                        data.setdefault("Price", cleaned)
            return


def _extract_meta_fallbacks(soup: BeautifulSoup, data: dict) -> None:
    if not data.get("Product Name"):
        for selector in ["h1", ".pdp-product-name", "[data-testid='product-title']"]:
            tag = soup.select_one(selector)
            if tag:
                name = clean_text(tag.get_text())
                if name:
                    data["Product Name"] = name.upper()
                    data.setdefault(
                        "Product Family Id",
                        extract_family_id(data["Product Name"]),
                    )
                    break

    if not data.get("Image URL"):
        for selector in [
            "meta[property='og:image']",
            "img[src]",
            "img[data-src]",
            "img[data-zoom-image]",
        ]:
            tag = soup.select_one(selector)
            if not tag:
                continue
            src = tag.get("content") or tag.get("data-zoom-image") or tag.get("data-src") or tag.get("src")
            src = _normalise_url(src or "")
            if src:
                data["Image URL"] = src
                break

    if not data.get("Description"):
        for selector in [
            "meta[name='description']",
            "meta[property='og:description']",
            ".product-description",
            ".description",
            "[data-testid='product-description']",
        ]:
            tag = soup.select_one(selector)
            if not tag:
                continue
            text = tag.get("content") or tag.get_text()
            text = clean_text(text)
            if text:
                data["Description"] = text
                break

    if "Price" not in data:
        candidates = []
        for selector in [
            "[data-testid='product-price']",
            ".price",
            ".product-price",
            ".sales .value",
        ]:
            for tag in soup.select(selector):
                text = clean_text(tag.get_text(" ", strip=True))
                if text:
                    candidates.append(text)
        for text in candidates:
            price = clean_price(text)
            if price is not None:
                data["Price"] = price
                break


def _extract_pair_containers(soup: BeautifulSoup, data: dict) -> None:
    specs: list[str] = []

    for row in soup.select("tr"):
        label_tag = row.find(["th", "dt"])
        value_tag = row.find(["td", "dd"])
        if not label_tag or not value_tag:
            continue
        label = clean_text(label_tag.get_text(" ", strip=True))
        value = clean_text(value_tag.get_text(" ", strip=True))
        if not label or not value:
            continue
        _maybe_store_field(data, label, value)
        specs.append(f"{label}: {value}")

    dts = soup.select("dt")
    if dts:
        for dt in dts:
            dd = dt.find_next_sibling("dd")
            if not dd:
                continue
            label = clean_text(dt.get_text(" ", strip=True))
            value = clean_text(dd.get_text(" ", strip=True))
            if not label or not value:
                continue
            _maybe_store_field(data, label, value)
            specs.append(f"{label}: {value}")

    for block in soup.select(
        ".specification, .spec, .specs, .product-spec, .product-specification, "
        ".accordion-content, .pdp-details, .product-details, .details"
    ):
        # Some PDPs render label/value as small title + following paragraph/spans.
        labels = block.select("h2, h3, h4, h5, h6, dt, strong, b, .label, .name")
        for label_tag in labels:
            label = clean_text(label_tag.get_text(" ", strip=True)).rstrip(":")
            if not label or len(label) > 40:
                continue
            value_tag = label_tag.find_next_sibling(["p", "span", "div", "dd", "li"])
            if not value_tag:
                continue
            value = clean_text(value_tag.get_text(" ", strip=True))
            if not value or value.lower() == label.lower():
                continue
            _maybe_store_field(data, label, value)
            specs.append(f"{label}: {value}")

    for item in soup.select("li"):
        text = clean_text(item.get_text(" ", strip=True))
        if not text or ":" not in text:
            continue
        label, value = text.split(":", 1)
        label = clean_text(label)
        value = clean_text(value)
        if not label or not value or len(label) > 40:
            continue
        _maybe_store_field(data, label, value)
        specs.append(f"{label}: {value}")

    if specs:
        merged = parse_spec_block(" | ".join(specs))
        for key, value in merged.items():
            if key == "Dimensions":
                dims = parse_dimensions(value)
                for k, v in dims.items():
                    if v:
                        data.setdefault(k, v)
            else:
                data.setdefault(key, value)
        data.setdefault("Specifications", " | ".join(dict.fromkeys(specs)))


def _extract_script_backed_fields(html: str, data: dict) -> None:
    for raw_key, canonical in SCRIPT_FIELD_ALIASES.items():
        pattern = re.compile(
            rf'["\']{re.escape(raw_key)}["\']\s*[:=]\s*["\']([^"\']{{1,200}})["\']',
            re.IGNORECASE,
        )
        match = pattern.search(html)
        if not match:
            continue
        _maybe_store_field(data, canonical, match.group(1))

    if "Specifications" not in data:
        lines = []
        for raw_key, canonical in SCRIPT_FIELD_ALIASES.items():
            pattern = re.compile(
                rf'["\']{re.escape(raw_key)}["\']\s*[:=]\s*["\']([^"\']{{1,200}})["\']',
                re.IGNORECASE,
            )
            match = pattern.search(html)
            if match:
                lines.append(f"{canonical}: {clean_text(match.group(1))}")
        if lines:
            data["Specifications"] = " | ".join(lines)


def _extract_download_links(soup: BeautifulSoup, data: dict) -> None:
    for tag in soup.select("a[href]"):
        href = tag.get("href", "")
        label = clean_text(tag.get_text(" ", strip=True))
        if not href:
            continue
        href = _normalise_url(href)
        label_lower = label.lower()
        for alias, canonical in DOWNLOAD_LABEL_MAP.items():
            if alias in label_lower:
                data.setdefault(canonical, href)
                break
        if href.lower().endswith(".pdf") and "Tearsheet Link" not in data:
            if "tear" in href.lower() or "spec" in href.lower():
                data["Tearsheet Link"] = href


async def scrape_product(page, url: str) -> dict:
    data: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2200)
        await _accept_cookies(page)
    except Exception as exc:
        print(f"    [WARN] Product load failed: {url} - {exc}")
        return data

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    _extract_jsonld_product(soup, data)
    _extract_meta_fallbacks(soup, data)
    _extract_pair_containers(soup, data)
    _extract_script_backed_fields(html, data)
    _extract_download_links(soup, data)

    if data.get("Product Name") and not data.get("Product Family Id"):
        data["Product Family Id"] = extract_family_id(data["Product Name"])

    dims_from_text = data.get("Dimensions")
    if dims_from_text:
        for key, value in parse_dimensions(dims_from_text).items():
            if value:
                data.setdefault(key, value)

    if not data.get("SKU"):
        sku_patterns = [
            r'\bSKU\b[:\s#-]*([A-Z0-9-]{3,})',
            r'\bItem\b[:\s#-]*([A-Z0-9-]{3,})',
        ]
        text = clean_text(soup.get_text(" ", strip=True))
        for pattern in sku_patterns:
            match = re.search(pattern, text, re.I)
            if match:
                data["SKU"] = match.group(1).upper()
                break

    # Keep a derived tearsheet if the PDP exposes neither a direct link nor a download label.
    if "Tearsheet Link" not in data and data.get("SKU"):
        data["Tearsheet Link"] = f"{BASE_URL}/on/demandware.static/-/Sites-vc-master-catalog/default/dwec0e0000/pdfs/{data['SKU']}.pdf"

    return data


async def scrape_category(
    page,
    writer: ExcelWriter,
    category: dict,
    max_products: int | None = None,
) -> None:
    category_name = category["name"]
    category_links = category.get("links", [])
    studio_columns = category.get("studio_columns", [])
    primary_link = category_links[0] if category_links else ""

    mode_tag = f" [TEST: max {max_products} products]" if max_products else ""
    print(f"\n[Category] {category_name}{mode_tag}")
    writer.add_sheet(category_name, primary_link, studio_columns=studio_columns)

    all_urls: list[str] = []
    seen_urls: set[str] = set()

    for link in category_links:
        urls = await get_product_links(page, link, max_products=max_products)
        for url in urls:
            if url not in seen_urls:
                seen_urls.add(url)
                all_urls.append(url)
        if max_products and len(all_urls) >= max_products:
            break

    if max_products:
        all_urls = all_urls[:max_products]

    print(f"  [Category] {len(all_urls)} products to scrape")

    for idx, product_url in enumerate(all_urls, start=1):
        print(f"  [{idx}/{len(all_urls)}] {product_url}")
        try:
            data = await scrape_product(page, product_url)
            if not data.get("SKU"):
                data["SKU"] = generate_sku(VENDOR_NAME, category_name, idx)
                print(f"    [SKU generated] {data['SKU']}")
            if not data.get("Product Family Id") and data.get("Product Name"):
                data["Product Family Id"] = extract_family_id(data["Product Name"])
            writer.write_row(data, category_name=category_name)
        except Exception as exc:
            print(f"    [ERROR] {product_url}: {exc}")
        await async_polite_delay(1.0, 2.4)

    print(f"  [Category] Done - {len(all_urls)} rows buffered")


async def main() -> None:
    info_path = Path(__file__).parent / "vendor_info.json"
    if info_path.exists():
        vendor_info = json.loads(info_path.read_text(encoding="utf-8"))
    else:
        from vendor_parser import parse_vendor

        vendor_info = parse_vendor(VENDOR_NAME)

    categories = [cat for cat in vendor_info["categories"] if cat.get("links")]
    max_products: int | None = None
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        max_products = TEST_MAX_PRODUCTS

    print(f"\n[Scraper] Vendor  : {vendor_info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")
    print(f"[Scraper] Cats    : {len(categories)}")
    if TEST_MODE:
        print(f"[Scraper] Max products/cat: {TEST_MAX_PRODUCTS}")

    writer = ExcelWriter(OUTPUT_PATH, vendor_info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        page.set_default_timeout(TIMEOUT_MS)
        for category in categories:
            await scrape_category(page, writer, category, max_products=max_products)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
