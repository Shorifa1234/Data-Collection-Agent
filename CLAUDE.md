# AVS Agent — System Guide for Claude

## What this project does

**AVS Agent** is an agentic web-scraping system for interior design vendors, developed by Accord Tech Solutions.
You (Claude Code) act as the senior Python developer who generates and runs vendor-specific scrapers on demand.

## Folder layout

```
Agentic system/                     ← project root (this folder)
├── CLAUDE.md                       ← you are here
├── requirements.txt
├── vendor_parser.py                ← reads SD tracker → category/link/column JSON
├── base_scraper.py                 ← shared Playwright browser + dynamic ExcelWriter
├── orchestrator.py                 ← dispatcher (check → run or spec)
│
├── vendor sheet/
│   └── SD_Web Scraping - Status Tracker.xlsx   ← source of truth for all vendors
│
├── sample data/
│   └── *.xlsx                      ← reference outputs from real vendor scrapes (15+ vendors)
│
├── Log/
│   └── changelog.md                ← project changelog (version history)
│
└── <Vendor Name>/                  ← auto-created per vendor
    ├── Code/
    │   ├── vendor_info.json        ← categories + links + studio_columns (written by orchestrator)
    │   └── scraper.py              ← vendor-specific scraper (you write this)
    └── Data/
        ├── <Vendor Name>.xlsx      ← scraped output
        └── run_log.txt             ← timing log appended after every run
```

## Source of truth: the SD tracker

- **Tracker path**: `vendor sheet/SD_Web Scraping - Status Tracker.xlsx`
- Each vendor has its own sheet.
- Row parsing:
    - col B = product group label ("Furniture", "Lighting", "Seating", …)
    - col C = "Category" → name in col D; "Link" / "Link 2" → URL in col D
    - col F = "Studio Column Names" label; col G onwards = the column list for that category
- Categories with **no link** (empty/None) must be **skipped**.

## Column strategy — DYNAMIC, not fixed

**The output columns are NOT a fixed list. They vary per category and per vendor.**

The system works like this:

1. `vendor_parser.py` extracts per-category `studio_columns` from the tracker (cols G onwards in the "Category" row).
2. The scraper collects **every field available** on the product page — do not discard any field.
3. `ExcelWriter.add_sheet(studio_columns=cat["studio_columns"])` records the preferred column order.
4. `ExcelWriter.save()` automatically finalises columns = CORE_FIRST + studio_columns + anything extra discovered.

**You must NOT hardcode a fixed column list.** Collect everything, let the writer sort it out.

### Core columns (always present, always first)

`Index, Category, Manufacturer, Source URL, Image URL, Product Name, Product Family Id, Price, SKU, Description, Weight, Specifications, Material, Materials, Dimensions, Length, Width, Depth, Diameter, Height, Finish, Collection, Origin, Lead Time`

- **`Source URL`** — the product page URL (formerly `Source`, formerly `Product URL`)
- **`Price`** — numeric list price, no currency symbols (formerly `List Price`)
- **`Manufacturer`** — always the **vendor/seller name** (e.g. `"Cowtan & Tout"`), auto-populated from `vendor_name`. Never use a sub-brand or designer name here.
- **`Brand`** — if the vendor sells products from multiple sub-brands or designers (e.g. multi-brand distributors like Cowtan & Tout), store the specific brand/designer name in a separate `Brand` column. Leave this field out entirely if the vendor only sells its own products.

> **Note on `Material` vs `Materials`:** both spellings appear in vendor data — collect whichever the site uses; the writer keeps both if both exist.

### Category-specific extras (collect all that exist on the site)

- **Seating**: `Seat Height, Seat Depth, Seat Width, Seat Length, Arm Height, Back, Frame, Footrest, Base/Foot Type, Seat Construction, COM, COM Yardage, COM Available, Fabric, Fabric Selection Required, Upholstery, Color, Components, Timber`
- **Lighting**: `Wattage, Voltage, Socket, Socket Type, Socket Qty, Lamping, Lamp Type, Lamp Quantity, Bulb Type, Bulb Qty, Bulb Wattage, Lumens, Brightness, Color Temperature, CRI, Dimming, Mounting, Canopy, Hanging Length, Chain Length, Chain Finish, Min Drop, Shade Details, Shade Color, Shade Shape, Shade Included, Wiring Type, Cut Sheet, 2D CAD, 3D Model`
- **Structural/furniture**: `Base, Size, Assembly Required, Hardware Details, Install Guide, Spec Sheet`
- **Upholstery/finish**: `COL, Fabric, Upholstery, Color, Wood, Timber, Components`
- **Shipping/logistics**: `Pack, CBM, Carton Size, Box 1 Size, Box 1 Weight, Box 2 Size, Box 2 Weight, 20FT, 40GP, 40HQ`
- **Rugs/textiles**: `Pattern, Construction, Care Instructions, Horizontal Repeat, Vertical Repeat, Wallcovering Area, Pile Height, Thickness`
- **Pillows/decor**: `Fill, Pillow Size, Removable Cover, Closure Type, Use, Watertight`
- **Attribution**: `Designer, Maker, Production, Date, Tariff Disclaimer`
- **Always last**: `Tearsheet Link`
- **Plus any field not in the above list — add it at the end**

## The /scrape-vendor command

When the user invokes `/scrape-vendor <Vendor Name>`, follow the steps in `.claude/commands/scrape-vendor.md`.

## Writing a new vendor scraper

When orchestrator.py exits with code **3**, no scraper exists yet. You must:

1. Read `<Vendor Name>/Code/vendor_info.json` for categories, links, and studio_columns
2. Fetch the first category listing page to understand HTML structure
3. Fetch 2-3 product detail pages to understand ALL available data fields
4. Write `<Vendor Name>/Code/scraper.py` using the template below
5. Run the scraper: `python orchestrator.py "<Vendor Name>" --headless true`

### Scraper template skeleton

```python
import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, parse_spec_block, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Vendor Name Here")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

async def scrape_product(page, url: str) -> list[dict]:
    """
    Return a LIST of flat dicts — one dict per variant (finish/size/color).
    If the product has no selectable variants, return a single-element list.

    Variant strategy (try in order):
      1. JSON-LD offers array  — each offer = one variant row (fast, no clicking)
      2. UI swatch/option buttons — click each, capture updated SKU/price/image
      3. Single-row fallback   — return [base_data] if no variants found

    Fields that change per variant : SKU, Price, Finish (or Color/Size), Image URL
    Fields shared across variants  : Product Name, Product Family Id, Description,
                                     all dimensions, specs, Designer, Collection, …
    """
    base = {"Source URL": url}
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    # --- 1. collect base/shared fields ---
    # Product Name → from h1 or JSON-LD name
    # Product Family Id → extract_family_id(product_name)
    # Description, Designer, Collection → from description block
    # Height/Width/Depth/Diameter/Weight → from spec section (numbers only)
    # Specifications → joined spec block string
    # Tearsheet Link → if found

    # --- 2. collect variants ---
    variants = []  # fill from JSON-LD offers OR UI swatch clicks

    # JSON-LD approach:
    #   offers = obj.get("offers", [])  — if len > 1, each offer is a variant
    #   per offer: row["SKU"], row["Price"], row["Image URL"], row["Finish"]

    # UI swatch approach (fallback):
    #   swatches = await page.query_selector_all(".swatch-option")
    #   for swatch in swatches: click → wait → capture SKU/price/image/finish

    return variants if variants else [base]

async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a listing page (handle pagination)."""
    links = []
    # ... navigate, extract hrefs, handle ?paged=N or equivalent ...
    return links

async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
            if not cat["links"]: continue

            # Register the sheet (first link used as the canonical source URL)
            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],   # ← pass tracker columns
            )

            # Collect product URLs from ALL links for this category
            # (tracker may list Link, Link 2, Link 3 … for the same category)
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            global_idx = 1
            for url in all_product_urls:
                variant_rows = await scrape_product(page, url)
                for variant in variant_rows:
                    # Mandatory: generate SKU if vendor did not provide one
                    if not variant.get("SKU"):
                        variant["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    # Mandatory: ensure Product Family Id is set
                    if not variant.get("Product Family Id") and variant.get("Product Name"):
                        variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                    writer.write_row(variant, category_name=cat["name"])
                    global_idx += 1
                await async_polite_delay()

    writer.save()   # ← columns finalised here, fully dynamic

if __name__ == "__main__":
    asyncio.run(main())
```

## Field rules — follow these exactly

### Category name standards

Always use these exact spellings when naming categories — vendor JSONs and Excel sheets must match:

| Correct | Wrong |
|---|---|
| `Flush Mounted` | `Flush Mount`, `Flush Mounts` |
| `Table Lamps` | `Tables Lamps` |

If a `vendor_info.json` has the wrong spelling, correct it before running the scraper.

### Mandatory fields (must always be present)

`Index`, `Category`, `Manufacturer`, `Source URL`, `Image URL`, `Product Name`, `Product Family Id`, `SKU`

- These columns always appear in every sheet even if some rows have gaps.
- All other fields are optional — collect them if available, skip if not.

### SKU — generate when missing

If the vendor does not provide a SKU, generate one using `generate_sku()` from `base_scraper`:

```python
from base_scraper import generate_sku
# In scrape_category(), after getting the product data:
if not data.get("SKU"):
    data["SKU"] = generate_sku(vendor_name, category_name, product_index)
```

Formula: **first 3 alpha letters of vendor name** + **first 3 alpha letters of category** + **product index number**
Examples: `"VISCHA45"` (Visual Comfort / Chandeliers / 45), `"MADLOU9"` (Made Goods / Lounge Chairs / 9)

### Price — number only

Use `clean_price()` from `base_scraper`. It strips `$`, `USD`, commas, and takes the lower value from price ranges.

```python
from base_scraper import clean_price
data["Price"] = clean_price(raw_price_string)   # returns float or None
```

Never store `"$1,200"` or `"USD 1200"` — store `1200.0`.

### Dimensions — numbers only, no inch marks

Use `parse_dimensions()` from `base_scraper`. It returns:

- `Dimensions` — cleaned string with `"` and `in` removed
- `Width`, `Height`, `Depth`, `Length`, `Diameter` — pure numeric strings (e.g. `"22.5"` not `'22.5"'`)

`parse_dimensions` handles all common formats:
- Label before number: `W 25" x D 12" x H 22.5"`
- Number before label: `16.00" W x 5.00" H` (Hennepin Made style)
- Number directly followed by label (no quote): `30H x 18.5Dia` (SkLO style)

Fractions like `1/2` are converted to decimals (`0.5`). Never store inch marks (`"`) in any dimension field.

> **Auto-derivation in `write_row`:** `ExcelWriter.write_row()` automatically calls `parse_dimensions` on the `Dimensions` value and fills in any missing sub-fields (`Width`, `Height`, `Depth`, `Diameter`, `Length`). This means scrapers only need to store `Dimensions` — the individual fields are populated for free. Sub-fields already explicitly set in the row dict are never overwritten.

### Product Family Id — common base name, not an exact copy

Use `extract_family_id()` from `base_scraper`. It strips variant suffixes (size, colour, finish) from the product name to find the family group.

```python
from base_scraper import extract_family_id
data["Product Family Id"] = extract_family_id(product_name)
```

Examples:

- `"OSLO DINING CHAIR - GREY FABRIC"` → `"OSLO DINING CHAIR"`
- `"REEF TABLE LAMP - BRASS FINISH"` → `"REEF TABLE LAMP"`
- `"OAK NIGHTSTAND"` → `"OAK NIGHTSTAND"` (no variant detected, unchanged)

If the site provides an explicit family name or collection, prefer that over the derived value.

### Other rules

1. **Dynamic columns** — never hardcode a fixed column list in a scraper.
2. **Collect everything** — if a field exists on the product page, capture it as a dict key.
3. **Pass studio_columns** to `writer.add_sheet()` — they come from `vendor_info.json`.
4. **Skip categories with no links** (empty `cat["links"]`).
5. **Polite scraping** — use `async_polite_delay()` between requests.
6. **One sheet per category** in the output Excel.
7. **Tearsheet Link** — construct as `{BASE_URL}/tearsheet/product/{slug}` if not directly available.
8. Leave missing optional fields out of the dict — do not store `""` or `None` (ExcelWriter renders missing keys as blank cells).
9. Run scrapers `HEADLESS=true` unless the user says otherwise.
10. The scraper must be runnable **standalone** (`python scraper.py`) as well as via orchestrator.
11. **One row per variant** — `scrape_product` must return `list[dict]`. For each finish/size/color variant available on the product page, produce a separate row. Fields that change per variant: `SKU`, `Price`, `Finish` (or `Color`/`Size`), `Image URL`. All other fields (name, description, dimensions, specs) are shared and duplicated across variant rows. Use JSON-LD offers array as primary source; fall back to clicking UI swatches if only one offer is present.
12. **Image URL** — always take the product image from the **product detail page**, not from listing/thumbnail images. Prefer `data-zoom-image` or `data-src` attributes over `src` for highest resolution.
13. **Manufacturer** — always set `data["Manufacturer"] = vendor_name`. The `ExcelWriter.write_row()` auto-populates this, but set it explicitly in the scraper for clarity.

## Shopify vendor scraping patterns

Many vendors use Shopify. Apply these patterns for any Shopify-based site.

### Dimensions — three sources, all required

Shopify product pages can show dimensions in up to three separate places. **Always scrape all three** and merge with `setdefault` (first value wins):

| Source | Example | Notes |
|---|---|---|
| Page label (outside accordions) | `Size : 40" X 40" X 18"` | Static text shown beside swatches — **most commonly missed**. Scrape with a DOM walker before opening any accordion. |
| Description accordion bullets | `Size Width: 43.3`, `Size Height: 15.7` | Key:Value bullet lines inside the description `<details>` block |
| Description accordion text | `Item Dimensions: (LxWxH) 39-1/2 x 39-1/2 x 29-1/2` | Inline sentence, may use mixed fractions like `39-1/2` |

**Dimension string formats to handle:**

```
40" X 40" X 18"            → unlabeled W x D x H
(LxWxH) 39-1/2 x 39-1/2 x 29-1/2  → explicit LxWxH order
W: 58.5 D: 41.5 H: 41.5   → labeled with colon
W 25" x D 12" x H 22.5"   → labeled with space
```

Mixed fractions like `39-1/2` must be converted to decimals (`39.5`).

### Accordion opening

Use `startsWith` (not exact match) when clicking summary elements — theme icons or extra spans inside `<summary>` will break exact text matching:

```python
# Correct — robust to extra content inside <summary>
s.textContent.trim().startsWith(heading)

# Wrong — breaks if summary contains an icon span
s.textContent.trim() === heading
```

### Shopify JSON — variant data

Use this priority order to find the embedded product JSON:

1. `window.productJSON`
2. `window.__product__`
3. `window.theme?.product`
4. `window.ShopifyAnalytics?.meta?.product`
5. `<script type="application/json">` tags (parse each, find one with `.variants`)
6. `{product: {...}}` wrapper blobs in JSON script tags

The `options` array on the product object names each option slot (`option1`/`option2`/`option3`). Map them to proper column names (`Color`, `Size`, `Finish`, etc.).

Shopify prices from `ShopifyAnalytics` are in **cents** (divide by 100). Prices from full product JSON `variants[].price` may be in dollars as a string — use `clean_price()`.

### Variant Source URL

Each variant row must have its own URL:
```python
row["Source URL"] = f"{base_url}?variant={variant_id}"
```

### Multi-brand Shopify retailers

If the vendor sells products from multiple brands (e.g. Safavieh sells Theodore Alexander, Hooker Furniture, etc.):
- `Manufacturer` = always the **store name** (e.g. `"Safavieh"`)
- `Brand` = the sub-brand from the Specifications table (e.g. `"Theodore Alexander"`)
- Never skip the `Brand` row from the spec table — store it in a `Brand` column

```python
# In spec table processing:
if spec_table.get("Brand") and spec_table["Brand"].lower() != vendor_name.lower():
    base["Brand"] = spec_table.pop("Brand")
else:
    spec_table.pop("Brand", None)
```

### Listing page deduplication

Strip query parameters from listing hrefs before adding to `seen` set — the same product URL can appear with different `?sort_by=` params:

```python
base_h = h.split("?")[0]
if base_h not in seen:
    seen.add(base_h)
    links.append(base_h)
```

## Partial runs and merging Excel files

### Running only specific categories

All scrapers support a `SCRAPE_CATEGORIES` env var. Use it when a full run was interrupted and you only need to complete the missing categories — without re-scraping categories that already have good data.

```bash
# Run only the categories that are missing or incomplete
SCRAPE_CATEGORIES="Ottomans,Cabinets,Sofas & Loveseats" python orchestrator.py "Safavieh"
SCRAPE_CATEGORIES="Bar Stools,Lounge Chairs,Fabric,Outdoor Seating" python orchestrator.py "Shev Chair"
```

This produces a **new Excel containing only those categories**. The completed categories stay in the original Excel. Merge the two files afterwards using `merge_excel.py`.

### Checkpoint / resume system

Every scraper saves a checkpoint after each product (`<Vendor>_progress.json` + `<Vendor>_rows.jsonl` in the `Data/` folder). If a run is interrupted:

1. Check that the checkpoint files are still present:  `ls <Vendor>/Data/`
2. Simply re-run — the scraper will replay already-scraped rows and continue from where it stopped:
   ```bash
   python orchestrator.py "Safavieh"
   ```
3. On a successful full run the checkpoint files are deleted automatically.

If checkpoint files were accidentally deleted, use `SCRAPE_CATEGORIES` to run only the missing categories instead.

### Merging two Excel files — `merge_excel.py`

Use this when you have a partial Excel from a local run and a second Excel from a VPS run covering the remaining categories.

```bash
# Saves as <file1>_merged.xlsx automatically
python merge_excel.py "Safavieh/Data/Safavieh.xlsx" "Safavieh/Data/Safavieh_vps.xlsx"

# Custom output name
python merge_excel.py "Shev Chair/Data/Shev Chair.xlsx" "Shev Chair/Data/Shev Chair_vps.xlsx" "Shev Chair/Data/Shev Chair_FINAL.xlsx"
```

Rules:
- Sheets from **file1** come first and are kept exactly as-is
- Sheets from **file2** that are NOT already in file1 are appended
- If a sheet name exists in both files, file1 wins — file2's version is skipped
- Column widths, row heights, freeze panes, and styling are all copied

### Running on VPS without losing the session

Always use `screen` or `nohup` on VPS so the process keeps running if you disconnect:

```bash
# screen (recommended — you can reattach any time)
screen -S vendor_run
python orchestrator.py "Safavieh"
# Detach: Ctrl+A then D    Reattach: screen -r vendor_run

# nohup (fire and forget)
nohup python orchestrator.py "Safavieh" > safavieh_run.log 2>&1 &
tail -f safavieh_run.log
```

After a VPS run finishes, download the Excel and merge with your local file if needed.

## Useful commands

```bash
# List all vendors in the tracker
python vendor_parser.py --list

# Parse one vendor (returns JSON with studio_columns per category)
python vendor_parser.py "The Future Perfect"

# Run orchestrator — single vendor
python orchestrator.py "The Future Perfect"
python orchestrator.py "The Future Perfect" --headless false   # show browser
python orchestrator.py "The Future Perfect" --force-regen      # rewrite code

# Run orchestrator — multiple vendors in one session (timing summary shown at end)
python orchestrator.py "Wesley Hall" "Vanguard Designs" "Visual Comfort"
python orchestrator.py "Wesley Hall" "Vanguard Designs" --test --headless false

# Install dependencies
pip install -r requirements.txt
python -m playwright install chromium
```

### Multi-link categories

When the tracker lists multiple links for a single category (Link, Link 2, Link 3 …),
`vendor_info.json` will have all URLs in `cat["links"]`. The scraper template
`main()` already iterates every link and deduplicates product URLs before scraping.
No extra work is needed — the parser and template handle it automatically.
