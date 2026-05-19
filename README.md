# AVS Agent — Vendor Scraping System

**Version 1.8** · Built by Accord Tech Solutions

A Claude Code-powered agentic system that scrapes product data from interior design vendor websites and saves results to structured Excel files. Columns are **fully dynamic** — every field found on the product page becomes a column, ordered by the tracker's studio_columns definition.

**34 vendor scrapers** currently implemented across Playwright, Shopify API, Magento 2, WooCommerce, BigCommerce, NetSuite, EPiServer, and custom CMS platforms.

---

## How it works

```
You type:  /scrape-vendor "Wesley Hall"
               │
               ▼
    Claude reads the SD tracker Excel
    → finds all categories + links + studio_columns for that vendor
               │
               ▼
    Checks: does  Wesley Hall/Code/scraper.py  exist?
               │
       ┌───────┴───────┐
      YES              NO
       │               │
       ▼               ▼
   Run TEST first   Claude analyses the website HTML,
   (2 cats, 5 prods) writes scraper.py, then tests it
       │               │
       └───────┬───────┘
               ▼
    Show test results + timing — ask user to confirm full run
               │
               ▼
    Full run → all categories, all links per category, all products
               ▼
    Wesley Hall/
    ├── Code/
    │   ├── vendor_info.json
    │   └── scraper.py
    └── Data/
        ├── Wesley Hall_TEST.xlsx   ← test output
        └── Wesley Hall.xlsx        ← full output
```

---

## Setup (one-time)

```bash
# 1. Install Python dependencies
pip install -r requirements.txt

# 2. Install Playwright browser
python -m playwright install chromium

# 3. Open this folder in VS Code
# Claude Code slash commands are ready in .claude/commands/
```

---

## Slash Commands (use inside Claude Code)

### `/scrape-vendor <Vendor Name>`

The main command. Runs a test first, then asks for confirmation before the full run.

**Single vendor:**
```
/scrape-vendor The Future Perfect
/scrape-vendor Arteriors
/scrape-vendor Wesley Hall
```

**Multiple vendors in one go:**
```
/scrape-vendor Wesley Hall, Vanguard Designs, Visual Comfort
/scrape-vendor "Made Goods" "Warner House" "Wells Abbott"
```

Claude will process each vendor in sequence — generating or updating the scraper
code if needed, running the test, and reporting results before asking you to confirm
the full run.

**What Claude does for each vendor:**

1. Verifies the vendor exists in the SD tracker
2. Runs a **test scrape** (2 categories, max 5 products each) → `<Vendor>_TEST.xlsx`
3. Shows you the test results (columns, sample data, timing)
4. Asks: "Ready for the full run?"
5. On confirmation → full scrape of all categories and all listing links → `<Vendor>.xlsx`
6. Reports scrape time and total session time at the end

---

### `/test-scrape <Vendor Name>`

Run only the test — 2 categories, max 5 products. Use this to verify the scraper
without waiting for a full run.

```
/test-scrape The Future Perfect
/test-scrape Arteriors
```

**Output:** `<Vendor Name>/Data/<Vendor Name>_TEST.xlsx`

Claude will inspect the test output and tell you:

- Which columns were captured per category
- A sample product row
- Whether to proceed with the full scrape

---

## How to ask Claude to build scrapers for multiple vendors

Just list the vendor names — Claude handles everything in sequence:

```
/scrape-vendor Wesley Hall, Vanguard Designs
```

```
/scrape-vendor "The Future Perfect" "Made Goods" "Regina Andrew"
```

Or in plain English:

> "Please create scrapers for Wesley Hall, Vanguard Designs, and Visual Comfort"

> "Build and run the scraper for Warner House and Wells Abbott"

Claude will:
- Create `vendor_info.json` for each vendor from the tracker
- Write `scraper.py` for any vendor that doesn't have one yet
- Run the test for each and report results
- Ask for your confirmation before the full run

---

## Merging partial Excel files — `merge_excel.py`

When a run is split across two machines (local + VPS), or a run was interrupted and completed in two parts, merge the resulting Excel files:

```bash
# Output saved automatically as <file1>_merged.xlsx
python merge_excel.py "Safavieh/Data/Safavieh.xlsx" "Safavieh/Data/Safavieh_vps.xlsx"

# Custom output name
python merge_excel.py "Shev Chair/Data/Shev Chair.xlsx" "Shev Chair/Data/Shev Chair_vps.xlsx" "Shev Chair/Data/Shev Chair_FINAL.xlsx"
```

**Rules:**
- Sheets from **file1** come first — kept exactly as-is
- Sheets from **file2** not already in file1 are appended
- If the same sheet name exists in both, file1 wins
- Column widths, row heights, freeze panes, and cell styling are all preserved

---

## Checkpoint / resume system

Every scraper saves progress after each product to `<Vendor>/Data/<Vendor>_progress.json` and `<Vendor>/Data/<Vendor>_rows.jsonl`. If a run is interrupted:

1. Check checkpoint files are still there: `ls <Vendor>/Data/`
2. Re-run normally — scraper replays completed rows and continues from where it stopped:
   ```bash
   python orchestrator.py "Safavieh"
   ```
3. Checkpoint files are deleted automatically on a successful full run.

If checkpoints were lost, use `SCRAPE_CATEGORIES` to run only the missing categories.

---

## Running on VPS without losing the session

Always use `screen` or `nohup` on VPS so the process survives a disconnect:

```bash
# screen — recommended (reattach at any time)
screen -S vendor_run
python orchestrator.py "Safavieh"
# Detach: Ctrl+A then D    Reattach: screen -r vendor_run

# nohup — fire and forget
nohup python orchestrator.py "Safavieh" > safavieh.log 2>&1 &
tail -f safavieh.log
```

**Never run directly in an SSH session** — if the connection drops, the scraper is killed and you lose unsaved progress.

---

## Running vendors in parallel (VPS)

On a multi-core VPS (8+ cores / 16 GB RAM) you can run 2 vendors simultaneously using `run_parallel.py`:

```bash
# Run 2 vendors at the same time (recommended)
python run_parallel.py "Gabby" "Curry"

# Test mode first — always recommended before a full run
python run_parallel.py "Parker Southern" "Allied Maker" --test

# Run all vendors in batches of 2
python run_parallel.py --batch 2 "Parker Southern" "Gabby" "Curry" "Allied Maker" "Regina Andrew"

# Show help
python run_parallel.py --help
```

**Resource guide:**
- 2 vendors at once → safe for all scraper types
- 3 vendors → fine for lightweight scrapers; watch RAM for JS-heavy sites
- Each Playwright instance uses ~2–3 GB RAM

Output from all vendors is interleaved in real time, prefixed with `[VendorName]`. A timing summary is printed at the end.

---

## Python CLI Commands

### List all vendors in the tracker

```bash
python vendor_parser.py --list
```

### Parse a vendor (see categories + links + studio_columns)

```bash
python vendor_parser.py "Wesley Hall"
python vendor_parser.py "The Future Perfect"
```

### Test run — single vendor (2 categories, max 5 products)

```bash
python orchestrator.py "The Future Perfect" --test
python orchestrator.py "The Future Perfect" --test --headless false   # watch the browser
```

### Test run — multiple vendors

```bash
python orchestrator.py "Wesley Hall" "Vanguard Designs" --test
python orchestrator.py "Wesley Hall" "Vanguard Designs" "Visual Comfort" --test --headless false
```

### Full run — single vendor

```bash
python orchestrator.py "The Future Perfect"
python orchestrator.py "The Future Perfect" --headless false
```

### Partial run — specific categories only

Use when a full run was interrupted and only some categories need to be scraped:

```bash
SCRAPE_CATEGORIES="Ottomans,Cabinets,Sofas & Loveseats" python orchestrator.py "Safavieh"
SCRAPE_CATEGORIES="Bar Stools,Lounge Chairs,Fabric,Outdoor Seating" python orchestrator.py "Shev Chair"
```

This produces a new Excel with only those categories. Merge it with the original using `merge_excel.py`.

### Full run — multiple vendors

```bash
python orchestrator.py "Wesley Hall" "Vanguard Designs" "Visual Comfort"
python orchestrator.py "Wesley Hall" "Warner House" --headless false
```

A timing summary is printed at the end of every run:

```
============================================================
ORCHESTRATOR SUMMARY
============================================================
  Wesley Hall                     status=ok  scrape=4m 32s
  Vanguard Designs                status=ok  scrape=7m 18s

  Total session time: 11m 50s
============================================================
```

### Regenerate scraper code

```bash
python orchestrator.py "The Future Perfect" --force-regen
# Then /scrape-vendor to let Claude rewrite the code
```

---

## Output format

Each vendor gets its own folder:

```
<Vendor Name>/
├── Code/
│   ├── vendor_info.json   ← categories + links + studio_columns from tracker
│   └── scraper.py         ← auto-generated scraper
└── Data/
    ├── <Vendor Name>_TEST.xlsx   ← test output (2 cats, 5 products)
    └── <Vendor Name>.xlsx        ← full output
```

### Excel structure

- **Row 1:** Brand name
- **Row 2:** Category listing URL
- **Row 3:** (blank)
- **Row 4:** Column headers (dark blue, white text)
- **Row 5+:** Product data

### Column ordering (dynamic)

Columns are **not fixed** — they vary by category and by vendor. The ordering is:

1. **Core (always first):** Index, Category, Manufacturer, Source, Image URL, Product Name, Product Family Id, Price, SKU, Description, Weight, Specifications, Materials, Dimensions, Length, Width, Depth, Diameter, Height
2. **Tracker studio_columns** (category-specific, in tracker order)
3. **Known extras** in logical order: Seat Height, Seat Depth, Base, Canopy, COM, COL, Fabric, Finish, Illumination, Socket, Wattage, Designer, Maker, Collection, Lead Time, Origin, Production, Date, Tariff Disclaimer, Tearsheet Link
4. **Any new field** discovered on the product page (appended alphabetically)

### Multi-link categories

Some categories in the tracker have more than one listing URL (Link, Link 2, Link 3 …).
The scraper automatically visits all of them, deduplicates products, and writes
everything into a single sheet for that category.

---

## Vendor index

| Vendor | Tracker name | Platform |
|---|---|---|
| Alfonso Marina | `Alfonso Marina` | WooCommerce (JS-rendered listing) |
| Allied Maker | `Allied Maker` | NetSuite / custom CMS |
| Arteriors | `Arteriors` | Custom |
| Asthetic Decor | `Asthetic Decor` | WooCommerce |
| Basset Mirror | `Basset Mirror` | Custom ColdFusion CMS |
| Bernhardt | `Bernhardt` | Custom JSON API |
| Bunny Williams Home | `Bunny Williams Home (BWH)` | Shopify API |
| Carvers Guild | `Carvers Guild` | Custom |
| Century | `Century` | Shopify API |
| Curry & Company | **`Curry`** | EPiServer CMS |
| EQ3 | `EQ3` | Custom |
| English Georgian American | `English Georgian American` | Shopify |
| Fleur | `Fleur` | Shopify |
| Four Hands | `Four Hands` | React SPA *(login required — set `FH_EMAIL` / `FH_PASSWORD`)* |
| Gabby | `Gabby` | Shopify (`gabriellawhite.com`) |
| Hector Finch | `Hector Finch` | Custom |
| Hennepin Made | `Hennepin Made` | Custom |
| Highland House | `Highland House` | Custom ASP.NET |
| Kannoa | `Kannoa` | Shopify API |
| Kravet | `Kravet` | Magento 2 + Algolia |
| Palecek | `Palecek` | Custom CMS (Cloudflare) |
| Parker Southern | `Parker Southern` | Custom ASP.NET |
| Porta Romana | `Porta Romana` | Shopify API |
| Regina Andrew | `Regina Andrew` | SuiteCommerce |
| Remains | `Remains` | Shopify API |
| Safavieh | `Safavieh` | Shopify (multi-brand retailer) |
| Shev Chair | `Shev Chair` | WooCommerce |
| SkLO | `SkLo` | Custom |
| Sunpan | `Sunpan` | Shopify (requests + detail page HTML) |
| The Future Perfect | `The Future Perfect` | Custom |
| Verellen | `Verellen` | Custom |
| Villa & House | `Villa & House` | BigCommerce |
| Visual Comfort | `Visual Comfort` | Custom |
| Woodbridge Furniture | `Woodbridge Furniture` | Custom |

> **Curry & Company** is stored as **`Curry`** in the SD tracker. Always use `"Curry"` when calling orchestrator or vendor_parser.

---

## File reference

| File | Purpose |
| --- | --- |
| `CLAUDE.md` | Claude's full instruction manual for this system |
| `requirements.txt` | Python package dependencies |
| `vendor_parser.py` | Reads SD tracker → categories + links + studio_columns per vendor |
| `base_scraper.py` | Shared utilities: Playwright browser, dynamic ExcelWriter, parsers |
| `orchestrator.py` | Main dispatcher: supports 1 or many vendors, tracks timing per vendor |
| `run_parallel.py` | Runs 2+ vendor scrapers simultaneously on multi-core VPS |
| `merge_excel.py` | Merges two vendor Excel files into one (append missing sheets) |
| `.claude/commands/scrape-vendor.md` | `/scrape-vendor` slash command |
| `.claude/commands/test-scrape.md` | `/test-scrape` slash command |
| `vendor sheet/SD_Web Scraping - Status Tracker.xlsx` | Source of truth: vendor categories, links, column definitions |
| `Log/changelog.md` | Version history |

---

## Orchestrator exit codes

| Code | Meaning |
| --- | --- |
| 0 | Success — all vendors scraped |
| 1 | One or more vendors not found in tracker |
| 2 | One or more scrapers crashed (check traceback) |
| 3 | One or more scrapers not yet written — Claude must generate |

---

## Test mode env vars (used inside scraper.py)

These are set automatically by the orchestrator when `--test` is passed:

| Variable | Default | Description |
| --- | --- | --- |
| `TEST_MODE` | `false` | `true` activates test limits |
| `TEST_MAX_CATEGORIES` | `2` | Max categories to scrape in test mode |
| `TEST_MAX_PRODUCTS` | `5` | Max products per category in test mode |
| `HEADLESS` | `true` | `false` opens a visible browser window |
| `OUTPUT_PATH` | auto | Full path to output .xlsx (set by orchestrator) |
| `VENDOR_NAME` | auto | Vendor name string |

---

## Special vendor notes

| Vendor | Note |
|---|---|
| **Four Hands** | Login required. Set `FH_EMAIL` and `FH_PASSWORD` env vars before running. |
| **Curry & Company** | Tracker name is `Curry`. Use `python orchestrator.py "Curry"`. |
| **Gabby** | gabby.com redirects to gabriellawhite.com — Playwright follows automatically. |
| **Palecek** | Cloudflare-protected. Scraper uses custom browser context (fonts not blocked) to pass CF challenge on VPS. |
| **Sunpan** | Shopify store. Listing via `/products.json` API; product spec table fetched separately via requests. One row per finish variant. |
| **Villa & House** | BigCommerce trade portal — no price shown without login. Dimensions extracted from body text pattern. |
| **Basset Mirror** | No price (trade-only). SKU and dimensions are in the sub-header line e.g. `7086-LR-140 \| 52x24x16`. |
| **Alfonso Marina** | WooCommerce — listing pages are JS-rendered (Playwright). Product pages work via requests. SKU is extracted from the image filename. No price (trade-only). |
| **Safavieh** | Shopify multi-brand retailer. `Manufacturer` is always "Safavieh"; sub-brands (Theodore Alexander, Hooker Furniture, etc.) go in a separate `Brand` column from the spec table. Dimensions come from three places: page label, description bullets, and description text — all three are scraped. Full product data fetched via Shopify REST API (`/products/{handle}.json`). Supports `SCRAPE_CATEGORIES` for partial runs. |
| **Shev Chair** | WooCommerce. Pricing is contract-only ($0 on site). Dimensions parsed from plain text below the "Dimensions Width Depth Height" header. Fabric category has detailed specs in the `#tab-description` panel. Supports `SCRAPE_CATEGORIES` for partial runs. |

---

## Troubleshooting

**"Vendor not found"**
Run `python vendor_parser.py --list` — names must match the tracker sheet exactly (case-insensitive handled automatically).

**Test output is empty or mostly blank**
The CSS selectors in `scraper.py` are wrong for this site. Use `--force-regen` to let Claude rewrite the scraper.

**Scraper timeout / anti-bot block**
Run with `--headless false` to watch the browser. The site may need longer delays.

**Pagination stops early**
Console shows `+0 new products` — the next-page detection logic needs updating. Use `--force-regen`.

**Column is missing from output**
The field wasn't found on the product page. Check if it's JS-rendered (needs `wait_for_load_state`) or uses a different CSS selector.

**Multi-link category only shows products from the first link**
The scraper was written before V1.1. Use `--force-regen` so Claude rewrites it with the dedup loop.
