# P Kaufman Fabric Scraper

## Overview

P Kaufman is a fabric vendor selling through multiple brands:
- **P/K Fabrics** — main brand
- **P/K Lifestyles** — lifestyle collection  
- **Golding Fabrics** — premium collection

The scraper collects all fabrics (662 products) across these brands from **four different listing URLs** using two different methods:
1. **FastSimon API** for search results (induro, indurance)
2. **Playwright web scraping** for category pages (Golding Fabrics view-all)

## Key Vendor Features

### Multi-Link Category
The "Fabric" category in the tracker has **4 listing URLs**:
```
https://pkaufmann.com/pkfabrics/                    → homepage (no products)
https://pkaufmann.com/search-results/?q=induro      → 81 products (FastSimon API)
https://pkaufmann.com/search-results/?q=indurance   → 579 products (FastSimon API)
https://pkaufmann.com/p-k-fabrics/golding-fabrics/view-all-fabrics/ → 83 products (Playwright)
```

Total: **743 products before deduplication → 662 unique after deduplication**

### No Variants — Each Colorway = Separate URL
- Each product **colorway has its own URL** and product ID
- Example: "HUNT CLUB - Jade" has URL `/hunt-club-jade/` (SKU: 3000803)
- "HUNT CLUB - Natural" has URL `/hunt-club-natural/` (SKU: 3000800)
- No variant clicking needed — each colorway is scraped as a complete separate product

### FastSimon Search Results API
For URLs like `/search-results/?q=induro`, the scraper:
1. Extracts the query parameter (`q=induro` or `q=indurance`)
2. Calls the FastSimon full_text_search API directly instead of rendering the page
3. API endpoint: `https://api.fastsimon.com/full_text_search`
4. Handles pagination automatically (50 products per page)

**Why API instead of Playwright?** The search results pages require a trade account login to display products via the UI. The FastSimon API returns product data directly without login.

## Running the Scraper

### Test run (first 5 products per category)
```bash
python orchestrator.py "P Kaufman" --test --headless true
```

### Full run (all 662 products)
```bash
python orchestrator.py "P Kaufman" --headless true
```

### Show browser (debugging)
```bash
python orchestrator.py "P Kaufman" --headless false
```

## Output Structure

**File:** `P Kaufman/Data/P Kaufman.xlsx`

**Sheet:** Fabric (662 rows × 35 columns)

**Sample columns:**
- Index, Category, Manufacturer, Source URL, Image URL, Product Name, Product Family Id, SKU
- Collection, Pattern Name, Color, Width, Material, Origin, Finish, Care Instructions
- Construction, Design Status, Fabric Type, Fabric Weight, Flammability, Sustainability
- Abrasion-Wyzenbeek, Pattern Match, Repeat, Brand, Division Name, End Use
- Vertical/Horizontal Repeat, New Product, Discontinued, UPC Code

**Core fields always present:**
- **Manufacturer:** Always "P Kaufman" (vendor name)
- **Brand:** Designer/sub-brand name (e.g., "Studio NYC Design") when applicable
- **SKU:** From product page BCData JSON
- **Image URL:** Highest resolution via `data-zoom-image` attribute
- **Product Family Id:** Derived from Pattern Name (e.g., "HUNT CLUB", "GATSBY")

**Price:** Not captured (requires trade account login — "Call for Price")

## Product Detail Pages

All product URLs follow slug format: `https://pkaufmann.com/{pattern-name}-{color-name}/`

Example: `https://pkaufmann.com/hunt-club-jade/`

Page structure:
- Product specs in `.product-specifications-content` list (label/value pairs)
- Spec labels automatically mapped to standard columns (Collection Name → Collection, Pattern Name → Pattern Name, Color Name → Color, etc.)
- Image via `data-zoom-image` attribute
- SKU from BigCommerce `BCData` JSON variable

## Specifications Collected

The scraper extracts **all available** product specifications from each page, including:

### Standard textile fields
- Collection Name, Pattern Name, Color Name, Division Name
- Fabric Content (maps to "Material"), Fabric Width (maps to "Width")
- Fabric Weight, Finish, Fabric Care (maps to "Care Instructions")
- Print or Woven (maps to "Construction"), Pattern design (maps to "Pattern")
- Pattern Match, Repeat (Horizontal/Vertical)

### Performance & sustainability
- Abrasion-Wyzenbeek, Abrasion-Martindale, Flammability
- Sustainability (e.g., "Standard 100 by OEKO-TEX")

### Product status
- Design Status, New Product, Discontinued

### Additional
- Country of Origin, Brand (when vendor sells multiple brands), UPC Code, etc.

## Important Notes

1. **Deduplication:** Products appearing in multiple listing URLs (e.g., same product in induro AND indurance searches) are deduplicated by URL before scraping.

2. **FastSimon UUID hardcoded:** The scraper uses hardcoded FastSimon UUID and Store ID:
   ```python
   FAST_SIMON_UUID     = "1382dbe2-7b41-41ac-8cfe-bfbf7e3e3dd5"
   FAST_SIMON_STORE_ID = 1
   ```
   If P Kaufman changes their FastSimon configuration, this will need updating.

3. **Homepage URL gracefully skipped:** `/pkfabrics/` is a homepage, not a product listing. The scraper logs a warning and continues with other URLs.

4. **Dynamic columns:** The output columns are NOT fixed. The scraper collects all available fields from each product page, and `ExcelWriter` finalises column order automatically (core columns first, then tracker-defined studio_columns, then any extras found).

5. **Width cleaning:** Fabric width values (e.g., `54"`) are cleaned to numeric only (54.0) per core column specifications.

## Timing

- **Test run (5 products):** ~2m 11s
- **Full run (662 products):** ~1h 10m 51s
  - 81 induro products: ~20s (FastSimon API)
  - 579 indurance products: ~2m (FastSimon API)
  - 83 Golding products: ~1h 5m (Playwright scraping)

## Troubleshooting

| Issue | Solution |
|---|---|
| Search results page shows 0 products | This is expected without trade account login. Scraper uses FastSimon API instead. |
| FastSimon API returns 404 | Check UUID and Store ID hardcoded values. May need update if P Kaufman changes their search setup. |
| Golding Fabrics listing shows 0 products | Usually a timeout issue. Increase `wait_for_timeout()` from 4000ms to 6000ms in `get_product_links()`. |
| Some colorways missing | Check if they're in a different listing URL (induro vs indurance). The deduplication set will catch them if present in multiple URLs. |

## File Locations

```
P Kaufman/
├── README.md                              ← This file
├── Code/
│   ├── scraper.py                         ← Main scraper (you run this)
│   └── vendor_info.json                   ← Category/link/column metadata
└── Data/
    └── P Kaufman.xlsx                     ← Output (662 products × 35 cols)
```

## Last Updated

- Date: 2026-05-10
- Scraper version: 1.0 (FastSimon API + Playwright hybrid)
- Total products captured: 662
- Status: ✅ Fully tested and validated
