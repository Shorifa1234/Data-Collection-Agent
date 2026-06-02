import asyncio, sys
sys.path.insert(0, '.')
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, 'Forom/Code')
from scraper import scrape_product
from base_scraper import PlaywrightBrowser

PRODUCTS = [
    ('Strips Sofa (multi-size)',    'https://www.foromshop.com/products/strips-sofa'),
    ('Liona Sectional (collapsed)', 'https://www.foromshop.com/products/liona-curve-sectional'),
    ('Flowerpot VP1 (lighting)',    'https://www.foromshop.com/products/flowerpot-vp1-pendant'),
]

SHOW_KEYS = [
    'Product Name', 'Dimensions', 'Seat Height', 'Seat Width', 'Seat Depth',
    'Material', 'Structure', 'Springing', 'Structure Upholstery', 'Base/Foot Type', 'Cover',
    'Origin', 'Weight', 'Wattage', 'Socket Type', 'Cable Length', 'Cord Material',
    'Shade Dimension', 'IP Rating', 'Description', 'Care Instructions',
    'Tearsheet Link', 'Spec Sheet', '3D Model', 'Price', 'Compare Price', 'SKU',
]

async def debug():
    async with PlaywrightBrowser(headless=True) as page:
        for label, url in PRODUCTS:
            print(f'\n{"="*60}')
            print(f'  {label}')
            print(f'{"="*60}')
            try:
                rows = await scrape_product(page, url)
                row = rows[0]
                for k in SHOW_KEYS:
                    if k in row and row[k]:
                        print(f'  {k}: {str(row[k])[:120]}')
                print(f'  --- Total fields: {len(row)}  Variants: {len(rows)} ---')
            except Exception as e:
                import traceback; traceback.print_exc()

asyncio.run(debug())
