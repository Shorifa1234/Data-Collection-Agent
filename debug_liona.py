import asyncio, sys, traceback
sys.path.insert(0, '.')
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, 'Forom/Code')
from scraper import scrape_product
from base_scraper import PlaywrightBrowser

async def debug():
    url = 'https://www.foromshop.com/products/liona-curve-sectional'
    async with PlaywrightBrowser(headless=True) as page:
        try:
            rows = await scrape_product(page, url)
            print(f'Got {len(rows)} rows')
            for k, v in rows[0].items():
                print(f'  {k}: {repr(str(v)[:150])}')
        except Exception:
            traceback.print_exc()

asyncio.run(debug())
