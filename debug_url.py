import asyncio, sys
sys.path.insert(0, '.')
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, 'Forom/Code')
from scraper import get_shopify_json, _dismiss_cookie_banner
from base_scraper import PlaywrightBrowser

URLS = [
    'https://www.foromshop.com/products/strips-sofa',
    'https://www.foromshop.com/products/flowerpot-vp1-pendant',
]

async def debug():
    async with PlaywrightBrowser(headless=True) as page:
        # Accept cookies once
        await page.goto('https://www.foromshop.com', timeout=30000, wait_until='domcontentloaded')
        await page.wait_for_timeout(2000)
        await _dismiss_cookie_banner(page)

        for url in URLS:
            print(f'\n--- {url} ---')
            await page.goto(url, timeout=45000, wait_until='domcontentloaded')
            await page.wait_for_timeout(2000)

            # Check final URL (after any redirect)
            final_url = page.url
            print(f'  Final URL: {final_url}')

            # Check h1
            h1s = await page.query_selector_all('h1')
            for i, h in enumerate(h1s[:3]):
                print(f'  h1[{i}]: {(await h.inner_text()).strip()[:80]}')

            # Test JSON API
            shopify = await get_shopify_json(page, url)
            print(f'  API title: {shopify.get("title", "MISSING")}')
            print(f'  API variants: {len(shopify.get("variants", []))}')

asyncio.run(debug())
