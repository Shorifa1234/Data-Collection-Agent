"""Test GraphQL product listing with category UID."""
import requests, json

GQL_URL = "https://hookerfurnishings.com/graphql"
HEADERS = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}

# Test with Dining Tables (uid=NTA1)
query = """
query GetCategoryProducts($uid: String!, $page: Int!) {
  products(
    filter: { category_uid: { eq: $uid } }
    pageSize: 10
    currentPage: $page
    sort: { name: ASC }
  ) {
    total_count
    page_info { total_pages current_page }
    items {
      name
      sku
      url_key
      url_suffix
    }
  }
}
"""

for uid, name in [("NTA1", "Dining Tables"), ("NTI5", "Nightstands"), ("NTA4", "Dining Chairs")]:
    r = requests.post(GQL_URL, headers=HEADERS, json={
        "query": query,
        "variables": {"uid": uid, "page": 1}
    }, timeout=30)
    data = r.json()
    prods = data.get("data", {}).get("products", {})
    print(f"\n{name} (uid={uid}):")
    print(f"  Total: {prods.get('total_count')}, Pages: {prods.get('page_info', {}).get('total_pages')}")
    for p in prods.get("items", [])[:3]:
        url = f"https://hookerfurnishings.com/{p['url_key']}{p.get('url_suffix','')}"
        print(f"  {p['sku']} | {p['name'][:40]} | {url}")
