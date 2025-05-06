import requests
from openpyxl import Workbook
from bs4 import BeautifulSoup

# WooCommerce API credentials and base URL
base_url = "https://captronics.in/wp-json/wc/v3/products"
auth = (
    "ck_cc48427d868965fe0577daf88d5b96eb748ff2a7",
    "cs_ae8125f4daf0b10755ff659b77db67f4ddb830f8"
)

# Function to clean HTML
def clean_html(raw_html):
    return BeautifulSoup(raw_html, "html.parser").get_text(separator="\n").strip()

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Products"
ws.append([
    "Product Name",
    "Long Description",
    "Short Description",
    "Price",
    "Stock Quantity",
    "Categories",
    "Attributes"
])

page = 1
while True:
    print(f"📦 Fetching page {page}...")
    params = {"per_page": 100, "page": page}
    response = requests.get(base_url, auth=auth, params=params)

    if response.status_code != 200:
        print(f"❌ Failed on page {page}. Status code: {response.status_code}")
        break

    products = response.json()
    if not products:
        print("✅ All products fetched.")
        break

    for product in products:
        name = product.get("name", "")
        long_desc = clean_html(product.get("description", ""))
        short_desc = clean_html(product.get("short_description", ""))
        price = product.get("price", "")
        stock = product.get("stock_quantity", "")

        # Join category names
        categories = ", ".join([cat["name"] for cat in product.get("categories", [])])

        # Join attribute values
        attributes = []
        for attr in product.get("attributes", []):
            values = ", ".join(attr.get("options", []))
            attributes.append(f"{attr.get('name')}: {values}")
        attributes_str = " | ".join(attributes)

        ws.append([name, long_desc, short_desc, price, stock, categories, attributes_str])

    page += 1

# Save to Excel
wb.save("products_full_details.xlsx")
print("📁 Saved: products_full_details.xlsx")
