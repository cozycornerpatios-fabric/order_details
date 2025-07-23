from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
import pandas as pd
import requests
import urllib.parse

app = FastAPI()

KATANA_API_TOKEN = "e505d03b-c53f-4519-9d6a-69c7637b0f64"
KATANA_API_BASE = "https://api.katanamrp.com/v1"

@app.get("/")
def root():
    return {
        "message": "Katana Excel Export API is running. Use /generate_excel?order_number=YOUR_ORDER_NUMBER"
    }

def fetch_order_by_number(order_number):
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    response = requests.get(f"{KATANA_API_BASE}/sales_orders", headers=headers)

    if response.status_code != 200:
        raise HTTPException(status_code=500, detail="Failed to fetch data from Katana")

    orders = response.json().get("results", [])
    matching_order = next((order for order in orders if order.get("order_number") == order_number), None)

    if not matching_order:
        raise HTTPException(status_code=404, detail="Order number not found")

    return matching_order

@app.get("/generate_excel")
def generate_excel(order_number: str = Query(...)):
    order = fetch_order_by_number(order_number)

    # Protect against missing product or order_lines
    product_name = ""
    hs_code = ""
    if order.get("order_lines"):
        first_line = order["order_lines"][0]
        product = first_line.get("product", {})
        product_name = product.get("name", "")
        hs_code = product.get("hs_code", "")

    df = pd.DataFrame(columns=[ ... ])  # keep your column list

    df.loc[0] = [
        1,
        order.get("customer_name", ""),
        order.get("customer_company", ""),
        order.get("shipping_address_line1", ""),
        "", "",  # Address Line 2/3
        order.get("shipping_country", ""),
        order.get("shipping_city", ""),
        order.get("shipping_state", ""),
        order.get("shipping_postal_code", ""),
        order.get("customer_phone", ""),
        "", "",  # Phone ext, Tax number
        order.get("customer_email", ""),
        "", "", "", "",  # Reference, billing
        order["order_number"],
        order["created_date"],
        len(order["order_lines"]),
        sum(line["quantity"] for line in order["order_lines"]),
        "", "", "", "", "", "", "", "",
        order["total_price"],
        "", "", "", "",
        product_name,
        hs_code,
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
    ]

    safe_order_number = urllib.parse.quote(order_number)
    file_path = f"/tmp/order_{safe_order_number}.xlsx"
    df.to_excel(file_path, index=False)

    return FileResponse(file_path, filename=f"Customs_Template_{order_number}.xlsx")
