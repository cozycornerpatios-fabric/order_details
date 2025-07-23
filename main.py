from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
import pandas as pd
import requests
import urllib.parse

app = FastAPI()

# Replace with your real Katana API token
KATANA_API_TOKEN = "f6db7848-c6aa-47e9-8595-f07a402220e4"
KATANA_API_BASE = "https://api.katanamrp.com/v1"

@app.get("/")
def root():
    return {
        "message": "✅ Katana Excel Export API is running. Use /generate_excel?order_number=YOUR_ORDER_NUMBER"
    }

def fetch_order_by_number(order_number: str):
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    url = f"{KATANA_API_BASE}/sales_orders"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        raise HTTPException(status_code=500, detail="Failed to fetch data from Katana")

    orders = response.json().get("results", [])
    matching_order = next((order for order in orders if order.get("order_number") == order_number), None)

    if not matching_order:
        raise HTTPException(status_code=404, detail="Order not found")

    return matching_order

@app.get("/generate_excel")
def generate_excel(order_number: str = Query(...)):
    order = fetch_order_by_number(order_number)

    # Extract product info
    product_names = []
    skus = []
    total_qty = 0

    for line in order.get("order_lines", []):
        product = line.get("product", {})
        product_names.append(product.get("name", ""))
        skus.append(product.get("sku", ""))
        total_qty += line.get("quantity", 0)

    df = pd.DataFrame([{
        "Order Number": order.get("order_number", ""),
        "Customer Name": order.get("customer_name", ""),
        "Email": order.get("customer_email", ""),
        "Shipping Address Line 1": order.get("shipping_address_line1", ""),
        "Shipping State": order.get("shipping_state", ""),
        "Shipping ZIP": order.get("shipping_postal_code", ""),
        "Shipping Country": order.get("shipping_country", ""),
        "Phone": order.get("customer_phone", ""),
        "No. of Items": len(order.get("order_lines", [])),
        "Total Quantity": total_qty,
        "Product Names": ", ".join(product_names),
        "SKUs": ", ".join(skus),
        "Invoice Value": order.get("total_price", ""),
        "Fulfillment Status": order.get("fulfillment_status", ""),
    }])

    safe_order_number = urllib.parse.quote(order_number)
    file_path = f"/tmp/order_{safe_order_number}.xlsx"
    df.to_excel(file_path, index=False)

    return FileResponse(file_path, filename=f"Customs_Template_{order_number}.xlsx")

