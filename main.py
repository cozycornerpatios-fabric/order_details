from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
import pandas as pd
import requests
import urllib.parse
import os

app = FastAPI()

# === Katana API Credentials and Endpoints ===
KATANA_API_TOKEN = "4103d451-7217-42fa-9cca-0f8a9e70155e"
KATANA_API_BASE = "https://api.katanamrp.com/v1"
KATANA_SALES_ORDERS_URL = f"{KATANA_API_BASE}/sales_orders"
KATANA_CUSTOMERS_URL = f"{KATANA_API_BASE}/customers"

@app.get("/")
def root():
    return {
        "message": "✅ Katana Excel Export API is running. Use /generate_excel?order_number=...&customer_reference=..."
    }

def fetch_order_by_number(order_number: str):
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    response = requests.get(KATANA_SALES_ORDERS_URL, headers=headers)
    if response.status_code != 200:
        raise HTTPException(status_code=500, detail="Failed to fetch sales orders from Katana")

    orders = response.json().get("results", [])
    matching_order = next((order for order in orders if order.get("order_number") == order_number), None)

    if not matching_order:
        raise HTTPException(status_code=404, detail="Order not found")

    print(f"✔️ Found order: {order_number}")
    return matching_order

def fetch_customer_by_reference(reference_number: str):
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    response = requests.get(KATANA_CUSTOMERS_URL, headers=headers)
    if response.status_code != 200:
        raise HTTPException(status_code=500, detail="Failed to fetch customers from Katana")

    customers = response.json().get("results", [])
    matching_customer = next(
        (cust for cust in customers if cust.get("reference") == reference_number), None)

    if not matching_customer:
        raise HTTPException(status_code=404, detail="Customer not found")

    print(f"✔️ Found customer reference: {reference_number}")
    return matching_customer

@app.post("/generate_excel")
def generate_excel(
    order_number: str = Query(...),
    customer_reference: str = Query(...)
):
    order = fetch_order_by_number(order_number)
    customer = fetch_customer_by_reference(customer_reference)

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
        "Customer Name": customer.get("name", ""),
        "Email": customer.get("email", ""),
        "Shipping Address Line 1": order.get("shipping_address_line1", ""),
        "Shipping State": order.get("shipping_state", ""),
        "Shipping ZIP": order.get("shipping_postal_code", ""),
        "Shipping Country": order.get("shipping_country", ""),
        "Phone": customer.get("phone", ""),
        "No. of Items": len(order.get("order_lines", [])),
        "Total Quantity": total_qty,
        "Product Names": ", ".join(product_names),
        "SKUs": ", ".join(skus),
        "Invoice Value": order.get("total_price", ""),
        "Fulfillment Status": order.get("fulfillment_status", "")
    }])

    safe_order_number = urllib.parse.quote(order_number)
    file_path = f"/tmp/order_{safe_order_number}.xlsx"
    df.to_excel(file_path, index=False)

    return FileResponse(
        file_path,
        filename=f"Customs_Template_{order_number}.xlsx",
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
