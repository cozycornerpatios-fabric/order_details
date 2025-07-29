from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import StreamingResponse
import pandas as pd
import requests
from io import BytesIO
import urllib.parse
import os
from typing import List

app = FastAPI()

# Security: Use environment variable, or fallback to your token (for dev only)
KATANA_API_TOKEN = os.environ.get("KATANA_API_TOKEN", "1e2c6dd2-c8ed-490d-8c70-a4bb21152682")
KATANA_API_BASE = "https://api.katanamrp.com/v1"
KATANA_SALES_ORDERS_URL = f"{KATANA_API_BASE}/sales_orders"

def fetch_all_sales_orders():
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    results = []
    url = KATANA_SALES_ORDERS_URL
    while url:
        response = requests.get(url, headers=headers)
        print(f"Fetching {url} - Status {response.status_code} - Text: {response.text}")  # Debug for logs!
        if response.status_code != 200:
            raise HTTPException(
                status_code=500,
                detail=f"Failed to fetch sales orders from Katana: {response.status_code} - {response.text}"
            )
        data = response.json()
        results.extend(data.get("results", []))
        url = data.get("next")
    return results

@app.post("/generate_excel")
def generate_excel(order_numbers: List[str] = Query(...)):
    all_orders = fetch_all_sales_orders()
    selected_orders = [order for order in all_orders if order.get("order_number") in order_numbers]
    if not selected_orders:
        raise HTTPException(status_code=404, detail="No matching sales orders found")

    records = []
    for order in selected_orders:
        shipping = order.get("shipping_address", {}) or {}
        records.append({
            "Order ID": order.get("id", ""),
            "Recipient_Contact Name": order.get("customer_name", ""),
            "Recipient_Company Name": shipping.get("company_name", ""),
            "Recipient_Address Line 1": shipping.get("address_line1", ""),
            "Recipient_Address Line 2": shipping.get("address_line2", ""),
            "Recipient_Address Line 3": shipping.get("address_line3", ""),
            "Recipient_Country": shipping.get("country", ""),
            "Recipient_City": shipping.get("city", ""),
            "Recipient_State": shipping.get("state", ""),
            "Recipient_Postal code": shipping.get("postal_code", ""),
            "Recipient_Phone Number": shipping.get("phone", ""),
            "Recipient_Email": order.get("customer_email", ""),
            "Reference_1": order.get("order_number", ""),
            "Invoice Value": order.get("total_price", ""),
        })

    columns = [
        "Order ID",
        "Recipient_Contact Name", "Recipient_Company Name", "Recipient_Address Line 1",
        "Recipient_Address Line 2", "Recipient_Address Line 3", "Recipient_Country",
        "Recipient_City", "Recipient_State", "Recipient_Postal code", "Recipient_Phone Number",
        "Recipient_Email", "Reference_1", "Invoice Value"
    ]
    df = pd.DataFrame(records, columns=columns)

    # Create in-memory Excel file
    excel_io = BytesIO()
    df.to_excel(excel_io, index=False)
    excel_io.seek(0)

    safe_names = "_".join([urllib.parse.quote(num) for num in order_numbers])
    filename = f"SalesOrders_{safe_names}.xlsx"

    return StreamingResponse(
        excel_io,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
