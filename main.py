from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
import pandas as pd
import requests

app = FastAPI()

KATANA_API_TOKEN = "e505d03b-c53f-4519-9d6a-69c7637b0f64"
KATANA_API_BASE = "https://api.katanamrp.com/v1"

def fetch_order_by_number(order_number):
    headers = {"Authorization": f"Bearer {KATANA_API_TOKEN}"}
    response = requests.get(f"{KATANA_API_BASE}/sales_orders", headers=headers)

    if response.status_code != 200:
        raise HTTPException(status_code=500, detail="Failed to fetch data from Katana")

    orders = response.json().get("results", [])

    # Manually find the correct order by order_number (e.g. "#43-19905-21")
    matching_order = next((order for order in orders if order.get("order_number") == order_number), None)

    if not matching_order:
        raise HTTPException(status_code=404, detail="Order number not found")

    return matching_order


@app.get("/generate_excel")
def generate_excel(order_number: str = Query(...)):
    order = fetch_order_by_number(order_number)

    df = pd.DataFrame(columns=[
        "Sequence_Number", "Recipient_Contact Name", "Recipient_Company Name", "Recipient_Address Line 1",
        "Recipient_Address Line 2", "Recipient_Address Line 3", "Recipient_Country", "Recipient_City",
        "Recipient_State", "Recipient_Postal code", "Recipient_Phone Number", "Recipient_Phone_Ext.",
        "Recipient_Tax Number", "Recipient_Email", "Reference_1", "Bill Shipment To", "Account # (Shipment)",
        "Bill Duties & Taxes To", "Account # (Duties)", "Invoice Number", "Invoice Date", "Total No of Package",
        "Total Shipment weight", "Pkg_length", "Pkg_width", "Pkg_height", "Freight_charges",
        "Insurance_charges", "Other_charges", "Total GST Amt", "FOB Value", "Carriage Value",
        "Invoice Value", "CURRENCY", "Country of Manufacture", "COMMODITY", "HS CODE 1",
        "St. of Origin of goods", "Dis. Of Origin of goods", "QUANTITY 1", "UOM1", "UNIT_VALUE 1",
        "UNIT_Weight 1", "GST _%", "GST_Amount", "Additional Shipment/Invoice info. (If any)",
        "Packaging", "Terms_Of_Sales", "user_field_1", "user_field_2"
    ])

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
        "", "", "", "", "", "", "", "",  # Pkg dims, charges
        order["total_price"],
        "", "", "", "",  # Currency, origin, etc.
        order["order_lines"][0]["product"]["name"],
        order["order_lines"][0]["product"].get("hs_code", ""),
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
    ]

    file_path = f"/tmp/order_{order_number}.xlsx"
    df.to_excel(file_path, index=False)

    return FileResponse(file_path, filename=f"Customs_Template_{order_number}.xlsx")
