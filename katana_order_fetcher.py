import sys
import json
import requests
import pandas as pd
import traceback

# === CONFIGURATION ===
API_TOKEN = '1e2c6dd2-c8ed-490d-8c70-a4bb21152682'
USER_TEMPLATE_PATH = 'format.xls'
OUTPUT_FILE = 'formatted_sales_orders.xlsx'

BASE_URL = 'https://api.katanamrp.com/v1'
HEADERS = {
    'Authorization': f'Bearer {API_TOKEN}',
    'Content-Type': 'application/json'
}

# === FETCH ALL SALES ORDERS ===
def fetch_all_sales_orders():
    orders = []
    endpoint = f'{BASE_URL}/sales_orders'
    page = 1
    while True:
        params = {'page_size': 100, 'page': page}
        response = requests.get(endpoint, headers=HEADERS, params=params)
        if response.status_code != 200:
            raise Exception(f"Error {response.status_code}: {response.text}")
        data = response.json().get('results', [])
        if not data:
            break
        orders.extend(data)
        page += 1
    return orders

# === FETCH CUSTOMER DETAILS ===
def fetch_customer(customer_id):
    url = f'{BASE_URL}/customers/{customer_id}'
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        return response.json()
    return {}

# === GET SHIPPING ADDRESS ===
def get_shipping_address(addresses, shipping_id):
    for addr in addresses:
        if addr['id'] == shipping_id and addr['entity_type'] == 'shipping':
            return addr
    return {}

# === FILTER ORDERS BY ORDER NUMBER ===
def filter_orders(all_orders, order_numbers):
    return [order for order in all_orders if order['order_no'] in order_numbers]

# === MAP TO TEMPLATE FORMAT ===
def map_to_template(filtered_orders, template_path):
    user_template = pd.read_excel(template_path)
    output_data = pd.DataFrame(columns=user_template.columns)

    for order in filtered_orders:
        customer = fetch_customer(order['customer_id'])
        shipping = get_shipping_address(order.get('addresses', []), order.get('shipping_address_id'))
        line = order.get('sales_order_rows', [{}])[0]
        
        row = {}
        for col in user_template.columns:
            if col == 'Reference_1':
                row[col] = order.get('order_no', '')
            elif col == 'Recipient_Company Name':
                row[col] = customer.get('company_name', '')
            elif col == 'Recipient_Contact Name':
                row[col] = customer.get('full_name', '')
            elif col == 'Recipient_Email':
                row[col] = customer.get('email', '')
            elif col == 'Recipient_Address Line 1':
                row[col] = shipping.get('line_1', '')
            elif col == 'Recipient_Address Line 2':
                row[col] = shipping.get('line_2', '')
            elif col == 'Recipient_Phone Number':
                row[col] = shipping.get('phone', '')
            elif col == 'Recipient_Country':
                row[col] = shipping.get('country', '')
            elif col == 'Recipient_City':
                row[col] = shipping.get('city', '')
            elif col == 'Recipient_State':
                row[col] = shipping.get('state', '')
            elif col == 'Recipient_Postal code':
                row[col] = shipping.get('zip', '')
            elif col == 'Invoice Date':
                row[col] = order.get('order_created_date', '')
            elif col == 'Total No of Package':
                row[col] = line.get('quantity', '')
            else:
                row[col] = ''
        output_data = output_data.append(row, ignore_index=True)

    return output_data

# === EXPORT TO EXCEL ===
def export_to_excel(df, output_file):
    df.to_excel(output_file, index=False)
    print(json.dumps({"status": "success", "output_file": output_file, "records": len(df)}))

# === MAIN EXECUTION ===
if __name__ == '__main__':
    try:
        if len(sys.argv) < 2:
            raise ValueError("No Order IDs provided")
        try:
            order_numbers = json.loads(sys.argv[1])
        except json.JSONDecodeError:
            raise ValueError("Invalid input format: not valid JSON")

        all_orders = fetch_all_sales_orders()
        filtered_orders = filter_orders(all_orders, order_numbers)
        formatted_data = map_to_template(filtered_orders, USER_TEMPLATE_PATH)
        export_to_excel(formatted_data, OUTPUT_FILE)

    except Exception as e:
        print(json.dumps({
            "status": "error",
            "message": str(e),
            "details": traceback.format_exc()
        }))
