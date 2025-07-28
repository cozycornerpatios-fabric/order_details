import sys
import json
import requests
import pandas as pd
import traceback

# === CONFIGURATION ===
API_TOKEN = '1e2c6dd2-c8ed-490d-8c70-a4bb21152682'  # Replace with your actual API token
USER_TEMPLATE_PATH = 'format.xls'  # Excel template with desired format
OUTPUT_FILE = 'formatted_sales_orders.xlsx'

BASE_URL = 'https://api.katanamrp.com/v1'
HEADERS = {
    'Authorization': f'Bearer {API_TOKEN}',
    'Content-Type': 'application/json'
}

# === FUNCTION TO FETCH ALL SALES ORDERS ===
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

# === FILTER ORDERS BY REFERENCE ID ===
def filter_orders(all_orders, order_numbers):
    return [order for order in all_orders if order['reference'] in order_numbers]

# === MAP TO TEMPLATE FORMAT ===
def map_to_template(filtered_orders, template_path):
    user_template = pd.read_excel(template_path)
    column_mapping = {
        '#SO': 'reference',
        'Customer Name': 'customer_name',
        'Customer Email': 'customer_email',
        'Order Date': 'created_date',
        'Delivery Date': 'delivery_date',
        'Item Name': 'order_lines.name',
        'Quantity': 'order_lines.quantity',
        'Variant Code': 'order_lines.variant_code',
        'Warehouse': 'warehouse_name',
        'Tracking Number': 'tracking_number',
        'Shipping Carrier': 'shipping_carrier'
    }

    output_data = pd.DataFrame(columns=user_template.columns)

    for order in filtered_orders:
        row = {}
        for col in user_template.columns:
            if col in column_mapping:
                katana_col = column_mapping[col]
                if '.' in katana_col and 'order_lines' in order and order['order_lines']:
                    line = order['order_lines'][0]
                    row[col] = line.get(katana_col.split('.')[-1], '')
                else:
                    row[col] = order.get(katana_col, '')
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
            order_numbers = json.loads(sys.argv[1])  # Expecting a JSON array
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
