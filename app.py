from flask import Flask, request, jsonify
import subprocess
import json

app = Flask(__name__)

@app.route('/fetch-orders', methods=['POST'])
def fetch_orders():
    try:
        data = request.get_json()
        order_ids = data.get("order_ids", [])
        if not order_ids:
            return jsonify({"error": "No order_ids provided"}), 400

        # Call the fetcher script
        result = subprocess.run(
            ['python', 'katana_order_fetcher.py', json.dumps(order_ids)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        try:
            response = json.loads(result.stdout)
        except json.JSONDecodeError:
            response = {"status": "error", "message": "Script output not JSON", "details": result.stdout}

        return jsonify(response)

    except Exception as e:
        return jsonify({"error": str(e)}), 500
