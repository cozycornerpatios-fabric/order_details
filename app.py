from flask import Flask, request, jsonify
import subprocess
import json

app = Flask(__name__)

@app.route('/fetch-orders', methods=['POST'])
def fetch_orders():
    try:
        data = request.get_json()
        order_nos = data.get("order_no", [])
        if not order_nos:
            return jsonify({"error": "No order_no values provided"}), 400

        # Call the fetcher script with order_nos
        result = subprocess.run(
            ['python', 'katana_order_fetcher.py', json.dumps(order_nos)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        try:
            response = json.loads(result.stdout)
        except json.JSONDecodeError:
            response = {
                "status": "error",
                "message": "Script output not JSON",
                "details": result.stdout
            }

        return jsonify(response)

    except Exception as e:
        return jsonify({"error": str(e)}), 500
