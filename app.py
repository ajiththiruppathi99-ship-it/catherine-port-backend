from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Allow requests from the React frontend

# Path to the Excel file (relative to backend/ folder)
FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "contacts.xlsx")


# ──────────────────────────────────────────────
#  Contact Form  (6 fields)
# ──────────────────────────────────────────────
@app.route('/contact', methods=['POST'])
def save_contact():
    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "No data received"}), 400

        name     = data.get('name', '').strip()
        email    = data.get('email', '').strip()
        phone    = data.get('phone', '').strip()
        location = data.get('location', '').strip()
        company  = data.get('company', '').strip()
        purpose  = data.get('purpose', '').strip()
        date     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Basic validation
        if not name or not email:
            return jsonify({"error": "Name and Email are required"}), 400

        new_row = {
            "Name":     name,
            "Email":    email,
            "Phone":    phone,
            "Location": location,
            "Company":  company,
            "Purpose":  purpose,
            "Date":     date,
        }

        # Ensure data directory exists
        os.makedirs(os.path.dirname(FILE_PATH), exist_ok=True)

        # Append to existing sheet or create new one
        if os.path.exists(FILE_PATH):
            df = pd.read_excel(FILE_PATH)
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            df = pd.DataFrame([new_row])

        df.to_excel(FILE_PATH, index=False)

        return jsonify({"message": "Contact saved successfully!"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ──────────────────────────────────────────────
#  Message Form  (quick single-field message)
# ──────────────────────────────────────────────
@app.route('/message', methods=['POST'])
def save_message():
    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "No data received"}), 400

        message = data.get('message', '').strip()
        date    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not message:
            return jsonify({"error": "Message cannot be empty"}), 400

        new_row = {
            "Message": message,
            "Date":    date,
        }

        msg_file = os.path.join(os.path.dirname(__file__), "data", "messages.xlsx")
        os.makedirs(os.path.dirname(msg_file), exist_ok=True)

        if os.path.exists(msg_file):
            df = pd.read_excel(msg_file)
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            df = pd.DataFrame([new_row])

        df.to_excel(msg_file, index=False)

        return jsonify({"message": "Message saved successfully!"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
