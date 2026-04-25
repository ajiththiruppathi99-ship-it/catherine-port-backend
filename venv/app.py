from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend is running!"

FILE_PATH = "data/contacts.xlsx"

@app.route('/contact', methods=['POST'])
def save_contact():
    try:
        data = request.json

        name = data.get('name')
        email = data.get('email')
        message = data.get('message')
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        new_data = {
            "Name": name,
            "Email": email,
            "Message": message,
            "Date": date
        }

        if os.path.exists(FILE_PATH):
            df = pd.read_excel(FILE_PATH)
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        else:
            df = pd.DataFrame([new_data])

        df.to_excel(FILE_PATH, index=False)

        return jsonify({"message": "Saved successfully"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)