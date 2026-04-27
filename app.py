from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import io
from datetime import datetime

# ── Google Sheets ──────────────────────────────────────────────────────────────
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

app = Flask(__name__)
CORS(app)  # Allow requests from the React frontend

# ── Google Sheets setup ────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_gspread_client():
    """Build a gspread client from the service-account JSON stored in env vars."""
    import json
    service_account_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)

def get_or_create_sheet(client, spreadsheet_id: str, sheet_name: str, headers: list):
    """Return (worksheet, created_new). Creates the sheet with headers if missing."""
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(headers))
        ws.append_row(headers)
    return ws

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def home():
    return "Backend is running!"


@app.route("/contact", methods=["POST"])
def save_contact():
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
    HEADERS = ["Name", "Email", "Phone", "Location", "Company", "Purpose", "Date"]

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data received"}), 400

        name     = data.get("name", "").strip()
        email    = data.get("email", "").strip()
        phone    = data.get("phone", "").strip()
        location = data.get("location", "").strip()
        company  = data.get("company", "").strip()
        purpose  = data.get("purpose", "").strip()
        date     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not name or not email:
            return jsonify({"error": "Name and Email are required"}), 400

        client = get_gspread_client()
        ws = get_or_create_sheet(client, SPREADSHEET_ID, "Contacts", HEADERS)
        ws.append_row([name, email, phone, location, company, purpose, date])

        return jsonify({"message": "Contact saved successfully!"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/contacts", methods=["GET"])
def download_contacts():
    """Download the contacts sheet as an Excel file."""
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
    HEADERS = ["Name", "Email", "Phone", "Location", "Company", "Purpose", "Date"]

    try:
        client = get_gspread_client()
        ws = get_or_create_sheet(client, SPREADSHEET_ID, "Contacts", HEADERS)
        records = ws.get_all_records()

        df = pd.DataFrame(records) if records else pd.DataFrame(columns=HEADERS)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Contacts")
        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="contacts.xlsx",
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/message", methods=["POST"])
def save_message():
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
    HEADERS = ["Message", "Date"]

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data received"}), 400

        message = data.get("message", "").strip()
        date    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not message:
            return jsonify({"error": "Message cannot be empty"}), 400

        client = get_gspread_client()
        ws = get_or_create_sheet(client, SPREADSHEET_ID, "Messages", HEADERS)
        ws.append_row([message, date])

        return jsonify({"message": "Message saved successfully!"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
