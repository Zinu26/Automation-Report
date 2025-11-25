import os
import json
import chardet
import gspread
import xml.etree.ElementTree as ET
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
import re

# ==========================
# CONFIGURATION
# ==========================
TOKEN_FILE = "token.json"
CREDENTIALS_FILE = "<REDACTED_CREDENTIALS_FILE>.json"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]

SHARED_FOLDER_PATH = r"<REDACTED_SHARED_FOLDER_PATH>"
SPREADSHEET_ID = "<REDACTED_SPREADSHEET_ID>"
SHEET_NAME = "Sheet1"

# --- Exclusion Rules ---
EXCLUDED_GROUP_CODES = {"<REDACTED_CODE1>", "<REDACTED_CODE2>", "<REDACTED_CODE3>"}
EXCLUDED_TOWNSHIPS = {"", None, "<REDACTED_TOWNSHIP1>", "<REDACTED_TOWNSHIP2>"}
EXCLUDED_PROJECT_CODES = {"<REDACTED_PJ1>", "<REDACTED_PJ2>", "<REDACTED_PJ3>"}

# --- Email Notification ---
EMAIL_SENDER = "<REDACTED_EMAIL>"
EMAIL_RECIPIENTS = [
    "<REDACTED_EMAIL_1>",
    "<REDACTED_EMAIL_2>",
]

# ==========================
# AUTHENTICATION
# ==========================
def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            token_data = json.load(f)
        creds = Credentials.from_authorized_user_info(token_data, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                print("⚠️ Refresh failed, re-authenticating...")
                creds = None

        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=8080, access_type="offline", prompt="consent")

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds

# ==========================
# EMAIL FUNCTION
# ==========================
def send_email_notification(subject, body, to_email):
    try:
        # Load credentials and build Gmail service
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        service = build('gmail', 'v1', credentials=creds)

        # --- Format email body with HTML ---
        full_body = f"""
        <html>
            <body>
                <p><b>Hi,</b></p>
                <br><p>{body}</p><br>
                <p><b>Thank you and good day,</b><br>
                    <i>Automated Data Service</i></p>
            </body>
        </html>
        """

        # Create email
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        message.attach(MIMEText(full_body, 'html'))

        # Encode message for Gmail API
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        # Send email
        service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        print("📧 Email notification sent successfully.")

    except Exception as e:
        print(f"❌ Failed to send email: {e}")


# ==========================
# FETCH XML FILE (LOCAL ONLY)
# ==========================
def get_xml_file():
    today = datetime.now().strftime("%Y%m%d")
    filename = f"data_{today}.xml"
    local_path = os.path.join(SHARED_FOLDER_PATH, filename)

    if os.path.exists(local_path):
        print(f"📁 Found local XML file in shared folder: {local_path}")
        with open(local_path, "rb") as f:
            raw = f.read()
        result = chardet.detect(raw)
        encoding = result["encoding"] or "utf-8"
        return raw.decode(encoding, errors="replace"), True

    print(f"❌ XML file not found in shared folder: {local_path}")
    return None, False


# ==========================
# PARSE XML
# ==========================
def parse_custom_xml(xml_content):
    """Parse <Root><a ...><c .../></a>...</Root> into list of dictionaries"""
    rows = []
    try:
        root = ET.fromstring(xml_content)
    except Exception as e:
        print("❌ XML parsing error:", e)
        return rows

    for a_elem in root.findall("a"):
        a_attrs = a_elem.attrib.copy()
        c_elem = a_elem.find("c")
        if c_elem is not None:
            a_attrs.update(c_elem.attrib)
        rows.append(a_attrs)

    return rows

# ==========================
# CLASSIFY UNIT DESCRIPTION
# ==========================
def classify_description(desc):
    desc = (desc or "").upper()
    if re.search(r"STUDIO", desc):
        return "STUDIO"
    elif re.search(r"1BR|1-BR", desc):
        return "1BR"
    elif re.search(r"2BR|2-BR", desc):
        return "2BR"
    elif re.search(r"3BR|4BR|5BR|6BR|3-BR|4-BR|5-BR|6-BR", desc):
        return "3BR AND UP"
    else:
        return "OTHERS"


# ==========================
# MAIN
# ==========================
def main():
    creds = get_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        sheet = spreadsheet.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows="1000", cols="50")

    content, found = get_xml_file()
    date_today = datetime.now().strftime("%Y-%m-%d")

    if not found:
        send_email_notification(
            f"Dataset Update Failed ({date_today})",
            "No XML file detected in shared directory.",
            ", ".join(EMAIL_RECIPIENTS),
        )
        return

    xml_rows = parse_custom_xml(content)
    if not xml_rows:
        send_email_notification(
            f"Dataset Update Failed ({date_today})",
            "XML file found but contains no readable content.",
            ", ".join(EMAIL_RECIPIENTS),
        )
        return

    headers = sheet.row_values(1)
    if not headers:
        print("⚠️ No headers found in Sheet1.")
        return

    mapping = {
        "Group Code": "GroupCode",
        "Township": "Township",
        "Project Code": "ProjectCode",
        "Tower": "Tower",
        "Cluster": "Cluster",
        "Floor": "Floor",
        "Unit Code": "UnitCode",
        "Unit Description": "UnitDesc",
        "Area": "Area",
        "UnitPrice": "UnitPrice",
        "Launched Price": "Launched_Price",
        "Description": "Description"
    }

    cleaned_rows = []
    for item in xml_rows:
        if (
            (item.get("GroupCode") or "").strip() in EXCLUDED_GROUP_CODES or
            (item.get("Township") or "").strip() in EXCLUDED_TOWNSHIPS or
            (item.get("ProjectCode") or "").strip() in EXCLUDED_PROJECT_CODES
        ):
            continue
        cleaned_rows.append(item)

    data_rows = []
    for item in cleaned_rows:
        row = []
        for header in headers:
            xml_key = mapping.get(header)
            if header == "Description":
                row.append(classify_description(item.get("UnitDesc", "")))
            else:
                row.append(item.get(xml_key, "") if xml_key else "")
        data_rows.append(row)

    sheet.update(f"A2", data_rows)

    send_email_notification(
        f"Dataset Update Successful ({date_today})",
        f"Successfully processed and uploaded {len(data_rows)} rows.",
        ", ".join(EMAIL_RECIPIENTS),
    )

if __name__ == "__main__":
    main()