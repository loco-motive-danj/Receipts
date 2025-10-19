# --- imports ---
import io
import os
import json
import time
import requests
import pandas as pd
import glob
from flask import Flask, send_file
import threading
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- load token from Replit secret or local file ---
print("🔐 Loading token from Replit environment...")

token_json = os.getenv("GOOGLE_TOKEN")  # Replit secret
if not token_json and os.path.exists("token.json"):
  with open("token.json") as f:
    token_json = f.read()

if not token_json:
  raise Exception(
      "❌ No Google token found. Add GOOGLE_TOKEN in Replit Secrets!")

creds = Credentials.from_authorized_user_info(json.loads(token_json))

drive = build("drive", "v3", credentials=creds)
print("✅ Google Drive connected successfully!")

# ---- Load config ----
with open("config.json") as f:
  cfg = json.load(f)

AZURE_ENDPOINT = "https://receiptinvoiceaid.cognitiveservices.azure.com/"
AZURE_KEY = os.getenv("AZURE_KEY")
MODEL = "prebuilt-receipt"
FOLDER_ID = "1gBOXAU9b1zSt06c-1YPQcmPiu02zTdXZ"

# ---- Google Drive auth ----

SCOPES = ["https://www.googleapis.com/auth/drive"]


from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "service_account.json"  

def get_service_account_drive():
  creds_json = os.getenv("SERVICE_ACCOUNT_JSON")
  if not creds_json:
      raise Exception("❌ SERVICE_ACCOUNT_JSON not found in environment variables!")

  creds_info = json.loads(creds_json)
  creds = service_account.Credentials.from_service_account_info(
      creds_info,
      scopes=["https://www.googleapis.com/auth/drive"]
  )

  print("✅ Connected to Google Drive via Service Account")
  return build("drive", "v3", credentials=creds)

drive = get_service_account_drive()



def list_files(folder_id):
  q = f"'{folder_id}' in parents and trashed=false"
  res = drive.files().list(q=q).execute()
  return res.get("files", [])


def download_file(file_id, name):
  req = drive.files().get_media(fileId=file_id)
  fh = io.BytesIO()
  downloader = MediaIoBaseDownload(fh, req)
  done = False
  while not done:
    _, done = downloader.next_chunk()
  fh.seek(0)
  return fh.read()


def analyze_receipt(file_bytes):
  headers = {
      "Ocp-Apim-Subscription-Key": AZURE_KEY,
      "Content-Type": "application/octet-stream"
  }
  params = {"api-version": "2023-07-31"}
  resp = requests.post(
      f"{AZURE_ENDPOINT}formrecognizer/documentModels/{MODEL}:analyze",
      headers=headers,
      params=params,
      data=file_bytes)
  if resp.status_code not in (200, 202):
    print("Azure POST failed:", resp.status_code, resp.text)
    raise Exception(f"Azure request failed ({resp.status_code})")

  op = resp.headers.get("operation-location")
  if not op:
    raise Exception(
        "No operation-location header returned from Azure. Check endpoint, key, and Content-Type."
    )
  print("✅ Operation location:", op)

  # Poll for results
  while True:
    r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_KEY})
    data = r.json()
    if data.get("status") == "succeeded":
      return data
    elif data.get("status") == "failed":
      raise Exception("Azure analysis failed.")
    time.sleep(2)


def parse_and_save(data, name):
  docs = data["analyzeResult"]["documents"]
  if not docs:
    return None
  items = docs[0]["fields"]["Items"]["valueArray"]
  rows = []
  for it in items:
    obj = it["valueObject"]
    rows.append({
        "Description":
        obj.get("Description", {}).get("valueString", ""),
        "Quantity":
        obj.get("Quantity", {}).get("valueNumber", 1),
        "Total":
        obj.get("TotalPrice", {}).get("valueNumber", 0)
    })
  df = pd.DataFrame(rows)
  os.makedirs("outputs", exist_ok=True)
  out_path = "outputs/All_Receipts_Combined.xlsx"
  df_combined.to_excel(out_path, index=False)
  return out_path


def upload_to_drive(local_path, folder_id):
  file_metadata = {
      "name":
      os.path.basename(local_path),
      "parents": [folder_id],
      "mimeType":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }

  media = MediaFileUpload(
      local_path,
      mimetype=
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

  try:
    drive.files().create(body=file_metadata,
                         media_body=media,
                         fields="id, parents").execute()
    print(
        f"✅ Uploaded {os.path.basename(local_path)} to Drive folder {folder_id}"
    )
  except Exception as e:
    print(f"⚠️ Upload failed: {e}")


def merge_excels(output_dir="outputs"):
  all_files = glob.glob(os.path.join(output_dir, "*_parsed.xlsx"))
  dfs = [pd.read_excel(f) for f in all_files]
  merged = pd.concat(dfs, ignore_index=True)
  merged.to_excel(os.path.join(output_dir, "All_Receipts_Combined.xlsx"), index=False)
  print("🧾 Combined Excel saved as All_Receipts_Combined.xlsx")

def main():
  files = list_files(FOLDER_ID)
  for f in files:
      name = f["name"]
      if name.endswith(".xlsx") or "_parsed" in name:
          print(f"⏭️  Skipping {name} (not a receipt image/PDF)")
          continue

      print(f"Processing {name}...")
      content = download_file(f["id"], f["name"])
      parsed = analyze_receipt(content)
      out_path = parse_and_save(parsed, f["name"])
      if out_path:
          print(f"✅ Parsed and saved locally: {out_path}")
      else:
          print(f"⚠️ No data found for {f['name']}")

  # 🧩 Merge everything into one file at the end
  merge_excels()



app = Flask(__name__)

@app.route('/')
def home():
    return """
    <h2>🧾 Receipt Parser Running</h2>
    <p>The app is actively monitoring your Drive folder for receipts.</p>
    <p><a href='/download'>Download merged Excel results</a></p>
    """

@app.route('/download')
def download_results():
    path = "outputs/All_Receipts_Combined.xlsx"
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    else:
        return "<p>No results found yet. Please upload some receipts first.</p>"

def run_flask():
    app.run(host='0.0.0.0', port=8080)

if __name__ == "__main__":
    # Run the Flask app alongside your existing pipeline
    threading.Thread(target=run_flask).start()
    main()
