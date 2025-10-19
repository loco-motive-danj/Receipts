import os
import io
import time
import pandas as pd
import requests
from flask import Flask, request, send_file, render_template_string

# -----------------------------
# Flask setup
# -----------------------------
app = Flask(__name__)

# -----------------------------
# Azure Form Recognizer setup
# -----------------------------
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT")
AZURE_KEY = os.getenv("AZURE_KEY")
MODEL = "prebuilt-receipt"

# Simple upload page HTML
UPLOAD_HTML = """
<!DOCTYPE html>
<html>
  <head>
    <title>Receipt Parser</title>
    <style>
      body { font-family: sans-serif; margin: 60px; }
      h1 { color: #2a7ae2; }
      form { margin-top: 30px; }
      input[type=file] { margin-bottom: 15px; }
      button { padding: 8px 14px; background: #2a7ae2; color: white; border: none; border-radius: 4px; }
    </style>
  </head>
  <body>
    <h1>Upload a Receipt</h1>
    <form method="POST" enctype="multipart/form-data">
      <input type="file" name="file" accept=".jpg,.jpeg,.png,.pdf" required><br>
      <button type="submit">Parse Receipt</button>
    </form>
  </body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def upload_receipt():
    if request.method == "GET":
        return render_template_string(UPLOAD_HTML)

    # Get file from upload
    file = request.files["file"]
    file_bytes = file.read()

    # Send to Azure Form Recognizer
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_KEY,
        "Content-Type": "application/octet-stream"
    }
    params = {"api-version": "2023-07-31"}

    resp = requests.post(
        f"{AZURE_ENDPOINT}formrecognizer/documentModels/{MODEL}:analyze",
        headers=headers, params=params, data=file_bytes
    )

    if resp.status_code not in (200, 202):
        return f"<p>❌ Azure request failed: {resp.status_code}<br>{resp.text}</p>", 400

    op = resp.headers.get("operation-location")
    if not op:
        return "<p>❌ No operation-location header returned from Azure.</p>", 400

    # Poll until complete
    while True:
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_KEY})
        data = r.json()
        if data.get("status") == "succeeded":
            break
        elif data.get("status") == "failed":
            return "<p>❌ Azure analysis failed.</p>", 500
        time.sleep(2)

    # Extract item data
    try:
        docs = data["analyzeResult"]["documents"]
        items = docs[0]["fields"]["Items"]["valueArray"]
        rows = []
        for it in items:
            obj = it["valueObject"]
            rows.append({
                "Description": obj.get("Description", {}).get("valueString", ""),
                "Quantity": obj.get("Quantity", {}).get("valueNumber", 1),
                "Total": obj.get("TotalPrice", {}).get("valueNumber", 0)
            })

        df = pd.DataFrame(rows)
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"{os.path.splitext(file.filename)[0]}_parsed.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return f"<p>⚠️ Error parsing receipt: {str(e)}</p>", 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
