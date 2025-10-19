from flask import Flask, render_template_string, send_file
import pandas as pd
import os
from glob import glob

app = Flask(__name__)

@app.route('/')
def home():
    return render_template_string("""
        <h2>📊 Receipt Parser Dashboard</h2>
        <p>Receipts are automatically parsed in the background.</p>
        <form action="/merge" method="get">
            <button type="submit">📥 Download Combined Excel</button>
        </form>
    """)

@app.route('/merge')
def merge():
    files = glob("outputs/*.xlsx")
    if not files:
        return "No parsed files available yet.", 404

    all_data = []
    for f in files:
        df = pd.read_excel(f)
        df["Source"] = os.path.basename(f)
        all_data.append(df)

    merged = pd.concat(all_data, ignore_index=True)
    out_path = "outputs/All_Receipts_Combined.xlsx"
    merged.to_excel(out_path, index=False)

    return send_file(out_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
