from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import os

app = Flask(__name__)

# Path to your Excel template
TEMPLATE_PATH = 'MTV-QC-FM-013A_Rev.00 - MTC.xlsx'
OUTPUT_PATH = 'generated_file.xlsx'

# Utility to write to a cell safely (avoid merged cells)
def safe_write(ws, cell, value):
    if isinstance(ws[cell], MergedCell):
        print(f"⚠️ Skipped writing to merged cell: {cell}")
        return
    ws[cell] = value or 'N/A'  # fallback to 'N/A' if None

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    print("Received data:", data)

    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Mapping data fields to Excel cells
        safe_write(ws, 'C4', data.get('CUSTOMER_NAME'))
        safe_write(ws, 'C5', data.get('CUSTOMER_PURCHASE_ORDER_NUMBER'))
        safe_write(ws, 'C6', data.get('MTV_ORDER_NUMBER'))
        safe_write(ws, 'C7', data.get('MTV_ORDER_ITEM_NUMBER'))
        safe_write(ws, 'C9', data.get('TYPE'))
        safe_write(ws, 'C10', data.get('SIZE'))
        safe_write(ws, 'C11', data.get('CLASS'))
        safe_write(ws, 'C12', data.get('CONFIGURATION'))
        safe_write(ws, 'C13', data.get('OPERATOR'))
        safe_write(ws, 'C14', data.get('ACCEPTED_QUANTITY'))

        wb.save(OUTPUT_PATH)

        return jsonify({"url": f"https://{request.host}/download"})
    except Exception as e:
        print("Error generating Excel:", e)
        return jsonify({"error": "Failed to generate Excel file"}), 500

@app.route('/download')
def download_file():
    return send_file(OUTPUT_PATH, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
