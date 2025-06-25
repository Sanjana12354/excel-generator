from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    print("Received data:", data)

    template_path = "MTV-QC-FM-013A_Rev.00 - MTC.xlsx"
    output_path = "generated_filled_file.xlsx"

    wb = load_workbook(template_path)
    ws = wb.active

    # Fill Excel with API values
    ws['C4'] = data.get('CUSTOMER_NAME', 'N/A')
    ws['C5'] = data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', 'N/A')
    ws['C6'] = data.get('MTV_ORDER_NUMBER', 'N/A')
    ws['C7'] = data.get('MTV_ORDER_ITEM_NUMBER', 'N/A')
    ws['C9'] = data.get('TYPE', 'N/A')
    ws['C10'] = data.get('SIZE', 'N/A')
    ws['C11'] = data.get('CLASS', 'N/A')
    ws['C12'] = data.get('CONFIGURATION', 'N/A')
    ws['C13'] = data.get('OPERATOR', 'N/A')
    ws['C14'] = data.get('ACCEPTED_QUANTITY', 'N/A')

    wb.save(output_path)

    return jsonify({"url": f"https://{request.host}/download"})

@app.route('/download')
def download_file():
    return send_file("generated_filled_file.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
