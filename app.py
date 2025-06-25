from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook
import os

app = Flask(__name__)

EXCEL_FILENAME = "MTV-QC-FM-013A_Rev.00 - MTC.xlsx"
EXCEL_PATH = os.path.join(os.getcwd(), EXCEL_FILENAME)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    print("Received data:", data)

    wb = Workbook()
    ws = wb.active
    ws.append(['Key', 'Value'])
    for key, value in data.items():
        ws.append([key, value])

    wb.save(EXCEL_PATH)

    return jsonify({"url": f"https://{request.host}/download"})

@app.route('/download')
def download_file():
    return send_file(EXCEL_PATH, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
