from flask import Flask, request, send_file, jsonify
import openpyxl
from openpyxl.utils import get_column_letter
import tempfile
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    print("Received data:", data)

    # Load template
    template_path = 'template.xlsx'
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # Safe function to avoid writing to merged cells
    def safe_write(cell, value):
        try:
            if isinstance(ws[cell], openpyxl.cell.cell.Cell):
                ws[cell].value = value
        except Exception as e:
            print(f"Skip cell {cell} due to error: {e}")

    # Write data to Excel
    safe_write('C4', data.get('CUSTOMER_NAME', 'N/A'))
    safe_write('C5', data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', 'N/A'))
    safe_write('C6', data.get('MTV_ORDER_NUMBER', 'N/A'))
    safe_write('C7', data.get('MTV_ORDER_ITEM_NUMBER', 'N/A'))
    safe_write('C9', data.get('TYPE', 'N/A'))
    safe_write('C10', data.get('SIZE', 'N/A'))
    safe_write('C11', data.get('CLASS', 'N/A'))
    safe_write('C12', data.get('CONFIGURATION', 'N/A'))
    safe_write('C13', data.get('OPERATOR', 'N/A'))
    safe_write('C14', data.get('ACCEPTED_QUANTITY', 'N/A'))

    # Save to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    return send_file(tmp_path, as_attachment=True, download_name="generated_excel.xlsx")

@app.route('/')
def root():
    return 'Excel API Running'

if __name__ == '__main__':
    app.run(debug=True)
