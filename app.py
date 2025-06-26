import os
from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from io import BytesIO

app = Flask(__name__)

TEMPLATE_FILENAME = "MTV-QC-FM-013A_Rev.00 - MTC.xlsx"

@app.route("/generate-excel", methods=["POST"])
def generate_excel():
    try:
        # Load the Excel template
        template_path = os.path.join(os.getcwd(), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template file not found at {template_path}"}), 500

        wb = load_workbook(template_path)
        ws = wb.active

        data = request.json

        def safe_write(cell, value):
            ws[cell] = value if value else "N/A"

        # Fill in values
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

        # Save to a BytesIO stream
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Send file back
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='GeneratedExcel.xlsx'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
