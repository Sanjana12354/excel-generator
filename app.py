# app.py
from flask import Flask, request, send_file, jsonify
import io
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    try:
        data = request.get_json()

        # Load the template
        template_path = 'MTV-QC-FM-013A_Rev.00 - MTC.xlsx'
        workbook = load_workbook(template_path)
        sheet = workbook.active

        # Safe write helper
        def safe_write(cell, value):
            try:
                sheet[cell] = value
            except:
                pass

        # Fill data into template
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

        # Save to BytesIO
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='GeneratedExcel.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500
