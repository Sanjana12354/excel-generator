from flask import Flask, request, send_file, jsonify
import openpyxl
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    try:
        # Get JSON payload
        data = request.get_json()

        # Create a new Excel workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Order Data"

        # Optional: Set headers
        ws.append([
            'CUSTOMER_NAME',
            'CUSTOMER_PURCHASE_ORDER_NUMBER',
            'MTV_ORDER_NUMBER',
            'MTV_ORDER_ITEM_NUMBER',
            'TYPE',
            'SIZE',
            'CLASS',
            'CONFIGURATION',
            'OPERATOR',
            'ACCEPTED_QUANTITY'
        ])

        # Append data row
        ws.append([
            data.get('CUSTOMER_NAME', 'N/A'),
            data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', 'N/A'),
            data.get('MTV_ORDER_NUMBER', 'N/A'),
            data.get('MTV_ORDER_ITEM_NUMBER', 'N/A'),
            data.get('TYPE', 'N/A'),
            data.get('SIZE', 'N/A'),
            data.get('CLASS', 'N/A'),
            data.get('CONFIGURATION', 'N/A'),
            data.get('OPERATOR', 'N/A'),
            data.get('ACCEPTED_QUANTITY', 'N/A')
        ])

        # Save the file temporarily
        file_path = 'MTV-QC-FM-013A_Rev.00 - MTC.xlsx'
        wb.save(file_path)

        # Return the file as response
        return send_file(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='GeneratedExcel.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
