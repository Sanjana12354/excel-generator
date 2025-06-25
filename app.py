from flask import Flask, request, send_file, jsonify
import openpyxl
import os
import uuid

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    try:
        data = request.get_json()

        # Load the existing template
        template_path = "MTV-QC-FM-013A_Rev.00 - MTC.xlsx"
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        # Fill in the values (update specific cells as needed)
        ws["A2"] = data.get('CUSTOMER_NAME', 'N/A')
        ws["B2"] = data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', 'N/A')
        ws["C2"] = data.get('MTV_ORDER_NUMBER', 'N/A')
        ws["D2"] = data.get('MTV_ORDER_ITEM_NUMBER', 'N/A')
        ws["E2"] = data.get('TYPE', 'N/A')
        ws["F2"] = data.get('SIZE', 'N/A')
        ws["G2"] = data.get('CLASS', 'N/A')
        ws["H2"] = data.get('CONFIGURATION', 'N/A')
        ws["I2"] = data.get('OPERATOR', 'N/A')
        ws["J2"] = data.get('ACCEPTED_QUANTITY', 'N/A')

        # Save to a unique filename in static folder
        if not os.path.exists("static"):
            os.makedirs("static")

        filename = f"static/excel_{uuid.uuid4().hex}.xlsx"
        wb.save(filename)

        # Return the file URL
        return jsonify({"url": f"https://excel-generator-pbcg.onrender.com/{filename}"})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
