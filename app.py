from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    print("Received data:", data)

    try:
        # Load your actual Excel template
        template_path = "MTV-QC-FM-013A_Rev.00 - MTC.xlsx"
        output_path = "generated_file.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active

        # Safely write to starting cells of merged ranges
        if ws['C4'].coordinate == ws.merged_cells.ranges[0].start_cell.coordinate:
            ws['C4'] = data.get('CUSTOMER_NAME', 'N/A')

        ws['C5'] = data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', 'N/A')  # ensure C5 is not a merged range
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

    except Exception as e:
        print("ERROR:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route('/download')
def download_file():
    return send_file("generated_file.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
