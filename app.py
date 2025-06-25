from flask import Flask, request, send_file, abort
import openpyxl
import tempfile

app = Flask(__name__)
API_KEY = "your-secret-api-key"  # Replace this with something secure

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    auth = request.headers.get("Authorization", "")
    if auth != f"Bearer {API_KEY}":
        abort(401, description="Unauthorized")

    data = request.json.get("data", {})

    wb = openpyxl.load_workbook('generated_template.xlsx')
    ws = wb.active

    ws['C4'] = data.get('CUSTOMER_NAME', '')
    ws['C5'] = data.get('CUSTOMER_PURCHASE_ORDER_NUMBER', '')
    ws['C6'] = data.get('MTV_ORDER_NUMBER', '')
    ws['C7'] = data.get('MTV_ORDER_ITEM_NUMBER', '')
    ws['C9'] = data.get('TYPE', '')
    ws['C10'] = data.get('SIZE', '')
    ws['C11'] = data.get('CLASS', '')
    ws['C12'] = data.get('CONFIGURATION', '')
    ws['C13'] = data.get('OPERATOR', '')
    ws['C14'] = data.get('ACCEPTED_QUANTITY', '')

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp.name)
    return send_file(temp.name, as_attachment=True, download_name="MTV_Template.xlsx")

if __name__ == '__main__':
    app.run()
