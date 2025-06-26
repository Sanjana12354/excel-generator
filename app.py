from flask import Flask, request, send_file, jsonify
import openpyxl
import os
import tempfile
import shutil

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    try:
        data = request.get_json()

        template_path = 'MTV-QC-FM-013A_Rev.00 - MTC.xlsx'
        if not os.path.exists(template_path):
            return jsonify({'error': 'Template file not found.'}), 500

        temp_dir = tempfile.mkdtemp()
        temp_file = os.path.join(temp_dir, 'FilledTemplate.xlsx')
        shutil.copy(template_path, temp_file)

        wb = openpyxl.load_workbook(temp_file)
        ws = wb.active

        def safe_write(cell, value):
            try:
                ws[cell] = value
            except:
                pass

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

        wb.save(temp_file)

        return send_file(
            temp_file,
            as_attachment=True,
            download_name='GeneratedExcel.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

    finally:
        if 'temp_dir' in locals() and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == '__main__':
    app.run(debug=True)
