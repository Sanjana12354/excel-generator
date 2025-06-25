from flask import Flask, request, jsonify
from openpyxl import Workbook
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    # You can print to logs to verify payload
    print("Received data:", data)

    wb = Workbook()
    ws = wb.active

    ws.append(['Key', 'Value'])
    for key, value in data.items():
        ws.append([key, value])

    # Save to file
    file_path = "/tmp/generated_file.xlsx"
    wb.save(file_path)

    # Return download link (you can customize this if needed)
    return jsonify({"url": f"https://{request.host}/download"})


@app.route('/download')
def download_file():
    return app.send_static_file("generated_file.xlsx")


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
