from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# This will receive the POST request from Salesforce
@app.route('/start-excel-job', methods=['POST'])
def start_excel_job():
    data = request.get_json()
    record_id = data.get('recordId', 'UNKNOWN')
    
    print(f"Received recordId: {record_id}")
    
    # Simulate async Excel generation
    return jsonify({
        "jobId": "fake-job-id-001",
        "status": "started"
    })

# Salesforce will poll this endpoint to get job status
@app.route('/get-excel-job-status', methods=['GET'])
def get_excel_job_status():
    job_id = request.args.get('jobId')
    
    # Simulated ready file URL
    download_url = "https://excel-generator-pbcg.onrender.com/fake-download-url.xlsx"

    return jsonify({
        "status": "completed",
        "downloadUrl": download_url
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
