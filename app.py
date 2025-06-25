from flask import Flask, request, jsonify, send_file
import uuid
import os
import pandas as pd

app = Flask(__name__)

# Store job results in memory for simplicity
jobs = {}

@app.route('/')
def home():
    return '✅ Excel Generator API is up and running!'

@app.route('/start-excel-job', methods=['POST'])
def start_excel_job():
    try:
        data = request.json.get('data')
        if not data:
            return jsonify({'error': 'Missing data'}), 400

        # Create job ID
        job_id = str(uuid.uuid4())
        file_path = f'{job_id}.xlsx'

        # Convert data to DataFrame and save as Excel
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)

        # Store job
        jobs[job_id] = {
            'status': 'completed',
            'file': file_path
        }

        return jsonify({'job_id': job_id}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/status/<job_id>', methods=['GET'])
def get_job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify({'status': job['status']}), 200

@app.route('/download/<job_id>', methods=['GET'])
def download_excel(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    file_path = job['file']
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
