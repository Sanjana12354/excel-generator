# app.py
from flask import Flask, request, jsonify
import uuid
import threading
import time

app = Flask(__name__)

# Store job results in-memory (for demo; use Redis or DB in production)
results = {}

def generate_excel_job(job_id, data):
    # Simulate long-running Excel generation (replace with real logic)
    time.sleep(150)  # 2.5 minutes
    # After generation, set the result
    results[job_id] = {
        "status": "completed",
        "download_url": f"https://your-bucket.com/excels/{job_id}.xlsx"
    }

@app.route('/start-excel-job', methods=['POST'])
def start_excel_job():
    data = request.json
    job_id = str(uuid.uuid4())
    results[job_id] = { "status": "processing" }
    thread = threading.Thread(target=generate_excel_job, args=(job_id, data))
    thread.start()
    return jsonify({ "job_id": job_id })

@app.route('/get-excel-job/<job_id>', methods=['GET'])
def get_excel_job(job_id):
    result = results.get(job_id)
    if not result:
        return jsonify({ "error": "Job not found" }), 404
    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True)
