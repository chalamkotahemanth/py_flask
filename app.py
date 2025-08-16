from flask import Flask, request, render_template, send_file
from pptx import Presentation
import pandas as pd
import re
import os
import json
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Load dynamic KPI patterns
with open('kpi_patterns.json') as f:
    KPI_PATTERNS = json.load(f)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('pptfiles')
        all_data = []

        for ppt_file in files:
            if ppt_file:
                filename = ppt_file.filename
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                ppt_file.save(file_path)
                data = extract_kpis(ppt_file=file_path, filename=filename)
                all_data.append(data)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(UPLOAD_FOLDER, f'KPI_Results_{timestamp}.xlsx')
            final_df.to_excel(output_file, index=False)
            return send_file(output_file, as_attachment=True)

    return render_template('index.html')

def extract_kpis(ppt_file, filename):
    prs = Presentation(ppt_file)
    data = []
    manager = os.path.splitext(filename)[0]

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text
                for kpi_name, pattern in KPI_PATTERNS.items():
                    match = re.search(pattern, text, re.I)
                    if match:
                        data.append({"Manager/Sector": manager, "KPI": kpi_name, "Value": match.group(1)})

    return pd.DataFrame(data)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
