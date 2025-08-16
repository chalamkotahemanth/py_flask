from flask import Flask, request, render_template, send_file
from pptx import Presentation
import pandas as pd
import re
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- Helper function ---
def extract_runs_text(shape):
    text_parts = []
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                text_parts.append(run.text.strip())
    return " ".join(text_parts).strip()

# --- Dynamic KPI extractor ---
def extract_kpis(ppt_file, filename):
    prs = Presentation(ppt_file)
    data = []
    manager = os.path.splitext(filename)[0]

    for slide in prs.slides:
        for shape in slide.shapes:
            text = extract_runs_text(shape)
            if not text:
                continue

            # Regex: KPI name + value (percentages, ₹ values, plain numbers)
            kpi_pattern = re.compile(
                r"([A-Za-z][A-Za-z\s%]+?)[:\s]\s*(₹?\s*[\d,]+(?:\.\d+)?%?|[\d\.]+%)",
                re.I
            )

            matches = kpi_pattern.findall(text)
            for kpi, value in matches:
                clean_kpi = kpi.strip().replace("\n", " ")
                clean_value = value.strip()
                data.append({
                    "Manager/Sector": manager,
                    "KPI": clean_kpi,
                    "Value": clean_value
                })

    return pd.DataFrame(data)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('pptfiles')
        all_data = []

        for ppt_file in files:
            if ppt_file and ppt_file.filename.endswith(".pptx"):
                filename = ppt_file.filename
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                ppt_file.save(file_path)
                df = extract_kpis(file_path, filename)
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(UPLOAD_FOLDER, f"KPI_Results_{timestamp}.xlsx")
            final_df.to_excel(output_file, index=False)
            return send_file(output_file, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
