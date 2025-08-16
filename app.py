from flask import Flask, request, render_template, send_file
from pptx import Presentation
import pandas as pd
import re
KPI_FINDER = re.compile(r"""
    (?P<kpi>[A-Za-z][A-Za-z0-9\s/&\-()]+?)      # KPI name
    (?:\s*[:\-–—]\s*|\s+is\s+|\s+was\s+|\s+)    # separator: colon/dash/verb/space
    (?P<value>
        \(?\s*                                   # optional opening parenthesis
        (?:[₹$€£]\s*)?                           # optional currency symbol
        [\d,]+(?:\.\d+)?                         # number with , and . support
        (?:\s?(?:cr|crore|lakh|m|mn|b|bn))?      # optional magnitude word
        %?                                       # optional percent sign
        \s*\)?                                   # optional closing parenthesis
    )
""", re.IGNORECASE | re.VERBOSE)

import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- Helper function ---
def extract_runs_text(shape):
    parts = []
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if t:
                    parts.append(t)
    return " ".join(parts).strip()
 
# --- Dynamic KPI extractor ---
def extract_kpis(ppt_file, filename):
    prs = Presentation(ppt_file)
    rows = []
    manager = os.path.splitext(filename)[0]

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            text = extract_runs_text(shape)
            if not text:
                continue
            for m in KPI_FINDER.finditer(text):
                kpi = m.group("kpi").strip().replace("\n", " ")
                value = m.group("value").strip()
                rows.append({
                    "Manager/Sector": manager,
                    "Slide": slide_idx,
                    "KPI": kpi,
                    "Value": value
                })

    if rows:
        df = pd.DataFrame(rows).drop_duplicates(
            subset=["Manager/Sector", "Slide", "KPI", "Value"]
        )
        return df
    return pd.DataFrame(columns=["Manager/Sector", "Slide", "KPI", "Value"])
