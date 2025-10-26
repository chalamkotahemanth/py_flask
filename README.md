**PPT KPI Extractor
**
**PPT KPI Extractor is a Python-based backend application that extracts Key Performance Indicators (KPIs) from PowerPoint presentations and provides a downloadable Excel report. The application is designed to be modular, reusable, and deployment-ready.

Features**

Extract KPIs from PowerPoint slides (tables, charts, and text).

Export extracted KPIs to a structured Excel file (.xlsx) for reporting.

Handles multiple PPTX files in batch.

Clean, modular code with reusable functions.

Ready for deployment with Procfile (e.g., Render, Heroku).

Tech Stack

Python 3.10+

python-pptx – Read and parse PowerPoint slides

pandas – Data manipulation

openpyxl – Excel export

Flask / FastAPI – Backend API endpoints

gunicorn – Production-ready WSGI server

Project Structure
ppt-kpi-extractor/
│
├── app/
│   ├── main.py             # Entry point (Flask/FastAPI server)
│   ├── extractor.py        # Core KPI extraction logic
│   ├── utils.py            # Helper functions (cleaning, parsing)
│   ├── routes.py           # API routes/endpoints
│   └── __init__.py
│
├── output/                 # Folder for exported Excel files
│   └── KPI_Report.xlsx
│
├── Procfile                # For deployment (e.g., Render / Heroku)
├── requirements.txt        # Python dependencies
└── README.md

Installation

Clone the repo:

git clone https://github.com/yourusername/ppt-kpi-extractor.git
cd ppt-kpi-extractor


Create a virtual environment:

python -m venv venv
source venv/bin/activate    # Linux/Mac
venv\Scripts\activate       # Windows


Install dependencies:

pip install -r requirements.txt

Running Locally

For FastAPI backend:

uvicorn app.main:app --reload


For Flask backend:

python app/main.py


Open http://127.0.0.1:8000 (FastAPI) or http://127.0.0.1:5000 (Flask) to test.

Upload PPTX files via API or endpoint.

Download the generated Excel report from the /output folder.

Deployment

Procfile (example for Render/Heroku):

web: gunicorn app.main:app


Make sure all dependencies are in requirements.txt.

Push to GitHub and connect to Render/Heroku for automatic deployment.

The backend will provide endpoints to upload PPTX and download Excel KPI reports.

Usage Example (Python Script)
from app.extractor import extract_kpis
from pathlib import Path

ppt_file = Path("samples/SamplePresentation.pptx")
output_file = Path("output/KPI_Report.xlsx")

# Extract KPIs and save to Excel
extract_kpis(ppt_file, output_file)
print(f"KPI report saved to {output_file}")

Code Quality & Best Practices

Modular: extractor.py handles all logic; routes.py handles API routes.

Reusable: Functions designed to handle multiple file formats and batch processing.

Clean: PEP8-compliant, meaningful variable names, structured error handling.

Deployable: Works with Procfile for easy deployment to cloud platforms.

Contribution

Fork the repository.

Create a new branch (feature-xyz).

Commit your changes (git commit -m "Add feature").

Push to branch (git push origin feature-xyz).

Open a Pull Request.


