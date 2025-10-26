PPT KPI Extractor

PPT KPI Extractor is a Python-based backend application that extracts Key Performance Indicators (KPIs) from PowerPoint presentations and provides a downloadable Excel report. The application is designed to be modular, reusable, and deployment-ready.

Features

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

