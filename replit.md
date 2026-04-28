# Generator Instrukcji BHP

Streamlit application that generates Polish BHP (workplace safety) instruction Excel files from PDF safety data sheets.

## Stack
- Python 3.12
- Streamlit (frontend)
- pdfplumber (PDF parsing)
- openpyxl (Excel generation)
- Pillow (image handling)

## Project Layout
- `app.py` — Streamlit UI entry point
- `main.py` — PDF data extraction and Excel workbook population logic
- `assets/` — Static assets used in generated workbooks
- `DOCS/` — Sample input PDFs
- `temp_images/` — Temporary image storage during processing
- `wzor.xlsx` — Excel template used as the basis for generated instructions
- `.streamlit/config.toml` — Streamlit server configuration (port 5000, host 0.0.0.0)

## Running Locally
The "Start application" workflow runs:
```
streamlit run app.py --server.port 5000 --server.address 0.0.0.0 --server.headless true --server.enableCORS false --server.enableXsrfProtection false
```

## Deployment
Configured for autoscale deployment running the same Streamlit command.
