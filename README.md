# Bulk Certificate Generator

A Flask web app that generates certificates in bulk from:

- A PowerPoint template (`.pptx`) with placeholders like `{{student_name}}`
- An Excel file (`.xlsx/.xls`) containing participant details in matching columns

## Features

- Upload PPTX certificate template
- Upload participant Excel file
- Placeholder-column validation (`{{column_name}}` in PPTX must exist in Excel)
- Generates one certificate PPTX per participant
- Downloads a ZIP containing:
  - all generated certificate PPTX files
  - `certificates_manifest.xlsx` summary file

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open: `http://localhost:5000`

## Placeholder format

If your template contains:

- `{{student_name}}`
- `{{course_name}}`

your Excel must include columns:

- `student_name`
- `course_name`
