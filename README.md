# CV Certification Scanner (PDF-based)

This repository contains a deterministic Python tool that scans employee CVs in **PDF format** and generates an **Excel report** of Microsoft Azure certifications held by each employee.

The solution is designed to be **auditable, explainable, and stable**, avoiding GenAI/LLM guessing and relying on exact, case-insensitive text matching.

---

## Features

- Recursively scans all CV PDFs in a directory
- Reliable full-text extraction from PDFs (including content originally stored in Word textboxes)
- Case-insensitive, exact certification name matching
- Multi-vendor certification support:
  - Microsoft
  - AWS
  - Google Cloud
  - Databricks
- Deduplicated output (one row per employee per certification)
- Excel output with:
  - `employeeName`
  - `certName`
  - `certId`
  - `vendor`
- Console summary report:
  - Total CVs scanned
  - Total matched rows
  - Unique employees matched
  - Matches per certification

---

## Prerequisites

- Python 3.9+
- pip

Install dependencies using the provided `requirements.txt`:
```bash
pip install -r requirements.txt
````

---

## Folder Structure

```text
CVs/
├── pdfs/        # All CV PDFs (input)
├── simple-scan.py    # Main scanner script
├── requirements.txt    # Python dependencies
└── cert_name_matches.xlsx  # Output
```

---

## Usage

1. Place all CV PDFs under the `pdfs/` directory
2. Update `ROOT` in the script if needed:

   ```python
   ROOT = Path("/Users/your_path/pdfs")
   ```
3. Run:

   ```bash
   python simple-scan.py
   ```
4. Review the generated `ai_cert_employee.xlsx`

---

## Optional: Convert DOCX CVs to PDF (Recommended)

If your CVs are originally in **DOCX format**, converting them to PDF first is strongly recommended.
Many Word CV templates store content in textboxes or shapes, which are unreliable to extract directly from DOCX.

### Using LibreOffice (macOS / Linux)

Install LibreOffice:

```bash
brew install --cask libreoffice
```

From the parent directory:

```bash
find CVs -name "*.docx" -print0 | \
xargs -0 /Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless --convert-to pdf --outdir pdfs
```

This converts **all DOCX files recursively** into PDFs and saves them in the `pdfs/` folder.

---

## Matching Logic

* Case-insensitive exact substring matching
* No fuzzy matching or inference
* Certification IDs and vendors are sourced from a controlled catalog
* Vendor is dynamically populated per certification

This ensures results are predictable and suitable for audit or compliance reporting.

---

## Use Cases

* Skills inventory and certification audits
* Pre-sales capability mapping
* Workforce analytics and reporting
* Internal compliance checks

---

## Notes
* If an employee has duplicated CVs stored under different file naming conventions, the tool will treat them as separate inputs and count them multiple times
* This tool intentionally avoids GenAI/LLMs for reliability and explainability
* PDF-first processing is used to ensure maximum text coverage
