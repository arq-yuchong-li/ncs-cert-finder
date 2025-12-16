# CV Certification Scanner (PDF-based)

This repository contains a deterministic Python tool that scans employee CVs in **PDF format** and generates an **Excel report** of Microsoft Azure certifications held by each employee.

The solution is designed to be **auditable, explainable, and stable**, avoiding GenAI/LLM guessing and relying on exact, case-insensitive text matching.

---

## Features

- Recursively scans all CV PDFs in a directory
- Reliable text extraction from PDFs (handles content originally stored in Word textboxes)
- Case-insensitive, exact certification name matching
- Deduplicated output (one row per employee per certification)
- Excel output with:
  - `employeeName`
  - `certName`
  - `certId`
  - `vendor`
- Console summary report:
  - Total CVs scanned
  - Total matches
  - Matches per certification

---

## Prerequisites

- Python 3.9+
- pip

Python dependencies:
```bash
pip install pdfplumber pandas
````

---

## Folder Structure

```text
CVs/
├── pdfs/        # All CV PDFs (input)
├── simple-scan.py    # Main scanner script
└── cert_name_matches.xlsx  # Output
```

---

## Usage

1. Place all CV PDFs under the `pdfs/` directory
2. Update `ROOT` in the script if needed:

   ```python
   ROOT = Path("/Users/yourname/Downloads/CVs/pdfs")
   ```
3. Run:

   ```bash
   python script.py
   ```
4. Review the generated `cert_name_matches.xlsx`

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
* Certification IDs are derived from a controlled metadata map
* Vendor is fixed as **Microsoft Azure** in this script, change if needed

This ensures results are predictable and suitable for audit or compliance reporting.

---

## Use Cases

* Skills inventory and certification audits
* Pre-sales capability mapping
* Workforce analytics and reporting
* Internal compliance checks

---

## Notes

* This tool intentionally avoids GenAI/LLMs for reliability and explainability
* PDF-first processing is used to ensure maximum text coverage
