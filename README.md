# Universal 5-Day Inter-Run Precision App

A Streamlit starter application that extracts numeric assay results from instrument PDF reports and populates a five-day precision workbook based on an uploaded Excel template.

## Features
- Upload an original Excel precision template on the first run.
- Upload one or more PDF reports and assign a default day to each report.
- Embedded labels such as `INTER DAY2`, `DAY301`, and `DAY302` override the default day when detected.
- Detects analyte section headings dynamically rather than using a fixed analyte list.
- Separates Level 1 and Level 2 using default QC identifiers.
- Selects up to five results closest to the median for each analyte/day/level.
- Keeps extracted results in `Raw_Data_Audit` and a population summary in `Upload_Summary`.
- Upload the latest populated workbook in later sessions to append new days without losing earlier data.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy with GitHub + Streamlit Community Cloud
1. Create a GitHub repository, for example `universal-interrun-precision`.
2. Upload `app.py`, `requirements.txt`, and this `README.md`.
3. Open https://share.streamlit.io/ and sign in with GitHub.
4. Choose **New app**, select the repository, branch `main`, and file `app.py`.
5. Deploy. Streamlit will install the packages from `requirements.txt`.

## How to use
### First run
1. Upload your original `.xlsx` template.
2. Upload the PDF report(s).
3. Assign a default day to each PDF.
4. Click **Process PDFs** and download the populated workbook.

### Later uploads
1. Upload the latest populated workbook in the **latest populated workbook** field.
2. Upload new PDF report(s), assign their day(s), and process.
3. Download the updated workbook and keep it as the master file for the next upload.

## Important parsing notes
The supplied sample reports contain different naming patterns. The starter rules identify:
- Level 1: `101`, `301`, `AB1`, `K5972`, `6304961`
- Level 2: `102`, `302`, `AB2`, `K5973`, `6304962`

Review and adapt `infer_level()` for other QC material identifiers. If a PDF uses a different naming convention, the app may mark it `Unassigned`; do not use those results until the mapping is reviewed and corrected.

The parser reads text-based PDFs. Scanned/image-only PDFs require OCR and are not included in this starter package.

## Scientific caution
The app's requested selection method is the five observations closest to the median. This can reduce the observed SD/CV compared with using all valid replicates. The audit sheet preserves the extracted values, but the selected values and final precision statistics must be reviewed by a qualified laboratory professional and aligned with the approved protocol. Do not remove results solely because they are statistically inconvenient.

## Known limitations
- PDF layouts vary between instrument software versions; always inspect the audit sheet.
- The first worksheet in the uploaded template is used as the sheet layout to clone for newly detected analytes.
- The app creates a new worksheet for newly detected analytes, using the template's format.
- Formula recalculation is performed by Excel/LibreOffice when the downloaded workbook is opened.
- For a regulated validation workflow, add user authentication, audit logging, versioning, and formal verification/validation before production use.
