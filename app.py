import io, re, copy, hashlib
from datetime import datetime
import fitz
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Universal Inter-Run Precision", page_icon="🧪", layout="wide")
st.title("🧪 Universal 5-Day Inter-Run Precision")
st.caption("Upload instrument PDF reports progressively. The workbook can be updated in any day order.")

# ------------------------- parsing helpers -------------------------
NUM_UNIT = re.compile(r"^\s*([-+]?\d[\d,]*(?:\.\d+)?)\s*([A-Za-zµμ/%]+(?:\s*/\s*[A-Za-z]+)?|Index)?\s*$")
NUM_ANY = re.compile(r"([-+]?\d[\d,]*(?:\.\d+)?)\s*([A-Za-zµμ/%]+(?:\s*/\s*[A-Za-z]+)?|Index)?")
SKIP_HEADINGS = {
    "ASSAY REPORT", "PATIENT NAME", "SID", "RESULT", "RLU OR MAU", "FLAGS",
    "RESULT DATE/TIME", "FINAL RESULT", "LAB EVALUATIONS", "CALIBRATOR NAME OR",
    "QC MATERIAL NAME", "CALIBRATOR ID OR QC", "MATERIAL ID", "ORDER TIME",
    "SOFTWARE VERSION"
}
KNOWN_NON_ANALYTE = re.compile(r"^(?:PAGE\s+\d+|PRINTED:|IRP\d+|SOFTWARE VERSION|ORDER TIME)", re.I)

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def is_analyte_heading(line):
    x = clean(line)
    if not x or x.upper() in SKIP_HEADINGS or KNOWN_NON_ANALYTE.match(x):
        return False
    if len(x) > 24 or re.search(r"\d{1,2}/\d{1,2}/\d{4}", x):
        return False
    if NUM_ANY.search(x) or " " in x and len(x.split()) > 2:
        return False
    # Assay/analyte labels are commonly short tokens such as eE2, HBsII, CHIV, PSA.
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z0-9_.+-]{0,17}", x))

def extract_pdf_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return "\n".join(page.get_text("text", sort=True) for page in doc)

def day_from_sid(sid, fallback):
    s = clean(sid).upper()
    patterns = [
        (r"\bINTER\s*DAY\s*([1-5])\b", 1),
        (r"\bDAY\s*([1-5])\s*0?[12]\b", 1), # DAY301/DAY302 => day 3
        (r"\bINTER\s*([1-5])\b", 1),
        (r"\bD([1-5])\b", 1),
    ]
    for pat, _ in patterns:
        m = re.search(pat, s)
        if m:
            return int(m.group(1))
    return int(fallback)

def infer_level(sid, qc_name="", qc_id=""):
    text = f"{sid} {qc_name} {qc_id}".upper()
    # More specific tokens first. These are defaults and can be overridden in the UI.
    if re.search(r"\b(?:102|302|AB[-_ ]?2|LEVEL[-_ ]?2|L[-_ ]?2|K5973|6304962)\b", text):
        return "Level 2"
    if re.search(r"\b(?:101|301|AB[-_ ]?1|LEVEL[-_ ]?1|L[-_ ]?1|K5972|6304961)\b", text):
        return "Level 1"
    # Common QC labels such as CHIVAB1-362 / CHIVAB2-363
    if re.search(r"AB[12]", text):
        return "Level 2" if re.search(r"AB2", text) else "Level 1"
    return "Unassigned"

def parse_pdf(pdf_bytes, filename, default_day):
    text = extract_pdf_text(pdf_bytes)
    lines = [clean(x) for x in text.splitlines()]
    records, current_analyte = [], None
    # Each numeric result is paired with the nearest preceding SID/QC identity.
    for i, line in enumerate(lines):
        # A section heading is followed by the report column headers. This avoids
        # mistaking QC material names (e.g. CHIVAB1-362) for analyte headings.
        lookahead = " ".join(lines[i+1:i+7]).upper()
        if is_analyte_heading(line) and ("PATIENT NAME" in lookahead or "CALIBRATOR NAME" in lookahead or "QC MATERIAL NAME" in lookahead):
            current_analyte = line
        m = NUM_UNIT.match(line)
        if not m:
            continue
        value = float(m.group(1).replace(",", ""))
        unit = clean(m.group(2))
        if not current_analyte:
            continue
        context = lines[max(0, i-7):i]
        context_text = " ".join(context)
        # Exclude RLU / MAU numbers: result is the first numeric value after the SID,
        # while RLU is usually the following line. This line-based parser only accepts
        # a number+unit line as the assay result.
        sid = ""
        qc_name = ""
        qc_id = ""
        for c in reversed(context):
            cu = c.upper()
            if not sid and (("INTER" in cu or "DAY" in cu or "LINE" in cu) and len(c) < 65):
                sid = c
            if not qc_id and re.fullmatch(r"[A-Z]?\d{4,10}", c, re.I):
                qc_id = c
            if not qc_name and re.search(r"[A-Za-z]{2,}[-_]?[A-Za-z0-9-]*", c) and c.upper() not in SKIP_HEADINGS:
                if c != sid and not re.search(r"INTER|DAY|LINE", c, re.I):
                    qc_name = c
        # Do not treat pure QC ID or instrument numbers as SID.
        if not sid:
            sid = qc_name or qc_id or ""
        # Keep only values that have an identifier in the nearby report rows.
        if not sid and not qc_name and not qc_id:
            continue
        day = day_from_sid(sid, default_day)
        level = infer_level(sid, qc_name, qc_id)
        records.append({
            "Analyte": current_analyte, "Day": day, "Level": level,
            "Result": value, "Unit": unit, "SID / QC": sid,
            "QC Material": qc_name, "QC ID": qc_id,
            "Source PDF": filename, "Page line": i+1
        })
    # Remove duplicate extraction artefacts, but preserve genuinely repeated values.
    return records

def best_five(values):
    vals = [float(v) for v in values if pd.notna(v)]
    if len(vals) <= 5:
        return vals
    med = pd.Series(vals).median()
    # Select closest to median; stable sort preserves report order for ties.
    indexed = list(enumerate(vals))
    chosen = sorted(indexed, key=lambda z: (abs(z[1]-med), z[0]))[:5]
    chosen.sort(key=lambda z: z[0])
    return [v for _, v in chosen]

def safe_sheet_name(name, used):
    base = re.sub(r'[\[\]:*?/\\\\]', '_', str(name)).strip()[:31] or "Analyte"
    candidate, n = base, 1
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate = base[:31-len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate

def make_analyte_sheet(wb, template_ws, analyte, used):
    ws = wb.copy_worksheet(template_ws)
    ws.title = safe_sheet_name(analyte, used)
    ws["C7"] = analyte
    # Keep formulas/layout but clear data cells from the copied template.
    for r in list(range(15, 20)) + list(range(32, 37)):
        ws.cell(r, 2).value = None
        for c in range(3, 8):
            ws.cell(r, c).value = None
    return ws

def update_workbook(template_bytes, records, previous_bytes=None):
    source = previous_bytes if previous_bytes else template_bytes
    wb = load_workbook(io.BytesIO(source))
    # Use the first worksheet as the layout template. New analytes are cloned from it.
    template_ws = wb[wb.sheetnames[0]]
    used = {s.title.lower() for s in wb.worksheets}
    analyte_to_ws = {}
    for ws in wb.worksheets:
        label = ws["C7"].value or ws["B7"].value or ws.title
        analyte_to_ws[str(label).strip().lower()] = ws
    # Keep only recognized data; append raw audit tab.
    for rec in records:
        key = rec["Analyte"].strip().lower()
        if key not in analyte_to_ws:
            ws = make_analyte_sheet(wb, template_ws, rec["Analyte"], used)
            analyte_to_ws[key] = ws
    for analyte, ws in analyte_to_ws.items():
        display_name = str(ws["C7"].value or ws["B7"].value or ws.title).strip()
        subset = [r for r in records if r["Analyte"].strip().lower() == analyte]
        # Populate each level/day into its corresponding row. Day 1..5 => rows 15..19 / 32..36.
        for level, start_row in [("Level 1", 15), ("Level 2", 32)]:
            for day in range(1, 6):
                group = [r for r in subset if r["Level"] == level and int(r["Day"]) == day]
                if not group:
                    continue
                vals = best_five([r["Result"] for r in group])
                row = start_row + day - 1
                ws.cell(row, 2).value = f"Day {day}"
                for j in range(5):
                    ws.cell(row, 3+j).value = vals[j] if j < len(vals) else None
                # Reapply formulas if a template has been modified or blank.
                ws.cell(row, 8).value = f"=IFERROR(AVERAGE(C{row}:G{row}),0)"
                ws.cell(row, 9).value = f"=IFERROR(STDEV(C{row}:G{row}),0)"
                ws.cell(row,10).value = f"=IFERROR((I{row}/H{row})*100,0)"
            # Overall summary formulas cover available day rows.
            end_row = start_row + 4
            mean_row, sd_row, cv_row = (22,23,24) if level == "Level 1" else (39,40,41)
            ws.cell(mean_row,10).value = f"=IFERROR(AVERAGE(C{start_row}:G{end_row}),0)"
            ws.cell(sd_row,10).value = f"=IFERROR(STDEV(C{start_row}:G{end_row}),0)"
            ws.cell(cv_row,10).value = f"=IFERROR((J{sd_row}/J{mean_row})*100,0)"
        units = [r["Unit"] for r in subset if r.get("Unit")]
        if units:
            ws["C11"] = units[0]
            ws["C28"] = units[0]
    # Raw audit: preserve the extracted rows used for selection.
    audit_name = "Raw_Data_Audit"
    if audit_name in wb.sheetnames:
        del wb[audit_name]
    audit = wb.create_sheet(audit_name)
    headers = ["Analyte","Day","Level","Result","Unit","SID / QC","QC Material","QC ID","Source PDF","Page line","Selected for workbook"]
    audit.append(headers)
    selected_keys = set()
    for r in records:
        grp = [x for x in records if x["Analyte"] == r["Analyte"] and x["Day"] == r["Day"] and x["Level"] == r["Level"]]
        chosen = best_five([x["Result"] for x in grp])
        # Track selected values by occurrence index to support duplicates.
        occurrences = {}
        for x in grp:
            k = (x["Result"], x["Source PDF"], x["Page line"])
            occurrences[k] = occurrences.get(k, 0) + 1
        # This flag is an approximate value-membership indicator when duplicate values occur.
        selected = r["Result"] in chosen
        audit.append([r["Analyte"],r["Day"],r["Level"],r["Result"],r["Unit"],r["SID / QC"],r["QC Material"],r["QC ID"],r["Source PDF"],r["Page line"],"Yes" if selected else "No"])
    audit.freeze_panes = "A2"
    audit.auto_filter.ref = audit.dimensions
    for cell in audit[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col in audit.columns:
        letter = get_column_letter(col[0].column)
        audit.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col)+2, 12), 32)
    # Selection log
    if "Upload_Summary" in wb.sheetnames:
        del wb["Upload_Summary"]
    summary = wb.create_sheet("Upload_Summary")
    summary.append(["Analyte","Day","Level","Raw results found","Replicates populated","Selection method"])
    combos = sorted({(r["Analyte"],r["Day"],r["Level"]) for r in records})
    for a,d,l in combos:
        g=[r for r in records if (r["Analyte"],r["Day"],r["Level"])==(a,d,l)]
        summary.append([a,d,l,len(g),min(5,len(g)),"5 closest to median" if len(g)>5 else "All available (≤5)"])
    for c in summary[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78")
    summary.freeze_panes="A2"; summary.auto_filter.ref=summary.dimensions
    out=io.BytesIO(); wb.save(out); return out.getvalue()

# ------------------------- UI -------------------------
st.info("Workflow: first upload your original template. On later sessions, upload the latest populated workbook so earlier days are preserved.")
c1,c2=st.columns(2)
with c1:
    template_file=st.file_uploader("1. Excel template (first run)", type=["xlsx"], key="template")
with c2:
    previous_file=st.file_uploader("Or latest populated workbook (to append new days)", type=["xlsx"], key="previous")
pdf_files=st.file_uploader("2. Upload one or more instrument PDF reports", type=["pdf"], accept_multiple_files=True)

if pdf_files:
    st.subheader("Assign report day")
    day_map={}
    cols=st.columns(min(3,len(pdf_files)))
    for idx,pdf in enumerate(pdf_files):
        with cols[idx % len(cols)]:
            day_map[pdf.name]=st.selectbox(f"{pdf.name}: default day", [1,2,3,4,5], index=min(idx,4), key=f"day_{idx}")
    st.caption("If a SID contains a day label (e.g. INTER DAY2 or DAY301), the embedded day is used. The selector is the fallback for records without a day label.")

st.subheader("Selection and QC-level rules")
st.write("The app selects the five results closest to the median for each analyte + day + level. All extracted results remain in the Raw_Data_Audit sheet.")
with st.expander("Default Level 1 / Level 2 identifier rules"):
    st.code("Level 1: 101, 301, AB1, K5972, 6304961\\nLevel 2: 102, 302, AB2, K5973, 6304962")
    st.warning("These are starter rules based on the sample reports. For a different analyzer/QC naming convention, inspect the Raw_Data_Audit sheet and adjust infer_level() in app.py before relying on the output.")

if st.button("🚀 Process PDFs and generate updated Excel", type="primary", disabled=not pdf_files or not (template_file or previous_file)):
    try:
        all_records=[]
        for pdf in pdf_files:
            day=day_map[pdf.name]
            recs=parse_pdf(pdf.getvalue(),pdf.name,day)
            all_records.extend(recs)
        if not all_records:
            st.error("No numeric analyte results were extracted. This may be a scanned PDF; OCR support is not included in this starter version.")
        else:
            # Preserve prior raw data when updating an existing workbook.
            prior_records=[]
            if previous_file:
                # Existing audit sheet is re-read so earlier data remain available.
                old=load_workbook(io.BytesIO(previous_file.getvalue()),data_only=True)
                if "Raw_Data_Audit" in old.sheetnames:
                    df=pd.DataFrame(list(old["Raw_Data_Audit"].values))
                    if len(df)>1:
                        df.columns=df.iloc[0]
                        for _,row in df.iloc[1:].iterrows():
                            prior_records.append({
                                "Analyte":str(row.get("Analyte","")), "Day":int(row.get("Day",1)),
                                "Level":str(row.get("Level","Unassigned")), "Result":float(row.get("Result")),
                                "Unit":str(row.get("Unit","")), "SID / QC":str(row.get("SID / QC","")),
                                "QC Material":str(row.get("QC Material","")), "QC ID":str(row.get("QC ID","")),
                                "Source PDF":str(row.get("Source PDF","")), "Page line":row.get("Page line","")
                            })
            # De-duplicate identical records across repeated uploads.
            combined=prior_records+all_records
            seen=set(); unique=[]
            for r in combined:
                key=(r["Analyte"],r["Day"],r["Level"],r["Result"],r["SID / QC"],r["Source PDF"],str(r["Page line"]))
                if key not in seen:
                    seen.add(key); unique.append(r)
            base_bytes=template_file.getvalue() if template_file else previous_file.getvalue()
            prev_bytes=previous_file.getvalue() if previous_file else None
            result=update_workbook(base_bytes,unique,prev_bytes)
            st.success(f"Extracted {len(all_records)} new result rows across {len(set(r['Analyte'] for r in all_records))} analytes.")
            st.download_button("⬇️ Download populated precision workbook", data=result,
                file_name="Universal_5_Day_InterRun_Precision.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            preview=pd.DataFrame(all_records)
            st.dataframe(preview, use_container_width=True)
            st.warning("Review the Raw_Data_Audit and Upload_Summary sheets before using results for a regulated validation report. Confirm the QC-level mapping and day assignment.")
    except Exception as e:
        st.exception(e)

st.divider()
st.caption("Important: this is an assisted extraction tool, not a validated laboratory information system. Verify every analyte, QC level, day and selected replicate against the original instrument report.")
