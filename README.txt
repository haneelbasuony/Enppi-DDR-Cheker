**Prompt [compare_ddr_aconex.py :... | Try in Copilot Chat](https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/eyJpZCI6IlZYTmxjbFl4ZkdoMGRIQnpPaTh2YzNWaWMzUnlZWFJsTFdsdWRDNXZabVpwWTJVdVkyOXRMM3hQU1VRNk5qYzJNVEUzWVdFdFpEUXlZUzAwTXpGaUxXRXhZakF0WmpJMllqVTRaV00xWkdRM2ZEaGlOR0ZsWkRRM0xUSTRZakF0TkROaU15MDRNV0V6TFRnM016aGhaRFk1WkRVME9Id3lNREkyTFRBNExURTNWREExT2pFek9qQTBMakkzTnpFeU5UZGEiLCJzY2VuYXJpbyI6InNoYXJlTGlua1ZpYVJpY2hDb3B5IiwicHJvcGVydGllcyI6eyJwcm9tcHRTb3VyY2UiOiJ1c2VyIiwiY2xpY2tUaW1lc3RhbXAiOiIyMDI2LTA4LTE3VDA1OjEzOjA0LjMxNloifSwiY2hhdFR5cGUiOiJ3ZWIiLCJ2ZXJzaW9uIjoxLjF9)**
compare_ddr_aconex.py : 
"""
compare_ddr_aconex.py
<br><br>
Compares a DDR (Document Distribution Register) extract against:
  1) the PLIP SPO file (to confirm the PLIP ID exists)
  2) the Aconex document register export (to confirm the document exists,
     and that titles match)
<br><br>
and, for documents that need a NEW placeholder in Aconex, validates the
document number against the ENPPI/BIR numbering taxonomy.
<br><br>
OUTPUT: a single Excel file with one row per DDR document, showing
Document Number / Document Title / PLIP ID / Status / Notes, with the
Status column colour-coded so it's easy to scan.
<br><br>
--------------------------------------------------------------------------
HOW TO USE
--------------------------------------------------------------------------
1. Edit the CONFIG block below so the column names match your actual
   files EXACTLY as they appear in row 1 of each spreadsheet.
2. Run:
       python compare_ddr_aconex.py --ddr "DDR.xlsx" --plip "PLIP.xlsx" --aconex "Aconex.xlsx" --out "DDR_Comparison_Result.xlsx"
   (You can also just edit the DEFAULT_* paths below and run with no
   arguments if you'd rather not type the flags every time.)
<br><br>
--------------------------------------------------------------------------
DECISION LOGIC (matches the workflow you described)
--------------------------------------------------------------------------
For every document row in the DDR:
<br><br>
  Step 1 - PLIP CHECK
    - If the PLIP ID is NOT found in the PLIP SPO file:
        Status = "PLIP ID Not Found"  -> STOP (no further checks for this row)
    - If found -> go to Step 2
<br><br>
  Step 2 - EXISTENCE CHECK (in Aconex register, by Document Number)
    - If the Document Number does NOT exist in Aconex:
        -> go to Step 3 (taxonomy check), because a placeholder must be created
    - If it DOES exist -> go to Step 4 (title check)
<br><br>
  Step 3 - TAXONOMY CHECK (only for documents needing a new placeholder)
    - Validate the Document Number against the 7-part BIR-36-... numbering
      scheme.
    - If invalid:
        Status = "Placeholder Required - INVALID Doc Number Taxonomy"
    - If valid:
        Status = "Placeholder Required - Create in Aconex"
<br><br>
  Step 4 - TITLE CHECK (only for documents that already exist in Aconex)
    - If DDR title == Aconex title (case/whitespace-insensitive):
        Status = "OK - No Corrective Action Required"
    - Else:
        Status = "Title Mismatch - Update Title in Aconex"
"""
<br><br>
import argparse
import re
import sys
from pathlib import Path
import pyodbc
import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Border, Side
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
<br><br>
# ==========================================================================
# CONFIG - EDIT THESE TO MATCH YOUR FILES
# ==========================================================================
<br><br>
# --- Default file paths (used only if you don't pass --ddr/--plip/--aconex) ---
DEFAULT_DDR_PATH = "DDR.xlsx"
DEFAULT_PLIP_PATH = "PLIP.xlsx"
<br><br>
DEFAULT_OUTPUT_PATH = "DDR_Comparison_Result.xlsx"
<br><br>
# --- Sheet names (set to None to use the first/active sheet) ---
DDR_SHEET_NAME = "Deliverables"
PLIP_SHEET_NAME = None
<br><br><br>
# --- Header row (0-indexed). Set to None to auto-detect the row that
# contains your column titles -- useful because DDR/Aconex exports often
# have a title block, revision info, or logo rows above the real table.
# If auto-detect picks the wrong row, just hardcode the correct row number
# here (e.g. 3 means the headers are on the 4th row of the sheet).
DDR_HEADER_ROW = None
PLIP_HEADER_ROW = None
<br><br><br>
# How many rows from the top to scan when auto-detecting the header row
HEADER_SCAN_ROWS = 25
<br><br>
# --- Column names as they appear in row 1 of each file ---
# DDR file
DDR_COL_DOC_NUMBER = "DD Document/Drawing No"
DDR_COL_TITLE = "Title"
DDR_COL_PLIP_ID = "PLIP ID"
<br><br>
# PLIP SPO file (only the PLIP ID column is needed for the lookup)
PLIP_COL_PLIP_ID = "PLIP_ID"
<br><br>
# ==========================================================================
# SQL SERVER CONFIGURATION
# ==========================================================================
<br><br>
SQL_SERVER = "ES-MSSQL-01"
SQL_DATABASE = "ACONEX Reporting Data"
SQL_TABLE = "DocumentRegisterTest"
<br><br>
ACONEX_COL_DOC_NUMBER = "DocNo"
ACONEX_COL_TITLE = "Title"
ACONEX_COL_REVIEW_STATUS = "ReviewStatus"
ACONEX_COL_FILE_NAME = "FileName"
<br><br>
# ==========================================================================
# TAXONOMY RULES (Document Number = 7 dash-separated fields)
# Area - ProcessUnit - Originator - Discipline - DocType - Sequential - Sheet
# ==========================================================================
<br><br>
AREA_CODE_ALLOWED = {"BIR"}
PROCESS_UNIT_ALLOWED = {"36"}
<br><br>
ORIGINATOR_CODES = {
    "000000", "102396", "102961", "103440", "104262", "104332", "107625",
}
<br><br>
DISCIPLINE_CODES = {
    "AA", "BA", "CS", "EA", "FD", "HX", "IC", "IN", "JA", "KA", "LA", "MH",
    "MP", "MR", "MS", "NA", "OA", "PX", "QA", "RA", "SA", "TA", "VA", "ZP",
    "ZR", "ZT", "ZV", "ZW",
}
<br><br>
DOCUMENT_TYPE_CODES = set("""
0480 0706 0709 0712 0780 1439 1780 1905 2808 3004 3101 3180 3309 3311 3341 3347
3352 3365 3807 4301 4303 4322 4372 4880 5005 5101 5206 5402 5507 5706 5722
5753 5759 5760 5765 5766 5768 5775 5778 5780 5787 5792 5793 5798 5800 5802
5803 5804 5805 5880 5980 6006 6015 6044 6048 6180 6401 6402 6480 6603 6605
6612 6614 6619 6627 6680 6833 6834 6844 6857 6926 6944 6945 6999 7004 7179
7180 7205 7403 7415 7416 7417 7480 7507 7704 7739 7753 7772 7880 8203 8212
8380 8502 0603 1418 1422 1424 1426 2322 2341 2404 3314 3375 3581 4202 4380
4802 5011 5012 5503 5537 5711 5736 5789 5799 5801 6002 6008 6026 6033 6036
6063 6064 6065 6066 6067 6070 7756 8236 8382 0604 0711 1206 2307 2315 2331
2358 2386 2397 2401 4017 4018 4024 4038 4980 5680 5721 5729 6968 7303 7721
7770 8102 8225 8235 8604 0901 1380 1919 2105 2305 2335 2369 2373 2384 2398
2580 2901 2913 3323 3880 4005 4006 4012 4025 4180 4306 4308 4319 4324 4327
4329 4359 4606 5106 5772 6039 6858 7771 8251 8260 8804 8809 8880 3103 5180
5726 5732 5814 6406 7280 0702 0704 1601 2318 2356 2374 4016 4029 5511 5512
6613 6802 6807 6854 6873 6875 6881 7506 8223 8249 5306 980 1214 1409 1580
1926 2180 2310 2347 2351 2376 2378 4013 4019 4302 4328 4780 5505 5520 5531
5532 6056 6820 7680 7709 7711 8580 8606 0880 1702 1703 1979 3001 3357 5280
5718 5762 6017 6610 6611 6626 7712 3304 4803 0580 1903 2343 3356 4337 4363
5080 6958 2409 4033 4311 4809 6280 7402 7737 8005 0503 6061 6843 8226 1404
4354 4810 5796 6043 3201 3329 4304 4811 5517 5522 5533 5536 5701 5705 5794
5795 5876 5877 6019 6025 6032 6055 6059 6943 7582 8980 2364 2365 2366 2368
3363 4338 4368 4603 6967 8248 1419 4815 5733 6023 6049 6050 6876 1401 5009
5506 5779 6068 1448 6028 0902 0904 1203 1901 2314 2323 4680 6057 6170 7731
8420 8870 0403 1453 1701 3330 3359 3369 4305 4336 4349 4357 4373 5756 6042
6047 6623 6856 7003 7206 6409 A00 A01 A02 B00 B01 B02 B03 B04 B05 B06 B07
B08 B09 B10 B11 B12 B13 B14 B15 C01 C02 C03 C04 C05 C06 C07 C08 C09 C10
C11 C12 C13 C14 D01 D02 D03 D04 D05 D06 D07 D08 D09 D10 D11 D12 D13 D14
D15 D16 D17 D18 D19 D20 D21 D22 D23 D24 D25 D26 D27 E00 E01 E02 E03 E04
E05 E06 E07 E08 E09 E10 E11 E12 E13 F00 F01 F02 F03 F04 F05 F06 F07 F08
F09 F10 F11 F12 F13 F14 F15 F16 F17 F18 F19 F20 F21 F22 F23 G00 G01 G02
G03 G04 G05 G06 G07 G08 G09 G10 G11 G12 G13 G14 G15 G16 G17 G18 G19 G20
G21 G22 G23 G24 G25 G26 G27 G28 H00 H01 H02 H03 H04 H05 H06 H07 H08 H09
H10 H11 H12 H13 H14 H15 H16 H17 H18 H19 H20 H21 H22 H23 J00 J01 J02 J03
J04 J05 J06 J07 J08 J09 K00 K01 K02 K03 K04 K05 K06 K07 K08 K09 K10 K11
K12 K13 K14 K15 K16 K17 K18 K19 K20 K21 K22 K23 K24 K25 K26 K27 K28 K29
K30 K31 K32 K33 K34 K35 K36 L00 L01 L02 L03 L04 L05 L06 L07 L08 M00 M01
M02 M03 M04 M05 M06 M07 M08 M09 N00 N01 N02 N03 N04 N05 N06 N07 N08 N09
P00 P01 P02 P03 P04 P05 P06 P07 P08 P09 P10 P11 P12 P13 P14 P15 P16 P17
P18 P19 P20 P21 P22 P23 2394
""".split())
<br><br>
SEQ_NO_RE = re.compile(r"^\d{5}$")   # Sequential No. -> 5 numeric digits
SHEET_NO_RE = re.compile(r"^\d{4}$")  # Sheet No.      -> 4 numeric digits
<br><br><br><br>
def validate_taxonomy(doc_number: str):
    """
    Validates a document number against the 7-part BIR taxonomy.
    Returns (is_valid: bool, reason: str)
    reason is "" when valid, otherwise a short explanation of what failed.
    """
    if not doc_number or not isinstance(doc_number, str):
        return False, "Document number is empty/blank"
<br><br>
    parts = doc_number.strip().split("-")
<br><br>
    if len(parts) != 7:
        return False, f"Expected 7 dash-separated fields, found {len(parts)}"
<br><br>
    area, proc_unit, originator, discipline, doc_type, seq_no, sheet_no = parts
<br><br>
    if area not in AREA_CODE_ALLOWED:
        return False, f"Area code '{area}' should be BIR"
<br><br>
    if proc_unit not in PROCESS_UNIT_ALLOWED:
        return False, f"Process Unit code '{proc_unit}' should be 36"
<br><br>
    if originator not in ORIGINATOR_CODES:
        return False, f"Originator code '{originator}' not in allowed list"
<br><br>
    if discipline not in DISCIPLINE_CODES:
        return False, f"Discipline code '{discipline}' not in allowed list"
<br><br>
    if doc_type not in DOCUMENT_TYPE_CODES:
        return False, f"Document type '{doc_type}' not in allowed list"
<br><br>
    if not SEQ_NO_RE.match(seq_no):
        return False, f"Sequential No. '{seq_no}' must be exactly 5 digits"
<br><br>
    if not SHEET_NO_RE.match(sheet_no):
        return False, f"Sheet No. '{sheet_no}' must be exactly 4 digits"
<br><br>
    return True, ""
<br><br><br>
# ==========================================================================
# HELPERS
# ==========================================================================
<br><br>
def matches_doc_number_plip_pattern(doc_number, plip_id):
    """
    Checks whether the PLIP ID format corresponds to the
    Discipline + Document Type contained in the document number.
<br><br>
    Example:
<br><br>
    Document Number:
    BIR-36-103440-PX-2366-00001-0001
<br><br>
    Expected PLIP pattern:
    PX2366-XX
    """
<br><br>
    try:
<br><br>
        if not doc_number or not plip_id:
            return False
<br><br>
        parts = str(doc_number).strip().split("-")
<br><br>
        if len(parts) != 7:
            return False
<br><br>
        discipline = parts[3].strip().upper()
        doc_type = parts[4].strip().upper()
<br><br>
        expected_pattern = (
            f"^{re.escape(discipline)}"
            f"{re.escape(doc_type)}"
            r"-\d{2}$"
        )
<br><br>
        return bool(
            re.match(
                expected_pattern,
                str(plip_id).strip().upper()
            )
        )
<br><br>
    except Exception:
        return False
<br><br>
def norm(value):
    """Normalize a value for comparison: strip whitespace, uppercase, and
    turn NaN/None into an empty string."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip().upper()
<br><br><br>
def _cell_match(cell_value, required_cols_norm):
    """True if a single cell's text matches one of the required column names
    (case/whitespace-insensitive, and tolerant of partial/contains matches)."""
    if cell_value is None:
        return False
    text = str(cell_value).strip().upper()
    if not text or text.startswith("UNNAMED"):
        return False
    return any(text == req or req in text or text in req for req in required_cols_norm)
<br><br><br>
def find_header_row(path, sheet_name, required_cols, scan_rows=25):
    """Scans the first `scan_rows` rows of the sheet (no header assumed) and
    returns the 0-indexed row number that best matches the required column
    names, or None if nothing scores well."""
    raw = pd.read_excel(
        path, sheet_name=sheet_name, header=None, nrows=scan_rows
    ) if sheet_name else pd.read_excel(path, header=None, nrows=scan_rows)
<br><br>
    required_norm = [c.strip().upper() for c in required_cols]
<br><br>
    best_row, best_score = None, 0
    for row_idx in range(len(raw)):
        row_values = raw.iloc[row_idx].tolist()
        score = sum(1 for v in row_values if _cell_match(v, required_norm))
        if score > best_score:
            best_score, best_row = score, row_idx
<br><br>
    # Require at least half the required columns to be found in that row
    # before trusting it as the header row.
    if best_row is not None and best_score >= max(1, len(required_cols) // 2):
        return best_row, raw
    return None, raw
<br><br><br>
def load_sheet(path, sheet_name, required_cols, label, header_row=None, scan_rows=HEADER_SCAN_ROWS):
    try:
        if header_row is None:
            detected_row, raw_preview = find_header_row(path, sheet_name, required_cols, scan_rows)
            if detected_row is None:
                preview = raw_preview.head(scan_rows).to_string()
                raise Exception(
                    f"ERROR: could not auto-detect the header row in the {label} file " 
                    f"within the first {scan_rows} rows.\n"
                    f"Expected columns (or close matches): {required_cols}\n\n"
                    f"Preview of '{path}':\n{preview}\n\n"
                    f"-> Set the *_HEADER_ROW value for this file in the CONFIG block "
                    f"to the correct 0-indexed row number, or fix the required column "
                    f"names to match your file."
                )
            header_row = detected_row
            print(f"[{label}] Auto-detected header row: {header_row} (0-indexed)")
<br><br>
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row) if sheet_name \
            else pd.read_excel(path, header=header_row)
    except FileNotFoundError:
        raise Exception(f"ERROR: {label} file not found: {path}")
    except SystemExit:
        raise
    except Exception as e:
        raise Exception(f"ERROR: could not read {label} file '{path}': {e}")
<br><br>
    # Drop fully-blank / "Unnamed" columns that sometimes trail real data
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed.*$") | df.columns.isin(required_cols)]
<br><br>
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise Exception(
            f"ERROR: {label} file is missing expected column(s): {missing}\n"
            f"Columns found in '{path}' (header row {header_row}): {list(df.columns)}\n"
            f"-> Fix the CONFIG block at the top of this script: either correct the "
            f"column name constants, or set the *_HEADER_ROW value explicitly."
        )
    return df
<br><br>
def load_aconex_from_sql():
    """
    Loads the Aconex Register directly from SQL Server.
    """
<br><br>
    conn_string = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        "Trusted_Connection=yes;"
        "TrustServerCertificate=yes;"
    )
<br><br>
    query = f"""
    SELECT
        [{ACONEX_COL_DOC_NUMBER}],
        [{ACONEX_COL_TITLE}],
        [{ACONEX_COL_REVIEW_STATUS}],
        [{ACONEX_COL_FILE_NAME}]
    FROM [{SQL_TABLE}]
    WHERE [ProjectId] = '1342183550'
    """
<br><br>
    try:
        conn = pyodbc.connect(conn_string)
<br><br>
        df = pd.read_sql(query, conn)
<br><br>
        conn.close()
<br><br>
    except Exception as e:
        raise Exception(
            f"ERROR: Failed to load Aconex Register from SQL Server.\n{e}"
        )
<br><br>
    required_cols = [
        ACONEX_COL_DOC_NUMBER,
        ACONEX_COL_TITLE,
        ACONEX_COL_REVIEW_STATUS,
        ACONEX_COL_FILE_NAME
    ]
<br><br>
    missing = [
        c for c in required_cols
        if c not in df.columns
    ]
<br><br>
    if missing:
        raise Exception(
            f"ERROR: SQL table '{SQL_TABLE}' is missing columns: {missing}"
        )
<br><br>
    print(
        f"[Aconex SQL] Loaded {len(df):,} records from SQL Server"
    )
<br><br>
    return df
<br><br>
# ==========================================================================
# MAIN
# ==========================================================================
<br><br>
STATUS_STYLES = {
    "OK - No Corrective Action Required": "C6EFCE",
    "Title Mismatch - Update Title in Aconex": "FFEB9C",
    "Placeholder Required - Create in Aconex": "FFD8A8",
    "Placeholder Required - INVALID Doc Number Taxonomy": "FFC7CE",
    "PLIP ID Not Found": "D9D9D9",
    "InActive PLIP ID": "FFF2CC",
    "Delete from Aconex Document Register": "F4CCCC",
}
<br><br><br>
def main(ddr_path=None, plip_path=None, out_path=None):
<br><br>
    if ddr_path is None:
<br><br>
        parser = argparse.ArgumentParser(
            description="Compare DDR vs PLIP vs Aconex register"
        )
<br><br>
        parser.add_argument(
            "--ddr",
            default=DEFAULT_DDR_PATH
        )
<br><br>
        parser.add_argument(
            "--plip",
            default=DEFAULT_PLIP_PATH
        )
<br><br>
        parser.add_argument(
            "--out",
            default=DEFAULT_OUTPUT_PATH
        )
<br><br>
        args = parser.parse_args()
<br><br>
        ddr_path = args.ddr
        plip_path = args.plip
        out_path = args.out
<br><br>
    ddr_df = load_sheet(
        ddr_path, DDR_SHEET_NAME,
        [DDR_COL_DOC_NUMBER, DDR_COL_TITLE, DDR_COL_PLIP_ID], "DDR",
        header_row=DDR_HEADER_ROW,
    )
    plip_df = load_sheet(
        plip_path, PLIP_SHEET_NAME,
        [PLIP_COL_PLIP_ID], "PLIP SPO",
        header_row=PLIP_HEADER_ROW,
    )
    aconex_df = load_aconex_from_sql()
<br><br>
    # Build fast lookup sets/dicts
    plip_ids = {norm(v) for v in plip_df[PLIP_COL_PLIP_ID]}
<br><br>
    aconex_df["_key"] = aconex_df[ACONEX_COL_DOC_NUMBER].map(norm)
    aconex_lookup = {}
<br><br>
    for _, row in aconex_df.iterrows():
<br><br>
        aconex_lookup[row["_key"]] = {
            "title": row[ACONEX_COL_TITLE],
            "review_status": row[ACONEX_COL_REVIEW_STATUS],
            "file_name": row[ACONEX_COL_FILE_NAME]
        }
    
    # Build a set of all document numbers currently present in the DDR
    ddr_doc_numbers = {
        norm(v)
        for v in ddr_df[DDR_COL_DOC_NUMBER]
        if norm(v)
    }
<br><br><br>
    results = []
<br><br>
    for _, row in ddr_df.iterrows():
        doc_number_raw = row[DDR_COL_DOC_NUMBER]
        title_raw = row[DDR_COL_TITLE]
        plip_id_raw = row[DDR_COL_PLIP_ID]
<br><br>
        doc_number = str(doc_number_raw).strip() if pd.notna(doc_number_raw) else ""
        title = str(title_raw).strip() if pd.notna(title_raw) else ""
        plip_id = str(plip_id_raw).strip() if pd.notna(plip_id_raw) else ""
<br><br>
        status = ""
        note = ""
        review_status = ""
        document_type = ""
<br><br>
        # Step 1 - PLIP check
<br><br>
        if norm(plip_id) not in plip_ids:
<br><br>
            if matches_doc_number_plip_pattern(
                doc_number,
                plip_id
            ):
<br><br>
                status = "InActive PLIP ID"
                note = (
                    f"PLIP ID '{plip_id}' is not present in the "
                    f"PLIP register but matches the document "
                    f"number taxonomy."
                )
<br><br>
            else:
<br><br>
                status = "PLIP ID Not Found"
                note = (
                    f"PLIP ID '{plip_id}' not found in PLIP SPO file "
                    f"and does not match the document number taxonomy."
                )
<br><br>
        else:
            key = norm(doc_number)
            if key not in aconex_lookup:
                # Step 3 - needs a placeholder -> validate taxonomy
                is_valid, reason = validate_taxonomy(doc_number)
                if is_valid:
                    status = "Placeholder Required - Create in Aconex"
                    note = "Not found in Aconex register; document number passes taxonomy check"
                else:
                    status = "Placeholder Required - INVALID Doc Number Taxonomy"
                    note = f"Not found in Aconex register; taxonomy issue: {reason}"
            else:
                # Step 4 - exists -> compare titles
                aconex_record = aconex_lookup[key]
<br><br>
                aconex_title = aconex_record["title"]
                review_status = aconex_record["review_status"]
                file_name = aconex_record["file_name"]
<br><br>
                document_type = "Placeholder"
<br><br>
                if file_name is not None:
<br><br>
                    file_name_str = str(file_name).strip()
<br><br>
                    if (
                        file_name_str
                        and file_name_str.upper() != "NULL"
                        and not pd.isna(file_name)
                    ):
                        document_type = "PDF"
<br><br>
                if norm(title) == norm(aconex_title):
<br><br>
                    status = "OK - No Corrective Action Required"
                    note = "Exists in Aconex; title matches"
<br><br>
                else:
<br><br>
                    status = "Title Mismatch - Update Title in Aconex"
<br><br>
                    note = (
                        f"DDR title: '{title}' | "
                        f"Aconex title: '{aconex_title}'"
                    )
<br><br>
        results.append({
            "Document Number": doc_number,
            "Document Title (DDR)": title,
            "PLIP ID": plip_id,
            "Status": status,
            "Review Status": review_status,
            "Document Type": document_type,
            "Notes": note,
        })
<br><br>
    # ======================================================================
    # STEP 5 - FIND DOCUMENTS THAT EXIST IN ACONEX BUT NOT IN DDR
    # ======================================================================
    #
    # These documents are currently present in Aconex but are no longer
    # present in the latest DDR.
    #
    # Therefore they should be reviewed for deletion from Aconex.
    #
    # IMPORTANT:
    # This does NOT delete anything from Aconex.
    # It only reports the document as requiring deletion.
    # ======================================================================
<br><br>
    for aconex_doc_number, aconex_record in aconex_lookup.items():
<br><br>
        if not aconex_doc_number:
            continue
<br><br>
        # Document exists in Aconex but NOT in current DDR
        if aconex_doc_number not in ddr_doc_numbers:
<br><br>
            aconex_title = aconex_record["title"]
            review_status = aconex_record["review_status"]
            file_name = aconex_record["file_name"]
<br><br>
            review_status = (
                ""
                if pd.isna(review_status)
                else str(review_status)
            )
<br><br>
            document_type = "Placeholder"
<br><br>
            if file_name is not None:
                try:
                    if pd.notna(file_name) and str(file_name).strip():
                        document_type = "PDF"
                except Exception:
                    pass
<br><br>
            results.append({
                "Document Number": aconex_doc_number,
                "Document Title (DDR)": "",
                "PLIP ID": "",
                "Status": "Delete from Aconex Document Register",
                "Review Status": review_status,
                "Document Type": document_type,
                "Notes": (
                    f"Document exists in Aconex but is not present in the latest DDR. "
                    f"Aconex title: '{aconex_title}'"
                ),
            })
<br><br><br>
    out_df = pd.DataFrame(results)
<br><br>
    out_path = Path(out_path)
    out_df.to_excel(out_path, index=False, sheet_name="DDR Comparison")
    create_summary_sheet(out_path)
    style_output(out_path)
<br><br>
    print(f"Done. {len(out_df)} documents processed.")
    print(f"Output written to: {out_path.resolve()}")
    all_statuses = [
        "OK - No Corrective Action Required",
        "Title Mismatch - Update Title in Aconex",
        "Placeholder Required - Create in Aconex",
        "Placeholder Required - INVALID Doc Number Taxonomy",
        "PLIP ID Not Found",
        "InActive PLIP ID",
        "Delete from Aconex Document Register",
    ]
<br><br>
    status_counts = out_df["Status"].value_counts()
<br><br>
    print("\nStatus Summary")
    print("-" * 60)
<br><br>
    for status in all_statuses:
        print(f"{status:<55} {status_counts.get(status, 0)}")
<br><br><br>
def create_summary_sheet(path: Path):
<br><br>
    wb = load_workbook(path)
<br><br>
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
<br><br>
    ws = wb.create_sheet("Summary", 0)
<br><br>
    # ==========================================================
    # DASHBOARD TITLE
    # ==========================================================
    ws.merge_cells("A1:B1")
<br><br>
    ws["A1"] = "DDR Comparison Dashboard"
<br><br>
    ws["A1"].font = Font(
        bold=True,
        size=16,
        color="1F1F1F"
    )
<br><br>
    ws["A1"].alignment = Alignment(horizontal="center")
<br><br>
    # ==========================================================
    # READ DATA
    # ==========================================================
    df = pd.read_excel(path, sheet_name="DDR Comparison")
<br><br>
    status_counts = (
        df["Status"]
        .value_counts()
        .to_dict()
    )
<br><br>
    sorted_statuses = sorted(
        status_counts.items(),
        key=lambda x: x[1],
        reverse=True
    )
<br><br>
    # ==========================================================
    # TABLE HEADER
    # ==========================================================
    header_row = 3
<br><br>
    ws.cell(header_row, 1, "Status")
<br><br>
    header_fill = PatternFill(
        start_color="5B9BD5",
        end_color="5B9BD5",
        fill_type="solid"
    )
<br><br>
    header_font = Font(
        bold=True,
        color="FFFFFF",
        size=11
    )
<br><br>
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
<br><br>
    # ==========================================================
    # TABLE DATA
    # ==========================================================
    start_row = 4
<br><br>
    band_fill = PatternFill(
        start_color="F7F7F7",
        end_color="F7F7F7",
        fill_type="solid"
    )
<br><br>
    SUMMARY_STATUS_COLORS = {
    "OK - No Corrective Action Required": "C6EFCE",
    "Title Mismatch - Update Title in Aconex": "FFEB9C",
    "Placeholder Required - Create in Aconex": "FFD8A8",
    "Placeholder Required - INVALID Doc Number Taxonomy": "FFC7CE",
    "PLIP ID Not Found": "D9D9D9",
    "InActive PLIP ID": "FFF2CC",
    "Delete from Aconex Document Register": "F4CCCC",
    }
    
    for idx, (status, count) in enumerate(
        sorted_statuses,
        start=start_row
    ):
<br><br>
        ws.cell(idx, 1, status)
        ws.cell(idx, 2, count)
<br><br>
        color = SUMMARY_STATUS_COLORS.get(status)
<br><br>
        if color:
            fill = PatternFill(
                start_color=color,
                end_color=color,
                fill_type="solid"
            )
<br><br>
            ws.cell(idx, 1).fill = fill
            ws.cell(idx, 2).fill = fill
<br><br>
    # ==========================================================
    # GRAND TOTAL
    # ==========================================================
    grand_total = sum(status_counts.values())
<br><br>
    total_row = start_row + len(sorted_statuses)
<br><br>
    ws.cell(total_row, 1, "Grand Total")
    ws.cell(total_row, 2, grand_total)
<br><br>
    ws.cell(total_row, 1).font = Font(bold=True)
    ws.cell(total_row, 2).font = Font(bold=True)
<br><br>
    ws.cell(total_row, 1).fill = header_fill
    ws.cell(total_row, 2).fill = header_fill
<br><br>
    ws.cell(total_row, 1).font = header_font
    ws.cell(total_row, 2).font = header_font
<br><br>
    # ==========================================================
    # BORDERS
    # ==========================================================
    from openpyxl.styles import Border, Side
<br><br>
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
<br><br>
    for row in ws.iter_rows(
        min_row=header_row,
        max_row=total_row,
        min_col=1,
        max_col=2
    ):
        for cell in row:
            cell.border = thin_border
<br><br>
    # ==========================================================
    # COLUMN WIDTHS
    # ==========================================================
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 25
<br><br>
    # ==========================================================
    # CHART
    # ==========================================================
    chart = BarChart()
<br><br>
    chart.type = "bar"
    chart.style = 48
    chart.varyColors = True
<br><br>
    chart.height = 20
    chart.width = 30
<br><br>
    chart.title = "DDR Comparison Status Summary"
    chart.x_axis.title = "Document Number"
    chart.y_axis.title = "Status"
<br><br>
    try:
        chart.graphical_properties.line.noFill = True
    except Exception:
        pass
<br><br>
    last_data_row = start_row + len(sorted_statuses) - 1
<br><br>
    data = Reference(
        ws,
        min_col=2,
        min_row=header_row,
        max_row=last_data_row
    )
<br><br>
    categories = Reference(
        ws,
        min_col=1,
        min_row=start_row,
        max_row=last_data_row
    )
<br><br>
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
<br><br>
    # Orange bars like your sample image
    chart.varyColors = True
<br><br>
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
<br><br>
    ws.add_chart(chart, "D3")
<br><br>
    # ==========================================================
    # INSIGHT TEXT
    # ==========================================================
    if len(sorted_statuses) >= 2:
<br><br>
        top_one = sorted_statuses[0]
        top_two = sorted_statuses[1]
<br><br>
        ws["D22"] = (
            f"Top statuses:\n"
            f"'{top_one[0]}' ({top_one[1]})\n"
            f"'{top_two[0]}' ({top_two[1]})"
        )
<br><br>
        ws["D22"].font = Font(
            bold=True,
            size=12,
            color="44546A"
        )
<br><br>
    # ==========================================================
    # OPEN ON SUMMARY TAB
    # ==========================================================
    wb.active = 0
<br><br>
    wb.save(path)
    
def style_output(path: Path):
    """Colour-codes the Status column, bolds the header, and auto-fits columns."""
    wb = load_workbook(path)
    ws = wb["DDR Comparison"]
<br><br>
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
<br><br>
    headers = [cell.value for cell in ws[1]]
    status_col_idx = headers.index("Status") + 1 if "Status" in headers else None
<br><br>
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
<br><br>
    if status_col_idx:
        for row in ws.iter_rows(min_row=2, min_col=status_col_idx, max_col=status_col_idx):
            for cell in row:
                color = STATUS_STYLES.get(cell.value)
                if color:
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
<br><br>
    # Auto-fit-ish column widths
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, 12), 60)
<br><br>
    ws.freeze_panes = "A2"
    wb.save(path)
<br><br><br>
if __name__ == "__main__":
    main()
----------
launcher.pyw :
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import pyodbc
import threading
<br><br>
from compare_ddr_aconex import main
<br><br><br>
# ============================================================
# SQL CONFIGURATION
# ============================================================
<br><br>
SQL_SERVER = "ES-MSSQL-01"
SQL_DATABASE = "ACONEX Reporting Data"
<br><br><br>
# ============================================================
# MAIN WINDOW
# ============================================================
<br><br>
root = tk.Tk()
<br><br>
# ============================================================
# DARK THEME
# ============================================================
<br><br>
BG = "#1E1E1E"
CARD = "#252526"
TEXT = "#FFFFFF"
ACCENT = "#0078D4"
SUCCESS = "#107C10"
WARNING = "#FFB900"
ENTRY_BG = "#2D2D30"
<br><br>
root.configure(bg=BG)
root.title("DDR vs ACONEX Validator")
root.geometry("950x650")
root.resizable(False, False)
<br><br>
ddr_file = tk.StringVar()
plip_file = tk.StringVar()
output_folder = tk.StringVar()
<br><br>
status_text = tk.StringVar(value="Ready")
<br><br><br>
# ============================================================
# LOGO
# ============================================================
<br><br>
try:
    img = Image.open("enppi_logo.png")
    img = img.resize((220, 80))
    logo = ImageTk.PhotoImage(img)
<br><br>
    logo_label = tk.Label(
        root,
        image=logo,
        bg=BG
    )
    logo_label.pack(pady=(10, 5))
<br><br>
except Exception:
    pass
<br><br><br>
# ============================================================
# TITLE
# ============================================================
<br><br>
title_label = tk.Label(
    root,
    text="DDR vs ACONEX VALIDATOR",
    font=("Segoe UI", 22, "bold"),
    bg=BG,
    fg=TEXT
)
<br><br>
title_label.pack(pady=(10, 20))
<br><br><br>
# ============================================================
# FUNCTIONS
# ============================================================
<br><br>
def create_button(parent, text, command, color):
<br><br>
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=color,
        fg="white",
        relief="flat",
        cursor="hand2",
        activebackground=color,
        activeforeground="white",
        font=("Segoe UI", 9, "bold")
    )
<br><br>
    return btn
<br><br>
def browse_ddr():
    path = filedialog.askopenfilename(
        title="Select DDR File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
<br><br>
    if path:
        ddr_file.set(path)
<br><br><br>
def browse_plip():
    path = filedialog.askopenfilename(
        title="Select PLIP File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
<br><br>
    if path:
        plip_file.set(path)
<br><br><br>
def browse_output():
    path = filedialog.askdirectory(
        title="Select Output Folder"
    )
<br><br>
    if path:
        output_folder.set(path)
<br><br><br>
def test_connection():
<br><br>
    try:
<br><br>
        conn_string = (
            "DRIVER={ODBC Driver 18 for SQL Server};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={SQL_DATABASE};"
            "Trusted_Connection=yes;"
            "TrustServerCertificate=yes;"
        )
<br><br>
        conn = pyodbc.connect(conn_string)
<br><br>
        cursor = conn.cursor()
        cursor.execute("SELECT 1")
<br><br>
        conn.close()
<br><br>
        messagebox.showinfo(
            "Success",
            "SQL Connection Successful."
        )
<br><br>
    except Exception as e:
<br><br>
        messagebox.showerror(
            "Connection Failed",
            str(e)
        )
<br><br><br>
def validation_worker():
<br><br>
    try:
<br><br>
        output_file = os.path.join(
            output_folder.get(),
            "DDR_Comparison_Result.xlsx"
        )
<br><br>
        root.after(
            0,
            lambda: status_text.set("Running Validation...")
        )
<br><br>
        from compare_ddr_aconex import main
<br><br>
        main(
            ddr_path=ddr_file.get(),
            plip_path=plip_file.get(),
            out_path=output_file
        )
<br><br>
        root.after(
            0,
            lambda: status_text.set("Completed Successfully")
        )
        
        root.after(0, progress.stop)
        
        root.after(
            0,
            lambda: messagebox.showinfo(
                "Success",
                f"Output generated successfully:\n\n{output_file}"
            )
        )
<br><br>
        if os.path.exists(output_file):
            os.startfile(output_file)
<br><br>
    except Exception as e:
<br><br>
        error_msg = str(e)
<br><br>
        root.after(
            0,
            lambda: status_text.set("Failed")
        )
<br><br>
        root.after(
            0,
            lambda msg=error_msg: messagebox.showerror(
                "Error",
                msg
            )
        )
<br><br>
def run_validation():
<br><br>
    if not ddr_file.get():
        messagebox.showerror(
            "Missing File",
            "Please select a DDR file."
        )
        return
<br><br>
    if not plip_file.get():
        messagebox.showerror(
            "Missing File",
            "Please select a PLIP file."
        )
        return
<br><br>
    if not output_folder.get():
        messagebox.showerror(
            "Missing Folder",
            "Please select an output folder."
        )
        return
    
    progress.start(10)
<br><br>
    threading.Thread(
        target=validation_worker,
        daemon=True
    ).start()
<br><br>
def add_hover(widget, normal, hover):
<br><br>
    widget.bind(
        "<Enter>",
        lambda e: widget.config(bg=hover)
    )
<br><br>
    widget.bind(
        "<Leave>",
        lambda e: widget.config(bg=normal)
    )
# ============================================================
# INPUT FRAME
# ============================================================
<br><br>
main_frame = tk.Frame(
    root,
    bg=CARD,
    padx=15,
    pady=15,
    highlightbackground="#3F3F46",
    highlightthickness=1
)
<br><br>
main_frame.pack(
    fill="x",
    padx=20,
    pady=10
)
<br><br>
# DDR
<br><br>
tk.Label(
    main_frame,
    text="DDR File",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=0, column=0, sticky="w", pady=10)
<br><br>
tk.Entry(
    main_frame,
    textvariable=ddr_file,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=0, column=1, padx=10)
<br><br>
create_button(
    main_frame,
    "Browse",
    browse_ddr,
    ACCENT
).grid(row=0, column=2)
<br><br>
# PLIP
<br><br>
tk.Label(
    main_frame,
    text="PLIP File",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=1, column=0, sticky="w", pady=10)
<br><br>
tk.Entry(
main_frame,
    textvariable=plip_file,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=1, column=1, padx=10)
<br><br>
create_button(
    main_frame,
    "Browse",
    browse_plip,
    ACCENT
).grid(row=1, column=2)
<br><br>
# OUTPUT
<br><br>
tk.Label(
    main_frame,
    text="Output Folder",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=2, column=0, sticky="w", pady=10)
<br><br>
tk.Entry(
    main_frame,
    textvariable=output_folder,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=2, column=1, padx=10)
<br><br>
create_button(
    main_frame,
    "Browse",
    browse_output,
    ACCENT
).grid(row=2, column=2)
<br><br><br>
# ============================================================
# SQL BUTTON
# ============================================================
<br><br>
sql_button = tk.Button(
    root,
    text="Test SQL Connection",
    command=test_connection,
    bg=ACCENT,
    fg="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 10, "bold"),
    padx=15,
    pady=8
)
<br><br>
sql_button.pack(pady=15)
<br><br>
style = ttk.Style()
style.theme_use("default")
<br><br>
style.configure(
    "blue.Horizontal.TProgressbar",
    troughcolor=BG,
    background=ACCENT
)
<br><br>
progress = ttk.Progressbar(
    root,
    style="blue.Horizontal.TProgressbar",
    mode="indeterminate",
    length=450
)
<br><br>
progress.pack(pady=10)
# ============================================================
# STATUS
# ============================================================
<br><br>
status_frame = tk.Frame(
    root,
    bg=BG
)
<br><br>
status_frame.pack(pady=10)
<br><br>
tk.Label(
    status_frame,
    text="Status:",
    bg=BG,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(side="left")
<br><br>
tk.Label(
    status_frame,
    textvariable=status_text,
    bg=BG,
    fg="#00D7FF",
    font=("Segoe UI", 10)
).pack(side="left", padx=5)
<br><br>
# ============================================================
# RUN BUTTON
# ============================================================
<br><br>
run_button = tk.Button(
    root,
    text="RUN VALIDATION",
    command=run_validation,
    bg=SUCCESS,
    fg="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 12, "bold"),
    width=25,
    height=2
)
<br><br>
run_button.pack(pady=25)
<br><br>
add_hover(run_button, SUCCESS, "#16A34A")
add_hover(sql_button, ACCENT, "#1A8CFF")
<br><br>
# ============================================================
# FOOTER
# ============================================================
<br><br>
footer = tk.Label(
    root,
    text="DDR / PLIP / ACONEX Validation Tool",
    bg=BG,
    fg="#808080",
    font=("Segoe UI", 9)
)
<br><br>
footer.pack(side="bottom", pady=10)
# ============================================================
# START APPLICATION
# ============================================================
<br><br>
root.mainloop()
------------
i will upload the code into github construct teh perfect read me file for this code and also note that i launch the program using .bat file named "Run DDR Validator"
<br>
**Response**
# DDR vs ACONEX Validator

A desktop validation tool that compares a **Document Distribution Register (DDR)** against:

* A **PLIP SPO Register**
* The **ACONEX Document Register** (loaded directly from SQL Server)

The tool validates document existence, title consistency, PLIP references, and ENPPI/BIR document numbering taxonomy, then generates a professional Excel report and dashboard summarizing all findings.

***

## Features

### DDR vs PLIP Validation

Checks whether the PLIP ID in the DDR exists in the PLIP SPO register.

Possible results:

* ✅ OK
* ⚠️ InActive PLIP ID
* ❌ PLIP ID Not Found

***

### DDR vs ACONEX Validation

Checks whether the document already exists in ACONEX using the document number.

Possible results:

* ✅ OK - No Corrective Action Required
* ⚠️ Title Mismatch - Update Title in Aconex
* 🟠 Placeholder Required - Create in Aconex
* 🔴 Placeholder Required - INVALID Doc Number Taxonomy

***

### ENPPI/BIR Taxonomy Validation

For documents not found in ACONEX, the tool validates the document number against the approved ENPPI/BIR numbering structure:

```text
Area - ProcessUnit - Originator - Discipline - DocumentType - SequentialNo - SheetNo
```

Example:

```text
BIR-36-103440-PX-2366-00001-0001
```

***

### ACONEX Cleanup Detection

Identifies documents that:

* Exist in ACONEX
* No longer exist in the latest DDR

These documents are flagged as:

```text
Delete from Aconex Document Register
```

> Note: The tool only reports these records. No deletion action is performed in ACONEX.

***

### Automated Excel Dashboard

Generates:

#### DDR Comparison Sheet

Contains:

* Document Number
* Document Title
* PLIP ID
* Status
* Review Status
* Document Type
* Notes

#### Summary Dashboard Sheet

Includes:

* Status Summary Table
* Counts per Status
* Grand Total
* Bar Chart Visualization
* Top Findings Summary

***

## System Requirements

### Python Version

```text
Python 3.10+
```

### Required Packages

Install dependencies:

```bash
pip install pandas openpyxl pyodbc pillow
```

Or:

```bash
pip install -r requirements.txt
```

***

## SQL Server Requirements

The application connects directly to:

```text
Server   : ES-MSSQL-01
Database : ACONEX Reporting Data
Table    : DocumentRegisterTest
```

Required SQL columns:

```text
DocNo
Title
ReviewStatus
FileName
ProjectId
```

Current project filter:

```sql
WHERE ProjectId = '1342183550'
```

***

## Folder Structure

```text
DDR Validator
│
├── compare_ddr_aconex.py
├── launcher.pyw
├── Run DDR Validator.bat
├── enppi_logo.png
└── README.md
```

***

## Running the Application

### Recommended Method

Launch the application using:

```text
Run DDR Validator.bat
```

The batch file starts the graphical interface without requiring command-line interaction.

***

### Alternative Method

Run directly with Python:

```bash
python launcher.pyw
```

***

## Using the Tool

### Step 1

Launch:

```text
Run DDR Validator.bat
```

### Step 2

Browse and select:

* DDR Excel File
* PLIP Excel File
* Output Folder

### Step 3

(Optional)

Click:

```text
Test SQL Connection
```

to verify connectivity to SQL Server.

### Step 4

Click:

```text
RUN VALIDATION
```

### Step 5

After processing completes:

* The Excel report is generated
* The report opens automatically
* The Summary Dashboard is displayed as the default sheet

***

## Validation Logic

### Step 1 – PLIP Check

If PLIP ID is not found:

```text
PLIP ID Not Found
```

or

```text
InActive PLIP ID
```

***

### Step 2 – ACONEX Existence Check

If the document is not found in ACONEX:

```text
Placeholder Required
```

***

### Step 3 – Taxonomy Validation

If the document number fails validation:

```text
Placeholder Required - INVALID Doc Number Taxonomy
```

***

### Step 4 – Title Validation

If DDR title matches ACONEX title:

```text
OK - No Corrective Action Required
```

Otherwise:

```text
Title Mismatch - Update Title in Aconex
```

***

### Step 5 – ACONEX Cleanup Check

Documents found in ACONEX but missing from DDR:

```text
Delete from Aconex Document Register
```

***

## Output Example

| Document Number                  | Status                                  |
| -------------------------------- | --------------------------------------- |
| BIR-36-103440-PX-2366-00001-0001 | OK - No Corrective Action Required      |
| BIR-36-103440-PX-2366-00002-0001 | Title Mismatch - Update Title in Aconex |
| BIR-36-103440-PX-2366-00003-0001 | Placeholder Required - Create in Aconex |

***

## GUI Features

* Modern Dark Theme
* SQL Connectivity Test
* Progress Indicator
* Automatic Output Opening
* Excel Dashboard Generation
* Status Color Coding
* Multi-threaded Processing

***

## Notes

* No updates are written back to ACONEX.
* No records are deleted from ACONEX.
* SQL access requires appropriate permissions.
* All outputs are generated locally as Excel workbooks.

***

## Author

**Hane Hatem**

DDR / PLIP / ACONEX Validation Tool for automated document register verification and reporting.
