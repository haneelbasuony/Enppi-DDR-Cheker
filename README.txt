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
