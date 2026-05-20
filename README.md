# KPI Google Sheets to OneDrive Automation

Automation nay tai KPI/chi phi nhan su tu Google Sheets, chuan hoa du lieu ACCA/CMA/IT ve mot bang staging, va luu ket qua vao folder OneDrive local.

Project da tach ro input, output, log va config de tranh tinh trang file bi tron trong root folder.

## Workflow

```text
config/sources.json
  -> data/input/raw/*.xlsx
  -> automate_kpi.py
  -> data/output/staging/normalized_output_*.xlsx
  -> data/output/staging/normalized_output_*.csv
  -> data/output/final/  (future final workbook output)
```

## Folder structure

```text
.
|-- config/
|   `-- sources.json
|-- data/
|   |-- input/
|   |   `-- raw/
|   |       |-- ACCA.xlsx
|   |       |-- CMA.xlsx
|   |       `-- IT.xlsx
|   `-- output/
|       |-- staging/
|       |   |-- normalized_output_*.xlsx
|       |   `-- normalized_output_*.csv
|       `-- final/
|-- logs/
|   `-- kpi_automation_*.log
|-- automate_kpi.py
|-- run_kpi_automation.ps1
|-- setup_kpi_scheduled_task.ps1
|-- cleanup_legacy_root.ps1
|-- requirements.txt
|-- USAGE.md
|-- REQUIREMENTS_AUTOMATION.md
`-- WORKFLOW_AUTOMATION.md
```

Files in `data/input/raw`, `data/output`, and `logs` are generated/local data and are ignored by git.

## Input

Input is configured in:

```text
config/sources.json
```

Each source has:

- `url`: Google Sheet URL.
- `file`: local exported file name under `data/input/raw`.
- `department`: department label written to output.

Example:

```json
{
  "ACCA": {
    "url": "https://docs.google.com/spreadsheets/d/16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ/edit",
    "file": "ACCA.xlsx",
    "department": "SX ACCA+CMA"
  }
}
```

When running with `--download`, the program converts each Google Sheet URL to:

```text
https://docs.google.com/spreadsheets/d/<sheet_id>/export?format=xlsx
```

Then it saves raw input files here:

```text
data/input/raw/ACCA.xlsx
data/input/raw/CMA.xlsx
data/input/raw/IT.xlsx
```

## Processing logic

### ACCA/CMA

The script supports two formats:

- `month_employee`: one sheet per month, many employees inside that sheet. Year/month are inferred from the sheet name, for example `Apr 26`.
- `employee_month`: one sheet per employee, with explicit `Nam` and `Thang` columns.

The script finds the input table by detecting headers such as employee name, product name, actual quantity, KPI standard, and total KPI.

### IT

The IT source is treated as a matrix:

- Row 13 contains project/category/system/month fields.
- Row 14 contains employee names across columns.
- Each non-zero employee cell is unpivoted into one normalized output row.

## Output

Current output is staging data:

```text
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.xlsx
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.csv
```

The staging table includes:

- `source_file`
- `source_sheet`
- `source_type`
- `year`
- `month`
- `department`
- `program`
- `position`
- `employee`
- `product_or_project`
- `reference_link`
- `new_or_old`
- `project_name`
- `subject_or_system`
- `product_feature`
- `deliverable`
- `component`
- `complexity`
- `unit`
- `actual_quantity`
- `kpi_standard`
- `total_kpi`

This staging output is the current final runnable result of the project. It is ready for review, Power Query import, or later mapping into the company workbook.

## Final workbook output

The future final output should be stored under:

```text
data/output/final/
```

To generate the final workbook automatically, the target workbook must be available and its destination columns must be mapped:

```text
Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx
sheet: Data SX ACCA+CMA
```

This direct-write step is not implemented yet because the target workbook was not present when the automation was built.

## Run from scratch

From PowerShell:

```powershell
cd "C:\Users\admin\OneDrive\Documents\Excel"
pip install -r requirements.txt
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

After the run:

- Raw Google Sheet exports are in `data/input/raw`.
- Staging output is in `data/output/staging`.
- Logs are in `logs`.

## Run Python directly

Download latest Google Sheets and build staging output:

```powershell
python automate_kpi.py --download
```

Use existing raw files from `data/input/raw`:

```powershell
python automate_kpi.py
```

## Schedule monthly run

Create or update the Windows Scheduled Task:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Default schedule:

- Day 3 of every month.
- 09:00 local time.

## Clean generated files

Delete generated staging/final/log files manually, or remove old root-level generated files with:

```powershell
powershell -ExecutionPolicy Bypass -File ".\cleanup_legacy_root.ps1"
```

## Data safety

The public GitHub repo is intended to contain only code, config template, and docs. Generated Excel/CSV/log files are ignored by git:

- `*.xlsx`
- `*.csv`
- `logs/`
- `data/input/raw/*.xlsx`
- `data/output/**/*.xlsx`
- `data/output/**/*.csv`

