# Automation Workflow

## Goal

Automatically collect KPI/personnel cost data from department Google Sheets, normalize the different source formats, and store staging output in the local OneDrive folder.

## End-to-end flow

```text
1. User edits config/sources.json
2. User runs run_kpi_automation.ps1
3. Script downloads Google Sheets as .xlsx files
4. Raw input is stored in data/input/raw
5. Python normalizes ACCA/CMA and IT formats
6. Staging output is stored in data/output/staging
7. IT template-compatible final output is stored in data/output/final
8. Logs are stored in logs
9. OneDrive syncs local files to cloud
```

## Source handling

### ACCA/CMA

Supported layouts:

- Monthly sheets: sheet name contains month/year, rows contain employees and tasks.
- Employee sheets: each row contains explicit year/month.

The script detects the table header and extracts employee, product, subject, component, actual quantity, KPI standard, and total KPI.

### IT

Supported layout:

- Rows contain project/category/system/month.
- Columns contain employees.
- Non-zero matrix cells are unpivoted into normalized rows.

## Output levels

### Raw input

```text
data/input/raw/*.xlsx
```

These are direct exports from Google Sheets.

### Staging output

```text
data/output/staging/normalized_output_*.xlsx
data/output/staging/normalized_output_*.csv
```

This is the current runnable output and the main deliverable of the automation.

### Final output

```text
data/output/final/IT_template_output_*.xlsx
```

For the IT template sample, this output preserves the same matrix structure as the source template.

The separate company workbook output is still future work because the target workbook and destination mapping are not available yet.

## Manual run

```powershell
cd "C:\Users\admin\OneDrive\Documents\Excel"
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

## Python run

```powershell
python automate_kpi.py --download
```

## Scheduled run

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Default:

- Day 3 every month.
- 09:00 local time.

## Operational notes

- The computer must be on when the scheduled task runs.
- OneDrive desktop app must be signed in.
- Google Sheet URLs must allow export/download.
- Source files can be open; the script can fall back to a temporary downloaded copy.
- The future destination workbook should be closed when direct write is implemented.

## Remaining work for final workbook automation

To write directly into `Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx`, add:

1. Target workbook path configuration.
2. Mapping from staging columns to `Data SX ACCA+CMA`.
3. Replace-by-period logic to avoid duplicates.
4. Validation report for written rows and rejected rows.
