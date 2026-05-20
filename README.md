# KPI Google Sheets to OneDrive Automation

Project nay tu dong lay du lieu KPI/nhan su tu Google Sheets, chuan hoa cac format khac nhau cua phong ban, va xuat ket qua vao thu muc OneDrive local theo cau truc ro rang.

## 1. Workflow tong quan

```text
Google Sheet links
        |
        v
config/sources.json
        |
        v
data/input/raw/*.xlsx
        |
        v
automate_kpi.py
        |
        v
data/output/staging/normalized_output_*.xlsx
data/output/staging/normalized_output_*.csv
        |
        v
data/output/final/
```

Trong do:

- `Google Sheet links`: link nguon phong ban cung cap.
- `config/sources.json`: noi khai bao link nguon, ten file raw, va phong ban.
- `data/input/raw/`: noi luu file `.xlsx` tai tu Google Sheets.
- `data/output/staging/`: bang da normalize, dung de review/mapping.
- `data/output/final/`: noi de file ket qua cuoi cung sau khi bo sung mapping vao workbook tong hop.

## 2. Input tu link duoc cung cap

Input duoc khai bao tai:

[config/sources.json](config/sources.json)

Vi du:

```json
{
  "ACCA": {
    "url": "https://docs.google.com/spreadsheets/d/16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ/edit",
    "file": "ACCA.xlsx",
    "department": "SX ACCA+CMA"
  }
}
```

Khi chay voi `--download`, chuong trinh tu dong chuyen link tren thanh link export:

```text
https://docs.google.com/spreadsheets/d/<sheet_id>/export?format=xlsx
```

Sau do file duoc tai ve:

```text
data/input/raw/ACCA.xlsx
data/input/raw/CMA.xlsx
data/input/raw/IT.xlsx
```

## 3. Xu ly du lieu

### SX ACCA/CMA

Chuong trinh tu nhan dien 2 dang file:

- `month_employee`: sheet theo thang, vi du `Apr 26`; nam/thang duoc suy ra tu ten sheet.
- `employee_month`: sheet theo nhan vien; nam/thang lay truc tiep tu cot `Nam` va `Thang`.

Header duoc tim bang cac cot nhu:

- `Ten nhan vien`
- `Ten san pham`
- `So luong actual`
- `KPI standard`
- `Total KPI`

### IT

File IT duoc xu ly theo dang ma tran:

- Dong 13: thong tin hang muc/du an/thang.
- Dong 14: danh sach nhan vien theo cot ngang.
- Moi o co gia tri khac 0 se duoc unpivot thanh mot dong output.

## 4. Output la gi?

Output hien tai la bang staging da chuan hoa:

```text
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.xlsx
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.csv
```

Bang staging gom cac cot:

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

Day la ket qua da duoc chuan hoa tu moi file nguon. File nay co the dung de:

- Review du lieu tong hop.
- Import vao Power Query/Excel.
- Mapping tiep vao workbook tong hop `Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx`.

## 5. Ket qua cuoi cung

Ket qua cuoi cung du kien nam trong:

```text
data/output/final/
```

Hien tai project da automate den buoc staging. De ghi thang vao file tong hop cuoi cung, can dat file dich vao project va bo sung mapping sheet/cot:

```text
Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx
sheet: Data SX ACCA+CMA
```

Ly do chua ghi thang vao file dich: file workbook dich chua co trong workspace luc build script, nen chua xac dinh duoc header/cot can ghi.

## 6. Cau truc thu muc

```text
.
├── config/
│   └── sources.json
├── data/
│   ├── input/
│   │   └── raw/
│   │       ├── ACCA.xlsx
│   │       ├── CMA.xlsx
│   │       └── IT.xlsx
│   └── output/
│       ├── staging/
│       │   ├── normalized_output_*.xlsx
│       │   └── normalized_output_*.csv
│       └── final/
├── logs/
│   └── kpi_automation_*.log
├── automate_kpi.py
├── run_kpi_automation.ps1
├── setup_kpi_scheduled_task.ps1
├── requirements.txt
├── USAGE.md
├── REQUIREMENTS_AUTOMATION.md
└── WORKFLOW_AUTOMATION.md
```

File Excel nguon, output, va log khong duoc commit len GitHub.

## 7. Cach chay

Cai dependency:

```powershell
pip install -r requirements.txt
```

Chay workflow day du:

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

Chay Python truc tiep:

```powershell
python automate_kpi.py --download
```

## 8. Chay tu dong

Tao scheduled task:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Mac dinh:

- Chay ngay 3 hang thang.
- Luc 09:00.

## 9. Bao mat du lieu

Repo public chi chua code/config/docs. Cac file sau bi ignore:

- `*.xlsx`
- `*.csv`
- `logs/`
- file output sinh ra

