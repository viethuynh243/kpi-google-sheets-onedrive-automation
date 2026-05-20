# KPI Google Sheets to OneDrive Automation

Tool tu dong tai du lieu KPI tu Google Sheets, chuan hoa cac format khac nhau cua phong ban, va xuat bang staging vao thu muc OneDrive local.

## Chuc nang

- Tai 3 Google Sheet nguon ve dang `.xlsx`.
- Doc va chuan hoa du lieu SX ACCA/CMA:
  - Sheet theo thang, trong sheet co nhieu nhan vien.
  - Sheet theo nhan vien, trong sheet co cot `Nam` va `Thang`.
- Doc va chuan hoa du lieu IT dang ma tran:
  - Dong la du an/hang muc/thang.
  - Cot ngang la nhan vien.
- Xuat output:
  - `normalized_output_YYYYMMDD_HHMMSS.xlsx`
  - `normalized_output_YYYYMMDD_HHMMSS.csv`
- Co runner PowerShell de chay tay.
- Co script tao lich Windows Task Scheduler de chay tu dong hang thang.
- Co log moi lan chay trong thu muc `logs/`.

## Cau truc project

```text
.
├── automate_kpi.py
├── run_kpi_automation.ps1
├── setup_kpi_scheduled_task.ps1
├── requirements.txt
├── REQUIREMENTS_AUTOMATION.md
├── WORKFLOW_AUTOMATION.md
└── README.md
```

## Requirement

- Windows.
- Python 3.10 tro len.
- OneDrive desktop app da login va sync folder chua project.
- Google Sheets nguon phai cho phep export/download.
- Python package:

```powershell
pip install -r requirements.txt
```

## Cach chay nhanh

Chay bang PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

Script se:

1. Tai Google Sheet moi nhat.
2. Extract du lieu tu ACCA/CMA/IT.
3. Tao file output `.xlsx` va `.csv`.
4. Ghi log trong `logs/`.

## Cach chay truc tiep bang Python

Dung file nguon da co san trong folder:

```powershell
python automate_kpi.py
```

Tai lai Google Sheets truoc khi xu ly:

```powershell
python automate_kpi.py --download
```

Chi dinh folder lam viec:

```powershell
python automate_kpi.py --work-dir "C:\Users\admin\OneDrive\Documents\Excel" --download
```

## Output staging

Bang output gom cac cot:

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

## Setup chay tu dong hang thang

Chay lenh:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Mac dinh task chay:

- Ngay 3 hang thang.
- Luc 09:00 theo gio may.

Task name:

```text
KPI GoogleSheet To OneDrive Automation
```

## Ghi chu ve file bi lock

Neu Excel hoac OneDrive dang lock file nguon, script van co the chay tiep bang ban download tam thoi. Vi du:

```text
Warning: ACCA.xlsx is locked. Using downloaded temporary copy.
```

Neu sau nay ghi truc tiep vao file tong hop Excel, nen dong file dich truoc khi job chay de tranh loi lock.

## Gioi han hien tai

Workflow hien tai automate den buoc staging output. De ghi thang vao file tong hop:

```text
Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx
```

can bo sung mapping cot vao sheet dich, vi file dich chua co trong workspace luc xay dung tool.

## Bao mat du lieu

Repo nay khong nen commit file Excel nguon, output CSV/XLSX, log, hoac file chua du lieu nhan su. `.gitignore` da loai tru cac file do.

