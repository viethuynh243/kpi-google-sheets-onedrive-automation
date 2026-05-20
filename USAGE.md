# Huong dan su dung

## 1. Input can chuan bi

Nguoi dung chi can khai bao Google Sheet link trong:

```text
config/sources.json
```

Moi source can co:

- `url`: link Google Sheet.
- `file`: ten file raw se luu trong `data/input/raw`.
- `department`: phong ban de gan vao output.

Vi du:

```json
{
  "IT": {
    "url": "https://docs.google.com/spreadsheets/d/1x9FBjRHISImjCII7GqOcTTn40_BmAnRN/edit",
    "file": "IT.xlsx",
    "department": "IT"
  }
}
```

## 2. Chay tu link Google Sheet ra raw input

Chay:

```powershell
python automate_kpi.py --download
```

Ket qua raw input:

```text
data/input/raw/CMA.xlsx
data/input/raw/ACCA.xlsx
data/input/raw/IT.xlsx
```

Neu file local dang bi lock, script se doc ban download tam thoi va tiep tuc chay.

## 3. Chay tu raw input ra staging output

Neu da co san file trong `data/input/raw`, co the chay:

```powershell
python automate_kpi.py
```

Ket qua staging output:

```text
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.xlsx
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.csv
```

## 4. Chay workflow day du bang PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

Workflow nay gom:

1. Tai Google Sheet moi nhat.
2. Luu raw input vao `data/input/raw`.
3. Normalize du lieu.
4. Luu staging output vao `data/output/staging`.
5. Ghi log vao `logs`.

## 5. Kiem tra output

Mo file moi nhat trong:

```text
data/output/staging/
```

Kiem tra cac cot quan trong:

- `year`
- `month`
- `department`
- `employee`
- `product_or_project`
- `actual_quantity`
- `total_kpi`

## 6. Ket qua cuoi cung nam o dau?

Thu muc ket qua cuoi cung:

```text
data/output/final/
```

Hien tai folder nay la noi de workbook tong hop sau khi bo sung buoc ghi vao file dich. Buoc nay can file:

```text
Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx
```

va can mapping sheet/cot cua file dich.

## 7. Setup lich chay tu dong

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Mac dinh task chay ngay 3 hang thang luc 09:00.

## 8. Sua lich chay

Mo file:

```text
setup_kpi_scheduled_task.ps1
```

Sua cac dong:

```powershell
/D 3
/ST 09:00
```

Sau do chay lai script setup.

## 9. Thu muc khong duoc sua tay

- Khong can sua file trong `data/input/raw` neu chay bang `--download`, vi script se tu tai lai.
- Khong nen sua file trong `data/output/staging`, vi day la output sinh ra.
- Neu can chinh input, sua `config/sources.json`.

