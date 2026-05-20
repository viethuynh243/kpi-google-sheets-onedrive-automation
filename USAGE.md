# Huong dan su dung

## 1. Chuan bi

1. Dat project trong folder OneDrive local, vi du:

```text
C:\Users\admin\OneDrive\Documents\Excel
```

2. Cai Python package:

```powershell
pip install -r requirements.txt
```

3. Dam bao 3 Google Sheet nguon co quyen export/download.

## 2. Chay thu mot lan

Mo PowerShell tai folder project va chay:

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

Sau khi chay xong, kiem tra:

- File output `normalized_output_*.xlsx`
- File output `normalized_output_*.csv`
- File log trong folder `logs`

## 3. Kiem tra ket qua

Mo file `normalized_output_*.xlsx`, sheet `normalized_data`.

Can kiem tra nhanh:

- So dong co hop ly khong.
- Cot `year` va `month` da dung ky chua.
- Cot `department` phan biet duoc `SX ACCA+CMA` va `IT`.
- Cot `employee` co day du nhan vien khong.
- Cot `actual_quantity` va `total_kpi` co gia tri khong.

## 4. Bat lich chay tu dong

Chay:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Sau khi tao task, co the kiem tra trong Windows Task Scheduler:

```text
Task Scheduler Library > KPI GoogleSheet To OneDrive Automation
```

## 5. Doi lich chay

Mo file `setup_kpi_scheduled_task.ps1` va sua phan:

```powershell
$Trigger = New-ScheduledTaskTrigger `
    -Monthly `
    -DaysOfMonth 3 `
    -At 9:00AM
```

Vi du doi sang ngay 5 hang thang luc 18:00:

```powershell
$Trigger = New-ScheduledTaskTrigger `
    -Monthly `
    -DaysOfMonth 5 `
    -At 6:00PM
```

Sau khi sua, chay lai:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

## 6. Troubleshooting

### File Excel bi lock

Neu thay warning:

```text
Warning: ACCA.xlsx is locked. Using downloaded temporary copy.
```

Day khong phai loi nghiem trong. Script dang dung file tam vua download de tiep tuc xu ly.

### Khong tai duoc Google Sheet

Kiem tra:

- Link Google Sheet co con dung khong.
- File co quyen view/export khong.
- May co internet khong.

### Khong thay output moi

Kiem tra file log moi nhat trong folder `logs/`.

### Ghi truc tiep vao file tong hop bi loi

Dong file Excel dich truoc khi chay job. Excel desktop thuong lock file khi dang mo.

