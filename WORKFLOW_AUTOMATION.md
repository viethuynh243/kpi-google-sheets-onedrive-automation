# Workflow automate de xuat

## Muc tieu workflow

Tu dong lay du lieu tu 3 Google Sheet phong ban, normalize ve mot bang chung, ghi output vao thu muc OneDrive local de OneDrive tu dong sync len cloud.

## Luong xu ly

1. Windows Task Scheduler kich hoat `run_kpi_automation.ps1`.
2. `run_kpi_automation.ps1` goi `automate_kpi.py --download`.
3. `automate_kpi.py` tai 3 file Google Sheet ve dang `.xlsx`.
4. Neu mot file local dang bi Excel/OneDrive lock, script dung ban download tam thoi de doc tiep.
5. Script extract:
   - ACCA/CMA dang sheet theo thang.
   - ACCA/CMA dang sheet theo nhan vien co cot nam/thang.
   - IT dang ma tran nhan vien theo cot ngang.
6. Script tao output:
   - `normalized_output_YYYYMMDD_HHMMSS.xlsx`
   - `normalized_output_YYYYMMDD_HHMMSS.csv`
7. Log moi lan chay duoc luu trong folder `logs`.
8. OneDrive desktop sync cac file output len cloud.

## File trong workflow

- `automate_kpi.py`: logic download va normalize.
- `run_kpi_automation.ps1`: file runner co ghi log.
- `setup_kpi_scheduled_task.ps1`: tao lich chay tu dong tren Windows Task Scheduler.
- `requirements.txt`: dependency Python.

## Cach kich hoat workflow tu dong

Chay PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\admin\OneDrive\Documents\Excel\setup_kpi_scheduled_task.ps1"
```

Mac dinh lich chay:

- Ngay 3 hang thang.
- Luc 09:00 sang theo gio may.

## Cach chay tay khi can

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\admin\OneDrive\Documents\Excel\run_kpi_automation.ps1"
```

## Dieu kien de workflow chay on dinh

- May tinh phai bat vao thoi diem Task Scheduler chay. Neu may tat, task se chay lai khi may bat nho setting `StartWhenAvailable`.
- OneDrive desktop app dang login va sync folder `C:\Users\admin\OneDrive\Documents\Excel`.
- Google Sheet source cho phep export/download.
- Khong bat buoc phai dong Excel source nua, vi script co fallback file tam. Tuy nhien file dich tong hop sau nay nen dong khi script ghi truc tiep vao sheet dich.

## Buoc tiep theo de hoan thien 100%

Hien workflow da automate den staging output. De ghi thang vao file tong hop can co file:

`Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx`

Sau khi co file nay, can bo sung:

1. Mapping cot output staging vao sheet `Data SX ACCA+CMA`.
2. Che do replace theo ky thang de tranh duplicate.
3. Log so dong ghi vao file dich va so dong bi loi mapping.
