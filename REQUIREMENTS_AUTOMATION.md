# Requirement va thiet ke automation

## 1. Muc tieu

Tu dong tong hop du lieu KPI/chi phi nhan su tu cac file Google Sheet phong ban vao mot bang chuan, sau do dua vao file tong hop tren OneDrive:

- Nguon SX ACCA/CMA: file co cau truc giong nhau nhung co 2 cach nhap khac nhau:
  - Gop cong viec theo thang, trong sheet thang co nhieu nhan vien.
  - Gop cong viec theo nhan vien, trong sheet nhan vien co cot nam va thang.
- Nguon IT: file dang ma tran, dong la hang muc/du an/thang, cot ngang la nhan vien.
- Dich: file `Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx`, sheet `Data SX ACCA+CMA` hoac sheet du lieu tong hop tuong ung.

## 2. Pham vi automate

Automation can thuc hien cac buoc:

1. Tai ban moi nhat cua cac Google Sheet ve may bang link export `.xlsx`.
2. Doc tung file Excel nguon.
3. Chuan hoa du lieu tu cac dang input khac nhau ve mot bang duy nhat.
4. Ghi output ra file trung gian `.xlsx`/`.csv`.
5. Khi co file dich that, ghi truc tiep vao sheet data cua file tong hop nam trong thu muc OneDrive local.
6. OneDrive desktop app tu dong sync file sau khi script ghi xong.

## 3. Requirement moi truong

- May Windows co cai OneDrive desktop app va da sync folder chua file tong hop.
- Python 3.10 tro len.
- Thu vien Python:
  - `openpyxl`
- Quyen truy cap Google Sheet:
  - Link phai cho phep download/export.
  - Neu file private, can dung file da download san hoac bo sung Google API credential.
- File Excel dich khong duoc dang bi lock boi Excel khi script ghi du lieu.

## 4. File nguon

| Phong ban | File local | Google Sheet ID |
| --- | --- | --- |
| SX ACCA/CMA | `CMA.xlsx` | `1jOaBolZ78dbelYkoFL5jSUE0-yJn5425` |
| SX ACCA/CMA | `ACCA.xlsx` | `16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ` |
| IT | `IT.xlsx` | `1x9FBjRHISImjCII7GqOcTTn40_BmAnRN` |

## 5. Bang output chuan

Output hien tai gom cac cot:

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

Bang nay duoc dung lam lop staging. Khi co file tong hop that, mapping cot tu staging sang sheet dich se duoc cau hinh them theo header cua file dich.

## 6. Logic xu ly

### SX ACCA/CMA

- Tim dong header co `Tên nhân viên` va `Tên sản phẩm`.
- Neu sheet co cot `Năm` va `Tháng`: lay truc tiep nam/thang tu dong du lieu.
- Neu sheet khong co cot `Năm` va `Tháng`: suy ra nam/thang tu ten sheet, vi du `Apr 26` -> thang 4 nam 2026.
- Bo qua dong rong.
- Lay cac cot cong viec chinh: chuong trinh, vi tri, nhan vien, san pham, link, bo mon, dac tinh san pham, san pham ban giao, cau phan, do kho, so luong actual, KPI standard, total KPI.

### IT

- Doc sheet theo nam, vi du `2026`.
- Dong 13 la thong tin hang muc: `TT`, `Tên dự án`, `Phân loại`, `Hệ thống`, `Tháng`.
- Dong 14 la danh sach nhan vien theo cot ngang.
- Tu dong unpivot ma tran: moi o nhan vien co gia tri khac 0 se thanh mot dong output.
- Lay vi tri tu suffix trong ten nhan vien neu co dang `Nhân viên demo 065 - BA`.

## 7. Cach chay

Chay voi file nguon da co san trong thu muc:

```powershell
python automate_kpi.py
```

Tai lai Google Sheet roi moi xu ly:

```powershell
python automate_kpi.py --download
```

Chay voi thu muc khac:

```powershell
python automate_kpi.py --work-dir "C:\Users\admin\OneDrive\Documents\Excel" --download
```

Sau khi chay, chuong trinh tao:

- `normalized_output_YYYYMMDD_HHMMSS.csv`
- `normalized_output_YYYYMMDD_HHMMSS.xlsx`

## 8. De xuat lich chay

Nen dung Windows Task Scheduler:

- Trigger: ngay 3 hang thang, sau deadline nhan su nhap KPI.
- Action:

```powershell
python "C:\Users\admin\OneDrive\Documents\Excel\automate_kpi.py" --work-dir "C:\Users\admin\OneDrive\Documents\Excel" --download
```

## 9. Gioi han va rui ro

- Power Automate cloud khong phai lua chon tot neu tai khoan OneDrive la personal/family/nhom, vi kha nang login/connector bi gioi han hon so voi tai khoan business.
- Neu Google Sheet doi format header, script co the can update mapping.
- Neu file dich dang mo trong Excel, thao tac ghi file co the fail.
- Neu cong thuc trong Google Sheet khong duoc luu cached value vao file export, mot so cot cong thuc nhu `KPI standard`/`Total KPI` co the trong. Khi do nen tinh lai trong script hoac lay tu bang KPI standard.

## 10. Viec can bo sung khi co file dich

1. Doc sheet `Data SX ACCA+CMA` trong file `Von_hoa_chi_phi_nhan_su_2026_ANON_Huong_dan (1).xlsx`.
2. Xac dinh header/cot can ghi.
3. Them mapping tu bang staging sang sheet dich.
4. Them che do append hoac replace du lieu theo ky thang.
5. Them log bao cao so dong doc, so dong ghi, va cac dong loi mapping.
