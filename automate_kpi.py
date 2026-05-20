from __future__ import annotations

import argparse
import csv
import re
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook


GOOGLE_SHEETS = {
    "CMA": {
        "id": "1jOaBolZ78dbelYkoFL5jSUE0-yJn5425",
        "file": "CMA.xlsx",
        "department": "SX ACCA+CMA",
    },
    "ACCA": {
        "id": "16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ",
        "file": "ACCA.xlsx",
        "department": "SX ACCA+CMA",
    },
    "IT": {
        "id": "1x9FBjRHISImjCII7GqOcTTn40_BmAnRN",
        "file": "IT.xlsx",
        "department": "IT",
    },
}

OUTPUT_COLUMNS = [
    "source_file",
    "source_sheet",
    "source_type",
    "year",
    "month",
    "department",
    "program",
    "position",
    "employee",
    "product_or_project",
    "reference_link",
    "new_or_old",
    "project_name",
    "subject_or_system",
    "product_feature",
    "deliverable",
    "component",
    "complexity",
    "unit",
    "actual_quantity",
    "kpi_standard",
    "total_kpi",
]

MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


@dataclass
class Source:
    key: str
    path: Path
    department: str


def clean(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        value = re.sub(r"\s+", " ", value).strip()
        return value or None
    return value


def norm_header(value: Any) -> str:
    text = clean(value)
    if text is None:
        return ""
    return str(text).lower()


def download_sources(work_dir: Path) -> list[Source]:
    sources: list[Source] = []
    for key, meta in GOOGLE_SHEETS.items():
        out_path = work_dir / meta["file"]
        tmp_path = work_dir / f".{Path(meta['file']).stem}.download.xlsx"
        url = f"https://docs.google.com/spreadsheets/d/{meta['id']}/export?format=xlsx"
        print(f"Downloading {key} -> {out_path.name}")
        urllib.request.urlretrieve(url, tmp_path)
        source_path = out_path
        try:
            tmp_path.replace(out_path)
        except PermissionError:
            source_path = tmp_path
            print(f"Warning: {out_path.name} is locked. Using downloaded temporary copy.")
        sources.append(Source(key=key, path=source_path, department=meta["department"]))
    return sources


def local_sources(work_dir: Path) -> list[Source]:
    sources = []
    for key, meta in GOOGLE_SHEETS.items():
        path = work_dir / meta["file"]
        if path.exists():
            sources.append(Source(key=key, path=path, department=meta["department"]))
    return sources


def infer_month_year(sheet_name: str) -> tuple[int | None, int | None]:
    text = sheet_name.lower().strip()
    year_match = re.search(r"\b(20\d{2}|\d{2})\b", text)
    year = None
    if year_match:
        raw_year = int(year_match.group(1))
        year = raw_year if raw_year > 100 else 2000 + raw_year

    month = None
    for name, number in MONTHS.items():
        if re.search(rf"\b{name}\b", text):
            month = number
            break
    if month is None:
        month_match = re.search(r"\b(?:tháng|month)\s*(\d{1,2})\b", text)
        if month_match:
            month = int(month_match.group(1))
    return year, month


def find_header_row(ws, max_row: int = 30) -> tuple[int | None, dict[str, int]]:
    for row_idx in range(1, min(ws.max_row or 0, max_row) + 1):
        headers = {norm_header(ws.cell(row_idx, col).value): col for col in range(1, (ws.max_column or 0) + 1)}
        if "tên nhân viên" in headers and ("tên sản phẩm" in headers or "năm" in headers):
            return row_idx, headers
    return None, {}


def value_by_header(ws, row: int, headers: dict[str, int], *names: str) -> Any:
    for name in names:
        col = headers.get(name.lower())
        if col:
            return clean(ws.cell(row, col).value)
    return None


def is_kpi_input_sheet(ws) -> bool:
    row, headers = find_header_row(ws)
    return bool(row and "tên nhân viên" in headers)


def extract_kpi_rows(source: Source) -> list[dict[str, Any]]:
    wb = load_workbook(source.path, data_only=True)
    rows: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        header_row, headers = find_header_row(ws)
        if not header_row:
            continue

        inferred_year, inferred_month = infer_month_year(ws.title)
        has_year_month = "năm" in headers and "tháng" in headers
        source_type = "employee_month" if has_year_month else "month_employee"

        for row_idx in range(header_row + 2, (ws.max_row or 0) + 1):
            employee = value_by_header(ws, row_idx, headers, "tên nhân viên")
            product = value_by_header(ws, row_idx, headers, "tên sản phẩm")
            actual = value_by_header(ws, row_idx, headers, "số lượng actual")
            total = value_by_header(ws, row_idx, headers, "total kpi")
            if not any([employee, product, actual, total]):
                continue

            year = value_by_header(ws, row_idx, headers, "năm") if has_year_month else inferred_year
            month = value_by_header(ws, row_idx, headers, "tháng") if has_year_month else inferred_month
            rows.append(
                {
                    "source_file": source.path.name,
                    "source_sheet": ws.title,
                    "source_type": source_type,
                    "year": year,
                    "month": month,
                    "department": source.department,
                    "program": value_by_header(ws, row_idx, headers, "chương trình"),
                    "position": value_by_header(ws, row_idx, headers, "vị trí"),
                    "employee": employee,
                    "product_or_project": product,
                    "reference_link": value_by_header(ws, row_idx, headers, "link tham chiếu (gắn link wework, workflow)"),
                    "new_or_old": value_by_header(ws, row_idx, headers, "sản phẩm mới/sản phẩm cũ"),
                    "project_name": value_by_header(ws, row_idx, headers, "tên dự án"),
                    "subject_or_system": value_by_header(ws, row_idx, headers, "bộ môn"),
                    "product_feature": value_by_header(ws, row_idx, headers, "đặc tính sản phẩm"),
                    "deliverable": value_by_header(ws, row_idx, headers, "sản phẩm bàn giao"),
                    "component": value_by_header(ws, row_idx, headers, "cấu phần"),
                    "complexity": value_by_header(ws, row_idx, headers, "độ khó"),
                    "unit": value_by_header(ws, row_idx, headers, "đơn vị tính"),
                    "actual_quantity": actual,
                    "kpi_standard": value_by_header(ws, row_idx, headers, "kpi standard"),
                    "total_kpi": total,
                }
            )
    return rows


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def extract_it_rows(source: Source) -> list[dict[str, Any]]:
    wb = load_workbook(source.path, data_only=True)
    rows: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        year = int(ws.title) if str(ws.title).isdigit() else None
        employee_cols = []
        for col in range(6, (ws.max_column or 0) + 1):
            employee = clean(ws.cell(14, col).value)
            if employee:
                employee_cols.append((col, employee))

        current = {"tt": None, "project": None, "category": None, "system": None}
        for row_idx in range(15, (ws.max_row or 0) + 1):
            tt = clean(ws.cell(row_idx, 1).value)
            project = clean(ws.cell(row_idx, 2).value)
            category = clean(ws.cell(row_idx, 3).value)
            system = clean(ws.cell(row_idx, 4).value)
            month = clean(ws.cell(row_idx, 5).value)

            if tt is not None:
                current["tt"] = tt
            if project is not None:
                current["project"] = project
            if category is not None:
                current["category"] = category
            if system is not None:
                current["system"] = system

            if not current["project"] or not is_number(month):
                continue

            for col, employee in employee_cols:
                actual = clean(ws.cell(row_idx, col).value)
                if not is_number(actual) or actual == 0:
                    continue
                rows.append(
                    {
                        "source_file": source.path.name,
                        "source_sheet": ws.title,
                        "source_type": "it_matrix",
                        "year": year,
                        "month": int(month),
                        "department": source.department,
                        "program": None,
                        "position": employee.split(" - ")[-1] if " - " in employee else None,
                        "employee": employee,
                        "product_or_project": current["project"],
                        "reference_link": None,
                        "new_or_old": None,
                        "project_name": current["project"],
                        "subject_or_system": current["system"],
                        "product_feature": current["category"],
                        "deliverable": None,
                        "component": None,
                        "complexity": None,
                        "unit": "hour_or_effort",
                        "actual_quantity": actual,
                        "kpi_standard": None,
                        "total_kpi": actual,
                    }
                )
    return rows


def extract_all(sources: Iterable[Source]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source in sources:
        print(f"Extracting {source.path.name}")
        if source.key == "IT":
            rows.extend(extract_it_rows(source))
        else:
            rows.extend(extract_kpi_rows(source))
    return rows


def write_csv(rows: list[dict[str, Any]], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(rows: list[dict[str, Any]], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "normalized_data"
    ws.append(OUTPUT_COLUMNS)
    for row in rows:
        ws.append([row.get(col) for col in OUTPUT_COLUMNS])
    ws.freeze_panes = "A2"
    wb.save(path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Normalize KPI Google Sheets into one table.")
    parser.add_argument("--work-dir", default=".", help="Folder containing source/output files.")
    parser.add_argument("--download", action="store_true", help="Download latest Google Sheets before extracting.")
    args = parser.parse_args()

    work_dir = Path(args.work_dir).resolve()
    sources = download_sources(work_dir) if args.download else local_sources(work_dir)
    if not sources:
        raise SystemExit("No source files found. Run with --download or place ACCA.xlsx/CMA.xlsx/IT.xlsx in work-dir.")

    rows = extract_all(sources)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = work_dir / f"normalized_output_{timestamp}.csv"
    xlsx_path = work_dir / f"normalized_output_{timestamp}.xlsx"
    write_csv(rows, csv_path)
    write_xlsx(rows, xlsx_path)
    print(f"Done: {len(rows)} rows")
    print(f"CSV : {csv_path}")
    print(f"XLSX: {xlsx_path}")


if __name__ == "__main__":
    main()
