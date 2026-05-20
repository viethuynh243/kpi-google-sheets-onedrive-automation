from __future__ import annotations

import argparse
import csv
import json
import re
import unicodedata
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook


DEFAULT_GOOGLE_SHEETS = {
    "CMA": {
        "id": "1jOaBolZ78dbelYkoFL5jSUE0-yJn5425",
        "url": "https://docs.google.com/spreadsheets/d/1jOaBolZ78dbelYkoFL5jSUE0-yJn5425/edit",
        "file": "CMA.xlsx",
        "department": "SX ACCA+CMA",
    },
    "ACCA": {
        "id": "16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ",
        "url": "https://docs.google.com/spreadsheets/d/16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ/edit",
        "file": "ACCA.xlsx",
        "department": "SX ACCA+CMA",
    },
    "IT": {
        "id": "1x9FBjRHISImjCII7GqOcTTn40_BmAnRN",
        "url": "https://docs.google.com/spreadsheets/d/1x9FBjRHISImjCII7GqOcTTn40_BmAnRN/edit",
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

TEMPLATE_SHEET_NAME = "3. vốn hóa"
TEMPLATE_FILE_NAME = "von_hoa_template.xlsx"

TEMPLATE_HEADER_ALIASES = {
    "year": ["nam", "year"],
    "month": ["thang", "month"],
    "department": ["phong ban", "department", "bo phan"],
    "program": ["chuong trinh", "program"],
    "position": ["vi tri", "position", "chuc danh"],
    "employee": ["nhan vien", "ten nhan vien", "ho ten", "employee"],
    "product_or_project": ["san pham/du an", "san pham", "du an", "ten san pham", "ten du an"],
    "reference_link": ["link tham chieu", "link"],
    "new_or_old": ["san pham moi/cu", "san pham moi/san pham cu"],
    "project_name": ["ten du an"],
    "subject_or_system": ["bo mon/he thong", "bo mon", "he thong"],
    "product_feature": ["dac tinh/phan loai", "dac tinh san pham", "phan loai"],
    "deliverable": ["san pham ban giao"],
    "component": ["cau phan"],
    "complexity": ["do kho"],
    "unit": ["don vi tinh"],
    "actual_quantity": ["so luong actual", "actual quantity", "actual"],
    "kpi_standard": ["kpi standard"],
    "total_kpi": ["total kpi"],
    "source_file": ["file nguon"],
    "source_sheet": ["sheet nguon"],
    "source_type": ["loai nguon"],
}

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


def ensure_project_dirs(work_dir: Path) -> dict[str, Path]:
    dirs = {
        "config": work_dir / "config",
        "raw_input": work_dir / "data" / "input" / "raw",
        "template_input": work_dir / "data" / "input" / "template",
        "staging_output": work_dir / "data" / "output" / "staging",
        "final_output": work_dir / "data" / "output" / "final",
        "logs": work_dir / "logs",
    }
    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)
    return dirs


def extract_sheet_id(meta: dict[str, Any]) -> str:
    if meta.get("id"):
        return str(meta["id"])
    url = str(meta.get("url", ""))
    match = re.search(r"/spreadsheets/d/([^/]+)", url)
    if not match:
        raise ValueError(f"Missing Google Sheet id/url for source: {meta}")
    return match.group(1)


def load_sources_config(work_dir: Path) -> dict[str, dict[str, Any]]:
    config_path = work_dir / "config" / "sources.json"
    if not config_path.exists():
        return DEFAULT_GOOGLE_SHEETS
    with config_path.open(encoding="utf-8") as f:
        return json.load(f)


def clean(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        value = re.sub(r"\s+", " ", value).strip()
        return value or None
    return value


def ascii_key(value: Any) -> str:
    text = clean(value)
    if text is None:
        return ""
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = re.sub(r"[^a-zA-Z0-9]+", " ", text).lower().strip()
    return text


def norm_header(value: Any) -> str:
    text = clean(value)
    if text is None:
        return ""
    return str(text).lower()


def download_sources(work_dir: Path, sources_config: dict[str, dict[str, Any]]) -> list[Source]:
    dirs = ensure_project_dirs(work_dir)
    raw_input_dir = dirs["raw_input"]
    sources: list[Source] = []
    for key, meta in sources_config.items():
        out_path = raw_input_dir / meta["file"]
        tmp_path = raw_input_dir / f".{Path(meta['file']).stem}.download.xlsx"
        sheet_id = extract_sheet_id(meta)
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
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


def local_sources(work_dir: Path, sources_config: dict[str, dict[str, Any]]) -> list[Source]:
    dirs = ensure_project_dirs(work_dir)
    raw_input_dir = dirs["raw_input"]
    sources = []
    for key, meta in sources_config.items():
        path = raw_input_dir / meta["file"]
        if not path.exists():
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


def find_template_sheet(wb):
    wanted = ascii_key(TEMPLATE_SHEET_NAME)
    for ws in wb.worksheets:
        if ascii_key(ws.title) == wanted:
            return ws
    return None


def find_template_header(ws) -> tuple[int | None, dict[int, str]]:
    alias_to_key = {}
    for key, aliases in TEMPLATE_HEADER_ALIASES.items():
        for alias in aliases:
            alias_to_key[ascii_key(alias)] = key

    best_row = None
    best_mapping: dict[int, str] = {}
    for row_idx in range(1, min(ws.max_row or 0, 80) + 1):
        mapping: dict[int, str] = {}
        for col_idx in range(1, (ws.max_column or 0) + 1):
            header = ascii_key(ws.cell(row_idx, col_idx).value)
            if header in alias_to_key:
                mapping[col_idx] = alias_to_key[header]
        if len(mapping) > len(best_mapping):
            best_row = row_idx
            best_mapping = mapping
    if len(best_mapping) < 3:
        return None, {}
    return best_row, best_mapping


def write_template_final(rows: list[dict[str, Any]], template_path: Path, output_path: Path) -> bool:
    if not template_path.exists():
        print(f"FINAL SKIPPED: template not found at {template_path}")
        return False

    wb = load_workbook(template_path)
    ws = find_template_sheet(wb)
    if ws is None:
        print(f"FINAL SKIPPED: sheet '{TEMPLATE_SHEET_NAME}' not found in {template_path}")
        return False

    header_row, column_mapping = find_template_header(ws)
    if header_row is None:
        print(f"FINAL SKIPPED: could not detect headers in sheet '{ws.title}'")
        return False

    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)

    for row in rows:
        out = [None] * ws.max_column
        for col_idx, key in column_mapping.items():
            out[col_idx - 1] = row.get(key)
        ws.append(out)
    ws.freeze_panes = ws.freeze_panes or f"A{header_row + 1}"
    wb.save(output_path)
    return True


def main() -> None:
    parser = argparse.ArgumentParser(description="Normalize KPI Google Sheets into one table.")
    parser.add_argument("--work-dir", default=".", help="Folder containing source/output files.")
    parser.add_argument("--download", action="store_true", help="Download latest Google Sheets before extracting.")
    args = parser.parse_args()

    work_dir = Path(args.work_dir).resolve()
    dirs = ensure_project_dirs(work_dir)
    sources_config = load_sources_config(work_dir)
    sources = download_sources(work_dir, sources_config) if args.download else local_sources(work_dir, sources_config)
    if not sources:
        raise SystemExit("No source files found. Run with --download or place ACCA.xlsx/CMA.xlsx/IT.xlsx in work-dir.")

    rows = extract_all(sources)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = dirs["staging_output"] / f"normalized_output_{timestamp}.csv"
    xlsx_path = dirs["staging_output"] / f"normalized_output_{timestamp}.xlsx"
    template_path = dirs["template_input"] / TEMPLATE_FILE_NAME
    final_path = dirs["final_output"] / f"von_hoa_output_{timestamp}.xlsx"
    write_csv(rows, csv_path)
    write_xlsx(rows, xlsx_path)
    final_written = write_template_final(rows, template_path, final_path)
    print(f"Done: {len(rows)} rows")
    print(f"CSV : {csv_path}")
    print(f"XLSX: {xlsx_path}")
    if final_written:
        print(f"FINAL CAPITALIZATION: {final_path}")


if __name__ == "__main__":
    main()
