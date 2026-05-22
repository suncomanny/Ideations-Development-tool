from __future__ import annotations

import platform
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook


def validate_workbook(path: Path, required_sheets: list[str]) -> list[str]:
    issues: list[str] = []
    if not path.exists():
        return [f"Workbook does not exist: {path}"]

    try:
        workbook = load_workbook(path, read_only=False, data_only=False)
    except Exception as exc:
        return [f"Workbook failed to open with openpyxl: {exc}"]

    for sheet in required_sheets:
        if sheet not in workbook.sheetnames:
            issues.append(f"Missing required sheet: {sheet}")

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and any(token in value for token in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")):
                    issues.append(f"Formula/error text found at {sheet.title}!{cell.coordinate}: {value}")
    workbook.close()
    issues.extend(validate_sheet_views(path))
    return issues


def validate_sheet_views(path: Path) -> list[str]:
    issues: list[str] = []
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    try:
        with ZipFile(path) as archive:
            sheet_files = [
                name for name in archive.namelist()
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
            ]
            for name in sheet_files:
                root = ET.fromstring(archive.read(name))
                for view in root.findall(".//x:sheetView", ns):
                    pane = view.find("x:pane", ns)
                    if pane is None:
                        continue
                    x_split = pane.get("xSplit")
                    y_split = pane.get("ySplit")
                    valid_panes = {"topLeft"}
                    if x_split:
                        valid_panes.add("topRight")
                    if y_split:
                        valid_panes.add("bottomLeft")
                    if x_split and y_split:
                        valid_panes.add("bottomRight")
                    seen: set[tuple[str, str, str]] = set()
                    for selection in view.findall("x:selection", ns):
                        selection_pane = selection.get("pane") or "topLeft"
                        if selection_pane not in valid_panes:
                            issues.append(f"Invalid worksheet view in {name}: selection references missing pane {selection_pane}.")
                        key = (
                            selection_pane,
                            selection.get("activeCell") or "",
                            selection.get("sqref") or "",
                        )
                        if key in seen:
                            issues.append(f"Duplicate worksheet view selection in {name}: {selection_pane} {key[1]} {key[2]}.")
                        seen.add(key)
    except Exception as exc:
        issues.append(f"Workbook XML view validation failed: {exc}")
    return issues


def try_excel_com_open_save(path: Path) -> tuple[bool, str]:
    if platform.system().lower() != "windows":
        return False, "Excel COM validation skipped: this is not Windows."

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        return False, f"Excel COM validation skipped: pywin32 is not available ({exc})."

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(path))
        workbook.Save()
        return True, "Excel COM open/save validation passed."
    except Exception as exc:
        return False, f"Excel COM validation failed: {exc}"
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
