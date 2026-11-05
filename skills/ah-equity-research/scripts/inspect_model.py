#!/usr/bin/env python3
"""Read-only Excel inventory and explicit cell extraction; never evaluates formulas.

Dependency: openpyxl >= 3.1. This is not an automatic valuation-model parser.
Mapping JSON example:
{"metadata":{"currency":"CNY","unit":"million"},
 "required_fields":["target_price"],
 "fields":[{"id":"target_price","sheet":"DCF","cell":"B12",
            "required":true,"unit":"CNY/share","period":"FY2027E"}]}
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from zipfile import BadZipFile, ZipFile

LIMITATIONS = [
    "Does not calculate Excel formulas or refresh external links, queries, or macros.",
    "Cached values may be stale; their presence does not establish freshness or correctness.",
    "Cell mappings and units must be verified against the workbook by a human or analyst agent.",
]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_value(value):
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else str(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    # openpyxl represents array formulas and data tables as objects.
    return {"type": type(value).__name__,
            **{key: json_value(item) for key, item in vars(value).items()}} if hasattr(value, "__dict__") else str(value)


def load_mapping(path: Path | None) -> dict:
    if path is None:
        return {"metadata": {}, "required_fields": [], "fields": []}
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise ValueError("Mapping must be a JSON object.")
    fields = value.get("fields", [])
    required = value.get("required_fields", [])
    metadata = value.get("metadata", {})
    if not isinstance(fields, list) or not isinstance(metadata, dict):
        raise ValueError("Mapping fields must be a list and metadata must be an object.")
    if not isinstance(required, list) or not all(isinstance(x, str) and x for x in required):
        raise ValueError("required_fields must be a list of nonempty field IDs.")
    ids = set()
    for field in fields:
        if not isinstance(field, dict):
            raise ValueError("Each mapping field must be an object.")
        if not all(isinstance(field.get(key), str) and field[key] for key in ("id", "sheet", "cell")):
            raise ValueError("Each mapping requires nonempty id, sheet, and cell strings.")
        if field["id"] in ids:
            raise ValueError(f"Duplicate mapping ID: {field['id']}")
        ids.add(field["id"])
        if not re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}", field["cell"]):
            raise ValueError(f"Mapping requires a single A1 cell address: {field['cell']}")
        if "required" in field and not isinstance(field["required"], bool):
            raise ValueError(f"required must be boolean for {field['id']}")
    return {"metadata": metadata, "required_fields": required, "fields": fields}


def inspect_workbook(path: Path, mapping: dict | None = None, max_cells: int = 2_000_000,
                     detail_limit: int = 200) -> dict:
    """Return a JSON-serializable inventory; status error means unsafe for report extraction."""
    from openpyxl import load_workbook
    from openpyxl.utils.cell import coordinate_to_tuple

    path = path.resolve(strict=True)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx and .xlsm workbooks are supported.")
    if max_cells < 1 or detail_limit < 1:
        raise ValueError("max_cells and detail_limit must be positive.")
    mapping = mapping or {"metadata": {}, "required_fields": [], "fields": []}
    before_hash = sha256(path)
    report = {
        "schema_version": 1,
        "file": str(path), "sha256": before_hash, "size_bytes": path.stat().st_size,
        "inspected_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "read_only": True, "cache_freshness": "unverified", "limitations": LIMITATIONS,
        "metadata": mapping.get("metadata", {}), "sheets": [], "defined_names": [],
        "mapped_fields": [], "issues": [], "scan_complete": True,
    }

    def issue(level, code, message, **details):
        report["issues"].append({"level": level, "code": code, "message": message, **details})

    with ZipFile(path) as archive:
        report["external_link_count"] = sum(
            bool(re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", name))
            for name in archive.namelist())
        report["contains_vba"] = "xl/vbaProject.bin" in archive.namelist()
    if report["external_link_count"]:
        issue("warning", "EXTERNAL_LINKS", "External links are present and have not been refreshed.")
    formulas = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    cached = None
    try:
        cached = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        calculation = getattr(formulas, "calculation", None)
        report["calculation_properties"] = {
            key: json_value(getattr(calculation, key, None))
            for key in ("calcId", "calcMode", "fullCalcOnLoad", "forceFullCalc")
        }
        report["defined_names"].extend(
            {"name": name, "scope": "workbook", "definition": entry.attr_text,
             "hidden": entry.hidden}
            for name, entry in formulas.defined_names.items())
        remaining = max_cells
        for ws in formulas.worksheets:
            cs = cached[ws.title]
            sheet = {"name": ws.title, "state": ws.sheet_state,
                     "declared_dimension": ws.calculate_dimension(),
                     "max_row": ws.max_row, "max_column": ws.max_column,
                     "scanned_cells": 0, "nonempty_cells": 0, "formula_count": 0,
                     "missing_formula_cache_count": 0, "error_count": 0,
                     "formula_samples": [], "missing_formula_cache_samples": [],
                     "error_samples": [], "scan_complete": True}
            report["defined_names"].extend(
                {"name": name, "scope": ws.title, "definition": entry.attr_text,
                 "hidden": entry.hidden}
                for name, entry in getattr(ws, "defined_names", {}).items())
            for formula_row, cache_row in zip(ws.iter_rows(), cs.iter_rows()):
                for cell, cached_cell in zip(formula_row, cache_row):
                    if remaining <= 0:
                        sheet["scan_complete"] = False
                        break
                    remaining -= 1
                    sheet["scanned_cells"] += 1
                    if cell.value is not None:
                        sheet["nonempty_cells"] += 1
                    if cell.data_type == "f":
                        sheet["formula_count"] += 1
                        if len(sheet["formula_samples"]) < detail_limit:
                            sheet["formula_samples"].append({"cell": cell.coordinate,
                                "formula": json_value(cell.value), "cached_value": json_value(cached_cell.value)})
                        if cached_cell.value is None:
                            sheet["missing_formula_cache_count"] += 1
                            if len(sheet["missing_formula_cache_samples"]) < detail_limit:
                                sheet["missing_formula_cache_samples"].append(cell.coordinate)
                    if cell.data_type == "e" or cached_cell.data_type == "e":
                        sheet["error_count"] += 1
                        if len(sheet["error_samples"]) < detail_limit:
                            sheet["error_samples"].append({"cell": cell.coordinate,
                                "value": json_value(cached_cell.value if cached_cell.data_type == "e" else cell.value)})
                if not sheet["scan_complete"]:
                    break
            if not sheet["scan_complete"]:
                report["scan_complete"] = False
                issue("warning", "SCAN_LIMIT", "Inventory scan was truncated; counts are lower bounds.", sheet=ws.title)
            if sheet["missing_formula_cache_count"]:
                issue("warning", "MISSING_FORMULA_CACHE", "Some formulas have no cached result.", sheet=ws.title)
            if sheet["error_count"]:
                issue("error", "EXCEL_ERROR", "Excel error values detected.", sheet=ws.title)
            report["sheets"].append(sheet)

        ids = {field["id"] for field in mapping["fields"]}
        for name in mapping["required_fields"]:
            if name not in ids:
                issue("error", "REQUIRED_MAPPING_MISSING", "Required field has no mapping.", field=name)
        for field in mapping["fields"]:
            is_required = field.get("required", True) or field["id"] in mapping["required_fields"]
            level = "error" if is_required else "warning"
            result = {"id": field["id"], "sheet": field["sheet"], "cell": field["cell"],
                      "required": is_required, "status": "unavailable",
                      "metadata": {**mapping["metadata"], **{key: value for key, value in field.items()
                          if key not in {"id", "sheet", "cell", "required"}}}}
            report["mapped_fields"].append(result)
            if field["sheet"] not in formulas.sheetnames:
                issue(level, "MAPPED_SHEET_MISSING", "Mapped sheet does not exist.", field=field["id"])
                continue
            row, column = coordinate_to_tuple(field["cell"].replace("$", ""))
            if row > 1_048_576 or column > 16_384:
                issue("error", "INVALID_CELL", "Cell address exceeds Excel limits.", field=field["id"])
                continue
            cell = formulas[field["sheet"]].cell(row, column)
            cached_cell = cached[field["sheet"]].cell(row, column)
            formula = cell.data_type == "f"
            result.update({"raw_value": json_value(cell.value),
                           "formula": json_value(cell.value) if formula else None,
                           "cached_value": json_value(cached_cell.value),
                           "number_format": cell.number_format,
                           "effective_value": json_value(cached_cell.value if formula else cell.value)})
            if cell.data_type == "e" or cached_cell.data_type == "e":
                issue("error", "MAPPED_EXCEL_ERROR", "Mapped cell has an Excel error.", field=field["id"])
            elif formula and cached_cell.value is None:
                issue(level, "MAPPED_CACHE_MISSING", "Mapped formula has no cached result; recalculate a copy in Excel.", field=field["id"])
            elif cell.value is None:
                issue(level, "MAPPED_CELL_EMPTY", "Mapped cell is empty.", field=field["id"])
            else:
                result["status"] = "cache_unverified" if formula else "read"
    finally:
        formulas.close()
        if cached is not None:
            cached.close()
    report["input_unchanged"] = sha256(path) == before_hash
    if not report["input_unchanged"]:
        issue("error", "INPUT_CHANGED", "Workbook changed during inspection; discard this snapshot and rerun.")
    report["status"] = "error" if any(item["level"] == "error" for item in report["issues"]) else "warning" if report["issues"] else "ok"
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", type=Path, help="Source .xlsx or .xlsm; never modified")
    parser.add_argument("--mapping", type=Path, help="JSON object containing metadata, required_fields, and explicit fields (see model-contract.md)")
    parser.add_argument("--output", type=Path, help="New JSON file; refuses to overwrite any existing file. Default: stdout")
    parser.add_argument("--max-cells", type=int, default=2_000_000, help="Workbook-wide inventory scan cap (default: 2000000)")
    parser.add_argument("--detail-limit", type=int, default=200, help="Maximum detail samples per category per sheet")
    args = parser.parse_args(argv)
    try:
        source = args.workbook.resolve(strict=True)
        if args.output:
            output = args.output.resolve()
            protected = [source] + ([args.mapping.resolve(strict=True)] if args.mapping else [])
            if output in protected or output.exists():
                raise ValueError("Output must be a new path; existing files and inputs will not be overwritten.")
        result = inspect_workbook(source, load_mapping(args.mapping), args.max_cells, args.detail_limit)
        payload = json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            with args.output.open("x", encoding="utf-8") as stream:
                stream.write(payload + "\n")
        else:
            print(payload)
        return 1 if result["status"] == "error" else 0
    except ImportError:
        print("openpyxl >= 3.1 is required. Use an existing environment that contains it, or install it explicitly.", file=sys.stderr)
        return 2
    except (OSError, ValueError, KeyError, TypeError, BadZipFile) as error:
        print(f"Inspection failed: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
