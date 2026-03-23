"""Fast streaming scanner for hardcoded numeric vectors.

Adapted from ExcelLineageDetector (github.com/nitishkthakur/ExcelLineageDetector).

A "vector" is a contiguous run of hardcoded (non-formula) numeric cells
in a single column or row, of length >= MIN_VECTOR_LEN.
Uses lxml iterparse for O(n) streaming — constant memory regardless of file size.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

MIN_VECTOR_LEN = 3

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_to_idx(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def _idx_to_col(idx: int) -> str:
    result = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(rem + 65) + result
    return result


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class HardcodedVector:
    """A contiguous run of hardcoded numeric cells in one row or column."""
    sheet: str
    cell_range: str         # e.g. "B3:B15"
    direction: str          # "column" or "row"
    length: int
    start_cell: str         # e.g. "B3"
    end_cell: str           # e.g. "B15"
    values: tuple           # all numeric values
    sample_values: list = field(default_factory=list)   # first ≤5 values


# ---------------------------------------------------------------------------
# Streaming cell parser
# ---------------------------------------------------------------------------

def _stream_hardcoded_numerics(data: bytes) -> list[tuple[int, int, float]]:
    """Stream-parse sheet XML and return (row, col, value) for non-formula numeric cells."""
    results: list[tuple[int, int, float]] = []
    in_cell = False
    has_formula = False
    cell_ref = ""
    cell_type = "n"
    pending_value: float | None = None

    try:
        context = etree.iterparse(
            io.BytesIO(data),
            events=("start", "end"),
            recover=True,
            no_network=True,
        )
        for event, elem in context:
            ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

            if event == "start":
                if ltag == "c":
                    in_cell = True
                    has_formula = False
                    pending_value = None
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "n")
                elif ltag == "f" and in_cell:
                    has_formula = True
            else:
                if ltag == "v" and in_cell and not has_formula:
                    if cell_type not in ("s", "b", "e", "str") and elem.text:
                        try:
                            pending_value = float(elem.text)
                        except ValueError:
                            pass
                elif ltag == "c":
                    if in_cell and not has_formula and pending_value is not None and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            results.append((int(m.group(2)), _col_to_idx(m.group(1)), pending_value))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()
        del context
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# Run detection
# ---------------------------------------------------------------------------

def _runs(
    items: list[tuple[int, float]], min_len: int
) -> list[tuple[int, int, list[float]]]:
    if not items:
        return []
    runs = []
    run_start = items[0][0]
    run_vals = [items[0][1]]
    for i in range(1, len(items)):
        idx, val = items[i]
        if idx == items[i - 1][0] + 1:
            run_vals.append(val)
        else:
            if len(run_vals) >= min_len:
                runs.append((run_start, items[i - 1][0], run_vals[:]))
            run_start = idx
            run_vals = [val]
    if len(run_vals) >= min_len:
        runs.append((run_start, items[-1][0], run_vals[:]))
    return runs


def _find_vectors(
    cells: list[tuple[int, int, float]],
    sheet: str,
    min_len: int = MIN_VECTOR_LEN,
) -> list[HardcodedVector]:
    vectors: list[HardcodedVector] = []

    by_col: dict[int, list[tuple[int, float]]] = {}
    for row, col, val in cells:
        by_col.setdefault(col, []).append((row, val))

    for col, items in by_col.items():
        items.sort()
        col_letter = _idx_to_col(col)
        for start_r, end_r, vals in _runs(items, min_len):
            vectors.append(HardcodedVector(
                sheet=sheet,
                cell_range=f"{col_letter}{start_r}:{col_letter}{end_r}",
                direction="column",
                length=len(vals),
                start_cell=f"{col_letter}{start_r}",
                end_cell=f"{col_letter}{end_r}",
                values=tuple(vals),
                sample_values=vals[:5],
            ))

    by_row: dict[int, list[tuple[int, float]]] = {}
    for row, col, val in cells:
        by_row.setdefault(row, []).append((col, val))

    for row, items in by_row.items():
        items.sort()
        for start_c, end_c, vals in _runs(items, min_len):
            vectors.append(HardcodedVector(
                sheet=sheet,
                cell_range=f"{_idx_to_col(start_c)}{row}:{_idx_to_col(end_c)}{row}",
                direction="row",
                length=len(vals),
                start_cell=f"{_idx_to_col(start_c)}{row}",
                end_cell=f"{_idx_to_col(end_c)}{row}",
                values=tuple(vals),
                sample_values=vals[:5],
            ))

    return vectors


# ---------------------------------------------------------------------------
# Sheet map (shared with excel_scanner but kept self-contained for import independence)
# ---------------------------------------------------------------------------

def _get_sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    sheet_map: dict[str, str] = {}
    try:
        wb_root = etree.fromstring(zf.read("xl/workbook.xml"))
        rels: dict[str, str] = {}
        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path in zf.namelist():
            for rel in etree.fromstring(zf.read(rels_path)).iter():
                r_id = rel.get("Id") or rel.get("id", "")
                target = rel.get("Target", "")
                if r_id and target:
                    rels[r_id] = target
        sheets = (
            wb_root.findall(f".//{{{NS}}}sheet")
            or wb_root.findall(".//sheet")
            or wb_root.findall(".//{*}sheet")
        )
        for sh in sheets:
            name = sh.get("name", "")
            rid = (
                sh.get(f"{{{REL_NS}}}id")
                or sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            )
            target = rels.get(rid, "")
            if target:
                # Normalise: strip leading "/" and ensure "xl/" prefix
                target = target.lstrip("/")
                if not target.startswith("xl/"):
                    target = f"xl/{target}"
                sheet_map[name] = target
    except Exception:
        pass
    if not sheet_map:
        for n in zf.namelist():
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                m = re.search(r"sheet(\d+)\.xml", n)
                sheet_map[f"Sheet{m.group(1)}"] = n
    return sheet_map


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scan_hardcoded_vectors(
    path: Path,
    min_len: int = MIN_VECTOR_LEN,
    sheet_names: list[str] | None = None,
) -> dict[str, list[HardcodedVector]]:
    """Scan an Excel file and return hardcoded numeric vectors per sheet.

    Parameters
    ----------
    path:
        Path to .xlsx or .xlsm file.
    min_len:
        Minimum run length to qualify as a vector (default 3).
    sheet_names:
        If provided, only scan these sheets.

    Returns
    -------
    dict mapping sheet_name → list of HardcodedVector.
    """
    path = Path(path)
    result: dict[str, list[HardcodedVector]] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            for name, zip_path in sheet_map.items():
                if sheet_names and name not in sheet_names:
                    continue
                if zip_path not in zf.namelist():
                    continue
                data = zf.read(zip_path)
                cells = _stream_hardcoded_numerics(data)
                vectors = _find_vectors(cells, name, min_len)
                result[name] = vectors
    except zipfile.BadZipFile:
        pass
    return result
