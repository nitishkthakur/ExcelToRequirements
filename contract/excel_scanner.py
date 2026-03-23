"""Fast streaming Excel scanner.

Reads every cell (formula, hardcoded value, or string) from an XLSX/XLSM
file using lxml iterparse – O(n) memory regardless of file size.

Returns structured CellInfo records suitable for contract building.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class CellInfo:
    """All information extracted from a single Excel cell."""
    sheet: str
    cell_ref: str       # e.g. "B3"
    row: int
    col: int            # 1-based column index
    col_letter: str     # e.g. "B"
    value: object       # raw numeric or string value (None if formula-only)
    formula: str        # formula text without leading "=" (empty if not a formula)
    cell_type: str      # "formula" | "hardcoded_number" | "string" | "boolean" | "empty"
    # External workbook file names referenced in the formula (may be empty)
    ext_refs: list[str] = field(default_factory=list)


@dataclass
class SheetSummary:
    """Summary of all cells extracted from one sheet."""
    name: str
    cells: list[CellInfo] = field(default_factory=list)

    # Convenience views (populated after cells are set)
    formulas: list[CellInfo] = field(default_factory=list)
    hardcoded: list[CellInfo] = field(default_factory=list)
    strings: list[CellInfo] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Column helpers
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
# Shared strings reader
# ---------------------------------------------------------------------------

def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Parse xl/sharedStrings.xml and return an ordered list of strings."""
    shared: list[str] = []
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return shared
    try:
        root = etree.fromstring(zf.read(path))
        for si in root.iter():
            tag = si.tag.split("}")[-1] if "}" in si.tag else si.tag
            if tag == "si":
                # Collect all <t> text within this <si>
                parts: list[str] = []
                for t_el in si.iter():
                    t_tag = t_el.tag.split("}")[-1] if "}" in t_el.tag else t_el.tag
                    if t_tag == "t" and t_el.text:
                        parts.append(t_el.text)
                shared.append("".join(parts))
    except Exception:
        pass
    return shared


# ---------------------------------------------------------------------------
# Sheet map
# ---------------------------------------------------------------------------

def _get_sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return ordered {sheet_name: zip_path} from workbook.xml."""
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
                or sh.get("r:id", "")
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
    # Fallback
    if not sheet_map:
        for n in zf.namelist():
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                m = re.search(r"sheet(\d+)\.xml", n)
                sheet_map[f"Sheet{m.group(1)}"] = n
    return sheet_map


# ---------------------------------------------------------------------------
# External link map (index → workbook filename)
# ---------------------------------------------------------------------------

_EXT_LINK_RE = re.compile(r"\[(\d+)\]")
_EXT_WB_RE = re.compile(r"\[([^\]]+\.(?:xlsx?|xlsm|xlsb))\]", re.IGNORECASE)
_EXT_PATH_WB_RE = re.compile(
    r"'?(?:[A-Za-z]:\\[^'\[]*|\\\\[^'\[]*|https?://[^'\[]+)?\[([^\]]+\.(?:xlsx?|xlsm|xlsb))\]",
    re.IGNORECASE,
)


def _get_external_link_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return {link_index_str: filename} from xl/externalLinks/."""
    result: dict[str, str] = {}
    for name in zf.namelist():
        m = re.match(r"xl/externalLinks/externalLink(\d+)\.xml", name)
        if not m:
            continue
        idx = m.group(1)
        try:
            root = etree.fromstring(zf.read(name))
            for el in root.iter():
                tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if tag == "externalBook":
                    href = (
                        el.get(f"{{{REL_NS}}}id")
                        or el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                    )
                    # Try to get actual filename from the rels file
                    rels_path = f"xl/externalLinks/_rels/externalLink{idx}.xml.rels"
                    if rels_path in zf.namelist():
                        for rel in etree.fromstring(zf.read(rels_path)).iter():
                            rid = rel.get("Id") or rel.get("id", "")
                            if rid == href:
                                target = rel.get("Target", "")
                                if target:
                                    result[idx] = Path(target).name
                    break
        except Exception:
            pass
    return result


def _extract_ext_refs(formula: str, ext_link_map: dict[str, str]) -> list[str]:
    """Extract external workbook filenames referenced in a formula string."""
    refs: list[str] = []
    # Numeric index references: [1]Sheet!A1
    for m in _EXT_LINK_RE.finditer(formula):
        idx = m.group(1)
        if idx in ext_link_map:
            refs.append(ext_link_map[idx])
    # Literal file references: '[file.xlsx]Sheet'!A1
    for m in _EXT_PATH_WB_RE.finditer(formula):
        refs.append(m.group(1))
    return list(dict.fromkeys(refs))  # deduplicate preserving order


# ---------------------------------------------------------------------------
# Sheet cell streamer
# ---------------------------------------------------------------------------

def _stream_sheet(
    data: bytes,
    sheet_name: str,
    shared_strings: list[str],
    ext_link_map: dict[str, str],
) -> list[CellInfo]:
    """Stream-parse one sheet XML and return CellInfo for every cell."""
    cells: list[CellInfo] = []

    in_cell = False
    in_is = False          # inside <is> (inlineStr)
    has_formula = False
    is_array = False
    cell_ref = ""
    cell_type = "n"
    formula_text = ""
    pending_value: object = None
    inline_parts: list[str] = []

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
                    is_array = False
                    in_is = False
                    formula_text = ""
                    pending_value = None
                    inline_parts = []
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "n")
                elif ltag == "f" and in_cell:
                    has_formula = True
                    is_array = elem.get("t", "") == "array"
                elif ltag == "is" and in_cell and cell_type == "inlineStr":
                    in_is = True

            else:  # "end"
                if ltag == "f" and in_cell:
                    formula_text = elem.text or ""
                elif ltag == "t" and in_cell and in_is:
                    # Inline string text part
                    if elem.text:
                        inline_parts.append(elem.text)
                elif ltag == "is" and in_cell:
                    in_is = False
                    if inline_parts:
                        pending_value = "".join(inline_parts)
                elif ltag == "v" and in_cell:
                    if elem.text:
                        raw = elem.text
                        if cell_type == "s":
                            try:
                                pending_value = shared_strings[int(raw)]
                            except (ValueError, IndexError):
                                pending_value = raw
                        elif cell_type == "b":
                            pending_value = bool(int(raw))
                        elif cell_type in ("e", "str"):
                            pending_value = raw
                        else:
                            try:
                                v = float(raw)
                                pending_value = int(v) if v == int(v) else v
                            except ValueError:
                                pending_value = raw

                elif ltag == "c":
                    if in_cell and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            col_letter = m.group(1).upper()
                            row = int(m.group(2))
                            col = _col_to_idx(col_letter)

                            if has_formula:
                                ctype = "formula"
                                ext_refs = _extract_ext_refs(formula_text, ext_link_map)
                            elif cell_type == "b":
                                ctype = "boolean"
                                ext_refs = []
                            elif cell_type in ("s", "str", "inlineStr"):
                                ctype = "string"
                                ext_refs = []
                            elif cell_type == "e":
                                ctype = "error"
                                ext_refs = []
                            elif pending_value is not None:
                                ctype = "hardcoded_number"
                                ext_refs = []
                            else:
                                ctype = "empty"
                                ext_refs = []

                            cells.append(CellInfo(
                                sheet=sheet_name,
                                cell_ref=cell_ref,
                                row=row,
                                col=col,
                                col_letter=col_letter,
                                value=pending_value,
                                formula=formula_text,
                                cell_type=ctype,
                                ext_refs=ext_refs,
                            ))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()

        del context
    except Exception:
        pass

    return cells


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scan_workbook(
    path: Path,
    sheet_names: list[str] | None = None,
) -> list[SheetSummary]:
    """Scan an Excel workbook and return CellInfo for every cell on every sheet.

    Parameters
    ----------
    path:
        Path to an .xlsx or .xlsm file.
    sheet_names:
        If provided, only scan these sheets (by name).  If None, scan all.

    Returns
    -------
    List of SheetSummary, one per scanned sheet.
    """
    path = Path(path)
    results: list[SheetSummary] = []

    try:
        with zipfile.ZipFile(path) as zf:
            shared_strings = _read_shared_strings(zf)
            sheet_map = _get_sheet_map(zf)
            ext_link_map = _get_external_link_map(zf)

            for name, zip_path in sheet_map.items():
                if sheet_names and name not in sheet_names:
                    continue
                if zip_path not in zf.namelist():
                    continue

                data = zf.read(zip_path)
                cells = _stream_sheet(data, name, shared_strings, ext_link_map)

                summary = SheetSummary(name=name, cells=cells)
                summary.formulas = [c for c in cells if c.cell_type == "formula"]
                summary.hardcoded = [c for c in cells if c.cell_type == "hardcoded_number"]
                summary.strings = [c for c in cells if c.cell_type == "string"]
                results.append(summary)

    except (zipfile.BadZipFile, FileNotFoundError, OSError):
        pass

    return results


def get_sheet_names(path: Path) -> list[str]:
    """Return the ordered list of sheet names in the workbook."""
    path = Path(path)
    try:
        with zipfile.ZipFile(path) as zf:
            return list(_get_sheet_map(zf).keys())
    except Exception:
        return []
