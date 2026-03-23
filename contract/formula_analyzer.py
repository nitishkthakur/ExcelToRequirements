"""Formula analyzer – parses Excel formula strings.

Extracts:
- Cross-sheet references:  Sheet1!A1, 'My Sheet'!B2:B10
- Cross-workbook references: [file.xlsx]Sheet1!A1, =[1]Sheet1!A1
- Named-range references: MyRange, ExternalWorkbook!NamedRange
- All Excel functions used in a formula
- A simplified plain-English description of what the formula computes

The analysis is purely syntactic (no AST evaluation).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# Cross-workbook with literal path:  '[path\to\[wb.xlsx]Sheet'!A1
_WB_PATH_RE = re.compile(
    r"'?(?:[A-Za-z]:\\[^'\[]*|\\\\[^'\[]*|https?://[^'\[]+)?\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv|xls))\]([^'!]*)'?!([A-Za-z$\d:]+)",
    re.IGNORECASE,
)
# Numeric index:  =[1]SheetName!A1
_WB_IDX_RE = re.compile(
    r"\[(\d+)\]([^!]+)!(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)",
    re.IGNORECASE,
)
# Cross-sheet reference without workbook:  Sheet1!A1 or 'My Sheet'!B2
_SHEET_REF_RE = re.compile(
    r"(?<!\[)(?:'([^']+)'|([A-Za-z_\u00C0-\u024F][A-Za-z_0-9\u00C0-\u024F ]*))!(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)",
    re.IGNORECASE,
)
# Excel function names
_FUNC_RE = re.compile(r"\b([A-Z][A-Z0-9_\.]+)\s*\(", re.IGNORECASE)
# Named range in formula (bare word not followed by "(" — likely a defined name)
_NAMED_RANGE_RE = re.compile(r"\b([A-Za-z_\u00C0-\u024F][A-Za-z_0-9\u00C0-\u024F\.]{2,})\b(?!\s*\()")

# Aggregation functions
_AGG_FUNCS = frozenset([
    "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "MAX", "MIN", "LARGE", "SMALL",
    "MEDIAN", "STDEV", "STDEVP", "VAR", "VARP",
    "PRODUCT",
])
_LOOKUP_FUNCS = frozenset([
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
    "XLOOKUP", "XMATCH", "LOOKUP", "OFFSET", "INDIRECT",
    "CHOOSE",
])
_LOGIC_FUNCS = frozenset([
    "IF", "IFS", "IFERROR", "IFNA", "AND", "OR", "NOT",
    "SWITCH",
])
_TEXT_FUNCS = frozenset([
    "CONCATENATE", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID",
    "LEN", "TRIM", "SUBSTITUTE", "REPLACE", "FIND", "SEARCH",
    "TEXT", "VALUE", "UPPER", "LOWER", "PROPER",
    "&",
])
_DATE_FUNCS = frozenset([
    "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY",
    "DATEVALUE", "EOMONTH", "EDATE", "NETWORKDAYS", "WORKDAY",
    "DATEDIF",
])
_FINANCIAL_FUNCS = frozenset([
    "NPV", "IRR", "XIRR", "XNPV", "PMT", "PV", "FV", "RATE",
    "NPER", "CUMIPMT", "CUMPRINC",
])


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class FormulaRef:
    """A single reference found in a formula."""
    workbook: str           # empty string if same workbook
    sheet: str              # empty string if same sheet
    range_ref: str          # e.g. "A1:B10" or "MyRange"
    ref_type: str           # "cell" | "range" | "named"


@dataclass
class FormulaAnalysis:
    """Result of analyzing one formula string."""
    raw: str
    refs: list[FormulaRef] = field(default_factory=list)
    functions: list[str] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)   # aggregation / lookup / logic / etc.
    is_hardcoded_scalar: bool = False    # formula is a bare literal: =42 or ="text"
    description: str = ""               # concise plain-English description


# ---------------------------------------------------------------------------
# Analyzer
# ---------------------------------------------------------------------------

def analyze_formula(formula: str, current_sheet: str = "") -> FormulaAnalysis:
    """Parse a formula string and return a FormulaAnalysis.

    Parameters
    ----------
    formula:
        Formula text, with or without a leading "=".
    current_sheet:
        The sheet containing this formula (used to fill in blank sheet fields).
    """
    raw = formula.lstrip("=")
    result = FormulaAnalysis(raw=formula)

    # Check if it is a bare literal
    stripped = raw.strip()
    if _is_literal(stripped):
        result.is_hardcoded_scalar = True
        result.description = f"Hardcoded value: {stripped}"
        return result

    # --- Extract cross-workbook refs (literal path) ---
    for m in _WB_PATH_RE.finditer(raw):
        wb_file, sheet, rng = m.group(1), m.group(2), m.group(3)
        result.refs.append(FormulaRef(
            workbook=wb_file,
            sheet=sheet.strip("'"),
            range_ref=rng.replace("$", ""),
            ref_type=_range_type(rng),
        ))

    # --- Extract cross-workbook refs (numeric index) ---
    for m in _WB_IDX_RE.finditer(raw):
        idx, sheet, rng = m.group(1), m.group(2), m.group(3)
        result.refs.append(FormulaRef(
            workbook=f"[external#{idx}]",
            sheet=sheet.strip("'"),
            range_ref=rng.replace("$", ""),
            ref_type=_range_type(rng),
        ))

    # --- Extract cross-sheet refs (same workbook) ---
    for m in _SHEET_REF_RE.finditer(raw):
        sheet = (m.group(1) or m.group(2) or "").strip()
        rng = m.group(3).replace("$", "")
        if sheet and sheet.upper() not in {"TRUE", "FALSE"}:
            result.refs.append(FormulaRef(
                workbook="",
                sheet=sheet,
                range_ref=rng,
                ref_type=_range_type(rng),
            ))

    # --- Functions ---
    funcs = [m.group(1).upper() for m in _FUNC_RE.finditer(raw)]
    result.functions = list(dict.fromkeys(funcs))  # deduplicate, preserve order

    categories: list[str] = []
    for f in result.functions:
        if f in _AGG_FUNCS:
            categories.append("aggregation")
        if f in _LOOKUP_FUNCS:
            categories.append("lookup")
        if f in _LOGIC_FUNCS:
            categories.append("logic")
        if f in _TEXT_FUNCS:
            categories.append("text")
        if f in _DATE_FUNCS:
            categories.append("date")
        if f in _FINANCIAL_FUNCS:
            categories.append("financial")
    result.categories = list(dict.fromkeys(categories))

    # --- Plain-English description ---
    result.description = _describe(result)

    return result


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_literal(s: str) -> bool:
    if not s:
        return False
    # Numeric literal
    try:
        float(s)
        return True
    except ValueError:
        pass
    # String literal
    if s.startswith('"') and s.endswith('"') and len(s) > 1:
        return True
    # Boolean
    if s.upper() in ("TRUE", "FALSE"):
        return True
    return False


def _range_type(rng: str) -> str:
    rng = rng.replace("$", "")
    if ":" in rng:
        return "range"
    if re.match(r"^[A-Za-z]+\d+$", rng):
        return "cell"
    return "named"


def _describe(a: FormulaAnalysis) -> str:
    """Build a brief description from function categories and refs."""
    parts: list[str] = []
    if a.functions:
        parts.append(f"Uses {', '.join(a.functions[:3])}")
    if a.categories:
        parts.append(f"[{'/'.join(sorted(set(a.categories)))}]")
    ext_wbs = list(dict.fromkeys(r.workbook for r in a.refs if r.workbook))
    if ext_wbs:
        parts.append(f"References external workbook(s): {', '.join(ext_wbs[:3])}")
    sheets = list(dict.fromkeys(r.sheet for r in a.refs if r.sheet))
    if sheets:
        parts.append(f"Pulls from sheet(s): {', '.join(sheets[:3])}")
    return "; ".join(parts) if parts else "Direct cell reference or arithmetic"


def extract_source_sheets(formula: str) -> list[str]:
    """Return the list of sheet names referenced by a formula (same workbook)."""
    a = analyze_formula(formula)
    return list(dict.fromkeys(r.sheet for r in a.refs if r.sheet and not r.workbook))


def extract_source_workbooks(formula: str) -> list[str]:
    """Return external workbook file names referenced by a formula."""
    a = analyze_formula(formula)
    return list(dict.fromkeys(r.workbook for r in a.refs if r.workbook))
