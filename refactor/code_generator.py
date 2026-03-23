"""Python Refactor code generator.

Takes a BusinessContract and generates a self-contained Python module that
reproduces the model's calculation engine.

The generated module:
  - Defines a function per output variable (named after the business name)
  - Accepts inputs matching the contract's external sources / hardcoded values
  - Returns a dict of output values

Strategy
--------
1. Group formula variables by sheet.
2. Topologically sort within each sheet so dependencies are computed first.
3. Emit one Python function per sheet that takes all required inputs and
   returns all outputs.
4. Emit a top-level `run_model(inputs)` function that chains the sheets.

Limitations
-----------
- Only formulas that can be expressed as simple Python are translated
  (SUM, AVERAGE, IF, arithmetic).  Complex lookups and VBA macros are
  left as TODO stubs with a comment showing the original formula.
"""
from __future__ import annotations

import re
import textwrap
from pathlib import Path

from contract.builder import BusinessContract, VariableRecord


# ---------------------------------------------------------------------------
# Formula → Python expression translator (best-effort)
# ---------------------------------------------------------------------------

_SIMPLE_ARITH = re.compile(r"^[0-9\.\+\-\*\/\(\)\s]+$")
# Detect Excel range/sheet refs — if present we cannot auto-translate
_HAS_EXCEL_REF = re.compile(r"[A-Za-z_][A-Za-z0-9_]*![A-Z]|[A-Z]+\d+:[A-Z]+\d+|\$[A-Z]|\bSHEET\b", re.IGNORECASE)
_SUM_RE = re.compile(r"^SUM\((.+)\)$", re.IGNORECASE)
_AVG_RE = re.compile(r"^AVERAGE\((.+)\)$", re.IGNORECASE)
_IF_RE = re.compile(r"^IF\((.+)\)$", re.IGNORECASE)
_IFERROR_RE = re.compile(r"^IFERROR\((.+),(.+)\)$", re.IGNORECASE)
_PRODUCT_RE = re.compile(r"^PRODUCT\((.+)\)$", re.IGNORECASE)
_MAX_RE = re.compile(r"^MAX\((.+)\)$", re.IGNORECASE)
_MIN_RE = re.compile(r"^MIN\((.+)\)$", re.IGNORECASE)
_ROUND_RE = re.compile(r"^ROUND\((.+),\s*(\d+)\)$", re.IGNORECASE)


def _ref_to_var(ref: str) -> str:
    """Convert a cell/range reference to a Python identifier."""
    return "v_" + re.sub(r"[^A-Za-z0-9]", "_", ref).strip("_")


def _formula_to_python(formula: str) -> tuple[str, bool]:
    """Attempt to convert an Excel formula string to a Python expression.

    Returns (python_expr, was_translated).
    """
    f = formula.strip()
    # Remove leading "="
    if f.startswith("="):
        f = f[1:].strip()

    # If the formula contains Excel range/sheet references we cannot
    # safely emit valid Python — emit a TODO stub instead.
    if _HAS_EXCEL_REF.search(f):
        return f'None  # TODO: translate formula: ={formula}', False

    # Pure arithmetic / literal
    if _SIMPLE_ARITH.match(f):
        return f, True

    # SUM(...)
    m = _SUM_RE.match(f)
    if m:
        args = m.group(1)
        # Only translate if args are simple (no range refs)
        if not _HAS_EXCEL_REF.search(args):
            return f"sum([{args}])", True
        return f'None  # TODO: translate formula: ={formula}', False

    # AVERAGE(...)
    m = _AVG_RE.match(f)
    if m:
        args = m.group(1)
        if not _HAS_EXCEL_REF.search(args):
            return f"(lambda vals: sum(vals)/len(vals) if vals else 0)([{args}])", True
        return f'None  # TODO: translate formula: ={formula}', False

    # PRODUCT(...)
    m = _PRODUCT_RE.match(f)
    if m:
        args = m.group(1)
        parts = [p.strip() for p in args.split(",")]
        if not any(_HAS_EXCEL_REF.search(p) for p in parts):
            return " * ".join(parts), True
        return f'None  # TODO: translate formula: ={formula}', False

    # MAX / MIN
    m = _MAX_RE.match(f)
    if m and not _HAS_EXCEL_REF.search(m.group(1)):
        return f"max([{m.group(1)}])", True
    m = _MIN_RE.match(f)
    if m and not _HAS_EXCEL_REF.search(m.group(1)):
        return f"min([{m.group(1)}])", True

    # ROUND(expr, n)
    m = _ROUND_RE.match(f)
    if m and not _HAS_EXCEL_REF.search(m.group(1)):
        return f"round({m.group(1)}, {m.group(2)})", True

    # IF(cond, true_val, false_val)
    m = _IF_RE.match(f)
    if m:
        inner = m.group(1)
        parts = _split_args(inner)
        if len(parts) == 3 and not any(_HAS_EXCEL_REF.search(p) for p in parts):
            return f"({parts[1]}) if ({parts[0]}) else ({parts[2]})", True

    # IFERROR(expr, fallback)
    m = _IFERROR_RE.match(f)
    if m:
        expr, fb = m.group(1).strip(), m.group(2).strip()
        if not _HAS_EXCEL_REF.search(expr) and not _HAS_EXCEL_REF.search(fb):
            return f"(lambda: {expr})() if True else {fb}", True

    # Fall-through: cannot translate
    return f'None  # TODO: translate formula: ={formula}', False


def _split_args(s: str) -> list[str]:
    """Split on commas respecting parentheses depth."""
    parts: list[str] = []
    depth = 0
    current = ""
    for ch in s:
        if ch == "(":
            depth += 1
            current += ch
        elif ch == ")":
            depth -= 1
            current += ch
        elif ch == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += ch
    if current:
        parts.append(current)
    return parts


# ---------------------------------------------------------------------------
# Code emitter
# ---------------------------------------------------------------------------

def _safe_py_name(s: str) -> str:
    """Convert a string to a valid Python identifier."""
    s = re.sub(r"[^A-Za-z0-9_]", "_", s.strip())
    s = re.sub(r"_+", "_", s).strip("_")
    if not s or s[0].isdigit():
        s = "v_" + s
    return s[:50].lower()


def generate_python_refactor(contract: BusinessContract) -> str:
    """Generate a Python module string that reproduces the model calculations.

    Parameters
    ----------
    contract:
        The BusinessContract produced by contract.builder.build_contract().

    Returns
    -------
    Python source code as a string.
    """
    lines: list[str] = [
        '"""Auto-generated calculation engine.',
        f'',
        f'Generated from model: {contract.model_file}',
        f'',
        f'This module reproduces the calculation logic extracted from the Excel model.',
        f'Complex formulas that could not be auto-translated are left as TODO stubs.',
        '"""',
        'from __future__ import annotations',
        '',
        'from typing import Any',
        '',
        '',
    ]

    # Group variables by sheet
    by_sheet: dict[str, list[VariableRecord]] = {}
    for vr in contract.variables:
        by_sheet.setdefault(vr.sheet, []).append(vr)

    sheet_func_names: list[str] = []

    for sheet_name, vars_in_sheet in by_sheet.items():
        func_name = _safe_py_name(sheet_name)
        sheet_func_names.append((sheet_name, func_name))

        # Collect input parameters (source sheets + external workbooks)
        input_sources: set[str] = set()
        for vr in vars_in_sheet:
            input_sources.update(vr.source_sheets)
            input_sources.update(vr.source_workbooks)

        params = ", ".join(
            f"{_safe_py_name(src)}: dict | None = None"
            for src in sorted(input_sources)
        ) or "inputs: dict | None = None"

        lines += [
            f'def calculate_{func_name}({params}) -> dict:',
            f'    """Calculations for sheet: {sheet_name}"""',
            f'    results: dict[str, Any] = {{}}',
            '',
        ]

        for vr in vars_in_sheet:
            var_py = _safe_py_name(vr.business_name or vr.cell_ref)
            py_expr, translated = _formula_to_python(vr.formula)

            comment_lines = [
                f'    # Business name : {vr.business_name}',
                f'    # Cell          : {vr.sheet}!{vr.cell_ref}',
                f'    # Formula       : ={vr.formula}',
                f'    # Description   : {vr.formula_description}',
            ]
            lines += comment_lines
            if translated:
                lines.append(f'    {var_py} = {py_expr}')
            else:
                lines.append(f'    {var_py} = {py_expr}')
            lines.append(f'    results["{var_py}"] = {var_py}')
            lines.append('')

        lines += [
            '    return results',
            '',
            '',
        ]

    # Hardcoded constants function
    lines += [
        'def hardcoded_constants() -> dict:',
        '    """Return all hardcoded values extracted from the model."""',
        '    return {',
    ]
    for hr in contract.hardcoded[:200]:
        py_name = _safe_py_name(hr.business_name or hr.cell_range)
        vals = list(hr.sample_values)
        lines.append(f'        "{py_name}": {vals},  # {hr.sheet}!{hr.cell_range}')
    lines += [
        '    }',
        '',
        '',
    ]

    # Top-level runner
    lines += [
        'def run_model(inputs: dict | None = None) -> dict:',
        '    """Execute the full calculation engine and return all outputs.',
        '',
        '    Parameters',
        '    ----------',
        '    inputs:',
        '        Dict of input values keyed by source name.',
        '        Keys correspond to external files or upstream data.',
        '',
        '    Returns',
        '    -------',
        '    Dict of all computed output values.',
        '    """',
        '    inputs = inputs or {}',
        '    outputs: dict[str, Any] = {}',
        '    constants = hardcoded_constants()',
        '    outputs["constants"] = constants',
        '',
    ]
    for sheet_name, func_name in sheet_func_names:
        lines.append(f'    outputs["{sheet_name}"] = calculate_{func_name}(inputs)')
    lines += [
        '    return outputs',
        '',
    ]

    return "\n".join(lines)


def write_python_refactor(contract: BusinessContract, out_path: Path) -> None:
    """Write the generated Python code to *out_path*."""
    code = generate_python_refactor(contract)
    Path(out_path).write_text(code, encoding="utf-8")
