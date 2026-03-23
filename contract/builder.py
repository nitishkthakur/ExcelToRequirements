"""Business Contract builder.

Assembles all analysis results into a multi-sheet Excel workbook.

Contract sheets
---------------
1. Summary          – top-level overview (file name, sheet count, source types)
2. Variables        – all formula cells: sheet, cell, formula, description,
                      business name, source sheets/workbooks, function categories
3. Hardcoded Values – all hardcoded numeric vectors: sheet, range, values,
                      business name, upstream source match (if found)
4. External Sources – all external connections found (files, ODBC, OLEDB,
                      Power Query, Power Pivot, etc.)
5. Data Flow        – row-level trace: variable → source variable → source sheet
                      → source workbook → earliest traceable source
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

log = logging.getLogger("excel_to_req.contract.builder")


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class VariableRecord:
    sheet: str
    cell_ref: str
    formula: str
    formula_description: str
    business_name: str
    function_categories: list[str]
    source_sheets: list[str]                # same-workbook sheets referenced
    source_workbooks: list[str]             # external workbook files referenced
    value_sample: Any = None


@dataclass
class HardcodedRecord:
    sheet: str
    cell_range: str
    direction: str
    length: int
    sample_values: list
    business_name: str
    upstream_file: str = ""
    upstream_sheet: str = ""
    upstream_range: str = ""
    upstream_match_type: str = ""
    upstream_similarity: float = 0.0


@dataclass
class ExternalSourceRecord:
    source_type: str        # "Excel File" | "ODBC" | "OLEDB" | "Power Query" | "Hardcoded"
    name: str
    details: str
    sheets_referencing: list[str] = field(default_factory=list)


@dataclass
class DataFlowRecord:
    variable_sheet: str
    variable_cell: str
    variable_business_name: str
    source_type: str        # "Formula" | "Hardcoded" | "External"
    source_sheet: str
    source_range: str
    source_workbook: str
    earliest_source: str    # the most upstream traceable source


@dataclass
class BusinessContract:
    model_file: str
    summary: dict
    variables: list[VariableRecord] = field(default_factory=list)
    hardcoded: list[HardcodedRecord] = field(default_factory=list)
    external_sources: list[ExternalSourceRecord] = field(default_factory=list)
    data_flow: list[DataFlowRecord] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

def build_contract(
    model_path: Path,
    sheet_summaries,
    hardcoded_vectors_by_sheet: dict,
    upstream_matches: dict,
    namer: "LLMNamer | None" = None,   # noqa: F821
    upstream_paths: list[Path] | None = None,
) -> BusinessContract:
    """Build a BusinessContract from scanned model data.

    Parameters
    ----------
    model_path:
        Path to the model Excel file.
    sheet_summaries:
        List[SheetSummary] from excel_scanner.scan_workbook().
    hardcoded_vectors_by_sheet:
        dict[sheet_name, list[HardcodedVector]] from hardcoded_scanner.
    upstream_matches:
        dict["sheet|range", list[UpstreamMatch]] from upstream_matcher.
    namer:
        Optional LLMNamer for business-name inference.
    upstream_paths:
        Paths to upstream Excel files (for display in External Sources).
    """
    from contract.formula_analyzer import analyze_formula
    from contract.context_extractor import extract_vector_context
    from contract.excel_scanner import SheetSummary

    contract = BusinessContract(
        model_file=model_path.name,
        summary={
            "model_file": model_path.name,
            "total_sheets": len(sheet_summaries),
        },
    )

    # --- 1. Build variable records ---
    log.info("Building variable records …")
    var_records: list[VariableRecord] = []
    for ss in sheet_summaries:
        for cell in ss.formulas:
            analysis = analyze_formula(cell.formula, current_sheet=ss.name)
            var_records.append(VariableRecord(
                sheet=ss.name,
                cell_ref=cell.cell_ref,
                formula=cell.formula,
                formula_description=analysis.description,
                business_name="",           # filled in by LLM below
                function_categories=analysis.categories,
                source_sheets=list({r.sheet for r in analysis.refs if r.sheet and not r.workbook}),
                source_workbooks=list({r.workbook for r in analysis.refs if r.workbook}),
                value_sample=cell.value,
            ))

    # LLM business names for variables (batch)
    if namer and var_records:
        log.info(f"Inferring business names for {len(var_records)} formula cells …")
        descs = [
            {
                "sheet": r.sheet,
                "cell_ref": r.cell_ref,
                "formula": r.formula,
                "formula_description": r.formula_description,
                "value_sample": r.value_sample,
            }
            for r in var_records
        ]
        names = namer.name_formulas(descs)
        for r, name in zip(var_records, names):
            r.business_name = name
    else:
        for r in var_records:
            r.business_name = f"{r.sheet} Calculation"

    contract.variables = var_records

    # --- 2. Build hardcoded records ---
    log.info("Building hardcoded-value records …")
    # Build a sheet_summary lookup for context extraction
    ss_map = {ss.name: ss for ss in sheet_summaries}

    all_vectors: list = []
    all_contexts: list = []
    for sheet_name, vectors in hardcoded_vectors_by_sheet.items():
        ss = ss_map.get(sheet_name)
        for vec in vectors:
            all_vectors.append(vec)
            if ss:
                ctx = extract_vector_context(vec, ss)
                all_contexts.append(ctx)
            else:
                from contract.context_extractor import _empty_context
                all_contexts.append(_empty_context(vec))

    # LLM business names for hardcoded vectors (batch)
    if namer and all_contexts:
        log.info(f"Inferring business names for {len(all_contexts)} hardcoded vectors …")
        hc_names = namer.name_vectors(all_contexts)
    else:
        hc_names = [
            (ctx.labels_above + ctx.labels_left + ctx.adjacent_labels or [f"{ctx.sheet_name} Value"])[0]
            for ctx in all_contexts
        ]

    hc_records: list[HardcodedRecord] = []
    for vec, name in zip(all_vectors, hc_names):
        key = f"{vec.sheet}|{vec.cell_range}"
        matches = upstream_matches.get(key, [])
        top = matches[0] if matches else None
        hc_records.append(HardcodedRecord(
            sheet=vec.sheet,
            cell_range=vec.cell_range,
            direction=vec.direction,
            length=vec.length,
            sample_values=list(vec.values[:5]),
            business_name=name or f"{vec.sheet} Hardcoded",
            upstream_file=top.upstream_file if top else "",
            upstream_sheet=top.upstream_sheet if top else "",
            upstream_range=top.upstream_range if top else "",
            upstream_match_type=top.match_type if top else "",
            upstream_similarity=round(top.similarity, 4) if top else 0.0,
        ))
    contract.hardcoded = hc_records

    # --- 3. Build external source records ---
    log.info("Building external source records …")
    ext_files: dict[str, set[str]] = {}
    for vr in var_records:
        for wb in vr.source_workbooks:
            ext_files.setdefault(wb, set()).add(vr.sheet)
    for wb, sheets in ext_files.items():
        contract.external_sources.append(ExternalSourceRecord(
            source_type="Excel File",
            name=wb,
            details=f"Referenced by formulas in sheets: {', '.join(sorted(sheets))}",
            sheets_referencing=sorted(sheets),
        ))
    if upstream_paths:
        for p in upstream_paths:
            if p.name not in ext_files:
                contract.external_sources.append(ExternalSourceRecord(
                    source_type="Excel File (upstream)",
                    name=p.name,
                    details=f"Upstream source file for hardcoded value tracing: {str(p)}",
                ))
    # Hardcoded-value upstream sources
    hc_sources: dict[str, set[str]] = {}
    for hr in hc_records:
        if hr.upstream_file:
            hc_sources.setdefault(hr.upstream_file, set()).add(hr.sheet)
    for fname, sheets in hc_sources.items():
        if not any(s.name == fname for s in contract.external_sources):
            contract.external_sources.append(ExternalSourceRecord(
                source_type="Hardcoded Source",
                name=fname,
                details=f"Provides hardcoded values used in sheets: {', '.join(sorted(sheets))}",
                sheets_referencing=sorted(sheets),
            ))

    # --- 4. Build data flow records ---
    log.info("Building data flow records …")
    for vr in var_records:
        earliest = (
            ", ".join(vr.source_workbooks) if vr.source_workbooks
            else ", ".join(vr.source_sheets) if vr.source_sheets
            else vr.sheet
        )
        contract.data_flow.append(DataFlowRecord(
            variable_sheet=vr.sheet,
            variable_cell=vr.cell_ref,
            variable_business_name=vr.business_name,
            source_type="Formula",
            source_sheet=", ".join(vr.source_sheets),
            source_range=vr.cell_ref,
            source_workbook=", ".join(vr.source_workbooks),
            earliest_source=earliest,
        ))
    for hr in hc_records:
        if hr.upstream_file:
            earliest = f"{hr.upstream_file} | {hr.upstream_sheet} | {hr.upstream_range}"
        else:
            earliest = f"{hr.sheet} | {hr.cell_range} (no upstream found)"
        contract.data_flow.append(DataFlowRecord(
            variable_sheet=hr.sheet,
            variable_cell=hr.cell_range,
            variable_business_name=hr.business_name,
            source_type="Hardcoded",
            source_sheet=hr.upstream_sheet or hr.sheet,
            source_range=hr.upstream_range or hr.cell_range,
            source_workbook=hr.upstream_file,
            earliest_source=earliest,
        ))

    # Summary stats
    contract.summary.update({
        "formula_cells": len(var_records),
        "hardcoded_vectors": len(hc_records),
        "external_sources": len(contract.external_sources),
        "data_flow_rows": len(contract.data_flow),
    })
    return contract


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

def write_contract_excel(contract: BusinessContract, out_path: Path) -> None:
    """Write the BusinessContract to a multi-sheet Excel workbook."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise ImportError("openpyxl is required: pip install openpyxl")

    wb = Workbook()
    wb.remove(wb.active)

    # --- Helpers ---
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")

    def add_sheet(name: str, headers: list[str], rows: list[list]) -> None:
        ws = wb.create_sheet(title=name)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(wrap_text=True)
        for row in rows:
            ws.append([str(v) if v is not None else "" for v in row])
        # Auto-width (capped at 60)
        for col_idx, _ in enumerate(headers, 1):
            max_len = max(
                (len(str(ws.cell(r, col_idx).value or "")) for r in range(1, ws.max_row + 1)),
                default=10,
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    # 1. Summary
    add_sheet(
        "Summary",
        ["Property", "Value"],
        [[k, v] for k, v in contract.summary.items()],
    )

    # 2. Variables
    add_sheet(
        "Variables",
        [
            "Sheet", "Cell", "Formula", "Business Name",
            "Formula Description", "Function Categories",
            "Source Sheets (same WB)", "Source Workbooks",
        ],
        [
            [
                r.sheet, r.cell_ref, f"={r.formula}", r.business_name,
                r.formula_description, ", ".join(r.function_categories),
                ", ".join(r.source_sheets), ", ".join(r.source_workbooks),
            ]
            for r in contract.variables
        ],
    )

    # 3. Hardcoded Values
    add_sheet(
        "Hardcoded Values",
        [
            "Sheet", "Cell Range", "Direction", "Length",
            "Sample Values", "Business Name",
            "Upstream File", "Upstream Sheet", "Upstream Range",
            "Match Type", "Similarity",
        ],
        [
            [
                r.sheet, r.cell_range, r.direction, r.length,
                str(r.sample_values), r.business_name,
                r.upstream_file, r.upstream_sheet, r.upstream_range,
                r.upstream_match_type, r.upstream_similarity,
            ]
            for r in contract.hardcoded
        ],
    )

    # 4. External Sources
    add_sheet(
        "External Sources",
        ["Source Type", "Name", "Details", "Sheets Referencing"],
        [
            [r.source_type, r.name, r.details, ", ".join(r.sheets_referencing)]
            for r in contract.external_sources
        ],
    )

    # 5. Data Flow
    add_sheet(
        "Data Flow",
        [
            "Variable Sheet", "Variable Cell", "Business Name",
            "Source Type", "Source Sheet", "Source Range",
            "Source Workbook", "Earliest Traceable Source",
        ],
        [
            [
                r.variable_sheet, r.variable_cell, r.variable_business_name,
                r.source_type, r.source_sheet, r.source_range,
                r.source_workbook, r.earliest_source,
            ]
            for r in contract.data_flow
        ],
    )

    wb.save(str(out_path))
    log.info(f"Contract written to {out_path}")
