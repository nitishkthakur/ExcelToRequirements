"""Mermaid diagram generator.

Produces three types of flowcharts from a BusinessContract:

Diagram 1 – "Source-to-File"
    Input sources (Excel files / ODBC / hardcoded) → Model file → Output sheets.
    High-level topology only.

Diagram 2 – "Variable Lineage"
    Each formula variable → its source cells/ranges → intermediate sheets
    → external workbooks.  Shows the full transformation chain.

Diagram 3 – "End-to-End"
    Combined view: raw upstream inputs → hardcoded values and formula
    variables → output variables, with all intermediate transformations
    shown.

All diagrams are emitted as Mermaid flowchart syntax (TD = top-down).
"""
from __future__ import annotations

import re
from pathlib import Path

from contract.builder import BusinessContract, VariableRecord, HardcodedRecord


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _node_id(text: str) -> str:
    """Turn arbitrary text into a Mermaid-safe node identifier."""
    safe = re.sub(r"[^A-Za-z0-9_]", "_", text)
    safe = re.sub(r"_+", "_", safe).strip("_")
    return safe[:60] or "node"


def _label(text: str, max_len: int = 40) -> str:
    """Truncate and escape a display label for Mermaid."""
    t = text.replace('"', "'").replace("\n", " ").strip()
    if len(t) > max_len:
        t = t[:max_len - 1] + "…"
    return t


# ---------------------------------------------------------------------------
# Diagram 1 – Source-to-File
# ---------------------------------------------------------------------------

def diagram_source_to_file(contract: BusinessContract) -> str:
    """Return Mermaid code for the high-level source → model → output diagram."""
    lines = ["flowchart TD"]
    lines.append(f'    MODEL["{_label(contract.model_file)}"]')

    added_nodes: set[str] = set()

    def add_source_node(nid: str, display: str, src_type: str) -> None:
        if nid in added_nodes:
            return
        added_nodes.add(nid)
        shape_open, shape_close = {
            "Excel File": ("[(", ")]"),
            "Excel File (upstream)": ("[(", ")]"),
            "Hardcoded Source": ('["', '"(HC)"]'),
            "ODBC": ('["', '" (ODBC)"]'),
            "OLEDB": ('["', '" (OLEDB)"]'),
            "Power Query": ('["', '" (PQ)"]'),
        }.get(src_type, ('["', '"]'))
        lines.append(f"    {nid}{shape_open}{_label(display)}{shape_close}")

    # External sources → model
    for src in contract.external_sources:
        nid = _node_id(f"src_{src.name}")
        add_source_node(nid, src.name, src.source_type)
        lines.append(f"    {nid} --> MODEL")

    # Model → output sheets (unique sheets that have formula outputs)
    output_sheets: set[str] = {v.sheet for v in contract.variables}
    for sheet in sorted(output_sheets)[:20]:   # cap to avoid huge diagrams
        snid = _node_id(f"out_{sheet}")
        lines.append(f'    {snid}["{_label(sheet)}"]')
        lines.append(f"    MODEL --> {snid}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Diagram 2 – Variable Lineage
# ---------------------------------------------------------------------------

def diagram_variable_lineage(contract: BusinessContract) -> str:
    """Return Mermaid code showing formula variable transformation chains."""
    lines = ["flowchart TD"]
    added_edges: set[tuple[str, str]] = set()
    added_nodes: set[str] = set()

    def emit_node(nid: str, label: str, shape: str = "rect") -> None:
        if nid in added_nodes:
            return
        added_nodes.add(nid)
        if shape == "round":
            lines.append(f'    {nid}("{_label(label)}")')
        elif shape == "stadium":
            lines.append(f'    {nid}(["{_label(label)}"])')
        elif shape == "db":
            lines.append(f'    {nid}[("{_label(label)}")]')
        else:
            lines.append(f'    {nid}["{_label(label)}"]')

    def emit_edge(src: str, dst: str, label: str = "") -> None:
        key = (src, dst)
        if key in added_edges:
            return
        added_edges.add(key)
        if label:
            lines.append(f"    {src} -->|{_label(label, 25)}| {dst}")
        else:
            lines.append(f"    {src} --> {dst}")

    # Cap the number of variables shown to keep diagram readable
    vars_to_show = contract.variables[:80]

    for vr in vars_to_show:
        var_nid = _node_id(f"var_{vr.sheet}_{vr.cell_ref}")
        display = vr.business_name or f"{vr.sheet}!{vr.cell_ref}"
        emit_node(var_nid, display, "round")

        # Source sheets in same workbook
        for src_sheet in vr.source_sheets:
            src_nid = _node_id(f"sheet_{src_sheet}")
            emit_node(src_nid, src_sheet, "rect")
            emit_edge(src_nid, var_nid, "feeds")

        # External workbooks
        for wb in vr.source_workbooks:
            wb_nid = _node_id(f"wb_{wb}")
            emit_node(wb_nid, wb, "db")
            emit_edge(wb_nid, var_nid, "external ref")

        # If no source, mark as standalone calculation
        if not vr.source_sheets and not vr.source_workbooks:
            emit_node(var_nid, display, "round")

    # Add hardcoded value nodes that feed formulas (via sheet)
    hc_by_sheet: dict[str, list[HardcodedRecord]] = {}
    for hr in contract.hardcoded:
        hc_by_sheet.setdefault(hr.sheet, []).append(hr)

    for sheet_name, hrs in hc_by_sheet.items():
        sheet_nid = _node_id(f"sheet_{sheet_name}")
        if sheet_nid not in added_nodes:
            continue  # only show if it's actually a source for a variable
        for hr in hrs[:5]:   # show first 5 hardcoded vectors per sheet
            hc_nid = _node_id(f"hc_{hr.sheet}_{hr.cell_range}")
            emit_node(hc_nid, f"HC: {hr.business_name or hr.cell_range}", "stadium")
            emit_edge(hc_nid, sheet_nid, "hardcoded in")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Diagram 3 – End-to-End
# ---------------------------------------------------------------------------

def diagram_end_to_end(contract: BusinessContract) -> str:
    """Return Mermaid code showing the full end-to-end lineage."""
    lines = ["flowchart LR"]
    added_edges: set[tuple[str, str]] = set()
    added_nodes: set[str] = set()

    def emit_node(nid: str, label: str, shape: str = "rect") -> None:
        if nid in added_nodes:
            return
        added_nodes.add(nid)
        if shape == "round":
            lines.append(f'    {nid}("{_label(label)}")')
        elif shape == "stadium":
            lines.append(f'    {nid}(["{_label(label)}"])')
        elif shape == "db":
            lines.append(f'    {nid}[("{_label(label)}")]')
        else:
            lines.append(f'    {nid}["{_label(label)}"]')

    def emit_edge(src: str, dst: str, label: str = "") -> None:
        key = (src, dst)
        if key in added_edges:
            return
        added_edges.add(key)
        if label:
            lines.append(f"    {src} -->|{_label(label, 20)}| {dst}")
        else:
            lines.append(f"    {src} --> {dst}")

    # --- Layer 1: Upstream (raw) sources ---
    for src in contract.external_sources:
        nid = _node_id(f"raw_{src.name}")
        emit_node(nid, f"{src.name}\n({src.source_type})", "db")

    # --- Layer 2: Hardcoded value vectors ---
    for hr in contract.hardcoded[:50]:
        hc_nid = _node_id(f"hc_{hr.sheet}_{hr.cell_range}")
        emit_node(hc_nid, hr.business_name or hr.cell_range, "stadium")
        if hr.upstream_file:
            raw_nid = _node_id(f"raw_{hr.upstream_file}")
            emit_node(raw_nid, hr.upstream_file, "db")
            emit_edge(raw_nid, hc_nid, hr.upstream_match_type or "source")

    # --- Layer 3: Intermediate / formula variables ---
    vars_to_show = contract.variables[:80]
    for vr in vars_to_show:
        var_nid = _node_id(f"var_{vr.sheet}_{vr.cell_ref}")
        emit_node(var_nid, vr.business_name or vr.cell_ref, "round")

        # Connect from source sheets (which hold hardcoded values or intermediate calcs)
        for src_sheet in vr.source_sheets:
            sheet_nid = _node_id(f"sheet_{src_sheet}")
            emit_node(sheet_nid, src_sheet, "rect")
            emit_edge(sheet_nid, var_nid)
            # Hardcoded values in that source sheet feed into it
            for hr in contract.hardcoded:
                if hr.sheet == src_sheet:
                    hc_nid = _node_id(f"hc_{hr.sheet}_{hr.cell_range}")
                    emit_node(hc_nid, hr.business_name or hr.cell_range, "stadium")
                    emit_edge(hc_nid, sheet_nid)

        for wb in vr.source_workbooks:
            wb_nid = _node_id(f"raw_{wb}")
            emit_node(wb_nid, wb, "db")
            emit_edge(wb_nid, var_nid)

    # --- Layer 4: Output variables (no consumers in same model) ---
    # (heuristic: formula cells not referenced by any other formula)
    referenced_cells: set[str] = set()
    for vr in vars_to_show:
        referenced_cells.update(vr.source_sheets)  # rough proxy

    # Mark top-level outputs
    output_vars = [
        vr for vr in vars_to_show
        if not any(vr.sheet in vr2.source_sheets for vr2 in vars_to_show if vr2 is not vr)
    ][:20]
    for vr in output_vars:
        out_nid = _node_id(f"out_{vr.sheet}_{vr.cell_ref}")
        var_nid = _node_id(f"var_{vr.sheet}_{vr.cell_ref}")
        emit_node(out_nid, f"OUTPUT:\n{vr.business_name or vr.cell_ref}", "round")
        emit_edge(var_nid, out_nid)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_all_diagrams(contract: BusinessContract) -> dict[str, str]:
    """Generate all three Mermaid diagrams.

    Returns
    -------
    dict with keys "source_to_file", "variable_lineage", "end_to_end"
    mapping to Mermaid diagram strings.
    """
    return {
        "source_to_file": diagram_source_to_file(contract),
        "variable_lineage": diagram_variable_lineage(contract),
        "end_to_end": diagram_end_to_end(contract),
    }


def write_diagrams(
    contract: BusinessContract,
    out_dir: Path,
    stem: str = "diagrams",
) -> dict[str, Path]:
    """Write all diagrams to Markdown files in *out_dir*.

    Returns
    -------
    dict mapping diagram name → output file path.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    diagrams = generate_all_diagrams(contract)
    paths: dict[str, Path] = {}
    for name, code in diagrams.items():
        p = out_dir / f"{stem}_{name}.md"
        p.write_text(
            f"# Mermaid Diagram: {name.replace('_', ' ').title()}\n\n"
            f"```mermaid\n{code}\n```\n"
        )
        paths[name] = p
    return paths
