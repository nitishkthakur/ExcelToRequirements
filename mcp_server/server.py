"""Excel MCP Server.

Exposes Excel inspection and contract-building capabilities as MCP tools.
This allows LLM agents (e.g. Claude Desktop) to directly interrogate Excel
files during interactive contract-generation workflows.

Available tools
---------------
list_sheets(file_path)
    List all sheet names in an Excel workbook.

get_cell(file_path, sheet, cell_ref)
    Get the value and formula of a specific cell.

get_range(file_path, sheet, range_ref)
    Get all cell values in a range (returns 2-D list).

get_sheet_formulas(file_path, sheet)
    Get all formula cells on a sheet (cell_ref, formula, value).

get_hardcoded_vectors(file_path, sheet)
    Get all hardcoded numeric vectors on a sheet.

get_cell_context(file_path, sheet, cell_ref, radius)
    Get the surrounding cells for context (used by LLM for naming).

build_contract(file_path, upstream_dir, out_dir)
    Run the full contract-building pipeline on a model file and write outputs.

Usage
-----
    python -m mcp_server.server --transport stdio
    python -m mcp_server.server --transport sse --port 8765
"""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

log = logging.getLogger("excel_to_req.mcp_server")


# ---------------------------------------------------------------------------
# Tool implementations (no MCP dependency required for the logic itself)
# ---------------------------------------------------------------------------

def _list_sheets(file_path: str) -> list[str]:
    from contract.excel_scanner import get_sheet_names
    return get_sheet_names(Path(file_path))


def _get_cell(file_path: str, sheet: str, cell_ref: str) -> dict:
    from contract.excel_scanner import scan_workbook
    summaries = scan_workbook(Path(file_path), sheet_names=[sheet])
    if not summaries:
        return {"error": f"Sheet '{sheet}' not found"}
    cells = {c.cell_ref: c for c in summaries[0].cells}
    cell = cells.get(cell_ref.upper())
    if not cell:
        return {"error": f"Cell {cell_ref} not found", "cell_ref": cell_ref}
    return {
        "cell_ref": cell.cell_ref,
        "sheet": cell.sheet,
        "value": cell.value,
        "formula": f"={cell.formula}" if cell.formula else None,
        "cell_type": cell.cell_type,
    }


def _get_range(file_path: str, sheet: str, range_ref: str) -> list[list]:
    """Return a 2-D list of values for a range like 'A1:C5'."""
    from contract.excel_scanner import scan_workbook, _col_to_idx
    import re
    summaries = scan_workbook(Path(file_path), sheet_names=[sheet])
    if not summaries:
        return []

    # Parse range
    m = re.match(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", range_ref)
    if not m:
        return []
    c1, r1, c2, r2 = m.group(1).upper(), int(m.group(2)), m.group(3).upper(), int(m.group(4))
    col1 = _col_to_idx(c1)
    col2 = _col_to_idx(c2)

    cell_map: dict[tuple[int, int], Any] = {
        (c.row, c.col): (c.value if c.cell_type != "formula" else f"={c.formula}")
        for c in summaries[0].cells
    }
    grid: list[list] = []
    for row in range(r1, r2 + 1):
        grid.append([cell_map.get((row, col)) for col in range(col1, col2 + 1)])
    return grid


def _get_sheet_formulas(file_path: str, sheet: str) -> list[dict]:
    from contract.excel_scanner import scan_workbook
    summaries = scan_workbook(Path(file_path), sheet_names=[sheet])
    if not summaries:
        return []
    return [
        {
            "cell_ref": c.cell_ref,
            "formula": f"={c.formula}",
            "value": c.value,
            "ext_refs": c.ext_refs,
        }
        for c in summaries[0].formulas
    ]


def _get_hardcoded_vectors(file_path: str, sheet: str) -> list[dict]:
    from contract.hardcoded_scanner import scan_hardcoded_vectors
    vectors_by_sheet = scan_hardcoded_vectors(Path(file_path), sheet_names=[sheet])
    vecs = vectors_by_sheet.get(sheet, [])
    return [
        {
            "cell_range": v.cell_range,
            "direction": v.direction,
            "length": v.length,
            "sample_values": list(v.values[:10]),
        }
        for v in vecs
    ]


def _get_cell_context(file_path: str, sheet: str, cell_ref: str, radius: int = 3) -> dict:
    """Return a grid of cells surrounding *cell_ref* for context."""
    from contract.excel_scanner import scan_workbook, _col_to_idx, _idx_to_col
    import re
    summaries = scan_workbook(Path(file_path), sheet_names=[sheet])
    if not summaries:
        return {"error": f"Sheet '{sheet}' not found"}

    m = re.match(r"([A-Za-z]+)(\d+)", cell_ref.upper())
    if not m:
        return {"error": f"Invalid cell ref: {cell_ref}"}
    col = _col_to_idx(m.group(1))
    row = int(m.group(2))

    cell_map: dict[tuple[int, int], Any] = {
        (c.row, c.col): {"v": c.value, "f": c.formula, "t": c.cell_type}
        for c in summaries[0].cells
    }

    grid: dict[str, Any] = {}
    for dr in range(-radius, radius + 1):
        for dc in range(-radius, radius + 1):
            r2, c2 = row + dr, col + dc
            if r2 < 1 or c2 < 1:
                continue
            ref2 = f"{_idx_to_col(c2)}{r2}"
            cell_data = cell_map.get((r2, c2))
            if cell_data:
                grid[ref2] = cell_data
    return {"center": cell_ref, "radius": radius, "cells": grid}


def _build_contract(
    file_path: str,
    upstream_dir: str = "",
    out_dir: str = "",
    use_llm: bool = False,
) -> dict:
    """Run the full pipeline and return a summary."""
    from contract.excel_scanner import scan_workbook
    from contract.hardcoded_scanner import scan_hardcoded_vectors
    from contract.upstream_matcher import find_upstream_sources
    from contract.builder import build_contract, write_contract_excel
    from mermaid_gen.generator import write_diagrams
    from refactor.code_generator import write_python_refactor

    model_path = Path(file_path)
    out = Path(out_dir) if out_dir else model_path.parent
    out.mkdir(parents=True, exist_ok=True)
    stem = model_path.stem

    # Scan
    sheet_summaries = scan_workbook(model_path)
    hardcoded_by_sheet = scan_hardcoded_vectors(model_path)

    # Upstream matching
    upstream_paths: list[Path] = []
    if upstream_dir:
        upstream_paths = [
            p for p in Path(upstream_dir).iterdir()
            if p.suffix.lower() in (".xlsx", ".xlsm")
        ]

    all_vectors = [v for vecs in hardcoded_by_sheet.values() for v in vecs]
    matches = find_upstream_sources(all_vectors, upstream_paths)

    # LLM namer
    namer = None
    if use_llm:
        from contract.llm_namer import LLMNamer
        namer = LLMNamer()

    # Build contract
    contract = build_contract(
        model_path,
        sheet_summaries,
        hardcoded_by_sheet,
        matches,
        namer=namer,
        upstream_paths=upstream_paths,
    )

    # Write outputs
    contract_path = out / f"{stem}_contract.xlsx"
    write_contract_excel(contract, contract_path)

    diagram_paths = write_diagrams(contract, out, stem=stem)

    refactor_path = out / f"{stem}_refactor.py"
    write_python_refactor(contract, refactor_path)

    return {
        "contract": str(contract_path),
        "diagrams": {k: str(v) for k, v in diagram_paths.items()},
        "refactor": str(refactor_path),
        "summary": contract.summary,
    }


# ---------------------------------------------------------------------------
# MCP server (optional – requires the mcp package)
# ---------------------------------------------------------------------------

def create_mcp_server():
    """Create and configure the MCP server.

    Returns an mcp.Server instance, or raises ImportError if mcp is not installed.
    """
    try:
        from mcp.server import Server
        from mcp.server.models import InitializationOptions
        import mcp.types as types
    except ImportError:
        raise ImportError(
            "The 'mcp' package is required to run the MCP server. "
            "Install it with: pip install mcp"
        )

    server = Server("excel-to-requirements")

    @server.list_tools()
    async def list_tools() -> list[types.Tool]:
        return [
            types.Tool(
                name="list_sheets",
                description="List all sheet names in an Excel workbook.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string", "description": "Path to the Excel file"},
                    },
                    "required": ["file_path"],
                },
            ),
            types.Tool(
                name="get_cell",
                description="Get the value and formula of a specific cell.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                        "sheet": {"type": "string"},
                        "cell_ref": {"type": "string", "description": "e.g. 'B5'"},
                    },
                    "required": ["file_path", "sheet", "cell_ref"],
                },
            ),
            types.Tool(
                name="get_range",
                description="Get all cell values in a range as a 2-D list.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                        "sheet": {"type": "string"},
                        "range_ref": {"type": "string", "description": "e.g. 'A1:C10'"},
                    },
                    "required": ["file_path", "sheet", "range_ref"],
                },
            ),
            types.Tool(
                name="get_sheet_formulas",
                description="Get all formula cells on a sheet.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                        "sheet": {"type": "string"},
                    },
                    "required": ["file_path", "sheet"],
                },
            ),
            types.Tool(
                name="get_hardcoded_vectors",
                description="Get all hardcoded numeric vectors on a sheet.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                        "sheet": {"type": "string"},
                    },
                    "required": ["file_path", "sheet"],
                },
            ),
            types.Tool(
                name="get_cell_context",
                description=(
                    "Get surrounding cells around a target cell. "
                    "Useful for inferring business names from neighbouring labels."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                        "sheet": {"type": "string"},
                        "cell_ref": {"type": "string"},
                        "radius": {
                            "type": "integer",
                            "description": "Rows/cols to include on each side (default 3)",
                            "default": 3,
                        },
                    },
                    "required": ["file_path", "sheet", "cell_ref"],
                },
            ),
            types.Tool(
                name="build_contract",
                description=(
                    "Run the full contract-building pipeline on a model Excel file. "
                    "Writes contract.xlsx, Mermaid diagrams, and a Python refactor script."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string", "description": "Path to model .xlsx"},
                        "upstream_dir": {
                            "type": "string",
                            "description": "Directory containing upstream Excel files (optional)",
                        },
                        "out_dir": {
                            "type": "string",
                            "description": "Output directory (default: same as model file)",
                        },
                        "use_llm": {
                            "type": "boolean",
                            "description": "Whether to use LLM for business-name inference",
                            "default": False,
                        },
                    },
                    "required": ["file_path"],
                },
            ),
        ]

    @server.call_tool()
    async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
        try:
            if name == "list_sheets":
                result = _list_sheets(arguments["file_path"])
            elif name == "get_cell":
                result = _get_cell(
                    arguments["file_path"], arguments["sheet"], arguments["cell_ref"]
                )
            elif name == "get_range":
                result = _get_range(
                    arguments["file_path"], arguments["sheet"], arguments["range_ref"]
                )
            elif name == "get_sheet_formulas":
                result = _get_sheet_formulas(arguments["file_path"], arguments["sheet"])
            elif name == "get_hardcoded_vectors":
                result = _get_hardcoded_vectors(arguments["file_path"], arguments["sheet"])
            elif name == "get_cell_context":
                result = _get_cell_context(
                    arguments["file_path"],
                    arguments["sheet"],
                    arguments["cell_ref"],
                    int(arguments.get("radius", 3)),
                )
            elif name == "build_contract":
                result = _build_contract(
                    arguments["file_path"],
                    upstream_dir=arguments.get("upstream_dir", ""),
                    out_dir=arguments.get("out_dir", ""),
                    use_llm=bool(arguments.get("use_llm", False)),
                )
            else:
                result = {"error": f"Unknown tool: {name}"}
        except Exception as e:
            result = {"error": str(e)}

        return [types.TextContent(type="text", text=json.dumps(result, indent=2))]

    return server


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Excel-to-Requirements MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="MCP transport (default: stdio)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="Port for SSE transport (default: 8765)",
    )
    parser.add_argument(
        "--log-level",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
    )
    args = parser.parse_args()

    logging.basicConfig(level=getattr(logging, args.log_level))

    try:
        server = create_mcp_server()
    except ImportError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    if args.transport == "stdio":
        import asyncio
        from mcp.server.stdio import stdio_server

        async def run_stdio():
            from mcp.server.models import InitializationOptions
            async with stdio_server() as (read_stream, write_stream):
                await server.run(read_stream, write_stream, InitializationOptions(
                    server_name="excel-to-requirements",
                    server_version="1.0.0",
                    capabilities=server.get_capabilities(
                        notification_options=None,
                        experimental_capabilities={},
                    ),
                ))

        asyncio.run(run_stdio())
    else:
        import asyncio
        from mcp.server.sse import SseServerTransport
        from starlette.applications import Starlette
        from starlette.routing import Route, Mount
        import uvicorn

        sse = SseServerTransport("/messages")

        async def handle_sse(request):
            async with sse.connect_sse(request.scope, request.receive, request._send) as streams:
                await server.run(streams[0], streams[1], server.create_initialization_options())

        app = Starlette(routes=[
            Route("/sse", endpoint=handle_sse),
            Mount("/messages", app=sse.handle_post_message),
        ])
        uvicorn.run(app, host="0.0.0.0", port=args.port)


if __name__ == "__main__":
    main()
