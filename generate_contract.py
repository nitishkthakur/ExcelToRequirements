#!/usr/bin/env python3
"""generate_contract.py – main CLI for ExcelToRequirements.

Usage
-----
# Minimal – just scan the model file
python generate_contract.py model.xlsx

# With upstream files for hardcoded-value tracing
python generate_contract.py model.xlsx --upstream-dir ./sources/

# Enable LLM-based business-name inference (requires OPENAI_API_KEY env var)
python generate_contract.py model.xlsx --use-llm

# Specify output directory
python generate_contract.py model.xlsx --out-dir ./outputs/

# Only scan specific sheets
python generate_contract.py model.xlsx --sheets "Revenue" "Costs" "Summary"

# Skip specific outputs
python generate_contract.py model.xlsx --no-mermaid --no-refactor
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="generate_contract",
        description="Generate Business Contract, Mermaid diagrams, and Python refactor from an Excel model.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("file", help="Path to the Excel model file (.xlsx or .xlsm)")
    parser.add_argument(
        "--upstream-dir",
        metavar="DIR",
        help="Directory of upstream Excel files to trace hardcoded values against",
    )
    parser.add_argument(
        "--upstream",
        metavar="FILE",
        nargs="+",
        help="Individual upstream Excel files (alternative to --upstream-dir)",
    )
    parser.add_argument(
        "--out-dir",
        metavar="DIR",
        help="Output directory (default: same directory as the model file)",
    )
    parser.add_argument(
        "--sheets",
        metavar="SHEET",
        nargs="+",
        help="Only process these sheets (by name)",
    )
    parser.add_argument(
        "--use-llm",
        action="store_true",
        help="Use LLM to infer business names (requires OPENAI_API_KEY env var)",
    )
    parser.add_argument(
        "--llm-model",
        default="gpt-4o-mini",
        help="OpenAI model to use for naming (default: gpt-4o-mini)",
    )
    parser.add_argument(
        "--no-mermaid",
        action="store_true",
        help="Skip Mermaid diagram generation",
    )
    parser.add_argument(
        "--no-refactor",
        action="store_true",
        help="Skip Python refactor generation",
    )
    parser.add_argument(
        "--min-vector-len",
        type=int,
        default=3,
        help="Minimum length of a hardcoded vector to include (default: 3)",
    )
    parser.add_argument(
        "--min-similarity",
        type=float,
        default=0.85,
        help="Minimum similarity score for upstream matching (default: 0.85)",
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=3,
        help="Max upstream matches to report per vector (default: 3)",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    args = parser.parse_args(argv)

    # Logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    log = logging.getLogger("generate_contract")

    # Validate model file
    model_path = Path(args.file)
    if not model_path.exists():
        print(f"ERROR: File not found: {model_path}", file=sys.stderr)
        return 1
    if model_path.suffix.lower() not in (".xlsx", ".xlsm"):
        print(f"ERROR: Only .xlsx and .xlsm files are supported (got {model_path.suffix})",
              file=sys.stderr)
        return 1

    # Output directory
    out_dir = Path(args.out_dir) if args.out_dir else model_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = model_path.stem

    # Upstream paths
    upstream_paths: list[Path] = []
    if args.upstream_dir:
        upstream_paths += [
            p for p in Path(args.upstream_dir).iterdir()
            if p.suffix.lower() in (".xlsx", ".xlsm")
        ]
    if args.upstream:
        upstream_paths += [Path(p) for p in args.upstream]
    upstream_paths = [p for p in upstream_paths if p.exists()]

    t_start = time.perf_counter()

    # ── 1. Scan workbook ──────────────────────────────────────────────────
    from contract.excel_scanner import scan_workbook, get_sheet_names
    from contract.hardcoded_scanner import scan_hardcoded_vectors
    from contract.upstream_matcher import find_upstream_sources
    from contract.builder import build_contract, write_contract_excel

    log.info(f"Scanning model: {model_path.name}")
    sheet_names_in_file = get_sheet_names(model_path)
    log.info(f"  Found {len(sheet_names_in_file)} sheet(s)")

    target_sheets = args.sheets if args.sheets else None
    if target_sheets:
        missing = [s for s in target_sheets if s not in sheet_names_in_file]
        if missing:
            log.warning(f"Sheets not found in file: {missing}")
        target_sheets = [s for s in target_sheets if s in sheet_names_in_file]

    t1 = time.perf_counter()
    sheet_summaries = scan_workbook(model_path, sheet_names=target_sheets)
    log.info(f"  Cell scan complete in {time.perf_counter()-t1:.1f}s "
             f"({sum(len(s.cells) for s in sheet_summaries):,} cells)")

    # ── 2. Hardcoded vector detection ────────────────────────────────────
    t1 = time.perf_counter()
    hardcoded_by_sheet = scan_hardcoded_vectors(
        model_path,
        min_len=args.min_vector_len,
        sheet_names=target_sheets,
    )
    total_vecs = sum(len(v) for v in hardcoded_by_sheet.values())
    log.info(f"  Hardcoded vector detection: {total_vecs} vectors in {time.perf_counter()-t1:.1f}s")

    # ── 3. Upstream matching ─────────────────────────────────────────────
    upstream_matches: dict = {}
    if upstream_paths:
        log.info(f"  Matching hardcoded vectors against {len(upstream_paths)} upstream file(s) …")
        t1 = time.perf_counter()
        all_vectors = [v for vecs in hardcoded_by_sheet.values() for v in vecs]
        upstream_matches = find_upstream_sources(
            all_vectors,
            upstream_paths,
            min_similarity=args.min_similarity,
            top_n=args.top_n,
        )
        log.info(f"  Upstream matching: {len(upstream_matches)} matches in {time.perf_counter()-t1:.1f}s")
    else:
        log.info("  No upstream files provided; skipping upstream matching")

    # ── 4. LLM namer ────────────────────────────────────────────────────
    namer = None
    if args.use_llm:
        import os
        if not os.environ.get("OPENAI_API_KEY"):
            log.warning("OPENAI_API_KEY not set; falling back to heuristic naming")
        else:
            from contract.llm_namer import LLMNamer
            namer = LLMNamer(model=args.llm_model)
            log.info(f"  LLM naming enabled (model: {args.llm_model})")

    # ── 5. Build contract ────────────────────────────────────────────────
    log.info("Building Business Contract …")
    contract = build_contract(
        model_path,
        sheet_summaries,
        hardcoded_by_sheet,
        upstream_matches,
        namer=namer,
        upstream_paths=upstream_paths,
    )

    contract_path = out_dir / f"{stem}_contract.xlsx"
    write_contract_excel(contract, contract_path)
    log.info(f"  Contract written → {contract_path}")

    # ── 6. Mermaid diagrams ──────────────────────────────────────────────
    if not args.no_mermaid:
        from mermaid_gen.generator import write_diagrams
        log.info("Generating Mermaid diagrams …")
        diagram_paths = write_diagrams(contract, out_dir, stem=stem)
        for name, p in diagram_paths.items():
            log.info(f"  {name} → {p}")

    # ── 7. Python refactor ───────────────────────────────────────────────
    if not args.no_refactor:
        from refactor.code_generator import write_python_refactor
        refactor_path = out_dir / f"{stem}_refactor.py"
        log.info("Generating Python refactor …")
        write_python_refactor(contract, refactor_path)
        log.info(f"  Refactor → {refactor_path}")

    elapsed = time.perf_counter() - t_start
    print(f"\n✓ Done in {elapsed:.1f}s")
    print(f"  Model file       : {model_path}")
    print(f"  Sheets processed : {len(sheet_summaries)}")
    print(f"  Formula cells    : {contract.summary.get('formula_cells', 0):,}")
    print(f"  Hardcoded vectors: {contract.summary.get('hardcoded_vectors', 0):,}")
    print(f"  External sources : {contract.summary.get('external_sources', 0):,}")
    print(f"  Output directory : {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
