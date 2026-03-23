"""Tests for the ExcelToRequirements system.

Creates a small test workbook in memory and validates each module.
"""
from __future__ import annotations

import io
import os
import sys
import tempfile
from pathlib import Path

import pytest

# Make sure package root is on path
sys.path.insert(0, str(Path(__file__).parent.parent))


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_test_workbook(path: Path) -> None:
    """Create a minimal xlsx workbook for testing."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Revenue"

    # Headers
    ws["A1"] = "Month"
    ws["B1"] = "Sales"
    ws["C1"] = "Costs"
    ws["D1"] = "Profit"

    # Hardcoded values
    for i, val in enumerate([100, 120, 130, 90, 110, 140], start=2):
        ws[f"B{i}"] = val
    for i, val in enumerate([80, 85, 90, 70, 75, 95], start=2):
        ws[f"C{i}"] = val

    # Formula cells
    for i in range(2, 8):
        ws[f"D{i}"] = f"=B{i}-C{i}"

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Total Sales"
    ws2["B1"] = "=SUM(Revenue!B2:B7)"
    ws2["A2"] = "Total Costs"
    ws2["B2"] = "=SUM(Revenue!C2:C7)"
    ws2["A3"] = "Net Profit"
    ws2["B3"] = "=B1-B2"

    wb.save(str(path))


@pytest.fixture(scope="module")
def test_workbook(tmp_path_factory):
    d = tmp_path_factory.mktemp("data")
    p = d / "test_model.xlsx"
    _make_test_workbook(p)
    return p


# ---------------------------------------------------------------------------
# excel_scanner tests
# ---------------------------------------------------------------------------

class TestExcelScanner:
    def test_get_sheet_names(self, test_workbook):
        from contract.excel_scanner import get_sheet_names
        names = get_sheet_names(test_workbook)
        assert "Revenue" in names
        assert "Summary" in names

    def test_scan_workbook_sheets(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook)
        assert len(summaries) == 2
        names = {s.name for s in summaries}
        assert names == {"Revenue", "Summary"}

    def test_scan_formulas(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook, sheet_names=["Revenue"])
        rev = summaries[0]
        assert len(rev.formulas) > 0
        # D2:D7 should all be formulas
        formula_refs = {c.cell_ref for c in rev.formulas}
        assert "D2" in formula_refs

    def test_scan_hardcoded(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook, sheet_names=["Revenue"])
        rev = summaries[0]
        assert len(rev.hardcoded) > 0
        # B2:B7 are hardcoded
        hc_refs = {c.cell_ref for c in rev.hardcoded}
        assert "B2" in hc_refs

    def test_scan_strings(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook, sheet_names=["Revenue"])
        rev = summaries[0]
        str_values = {c.value for c in rev.strings if c.value}
        assert "Month" in str_values or "Sales" in str_values

    def test_filter_by_sheet_name(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook, sheet_names=["Summary"])
        assert len(summaries) == 1
        assert summaries[0].name == "Summary"

    def test_cross_sheet_formula(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        summaries = scan_workbook(test_workbook, sheet_names=["Summary"])
        fmls = summaries[0].formulas
        assert any("Revenue" in (c.formula or "") for c in fmls)

    def test_missing_file_returns_empty(self):
        from contract.excel_scanner import scan_workbook
        result = scan_workbook(Path("/nonexistent/path/file.xlsx"))
        assert result == []


# ---------------------------------------------------------------------------
# formula_analyzer tests
# ---------------------------------------------------------------------------

class TestFormulaAnalyzer:
    def test_cross_sheet_ref(self):
        from contract.formula_analyzer import analyze_formula
        a = analyze_formula("=SUM(Revenue!B2:B7)")
        assert "SUM" in a.functions
        assert any(r.sheet == "Revenue" for r in a.refs)
        assert "aggregation" in a.categories

    def test_arithmetic(self):
        from contract.formula_analyzer import analyze_formula
        a = analyze_formula("=B2-C2")
        assert not a.is_hardcoded_scalar
        assert a.functions == []

    def test_hardcoded_scalar(self):
        from contract.formula_analyzer import analyze_formula
        a = analyze_formula("=42")
        assert a.is_hardcoded_scalar

    def test_external_workbook_ref(self):
        from contract.formula_analyzer import analyze_formula
        a = analyze_formula("='[budget.xlsx]Sheet1'!A1")
        assert any(r.workbook == "budget.xlsx" for r in a.refs)

    def test_extract_source_sheets(self):
        from contract.formula_analyzer import extract_source_sheets
        sheets = extract_source_sheets("=SUM(Revenue!B2:B7)+Summary!B1")
        assert "Revenue" in sheets
        assert "Summary" in sheets

    def test_lookup_category(self):
        from contract.formula_analyzer import analyze_formula
        a = analyze_formula("=VLOOKUP(A1,Sheet2!A:B,2,0)")
        assert "lookup" in a.categories


# ---------------------------------------------------------------------------
# hardcoded_scanner tests
# ---------------------------------------------------------------------------

class TestHardcodedScanner:
    def test_scan_vectors(self, test_workbook):
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        result = scan_hardcoded_vectors(test_workbook)
        assert "Revenue" in result
        vecs = result["Revenue"]
        assert len(vecs) > 0

    def test_vector_direction(self, test_workbook):
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        result = scan_hardcoded_vectors(test_workbook)
        vecs = result.get("Revenue", [])
        # B2:B7 are column-direction
        col_vecs = [v for v in vecs if v.direction == "column"]
        assert len(col_vecs) > 0

    def test_vector_values_match(self, test_workbook):
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        result = scan_hardcoded_vectors(test_workbook)
        vecs = result.get("Revenue", [])
        all_vals = []
        for v in vecs:
            all_vals.extend(v.values)
        assert 100.0 in all_vals or 100 in all_vals

    def test_filter_by_sheet(self, test_workbook):
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        result = scan_hardcoded_vectors(test_workbook, sheet_names=["Summary"])
        # Summary has no hardcoded numbers (only formulas)
        assert "Revenue" not in result


# ---------------------------------------------------------------------------
# upstream_matcher tests
# ---------------------------------------------------------------------------

class TestUpstreamMatcher:
    def test_exact_match(self, tmp_path):
        """Create two workbooks with identical vectors and verify exact match."""
        from openpyxl import Workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.upstream_matcher import find_upstream_sources

        # Model workbook
        model_path = tmp_path / "model.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        for i, v in enumerate([10, 20, 30, 40, 50], start=1):
            ws[f"A{i}"] = v
        wb.save(str(model_path))

        # Upstream workbook (same values)
        upstream_path = tmp_path / "upstream.xlsx"
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "Source"
        for i, v in enumerate([10, 20, 30, 40, 50], start=1):
            ws2[f"B{i}"] = v
        wb2.save(str(upstream_path))

        vecs_by_sheet = scan_hardcoded_vectors(model_path)
        all_vecs = [v for vecs in vecs_by_sheet.values() for v in vecs]
        matches = find_upstream_sources(all_vecs, [upstream_path])

        assert len(matches) > 0
        top = list(matches.values())[0][0]
        assert top.match_type == "exact"
        assert top.similarity == 1.0

    def test_no_upstream_files(self, test_workbook):
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.upstream_matcher import find_upstream_sources
        vecs_by_sheet = scan_hardcoded_vectors(test_workbook)
        all_vecs = [v for vecs in vecs_by_sheet.values() for v in vecs]
        matches = find_upstream_sources(all_vecs, [])
        assert matches == {}


# ---------------------------------------------------------------------------
# context_extractor tests
# ---------------------------------------------------------------------------

class TestContextExtractor:
    def test_column_vector_labels(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.context_extractor import extract_vector_context

        summaries = scan_workbook(test_workbook, sheet_names=["Revenue"])
        vecs_by_sheet = scan_hardcoded_vectors(test_workbook, sheet_names=["Revenue"])
        ss = summaries[0]
        vecs = vecs_by_sheet.get("Revenue", [])

        for vec in vecs:
            ctx = extract_vector_context(vec, ss)
            assert ctx.sheet == "Revenue"
            assert ctx.vector_range == vec.cell_range

    def test_prompt_text_non_empty(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.context_extractor import extract_vector_context

        summaries = scan_workbook(test_workbook, sheet_names=["Revenue"])
        vecs_by_sheet = scan_hardcoded_vectors(test_workbook, sheet_names=["Revenue"])
        ss = summaries[0]
        vecs = vecs_by_sheet.get("Revenue", [])
        assert vecs, "No hardcoded vectors found in Revenue sheet"
        ctx = extract_vector_context(vecs[0], ss)
        text = ctx.to_prompt_text()
        assert "Revenue" in text
        assert len(text) > 10


# ---------------------------------------------------------------------------
# builder tests
# ---------------------------------------------------------------------------

class TestContractBuilder:
    def test_build_contract(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})

        assert contract.model_file == "test_model.xlsx"
        assert len(contract.variables) > 0
        assert len(contract.hardcoded) > 0
        assert len(contract.data_flow) > 0

    def test_write_contract_excel(self, test_workbook, tmp_path):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract, write_contract_excel

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})

        out_path = tmp_path / "contract.xlsx"
        write_contract_excel(contract, out_path)
        assert out_path.exists()
        assert out_path.stat().st_size > 0

        # Verify sheet names
        from openpyxl import load_workbook
        wb = load_workbook(str(out_path))
        assert "Summary" in wb.sheetnames
        assert "Variables" in wb.sheetnames
        assert "Hardcoded Values" in wb.sheetnames


# ---------------------------------------------------------------------------
# mermaid_gen tests
# ---------------------------------------------------------------------------

class TestMermaidGenerator:
    def test_source_to_file(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from mermaid_gen.generator import diagram_source_to_file

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        diagram = diagram_source_to_file(contract)

        assert "flowchart" in diagram
        assert "MODEL" in diagram

    def test_variable_lineage(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from mermaid_gen.generator import diagram_variable_lineage

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        diagram = diagram_variable_lineage(contract)
        assert "flowchart" in diagram

    def test_end_to_end(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from mermaid_gen.generator import diagram_end_to_end

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        diagram = diagram_end_to_end(contract)
        assert "flowchart" in diagram

    def test_write_diagrams(self, test_workbook, tmp_path):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from mermaid_gen.generator import write_diagrams

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        paths = write_diagrams(contract, tmp_path)

        assert "source_to_file" in paths
        assert "variable_lineage" in paths
        assert "end_to_end" in paths
        for p in paths.values():
            assert p.exists()
            content = p.read_text()
            assert "```mermaid" in content


# ---------------------------------------------------------------------------
# refactor tests
# ---------------------------------------------------------------------------

class TestRefactorGenerator:
    def test_generate_python_refactor(self, test_workbook):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from refactor.code_generator import generate_python_refactor

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        code = generate_python_refactor(contract)

        assert "def run_model" in code
        assert "def hardcoded_constants" in code
        assert "return outputs" in code

    def test_generated_code_is_valid_python(self, test_workbook):
        import ast
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from refactor.code_generator import generate_python_refactor

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})
        code = generate_python_refactor(contract)

        # Should parse without SyntaxError
        tree = ast.parse(code)
        assert tree is not None

    def test_write_python_refactor(self, test_workbook, tmp_path):
        from contract.excel_scanner import scan_workbook
        from contract.hardcoded_scanner import scan_hardcoded_vectors
        from contract.builder import build_contract
        from refactor.code_generator import write_python_refactor

        summaries = scan_workbook(test_workbook)
        hc = scan_hardcoded_vectors(test_workbook)
        contract = build_contract(test_workbook, summaries, hc, {})

        out = tmp_path / "refactor.py"
        write_python_refactor(contract, out)
        assert out.exists()
        assert out.stat().st_size > 0


# ---------------------------------------------------------------------------
# CLI integration test
# ---------------------------------------------------------------------------

class TestCLI:
    def test_cli_basic_run(self, test_workbook, tmp_path):
        from generate_contract import main
        ret = main([str(test_workbook), "--out-dir", str(tmp_path)])
        assert ret == 0
        contract_file = tmp_path / f"{test_workbook.stem}_contract.xlsx"
        assert contract_file.exists()

    def test_cli_no_mermaid_no_refactor(self, test_workbook, tmp_path):
        from generate_contract import main
        ret = main([
            str(test_workbook),
            "--out-dir", str(tmp_path),
            "--no-mermaid",
            "--no-refactor",
        ])
        assert ret == 0
        # No mermaid files
        md_files = list(tmp_path.glob("*.md"))
        assert len(md_files) == 0

    def test_cli_missing_file(self, tmp_path):
        from generate_contract import main
        ret = main([str(tmp_path / "nonexistent.xlsx")])
        assert ret == 1
