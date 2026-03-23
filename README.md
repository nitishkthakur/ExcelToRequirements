# ExcelToRequirements

**Transform any Excel model into a Business Contract, Mermaid lineage diagrams, and a Python calculation engine — automatically.**

---

## What it does

| Artifact | Description |
|---|---|
| **Business Contract** (`*_contract.xlsx`) | Multi-sheet workbook: Summary · Variables · Hardcoded Values · External Sources · Data Flow |
| **Mermaid Diagrams** (`*_diagrams_*.md`) | Three flowcharts: source-to-file topology, variable lineage, full end-to-end |
| **Python Refactor** (`*_refactor.py`) | A self-contained Python module that reproduces the model's calculation engine |

### What the Contract captures

- **Variables** – every formula cell: sheet, cell reference, formula text, formula description, business name (LLM-inferred), function categories, source sheets / external workbooks
- **Hardcoded Values** – every contiguous run of non-formula numeric cells: cell range, sample values, business name (LLM-inferred from neighbouring labels), upstream source file/sheet/range (via nearest-neighbour matching)
- **External Sources** – external Excel files, ODBC/OLEDB connections, Power Query sources, hardcoded-value upstream files
- **Data Flow** – row-by-row trace: variable → source → earliest traceable origin

---

## Installation

```bash
git clone <this repo>
cd ExcelToRequirements

python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

> **LLM naming** requires `openai` (included in requirements) and the `OPENAI_API_KEY` environment variable.
> **MCP server** requires `mcp` (`pip install mcp`).

---

## Quick start

```bash
# Minimal – scan model and generate all outputs
python generate_contract.py path/to/model.xlsx

# With upstream files for hardcoded-value source tracing
python generate_contract.py model.xlsx --upstream-dir ./source_files/

# Enable LLM-based business-name inference
export OPENAI_API_KEY=sk-...
python generate_contract.py model.xlsx --use-llm

# All options
python generate_contract.py model.xlsx \
  --upstream-dir ./sources/ \
  --out-dir ./outputs/ \
  --use-llm \
  --llm-model gpt-4o-mini \
  --sheets "Revenue" "Costs" "Summary" \
  --min-similarity 0.85 \
  --top-n 3 \
  --verbose
```

All outputs are written to `--out-dir` (default: same directory as the model file).

---

## MCP Server

The MCP server lets LLM agents (e.g. Claude Desktop) interactively interrogate Excel files.

```bash
# stdio transport (for Claude Desktop / MCP clients)
python -m mcp_server.server

# SSE transport (HTTP)
python -m mcp_server.server --transport sse --port 8765
```

### Available MCP tools

| Tool | Description |
|---|---|
| `list_sheets` | List all sheet names |
| `get_cell` | Get value + formula for a single cell |
| `get_range` | Get 2-D grid of values for a range |
| `get_sheet_formulas` | All formula cells on a sheet |
| `get_hardcoded_vectors` | All hardcoded numeric vectors on a sheet |
| `get_cell_context` | Surrounding cells (for LLM naming context) |
| `build_contract` | Run the full pipeline and write all outputs |

---

## Architecture

```
generate_contract.py          Main CLI entry point
contract/
  excel_scanner.py            O(n) streaming Excel parser (lxml iterparse)
  formula_analyzer.py         Formula string parser – refs, functions, categories
  hardcoded_scanner.py        Contiguous numeric-run detector (streaming)
  upstream_matcher.py         Nearest-neighbour matching for hardcoded value tracing
  context_extractor.py        Neighbouring-cell context for LLM naming
  llm_namer.py                OpenAI-compatible LLM interface with batching
  builder.py                  Business Contract assembler + Excel writer
mermaid_gen/
  generator.py                Three Mermaid diagram types
refactor/
  code_generator.py           Python calculation-engine code generator
mcp_server/
  server.py                   MCP server (stdio + SSE)
tests/
  test_all.py                 34 unit + integration tests
```

---

## Performance

- Uses **lxml iterparse** throughout – O(n) memory, works on 100 MB / 100-sheet files
- Upstream matching uses **vectorised numpy** (Pearson, cosine) with 2-D sliding-window batches
- Parallel upstream file scanning via `ProcessPoolExecutor`
- LLM calls are **batched** (up to 20 items per request) to minimise latency and cost

---

## LLM memory management

Business names are inferred in **stateless batch calls** – no conversation history is accumulated. Each batch contains up to `BATCH_SIZE=20` items with a compact context block (≤600 chars each). This keeps every request well within token limits regardless of model size.

---

## Running tests

```bash
pip install pytest
python -m pytest tests/ -v
```

---

## Upstream tracing algorithm

Hardcoded value vectors are matched against upstream files using:
1. **Exact hash match** – `tuple(round(v, 8) for v in values)` → O(1) lookup
2. **Vectorised approximate match** – Pearson correlation / cosine similarity via `numpy` sliding-window views; supports length-mismatched vectors (upstream longer than model)

See `contract/upstream_matcher.py` for details. Algorithm adapted from [ExcelLineageDetector](https://github.com/nitishkthakur/ExcelLineageDetector).
