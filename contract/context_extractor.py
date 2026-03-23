"""Context extractor – retrieve neighbouring cell content for LLM naming.

For a given hardcoded vector in an Excel sheet, this module fetches:
  - Row headers (string cells in the same row, columns to the left)
  - Column headers (string cells in the same column, rows above)
  - Adjacent labels (cells directly above, below, left, right of the vector)
  - Sheet name

This context is used by the LLM to infer a meaningful business name for the
hardcoded values.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from contract.excel_scanner import SheetSummary, CellInfo, _col_to_idx, _idx_to_col
from contract.hardcoded_scanner import HardcodedVector

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")

# How many rows above / cols to the left to look for headers
HEADER_ROWS_ABOVE = 5
HEADER_COLS_LEFT = 3


@dataclass
class VectorContext:
    """Context window around a hardcoded vector for LLM consumption."""
    sheet: str
    vector_range: str
    vector_direction: str
    vector_sample: list             # first ≤5 values
    vector_length: int

    # Labels found above the vector (column headers)
    labels_above: list[str]
    # Labels found to the left of the vector (row headers)
    labels_left: list[str]
    # Labels immediately adjacent (1 cell above start, 1 cell to left of start)
    adjacent_labels: list[str]
    # Sheet name context
    sheet_name: str

    def to_prompt_text(self) -> str:
        """Format context as a compact text block for the LLM prompt."""
        lines = [
            f"Sheet: {self.sheet_name}",
            f"Cell range: {self.vector_range} ({self.vector_direction}, {self.vector_length} values)",
            f"Sample values: {self.vector_sample}",
        ]
        if self.labels_above:
            lines.append(f"Labels above: {'; '.join(self.labels_above)}")
        if self.labels_left:
            lines.append(f"Labels to the left: {'; '.join(self.labels_left)}")
        if self.adjacent_labels:
            lines.append(f"Adjacent labels: {'; '.join(self.adjacent_labels)}")
        return "\n".join(lines)


def extract_vector_context(
    vector: HardcodedVector,
    sheet_summary: SheetSummary,
) -> VectorContext:
    """Extract labelling context around *vector* from a SheetSummary.

    Parameters
    ----------
    vector:
        The HardcodedVector to provide context for.
    sheet_summary:
        The pre-scanned SheetSummary for the same sheet.
    """
    # Build a fast lookup: (row, col) → CellInfo
    cell_map: dict[tuple[int, int], CellInfo] = {
        (c.row, c.col): c for c in sheet_summary.cells
    }

    # Parse start/end cells of the vector
    m_start = _CELL_RE.match(vector.start_cell)
    m_end = _CELL_RE.match(vector.end_cell)
    if not m_start or not m_end:
        return _empty_context(vector)

    start_col_letter = m_start.group(1).upper()
    start_row = int(m_start.group(2))
    end_row = int(m_end.group(2)) if vector.direction == "column" else start_row
    start_col = _col_to_idx(start_col_letter)
    end_col = _col_to_idx(m_end.group(1).upper()) if vector.direction == "row" else start_col

    labels_above: list[str] = []
    labels_left: list[str] = []
    adjacent_labels: list[str] = []

    if vector.direction == "column":
        # Look above the start cell for column headers
        for dr in range(1, HEADER_ROWS_ABOVE + 1):
            row_above = start_row - dr
            if row_above < 1:
                break
            ci = cell_map.get((row_above, start_col))
            if ci and ci.cell_type in ("string",) and isinstance(ci.value, str) and ci.value.strip():
                labels_above.append(ci.value.strip())
        # Look to the left for row labels (middle row of vector)
        mid_row = (start_row + end_row) // 2
        for dc in range(1, HEADER_COLS_LEFT + 1):
            col_left = start_col - dc
            if col_left < 1:
                break
            ci = cell_map.get((mid_row, col_left))
            if ci and ci.cell_type in ("string",) and isinstance(ci.value, str) and ci.value.strip():
                labels_left.append(ci.value.strip())
        # One cell directly above and directly left of start
        above = cell_map.get((start_row - 1, start_col))
        if above and isinstance(above.value, str) and above.value.strip():
            adjacent_labels.append(above.value.strip())
        left = cell_map.get((start_row, start_col - 1))
        if left and isinstance(left.value, str) and left.value.strip():
            adjacent_labels.append(left.value.strip())

    else:  # row
        # Look above the row for column headers (iterate across columns)
        for dc in range(end_col - start_col + 1):
            ci_above = cell_map.get((start_row - 1, start_col + dc))
            if ci_above and isinstance(ci_above.value, str) and ci_above.value.strip():
                labels_above.append(ci_above.value.strip())
        # Look further rows above for any header
        for dr in range(2, HEADER_ROWS_ABOVE + 1):
            row_above = start_row - dr
            if row_above < 1:
                break
            for dc in range(min(end_col - start_col + 1, 5)):
                ci = cell_map.get((row_above, start_col + dc))
                if ci and isinstance(ci.value, str) and ci.value.strip():
                    labels_above.append(ci.value.strip())
        # Look to the left of the row
        for dc in range(1, HEADER_COLS_LEFT + 1):
            col_left = start_col - dc
            if col_left < 1:
                break
            ci = cell_map.get((start_row, col_left))
            if ci and isinstance(ci.value, str) and ci.value.strip():
                labels_left.append(ci.value.strip())

    return VectorContext(
        sheet=vector.sheet,
        vector_range=vector.cell_range,
        vector_direction=vector.direction,
        vector_sample=list(vector.values[:5]),
        vector_length=vector.length,
        labels_above=labels_above[:HEADER_ROWS_ABOVE],
        labels_left=labels_left[:HEADER_COLS_LEFT],
        adjacent_labels=adjacent_labels[:4],
        sheet_name=sheet_summary.name,
    )


def _empty_context(vector: HardcodedVector) -> VectorContext:
    return VectorContext(
        sheet=vector.sheet,
        vector_range=vector.cell_range,
        vector_direction=vector.direction,
        vector_sample=list(vector.values[:5]),
        vector_length=vector.length,
        labels_above=[],
        labels_left=[],
        adjacent_labels=[],
        sheet_name=vector.sheet,
    )
