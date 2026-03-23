"""Nearest-neighbour upstream matching for hardcoded value vectors.

Adapted from ExcelLineageDetector (github.com/nitishkthakur/ExcelLineageDetector).

Given a model file's hardcoded vectors, identifies which upstream Excel files
likely contain the same data using exact hash matching and vectorized
approximate matching (Pearson, cosine, or Euclidean similarity).
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

import numpy as np

from contract.hardcoded_scanner import (
    HardcodedVector,
    _stream_hardcoded_numerics,
    _find_vectors,
    _get_sheet_map,
    MIN_VECTOR_LEN,
)

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class UpstreamMatch:
    """A match between a model vector and an upstream vector."""
    model_sheet: str
    model_range: str
    model_direction: str
    model_length: int
    model_sample: list[float]

    match_rank: int
    match_type: str         # "exact" | "approximate"
    similarity: float       # 1.0 for exact, <1.0 for approx

    upstream_file: str
    upstream_sheet: str
    upstream_range: str

    @property
    def upstream_source(self) -> str:
        return f"{self.upstream_file} | {self.upstream_sheet} | {self.upstream_range}"


# ---------------------------------------------------------------------------
# Upstream file scanner (all numerics including formula results)
# ---------------------------------------------------------------------------

def _stream_all_numerics(data: bytes) -> list[tuple[int, int, float]]:
    """Stream-parse sheet XML, returning (row, col, value) for ALL numeric cells."""
    import io
    from lxml import etree

    results: list[tuple[int, int, float]] = []
    in_cell = False
    cell_ref = ""
    cell_type = "n"
    pending_value: float | None = None

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
                    pending_value = None
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "n")
            else:
                if ltag == "v" and in_cell:
                    if cell_type not in ("s", "b", "e", "str") and elem.text:
                        try:
                            pending_value = float(elem.text)
                        except ValueError:
                            pass
                elif ltag == "c":
                    if in_cell and pending_value is not None and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            from contract.hardcoded_scanner import _col_to_idx
                            results.append((int(m.group(2)), _col_to_idx(m.group(1)), pending_value))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()
        del context
    except Exception:
        pass
    return results


def _scan_upstream_file(path: Path, min_len: int) -> list[HardcodedVector]:
    """Scan all numerics (formula + hardcoded) from an upstream file."""
    vectors: list[HardcodedVector] = []
    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            for name, zip_path in sheet_map.items():
                if zip_path not in zf.namelist():
                    continue
                data = zf.read(zip_path)
                cells = _stream_all_numerics(data)
                sheet_vectors = _find_vectors(cells, name, min_len)
                # Attach file name
                for v in sheet_vectors:
                    v.sheet = name
                vectors.extend(sheet_vectors)
    except Exception:
        pass
    return vectors


# ---------------------------------------------------------------------------
# Batch kernels (vectorized)
# ---------------------------------------------------------------------------

def _batch_pearson(model: np.ndarray, batch: np.ndarray) -> np.ndarray:
    if batch.ndim == 1:
        batch = batch.reshape(1, -1)
    if batch.shape[1] < 2:
        return np.zeros(batch.shape[0])
    model_c = model - model.mean()
    batch_c = batch - batch.mean(axis=1, keepdims=True)
    model_ss = np.dot(model_c, model_c)
    batch_ss = np.sum(batch_c ** 2, axis=1)
    if model_ss < 1e-30:
        return np.zeros(batch.shape[0])
    numerator = batch_c @ model_c
    denominator = np.sqrt(batch_ss * model_ss)
    safe = denominator > 1e-15
    result = np.where(safe, numerator / np.where(safe, denominator, 1.0), 0.0)
    return np.clip(result, -1.0, 1.0)


def _batch_cosine(model: np.ndarray, batch: np.ndarray) -> np.ndarray:
    if batch.ndim == 1:
        batch = batch.reshape(1, -1)
    model_norm = np.linalg.norm(model)
    if model_norm < 1e-15:
        return np.zeros(batch.shape[0])
    batch_norms = np.linalg.norm(batch, axis=1)
    dots = batch @ model
    denoms = batch_norms * model_norm
    safe = denoms > 1e-15
    result = np.where(safe, dots / np.where(safe, denoms, 1.0), 0.0)
    return np.clip(result, -1.0, 1.0)


_MAX_BATCH_ELEMENTS = 10_000_000


# ---------------------------------------------------------------------------
# Matcher
# ---------------------------------------------------------------------------

class UpstreamMatcher:
    """Match model hardcoded vectors against upstream file vectors.

    Supports exact (hash-based) and approximate (Pearson/cosine) matching.
    """

    def __init__(
        self,
        min_similarity: float = 0.85,
        top_n: int = 3,
        exact_decimal_places: int = 8,
        metric: str = "pearson",
    ):
        self.min_similarity = min_similarity
        self.top_n = top_n
        self.dp = exact_decimal_places
        self._metric = _batch_pearson if metric != "cosine" else _batch_cosine
        # {rounded_tuple → list of (upstream_file, HardcodedVector)}
        self._exact_index: dict[tuple, list[tuple[str, HardcodedVector]]] = {}
        # grouped by (file, length) for approx matching
        self._approx_by_len: dict[int, list[tuple[str, HardcodedVector, np.ndarray]]] = {}
        self._approx_arrays: dict[int, np.ndarray] = {}
        self._approx_meta: dict[int, list[tuple[str, HardcodedVector]]] = {}

    # ------------------------------------------------------------------ #

    def index_upstream_file(self, path: Path) -> int:
        """Scan and index one upstream file. Returns number of vectors indexed."""
        fname = path.name
        vectors = _scan_upstream_file(path, MIN_VECTOR_LEN)
        for v in vectors:
            key = tuple(round(x, self.dp) for x in v.values)
            self._exact_index.setdefault(key, []).append((fname, v))
            length = v.length
            self._approx_by_len.setdefault(length, []).append((fname, v, np.array(v.values, dtype=np.float64)))

        # Rebuild stacked arrays per length
        for length, items in self._approx_by_len.items():
            self._approx_arrays[length] = np.stack([x[2] for x in items])
            self._approx_meta[length] = [(x[0], x[1]) for x in items]
        return len(vectors)

    # ------------------------------------------------------------------ #

    def match(self, vec: HardcodedVector) -> list[UpstreamMatch]:
        """Find best upstream matches for a model vector."""
        matches: list[UpstreamMatch] = []
        seen: set[str] = set()
        model_key = tuple(round(x, self.dp) for x in vec.values)

        # --- Exact ---
        for (fname, uv) in self._exact_index.get(model_key, []):
            uid = f"{fname}|{uv.sheet}|{uv.cell_range}"
            if uid in seen:
                continue
            seen.add(uid)
            matches.append(UpstreamMatch(
                model_sheet=vec.sheet,
                model_range=vec.cell_range,
                model_direction=vec.direction,
                model_length=vec.length,
                model_sample=list(vec.values[:5]),
                match_rank=len(matches) + 1,
                match_type="exact",
                similarity=1.0,
                upstream_file=fname,
                upstream_sheet=uv.sheet,
                upstream_range=uv.cell_range,
            ))

        # --- Approximate ---
        model_arr = np.array(vec.values, dtype=np.float64)
        model_len = vec.length

        cands: list[tuple[float, str, HardcodedVector]] = []
        for length, arr in self._approx_arrays.items():
            if length < model_len or length > model_len * 2:
                continue
            meta = self._approx_meta[length]
            n = len(meta)
            if length == model_len:
                sims = self._metric(model_arr, arr)
                for i, sim in enumerate(sims):
                    if sim >= self.min_similarity:
                        fname, uv = meta[i]
                        uid = f"{fname}|{uv.sheet}|{uv.cell_range}"
                        if uid not in seen:
                            cands.append((float(sim), fname, uv))
            elif length > model_len:
                n_win = length - model_len + 1
                total_els = n * n_win * model_len
                if total_els <= _MAX_BATCH_ELEMENTS:
                    windows_3d = np.lib.stride_tricks.sliding_window_view(arr, model_len, axis=1)
                    flat = windows_3d.reshape(-1, model_len).copy()
                    sims_flat = self._metric(model_arr, flat)
                    sims_2d = sims_flat.reshape(n, n_win)
                    best_sims = sims_2d.max(axis=1)
                    for i, sim in enumerate(best_sims):
                        if sim >= self.min_similarity:
                            fname, uv = meta[i]
                            uid = f"{fname}|{uv.sheet}|{uv.cell_range}"
                            if uid not in seen:
                                cands.append((float(sim), fname, uv))

        # Sort by similarity desc, take top-N
        cands.sort(key=lambda x: -x[0])
        for sim, fname, uv in cands[:self.top_n]:
            uid = f"{fname}|{uv.sheet}|{uv.cell_range}"
            seen.add(uid)
            matches.append(UpstreamMatch(
                model_sheet=vec.sheet,
                model_range=vec.cell_range,
                model_direction=vec.direction,
                model_length=vec.length,
                model_sample=list(vec.values[:5]),
                match_rank=len(matches) + 1,
                match_type="approximate",
                similarity=sim,
                upstream_file=fname,
                upstream_sheet=uv.sheet,
                upstream_range=uv.cell_range,
            ))

        return matches


# ---------------------------------------------------------------------------
# Convenience function
# ---------------------------------------------------------------------------

def find_upstream_sources(
    model_vectors: list[HardcodedVector],
    upstream_paths: list[Path],
    min_similarity: float = 0.85,
    top_n: int = 3,
) -> dict[str, list[UpstreamMatch]]:
    """Match all model hardcoded vectors against upstream files.

    Returns
    -------
    dict mapping "sheet|range" → list[UpstreamMatch] (top matches).
    """
    matcher = UpstreamMatcher(min_similarity=min_similarity, top_n=top_n)
    for p in upstream_paths:
        if Path(p).suffix.lower() in (".xlsx", ".xlsm"):
            matcher.index_upstream_file(Path(p))

    result: dict[str, list[UpstreamMatch]] = {}
    for vec in model_vectors:
        key = f"{vec.sheet}|{vec.cell_range}"
        matches = matcher.match(vec)
        if matches:
            result[key] = matches
    return result
