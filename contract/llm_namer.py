"""LLM interface for business-name inference.

Uses an OpenAI-compatible API (GPT-4o, GPT-4, etc.) to infer:
  1. Business name / label for a hardcoded value vector
  2. Business description for a calculated variable (formula cell)

Memory management strategy
--------------------------
We use a **sliding context window** approach:
  - A small system prompt is kept permanently.
  - Each inference request is a standalone call (no accumulated history) —
    this avoids unbounded context growth when processing hundreds of sheets.
  - For batching, we send up to BATCH_SIZE items in a single request and
    ask the LLM to return a JSON list with one label per item.  This
    reduces round-trips while keeping each request well within token limits.

If no API key / client is available the module falls back to heuristic names
derived from the context labels.
"""
from __future__ import annotations

import json
import logging
import os
import re
from typing import Any

log = logging.getLogger("excel_to_req.llm_namer")

# Max items to batch in one LLM call
BATCH_SIZE = 20
# Max characters in a single context block sent to LLM
MAX_CONTEXT_CHARS = 600


# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """You are a financial-data analyst assistant.
Your task is to infer concise, professional business names for Excel data items.
When given context about a cell or range (sheet name, adjacent labels, sample values),
return a SHORT business name (2-6 words, title-case) that accurately describes
what the data represents.  If you cannot determine a good name, return "Unknown".
Do NOT include the cell reference or sheet name in the business name.
Return ONLY a JSON array of strings — one per item in the input list.""".strip()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _heuristic_name(labels: list[str], sheet: str) -> str:
    """Derive a name heuristically when no LLM is available."""
    candidates = [l for l in labels if l]
    if candidates:
        return candidates[0][:50].title()
    return f"{sheet} Value"


def _clean_name(raw: str) -> str:
    """Sanitize an LLM-returned name."""
    name = raw.strip().strip('"').strip("'")
    if not name or name.lower() in {"unknown", "n/a", "none", ""}:
        return ""
    # Remove cell-reference-like patterns
    name = re.sub(r"\b[A-Z]{1,3}\d+\b", "", name).strip()
    return name[:80]


# ---------------------------------------------------------------------------
# LLM client (lazy-loaded to avoid hard dependency)
# ---------------------------------------------------------------------------

def _get_client():
    """Return an OpenAI client, or None if openai is not installed."""
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    try:
        from openai import OpenAI
        return OpenAI(api_key=api_key)
    except ImportError:
        log.warning("openai package not installed; falling back to heuristic naming")
        return None


def _call_llm(client: Any, prompts: list[str], model: str = "gpt-4o-mini") -> list[str]:
    """Send a batch of context prompts to the LLM and return names."""
    numbered = "\n\n".join(f"[{i+1}] {p[:MAX_CONTEXT_CHARS]}" for i, p in enumerate(prompts))
    user_msg = (
        f"Below are {len(prompts)} data item(s). "
        "For each, return a concise business name.\n\n"
        + numbered
        + "\n\nRespond with a JSON array of strings — one name per item."
    )

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.1,
            max_tokens=200 + len(prompts) * 30,
        )
        raw = resp.choices[0].message.content or "[]"
        # Extract JSON array from response (may be wrapped in ```json ... ```)
        json_match = re.search(r"\[.*\]", raw, re.DOTALL)
        if json_match:
            names = json.loads(json_match.group())
            if isinstance(names, list):
                return [_clean_name(str(n)) for n in names]
    except Exception as e:
        log.warning(f"LLM call failed: {e}")
    return [""] * len(prompts)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class LLMNamer:
    """Infer business names for Excel data items using an LLM.

    Parameters
    ----------
    model:
        OpenAI model name.  Defaults to "gpt-4o-mini" (cheap & fast).
    batch_size:
        Items per LLM call.
    fallback_heuristic:
        If True, fall back to label-based heuristics when no LLM is available.
    """

    def __init__(
        self,
        model: str = "gpt-4o-mini",
        batch_size: int = BATCH_SIZE,
        fallback_heuristic: bool = True,
    ):
        self.model = model
        self.batch_size = batch_size
        self.fallback_heuristic = fallback_heuristic
        self._client = _get_client()

    def name_vectors(self, contexts: list["VectorContext"]) -> list[str]:  # noqa: F821
        """Infer business names for a list of VectorContext objects.

        Returns a list of names, one per input context.
        """
        if not contexts:
            return []

        prompts = [c.to_prompt_text() for c in contexts]
        fallback_labels = [
            c.labels_above + c.labels_left + c.adjacent_labels for c in contexts
        ]
        fallback_sheets = [c.sheet_name for c in contexts]

        if self._client is not None:
            names: list[str] = []
            for batch_start in range(0, len(prompts), self.batch_size):
                batch = prompts[batch_start: batch_start + self.batch_size]
                batch_names = _call_llm(self._client, batch, self.model)
                names.extend(batch_names)

            # Fill any blanks with heuristic
            result: list[str] = []
            for i, name in enumerate(names):
                if not name and self.fallback_heuristic:
                    name = _heuristic_name(fallback_labels[i], fallback_sheets[i])
                result.append(name or f"{fallback_sheets[i]} Value")
            return result

        # No LLM — use heuristics
        return [
            _heuristic_name(fallback_labels[i], fallback_sheets[i])
            for i in range(len(contexts))
        ]

    def name_formulas(self, descriptions: list[dict]) -> list[str]:
        """Infer business names for formula cells.

        Each item in *descriptions* is a dict with keys:
          - sheet, cell_ref, formula, formula_description, value_sample

        Returns a list of business names, one per item.
        """
        if not descriptions:
            return []

        prompts: list[str] = []
        for d in descriptions:
            lines = [
                f"Sheet: {d.get('sheet', '')}",
                f"Cell: {d.get('cell_ref', '')}",
                f"Formula: ={d.get('formula', '')}",
                f"Formula description: {d.get('formula_description', '')}",
            ]
            if d.get("value_sample") is not None:
                lines.append(f"Value: {d['value_sample']}")
            prompts.append("\n".join(lines)[:MAX_CONTEXT_CHARS])

        if self._client is not None:
            names = []
            for batch_start in range(0, len(prompts), self.batch_size):
                batch = prompts[batch_start: batch_start + self.batch_size]
                batch_names = _call_llm(self._client, batch, self.model)
                names.extend(batch_names)
            return [n or f"{d.get('sheet','')} Calc" for n, d in zip(names, descriptions)]

        return [
            _heuristic_name(
                [d.get("formula_description", ""), d.get("sheet", "")],
                d.get("sheet", ""),
            )
            for d in descriptions
        ]
