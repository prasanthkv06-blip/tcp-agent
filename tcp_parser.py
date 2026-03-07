"""
tcp_parser.py
=============
Handles:
  1. Extracting plain text from a TCP PDF (or plain text file).
  2. Calling an LLM (OpenAI cloud or local Ollama) to parse performance
     warranties from the extracted text and return them as structured JSON.

Dependencies: pdfplumber, openai, python-dotenv
"""

from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any

import requests                  # for local Ollama calls

logger = logging.getLogger(__name__)


# =============================================================================
# 1. PDF / text extraction
# =============================================================================

def extract_text_from_pdf(pdf_path: str | Path) -> str:
    """
    Extract all text from a PDF using pdfplumber (handles tables better than
    PyPDF2).  Falls back to PyPDF2 if pdfplumber is not installed.

    Parameters
    ----------
    pdf_path : str or Path
        Path to the TCP PDF file.

    Returns
    -------
    str
        Combined text of all pages, lightly cleaned.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"TCP PDF not found: {pdf_path}")

    # ── Try pdfplumber first ─────────────────────────────────────────────────
    try:
        import pdfplumber
        pages: list[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        raw = "\n\n".join(pages)
        logger.info("pdfplumber extracted %d characters from %s", len(raw), pdf_path.name)
        return _clean_text(raw)
    except ImportError:
        logger.warning("pdfplumber not installed – falling back to PyPDF2")

    # ── Fall back to PyPDF2 ──────────────────────────────────────────────────
    try:
        import PyPDF2
        pages = []
        with open(pdf_path, "rb") as fh:
            reader = PyPDF2.PdfReader(fh)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        raw = "\n\n".join(pages)
        logger.info("PyPDF2 extracted %d characters from %s", len(raw), pdf_path.name)
        return _clean_text(raw)
    except ImportError as exc:
        raise ImportError(
            "Neither pdfplumber nor PyPDF2 is installed.  "
            "Run:  pip install pdfplumber"
        ) from exc


def extract_text_from_file(file_path: str | Path) -> str:
    """
    Read plain-text TCP document.

    Parameters
    ----------
    file_path : str or Path

    Returns
    -------
    str
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"TCP text file not found: {file_path}")
    raw = file_path.read_text(encoding="utf-8", errors="replace")
    logger.info("Text file loaded: %d characters", len(raw))
    return _clean_text(raw)


def _clean_text(text: str) -> str:
    """Remove non-printable characters and collapse excessive whitespace."""
    text = re.sub(r"[^\x20-\x7E\n\t]", " ", text)   # keep printable ASCII + newline/tab
    text = re.sub(r"[ \t]{3,}", "  ", text)           # collapse horizontal whitespace
    text = re.sub(r"\n{4,}", "\n\n\n", text)          # max 3 consecutive blank lines
    return text.strip()


# =============================================================================
# 2. LLM warranty extraction
# =============================================================================

# ── Few-shot prompt ──────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
You are a specialist maritime lawyer assistant that extracts performance \
warranties from Time Charter Party (TCP) documents.

Your task is to identify ALL warranted performance figures and return them \
as a JSON array.  Each element must have exactly these keys:

  "metric"    – one of: speed, boil_off_rate_pct, nbog_m3,
                  lsmgo_me_cons_mt, lsmgo_total_cons_mt,
                  lsmgo_ae_cons_mt, lng_me_cons_m3, lng_total_cons_m3,
                  total_liquid_me_cons_mt, loading_rate, discharging_rate
  "value"     – numeric value (float)
  "unit"      – string describing the unit (e.g. "knots", "MT/day", "%/day",
                  "m3/day", "MT/hour")
  "condition" – FREE TEXT describing the condition under which this warranty
                  applies (e.g. "laden, good weather, BF <= 4",
                  "ballast voyage", "at sea").
                  Use "all conditions" if no condition is stated.
  "clause"    – optional clause or article reference from the document

Rules:
• Return ONLY the JSON array – no markdown, no explanation.
• If a figure appears multiple times for different conditions (e.g. laden vs
  ballast, or different speed options), include a separate entry for each.
• If NO performance warranties are found, return an empty array: []
• Do NOT invent values.  Only extract what is explicitly stated.
"""

_FEW_SHOT_USER = """\
Clause 24 – Performance Warranties
The vessel shall maintain an average speed of not less than 19.5 knots \
on a consumption of 150 MT per day of HFO for the main engine plus \
10 MT per day of LSMGO for auxiliaries, in good weather and calm seas \
(BF <= 4, Douglas sea state <= 3).

On ballast voyages: 20.0 knots on 140 MT HFO / 9 MT LSMGO per day.

The maximum permitted boil-off rate shall not exceed 0.12 % of total \
cargo volume per day.
"""

_FEW_SHOT_ASSISTANT = """\
[
  {
    "metric":    "speed",
    "value":     19.5,
    "unit":      "knots",
    "condition": "laden, good weather BF<=4, Douglas sea state<=3",
    "clause":    "Clause 24"
  },
  {
    "metric":    "total_liquid_me_cons_mt",
    "value":     150.0,
    "unit":      "MT/day",
    "condition": "laden, main engine HFO, good weather BF<=4, Douglas sea state<=3",
    "clause":    "Clause 24"
  },
  {
    "metric":    "lsmgo_ae_cons_mt",
    "value":     10.0,
    "unit":      "MT/day",
    "condition": "laden, auxiliary engines, good weather BF<=4, Douglas sea state<=3",
    "clause":    "Clause 24"
  },
  {
    "metric":    "speed",
    "value":     20.0,
    "unit":      "knots",
    "condition": "ballast",
    "clause":    "Clause 24"
  },
  {
    "metric":    "total_liquid_me_cons_mt",
    "value":     140.0,
    "unit":      "MT/day",
    "condition": "ballast, main engine HFO",
    "clause":    "Clause 24"
  },
  {
    "metric":    "lsmgo_ae_cons_mt",
    "value":     9.0,
    "unit":      "MT/day",
    "condition": "ballast, auxiliary engines",
    "clause":    "Clause 24"
  },
  {
    "metric":    "boil_off_rate_pct",
    "value":     0.12,
    "unit":      "%/day",
    "condition": "all conditions",
    "clause":    "Clause 24"
  }
]
"""


def _build_messages(tcp_text: str) -> list[dict]:
    """Construct the messages list for the chat completion."""
    return [
        {"role": "system",    "content": _SYSTEM_PROMPT},
        {"role": "user",      "content": _FEW_SHOT_USER},
        {"role": "assistant", "content": _FEW_SHOT_ASSISTANT},
        {"role": "user",      "content": f"Extract ALL warranties from the following TCP text:\n\n{tcp_text}"},
    ]


# =============================================================================
# 2a. Cloud LLM (OpenAI)
# =============================================================================

def _call_openai(tcp_text: str, model: str, temperature: int | float) -> str:
    """Call OpenAI chat completion and return the raw response string."""
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise ImportError("openai package not installed.  Run: pip install openai") from exc

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY not set in environment / .env file")

    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=_build_messages(tcp_text),
        temperature=temperature,
    )
    return response.choices[0].message.content.strip()


# =============================================================================
# 2b. Local LLM (Ollama)
# =============================================================================

def _call_ollama(tcp_text: str, model: str, base_url: str) -> str:
    """Call a locally running Ollama model and return the response string."""
    # Flatten messages into a single prompt string for Ollama
    messages = _build_messages(tcp_text)
    flat_prompt = "\n\n".join(
        f"[{m['role'].upper()}]\n{m['content']}" for m in messages
    )

    payload = {
        "model":  model,
        "prompt": flat_prompt,
        "stream": False,
        "options": {"temperature": 0},
    }
    resp = requests.post(base_url, json=payload, timeout=120)
    resp.raise_for_status()
    return resp.json().get("response", "").strip()


# =============================================================================
# 2c. Public interface
# =============================================================================

def call_llm(
    tcp_text: str,
    *,
    use_local: bool = False,
    model: str | None = None,
    temperature: float = 0,
    local_url: str = "http://localhost:11434/api/generate",
) -> str:
    """
    Unified LLM caller.  Switches between OpenAI and Ollama based on *use_local*.

    Parameters
    ----------
    tcp_text  : Extracted text from the TCP document.
    use_local : If True, use Ollama (local); otherwise use OpenAI.
    model     : Model name override.  Defaults to config.LLM_MODEL or
                "llama3" for Ollama.
    temperature : Sampling temperature (0 = deterministic).
    local_url : Ollama API endpoint.

    Returns
    -------
    str  Raw LLM response text.
    """
    if use_local:
        _model = model or "llama3"
        logger.info("Using local Ollama model: %s", _model)
        return _call_ollama(tcp_text, _model, local_url)
    else:
        from config import LLM_MODEL
        _model = model or LLM_MODEL
        logger.info("Using OpenAI model: %s", _model)
        return _call_openai(tcp_text, _model, temperature)


def parse_warranties_with_llm(
    tcp_text: str,
    **llm_kwargs: Any,
) -> list[dict]:
    """
    Send TCP text to an LLM and return a list of warranty dictionaries.

    Parameters
    ----------
    tcp_text    : Plain text extracted from the TCP document.
    **llm_kwargs: Forwarded to call_llm()  (use_local, model, …).

    Returns
    -------
    list[dict]  Each dict has keys: metric, value, unit, condition, clause.
                Returns [] if the LLM call fails or returns no warranties.
    """
    try:
        raw_response = call_llm(tcp_text, **llm_kwargs)
        logger.debug("Raw LLM response:\n%s", raw_response)
    except Exception as exc:
        logger.error("LLM call failed: %s", exc)
        return []

    warranties = _parse_json_response(raw_response)
    warranties = _validate_warranties(warranties)
    logger.info("Extracted %d warranty entries.", len(warranties))
    return warranties


def _parse_json_response(raw: str) -> list[dict]:
    """
    Safely parse the LLM response as JSON.

    The model may wrap the array in markdown fences; this function strips them.
    """
    # Strip markdown fences if present
    cleaned = re.sub(r"```(?:json)?", "", raw, flags=re.IGNORECASE).strip()
    cleaned = cleaned.strip("`").strip()

    try:
        data = json.loads(cleaned)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            # Model may return {"warranties": [...]}
            for key in ("warranties", "results", "data"):
                if isinstance(data.get(key), list):
                    return data[key]
        logger.warning("Unexpected JSON shape from LLM: %s", type(data))
        return []
    except json.JSONDecodeError as exc:
        logger.error("Failed to parse LLM JSON response: %s\nRaw: %s", exc, raw[:500])
        return []


def _validate_warranties(warranties: list[dict]) -> list[dict]:
    """
    Validate that each warranty has required keys and values are in
    reasonable ranges (from config.WARRANTY_RANGES).
    """
    try:
        from config import WARRANTY_RANGES
    except ImportError:
        WARRANTY_RANGES = {}

    valid: list[dict] = []
    required_keys = {"metric", "value", "unit"}

    for entry in warranties:
        if not required_keys.issubset(entry.keys()):
            logger.warning("Skipping incomplete warranty entry: %s", entry)
            continue

        # Coerce value to float
        try:
            entry["value"] = float(entry["value"])
        except (TypeError, ValueError):
            logger.warning("Non-numeric warranty value – skipping: %s", entry)
            continue

        # Range check
        metric = entry.get("metric", "")
        if metric in WARRANTY_RANGES:
            lo, hi = WARRANTY_RANGES[metric]
            if not (lo <= entry["value"] <= hi):
                logger.warning(
                    "Warranty value %.4f for '%s' is outside expected range [%s, %s]. "
                    "Keeping but flagging.",
                    entry["value"], metric, lo, hi,
                )
                entry["flagged"] = True

        # Fill optional keys with defaults
        entry.setdefault("condition", "all conditions")
        entry.setdefault("clause", "")
        valid.append(entry)

    return valid


# =============================================================================
# 3. Manual fallback
# =============================================================================

def manual_warranty_input() -> list[dict]:
    """
    Interactive CLI fallback for entering warranties manually when the LLM
    is unavailable or returns empty results.

    Returns
    -------
    list[dict]
    """
    print("\n=== Manual Warranty Entry ===")
    print("Enter warranties one at a time.  Leave 'metric' blank to stop.\n")
    warranties: list[dict] = []
    while True:
        metric = input("Metric (e.g. speed / boil_off_rate_pct / lsmgo_me_cons_mt): ").strip()
        if not metric:
            break
        try:
            value = float(input("Value: ").strip())
        except ValueError:
            print("  Invalid number – skipping.")
            continue
        unit      = input("Unit  (e.g. knots / MT/day / %/day): ").strip()
        condition = input("Condition (or press Enter for 'all conditions'): ").strip() or "all conditions"
        clause    = input("Clause reference (optional): ").strip()
        warranties.append({"metric": metric, "value": value, "unit": unit,
                           "condition": condition, "clause": clause})
        print(f"  Added: {metric} = {value} {unit}\n")
    return warranties


def load_manual_warranties_from_file(json_path: str | Path) -> list[dict]:
    """
    Load warranties from a pre-prepared JSON file (useful for testing without
    an LLM API key).

    Parameters
    ----------
    json_path : Path to a JSON file containing a list of warranty dicts.

    Returns
    -------
    list[dict]
    """
    json_path = Path(json_path)
    if not json_path.exists():
        raise FileNotFoundError(f"Warranty JSON file not found: {json_path}")
    with open(json_path, encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, list):
        raise ValueError("Warranty JSON file must contain a top-level array.")
    logger.info("Loaded %d warranties from %s", len(data), json_path.name)
    return _validate_warranties(data)
