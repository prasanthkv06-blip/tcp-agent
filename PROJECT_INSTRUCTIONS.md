# PROJECT INSTRUCTIONS — Vessel Performance Claim Calculator

> **For AI Agent / Developer Reference**
> Step-by-step task breakdown for building and improving the automated
> vessel performance claim calculation system.

---

## CURRENT STATE (What Already Exists)

The core pipeline is functional and has been tested against Al Sailiya (QET) noon-report data.

### Modules in Place

| Module | Status | What It Does |
|---|---|---|
| `main.py` | Working | CLI orchestrator — 4-step pipeline with argparse |
| `tcp_parser.py` | Working | PDF text extraction (pdfplumber/PyPDF2), LLM warranty parsing (OpenAI + Ollama), manual CLI fallback, JSON file loading |
| `data_extractor.py` | Working | Loads raw Excel, filters by date, validates columns, LLM column-matching fallback |
| `calculator.py` | Working | Derived columns (total liquid ME, boil-off rate, total BOG), daily DataFrame builder, actuals aggregation, text summary |
| `template_filler.py` | Working | Auto-detects header row, maps columns, writes daily rows + Summary sheet, creates sample template |
| `config.py` | Working | Column mappings for Al Sailiya 340-column format, validation ranges, LLM settings |

### Known Issues from First Test Run (2025-11-14 to 2025-11-30)

```
speed                    warranted=19.500  actual=13.073  dev=-33.0%   ← includes port/low-speed days
lsmgo_me_cons_mt         warranted= 3.500  actual= 0.000  dev=-100%   ← vessel uses LNG, not LSMGO for ME
boil_off_rate_pct        warranted= 0.120  actual= 0.203  dev=+69.5%  ← may include port idle BOG
```

These deviations reveal the need for smarter filtering (weather, sea-state, steaming-only, condition splits).

---

## PHASE 1: Data Quality & Filtering Improvements

> **Goal:** Make calculations accurate by filtering out non-representative data points.

### Task 1.1: Steaming-Only Filter for Speed Calculations

**Problem:** Speed averages include port days and slow-steaming, producing misleadingly low values.

**Approach:**
- In `data_extractor.py`, add a function `filter_steaming_only(df)` that:
  - Keeps rows where `Steaming Time (Hrs.)` > 0 (or a configurable threshold like > 6 hours).
  - Alternatively, checks `Report Type` for `NOON` at-sea reports vs port/anchor reports.
- In `calculator.py`, use this filter specifically for speed-related aggregation.

**Config change:** Add to `config.py`:
```python
STEAMING_TIME_THRESHOLD = 6.0  # hours — exclude partial-steaming days from speed calc
```

**Pitfall:** Some noon reports have steaming time = 0 on weather-delay days. Decide whether to include these. Check with the charter party clause — "good weather" days only may be the correct filter.

### Task 1.2: Good Weather Filter (Beaufort Scale)

**Problem:** TCP warranties typically apply "in good weather" (e.g., BF <= 4, swell <= 3m). The current calculator averages all days indiscriminately.

**Approach:**
- In `data_extractor.py`, add `filter_good_weather(df, max_bf=4, max_swell=3.0)`:
  ```python
  def filter_good_weather(df, max_bf=4, max_swell=3.0, col_map=None):
      from config import COLUMN_MAP
      _cm = col_map or COLUMN_MAP
      bf_col = _cm.get("wind_force_bft")
      swell_col = _cm.get("swell_height_m")
      mask = pd.Series(True, index=df.index)
      if bf_col and bf_col in df.columns:
          mask &= pd.to_numeric(df[bf_col], errors="coerce").fillna(0) <= max_bf
      if swell_col and swell_col in df.columns:
          mask &= pd.to_numeric(df[swell_col], errors="coerce").fillna(0) <= max_swell
      return df[mask].copy()
  ```
- Integrate into `calculator.compute_actuals()` — apply the good-weather filter when the warranty condition mentions "good weather" or "BF<=".
- Parse the condition string from the warranty to extract BF threshold.

**Config change:**
```python
GOOD_WEATHER_DEFAULTS = {
    "max_beaufort": 4,
    "max_swell_m": 3.0,
    "max_wind_sea_m": 3.0,
}
```

**Pitfall:** The `Wind Force(Bft.) (T)` column may contain text like "4-5" or be empty. Handle parsing gracefully.

### Task 1.3: Laden / Ballast Split

**Problem:** Speed and fuel warranties differ for laden vs ballast. The pipeline has `--condition` flag but it's basic.

**Approach:**
- The warranty JSON has `"condition": "laden, good weather BF<=4"` and `"condition": "ballast, ..."`.
- In `calculator.compute_actuals()`, for each warranty entry, parse the condition to detect "laden" or "ballast".
- Use `data_extractor.split_by_condition()` (already exists) to get per-condition DataFrames.
- Match each warranty against the right subset.

**Implementation detail:**
```python
def _extract_condition_type(condition_str: str) -> str | None:
    """Return 'laden', 'ballast', or None from a warranty condition string."""
    lower = condition_str.lower()
    if "laden" in lower:
        return "laden"
    if "ballast" in lower:
        return "ballast"
    return None
```

### Task 1.4: Exclude Port / Anchor Days

**Problem:** Fuel consumption during port operations shouldn't count toward at-sea warranty figures.

**Approach:**
- Check for a "Vessel status" or "Port/Sea" column in the raw data.
- Add `AT_SEA_STATUSES` to `config.py` (e.g., `["AT SEA", "STEAMING"]`).
- Filter in `data_extractor.py` before computing at-sea metrics.

---

## PHASE 2: Calculation Engine Enhancements

> **Goal:** Support all common TCP warranty types with correct formulas.

### Task 2.1: Loading / Discharging Rate Calculations

**Problem:** These metrics aren't computed yet. They require identifying port periods and cargo quantities.

**Approach:**
- Loading rate = cargo loaded (m3 or MT) / time in port (hours).
- Discharging rate = cargo discharged / time.
- Need columns: cargo quantity change, port duration.
- Check the raw data for `LNG ROB` changes between arrival and departure reports.
- This may require DEPARTURE and ARRIVAL report types (currently filtered out in `load_raw_data`).

**Implementation:**
1. Add `"loading_rate"` and `"discharging_rate"` to `COLUMN_MAP` if they exist directly.
2. If not direct columns, compute from ROB deltas:
   ```python
   def compute_loading_rate(df, rob_col, time_col):
       # Find port periods (consecutive rows at port)
       # Calculate ROB delta / time delta
       pass
   ```

**Pitfall:** Loading/discharging may span multiple noon reports. You need to aggregate across the entire port call, not per-day.

### Task 2.2: Fuel Consumption Normalization

**Problem:** The Al Sailiya uses LNG as primary fuel (gas fuel share often >90%). LSMGO ME consumption is near zero because the main engine burns gas, not liquid fuel. The TCP warranty may reference "total fuel" or "fuel oil equivalent".

**Approach:**
- Check what the TCP actually warranties — is it "total energy consumption" or specific fuel grades?
- The `total_liquid_me_cons_mt` derived column already sums liquid fuels.
- May need a `total_fuel_energy_equiv` that converts LNG m3 to MT equivalent using density and calorific value.
- Add conversion constants to `config.py`:
  ```python
  LNG_DENSITY_KG_M3 = 450  # typical, varies with composition
  LNG_CALORIFIC_VALUE_MJ_KG = 50
  LSMGO_CALORIFIC_VALUE_MJ_KG = 42.7
  ```

### Task 2.3: Per-Voyage Aggregation (vs Per-Day)

**Problem:** Some warranty metrics need per-voyage averages (e.g., average speed over the full laden voyage), not per-day averages.

**Approach:**
- Identify voyage segments: port-to-port.
- Weight speed by distance, not by day count: `avg_speed = total_distance / total_steaming_time`.
- In `calculator._aggregate()`:
  ```python
  def _aggregate(series, metric, df=None):
      if metric == "speed" and df is not None:
          dist_col = "distance_made_good"
          time_col = "steaming_time"
          if dist_col in df.columns and time_col in df.columns:
              total_dist = df[dist_col].sum()
              total_time = df[time_col].sum()
              if total_time > 0:
                  return total_dist / total_time
      return float(series.mean())
  ```

**Pitfall:** Distance-weighted speed is the industry-standard method for TCP claims. Simple arithmetic mean overstates the effect of slow days.

---

## PHASE 3: TCP Parsing Improvements

> **Goal:** Handle a wider variety of TCP clause formats reliably.

### Task 3.1: Improve LLM Prompt for Complex Clauses

**Problem:** TCPs vary enormously. Some have tabular warranties, conditional tiers (speed option A/B/C), or cross-referenced annexes.

**Approach:**
- Add more few-shot examples to `_SYSTEM_PROMPT` in `tcp_parser.py`:
  - A tabular warranty format.
  - A multi-tier speed/consumption option.
  - A boil-off warranty with re-liquefaction allowance.
- Add post-processing to merge duplicate metrics or flag conflicting entries.

**Example additional few-shot:**
```
Annex B – Performance Table
| Condition | Speed (kts) | ME Fuel (MT/d) | AE Fuel (MT/d) | BOG (%/d) |
| Laden     | 19.5        | 150            | 10             | 0.12      |
| Ballast   | 20.0        | 140            |  9             | 0.10      |
```

### Task 3.2: OCR Support for Scanned PDFs

**Problem:** Some TCP PDFs are scanned images, not searchable text.

**Approach:**
- Add optional OCR using `pytesseract` + `Pillow`.
- In `tcp_parser.py`, detect if pdfplumber returns empty/garbled text.
- Fall back to page-by-page image extraction → OCR.
- Add `--ocr` flag to `main.py`.

**Implementation sketch:**
```python
def _ocr_pdf(pdf_path):
    from pdf2image import convert_from_path
    import pytesseract
    images = convert_from_path(pdf_path)
    return "\n\n".join(pytesseract.image_to_string(img) for img in images)
```

**Dependencies:** Add `pytesseract`, `Pillow`, `pdf2image` to requirements.txt (optional group).

### Task 3.3: Warranty Condition Parser

**Problem:** Conditions like `"laden, good weather BF<=4, Douglas sea state<=3"` are free text. We need structured parsing.

**Approach:**
- Create `warranty_conditions.py` with a parser:
  ```python
  @dataclass
  class WarrantyCondition:
      vessel_status: str | None  # "laden", "ballast", None
      max_beaufort: float | None
      max_swell: float | None
      max_sea_state: float | None
      custom_text: str

  def parse_condition(text: str) -> WarrantyCondition:
      ...
  ```
- Use this in the calculator to automatically apply the right filters per warranty.

---

## PHASE 4: Template & Output Enhancements

> **Goal:** Produce claim-ready Excel files that match industry formats.

### Task 4.1: Support Multiple Template Formats

**Problem:** Different charterers/operators use different claim templates.

**Approach:**
- Make `TEMPLATE_HEADER_MAP` configurable (load from a JSON/YAML file instead of hardcoding).
- Add `--template-map` CLI argument.
- Provide 2-3 template mapping presets (Al Sailiya, generic LNG, BIMCO).

### Task 4.2: Conditional Formatting in Output

**Problem:** The Summary sheet works but could better highlight claim-worthy deviations.

**Approach:**
- Use openpyxl conditional formatting rules:
  - Red background for deviations > 5%.
  - Green for within warranty.
  - Yellow for marginal (2-5%).
- Add currency/financial impact column if claim rates are provided.

### Task 4.3: Multi-Voyage / Multi-Period Support

**Problem:** Currently handles one charter period at a time.

**Approach:**
- Accept a CSV/JSON file with multiple date ranges.
- Loop through each period, compute actuals, and write to separate sheets or a combined summary.
- Useful for long-term charters with monthly claim windows.

### Task 4.4: Chart Generation

**Problem:** Visual evidence strengthens claims.

**Approach:**
- Use openpyxl's chart API to add:
  - Speed time-series with warranty threshold line.
  - Fuel consumption bar chart (actual vs warranted).
  - Boil-off rate trend with upper limit.
- Add charts to a "Charts" sheet in the output workbook.

---

## PHASE 5: Testing & Validation

> **Goal:** Build confidence that calculations are correct and the pipeline handles edge cases.

### Task 5.1: Unit Tests for Calculator

**File:** Create `tests/test_calculator.py`

**Test cases:**
```python
def test_add_derived_columns_boil_off_rate():
    """BOG rate = NBOG / LNG_ROB * 100"""
    # Create minimal DataFrame with NBOG=100, LNG_ROB=100000 → rate=0.1%

def test_compute_actuals_speed_average():
    """Average speed over 5 days"""

def test_compute_actuals_empty_data():
    """Gracefully handle empty DataFrame"""

def test_aggregate_speed_distance_weighted():
    """Speed = total_distance / total_steaming_time"""
```

### Task 5.2: Unit Tests for TCP Parser

**File:** Create `tests/test_tcp_parser.py`

**Test cases:**
```python
def test_parse_json_response_clean():
    """Valid JSON array is parsed correctly"""

def test_parse_json_response_markdown_fences():
    """JSON wrapped in ```json ... ``` is handled"""

def test_validate_warranties_range_check():
    """Out-of-range values get flagged but kept"""

def test_validate_warranties_missing_keys():
    """Entries without required keys are skipped"""
```

### Task 5.3: Integration Test with Sample Data

**File:** Create `tests/test_integration.py`

- Use `sample_warranties.json` and a small synthetic Excel file.
- Run the full pipeline programmatically (call `main.run()` with mock args).
- Assert that the output Excel is created and contains expected sheets/rows.

### Task 5.4: Historical Claim Comparison

- Take a manually computed past claim.
- Run the same data through the pipeline.
- Compare results cell-by-cell.
- Document any discrepancies and adjust calculation logic.

---

## PHASE 6: Local LLM Integration

> **Goal:** Run warranty extraction without cloud API dependency.

### Task 6.1: Test with Ollama Models

- Install Ollama, pull `llama3`, `mistral`, and `qwen2`.
- Run each model against the same TCP text.
- Compare output quality with OpenAI GPT-4o.
- Document which model gives the best balance of accuracy and speed.

### Task 6.2: Improve Ollama Prompt Format

**Problem:** The current Ollama call uses a flat prompt. Ollama's chat API (`/api/chat`) supports structured messages.

**Fix in `tcp_parser.py`:**
```python
def _call_ollama_chat(tcp_text, model, base_url):
    """Use Ollama's chat API for better results."""
    messages = _build_messages(tcp_text)
    payload = {
        "model": model,
        "messages": messages,
        "stream": False,
        "options": {"temperature": 0},
    }
    url = base_url.replace("/api/generate", "/api/chat")
    resp = requests.post(url, json=payload, timeout=120)
    resp.raise_for_status()
    return resp.json().get("message", {}).get("content", "").strip()
```

### Task 6.3: Add Anthropic Claude Support

- Add `anthropic` to requirements.txt.
- Implement `_call_anthropic()` in `tcp_parser.py`.
- Add `--llm-provider` flag to `main.py` (`openai`, `ollama`, `anthropic`).

---

## PHASE 7: Error Handling & Robustness

> **Goal:** Make the pipeline production-ready.

### Task 7.1: Graceful Handling of Missing/Malformed Data

- In `data_extractor.load_raw_data()`:
  - Handle `.xls` files (suggest conversion or use `xlrd`).
  - Handle password-protected Excel files (detect and warn).
  - Handle multiple sheets (prompt user or use config).

### Task 7.2: Data Validation Report

- Before computing, generate a data quality report:
  - Count of null values per metric column.
  - Date gaps (missing noon reports).
  - Suspicious values (speed > 25 knots, negative fuel).
- Write to a "Data Quality" sheet in the output.

### Task 7.3: Retry Logic for LLM Calls

- Add exponential backoff for transient API failures.
- Cache successful LLM responses to avoid re-calling during re-runs.

### Task 7.4: Configuration Validation at Startup

- Validate `config.py` values at application start.
- Check that `COLUMN_MAP` keys are recognized metrics.
- Check that `WARRANTY_RANGES` bounds make sense.

---

## PHASE 8: Documentation & Deployment

### Task 8.1: User Guide

- Step-by-step guide with screenshots for non-technical users.
- Cover: installation, running first claim, interpreting output.

### Task 8.2: Configuration Guide

- How to adapt to a different vessel/report format.
- How to add new metrics.
- How to create a custom template.

### Task 8.3: Packaging

- Add `pyproject.toml` with proper metadata.
- Make installable via `pip install .`
- Consider a simple web UI (Streamlit or Gradio) as a future enhancement.

---

## QUICK REFERENCE: File Responsibilities

```
main.py               → CLI args, pipeline orchestration, validation
tcp_parser.py         → PDF extraction, LLM calls, warranty JSON parsing
data_extractor.py     → Excel loading, date filtering, column resolution
calculator.py         → Derived columns, daily DataFrame, actuals vs warranty
template_filler.py    → Write to Excel template, Summary sheet, charts
config.py             → All column mappings, thresholds, LLM settings
```

## QUICK REFERENCE: Column Mapping Pattern

When adding a new metric:
1. Add the raw Excel column name to `config.py → COLUMN_MAP`.
2. If it's derived (computed from other columns), add to `DERIVED_METRICS` and implement in `calculator.add_derived_columns()`.
3. Add the metric name to `tcp_parser._SYSTEM_PROMPT` allowed metrics list.
4. Add validation range to `config.py → WARRANTY_RANGES`.
5. Add template header mapping to `template_filler.py → TEMPLATE_HEADER_MAP`.

## PRIORITY ORDER

If working sequentially, the recommended order is:

1. **Phase 1** (filtering) — highest impact on calculation accuracy
2. **Phase 2** (calc engine) — correct formulas for industry standards
3. **Phase 5** (testing) — validate everything before adding more features
4. **Phase 3** (TCP parsing) — improve robustness of input handling
5. **Phase 4** (output) — polish the deliverable
6. **Phase 6** (local LLM) — cost/privacy optimization
7. **Phase 7** (robustness) — production hardening
8. **Phase 8** (docs) — handover readiness
