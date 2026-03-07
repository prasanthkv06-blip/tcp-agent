# TCP Agent — LNG Vessel Performance Claim Calculator

Automated Python pipeline for generating Time Charter Party (TCP) performance
claims for LNG carriers. Reads raw noon-report Excel data, applies maritime
industry rules, and produces a comprehensive Excel report with AI-powered
quality review.

**Vessel:** Id'Asah (IMO 9977220) — 174,221.98 cbm LNG Carrier

---

## Features

- **Automated voyage detection** — DEPARTURE → ARRIVAL boundary scanning
- **ROB-difference fuel accounting** — MGO, LNG, VLSFO (Rule 1)
- **Weather exclusion** — BF5+ hours proportional fuel deduction (Rule 2)
- **Multi-segment support** — voyage order revisions, port/fuel mode changes (Rule 3)
- **Boundary pro-rating** — speed-weighted ratio for shared boundary rows
- **Auxiliary data** — boiler, pilot flame, GCU, reliquefaction tracking (Rule 4)
- **Warranty interpolation** — speed/consumption table with ageing factors (Rule 5)
- **Speed anomaly detection** — flags rows below 10% of weighted average (Rule A)
- **Bunkering stop merging** — detects mid-voyage intermediate stops (Rule B)
- **AI Analyst** — Claude-powered expert review that cross-validates output vs raw data
- **Column auto-detection** — adapts to header name changes automatically
- **Dual interface** — CLI (`main.py`) and Web UI (`app.py` via Streamlit)

---

## Quick Start

### 1. Create a virtual environment

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Set your API keys

```bash
cp .env.example .env
# Edit .env and add your keys:
#   ANTHROPIC_API_KEY=sk-ant-...    (for AI Analyst)
#   OPENAI_API_KEY=sk-...           (optional, for TCP parsing)
```

### 3. Run via CLI

```bash
python main.py \
    --input "RawExcel Idasah 05Dec_22Dec.xlsx" \
    --output report.xlsx \
    --verbose
```

### 4. Run via Web UI

```bash
streamlit run app.py
```

---

## Project Structure

```
vessel_performance_app/
├── main.py               # CLI orchestrator — 5-step pipeline
├── app.py                # Streamlit web UI
├── config.py             # Column mappings, vessel params, speed tables
├── data_extractor.py     # Excel loading, voyage detection, Rule B
├── calculator.py         # Interpolation, segments, weather, Rule A
├── template_filler.py    # 3-sheet Excel output + AI Review sheet
├── ai_analyst.py         # Claude-powered expert review
├── tcp_parser.py         # PDF extraction + LLM warranty parsing
├── tests/
│   ├── test_calculator.py      # 42 tests
│   ├── test_data_extractor.py  # 28 tests
│   └── test_integration.py     # 12 tests
├── requirements.txt
├── .env.example
└── README.md
```

---

## Processing Rules

| Rule | Description |
|------|-------------|
| **Rule 1** | Fuel consumed via ROB difference: `ROB(DEP) - ROB(ARR)` |
| **Rule 2** | Weather exclusion when BF5 hours > 6: proportional fuel deduction |
| **Rule 3** | Segment splits on voyage order revision, next port change, fuel mode change |
| **Rule 4** | Auxiliary extraction: boiler, pilot flame, GCU, reliquefaction, subcooler |
| **Rule 5** | Warranty speed/consumption interpolation with ageing factors |
| **Rule A** | Speed anomaly: flag rows where avg speed < 10% of cumulative weighted avg |
| **Rule B** | Bunkering merge: combine intermediate ARRIVAL/DEPARTURE with same next port |

---

## AI Analyst

When `ANTHROPIC_API_KEY` is set, the app runs an expert Claude review that:

1. **Validates raw data structure** — column headers, data types, value ranges
2. **Cross-validates computed output** — fuel rates, distance/speed/time, ROB consistency
3. **Flags anomalies** — with severity levels (error / warning / info)

Output appears in both the Streamlit UI and a dedicated "AI Review" sheet in the Excel report.

Without the API key, the report generates normally — AI review is gracefully skipped.

---

## Output

The generated Excel report contains:

| Sheet | Contents |
|-------|----------|
| **CP Parameters** | Vessel specs, warranty tables |
| **Fuel Tables** | Speed/consumption warranty data |
| **Voyage N** | Per-voyage: daily breakdown, segments, weather exclusions, BOR, auxiliary data |
| **AI Review** | AI analyst findings with severity, category, and details |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Data loading and manipulation |
| `openpyxl` | Excel reading/writing |
| `anthropic` | AI Analyst (Claude API) |
| `streamlit` | Web UI |
| `pdfplumber` | PDF text extraction |
| `python-dotenv` | `.env` file support |
| `numpy` | Numeric operations |

---

## Tests

```bash
python -m pytest tests/ -v
```

82 tests covering interpolation, weather exclusion, segment detection, boundary
pro-rating, speed formula, speed anomaly detection, bunkering stop merging,
voyage detection, ROB-diff fuel, auxiliary data, and integration scenarios.

---

## License

Private — All rights reserved.
