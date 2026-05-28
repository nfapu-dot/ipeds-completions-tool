# IPEDS Completions Analysis Tool

Analysis tool for IPEDS Completions (C_A) and Institutional Characteristics (HD)
data. Joins five academic years of survey data, applies the IPEDS data-quality
rules, and produces a formatted Excel workbook with 5-year CAGR, program-count
growth, and market-level aggregation by CIP code × award level.

Two interfaces:
- **Web app** (recommended) — double-click to launch, point-and-click form,
  Excel download. No commands to type. See [Web app interface](#web-app-interface).
- **CLI** — for scripted/automated runs and power users. See
  [CLI reference](#cli-reference).

For the full data-handling specification, see [SPEC.md](SPEC.md).

---

## Web app interface

### Launching

Double-click **`Launch IPEDS Tool.command`** in the project folder. A Terminal
window opens (don't close it — it runs the app in the background), and your
browser opens to `http://localhost:8501`.

First-time load takes about 10 seconds (it reads ~1.4 million IPEDS rows).
After that, every filter change responds in under a second.

### Using the form

**Left sidebar — pick what to analyze:**
1. **Mode** — choose how to select institutions:
   - *By state* → state dropdown + Control/Level checkboxes
   - *By institution name* → live search box
   - *Specific UNITIDs* → searchable multi-select with friendly names
   - *From institutions.csv config* → uses the file in `config/`
2. **Program filters** — pick CIP codes (searchable by name) and award levels.
   Leave empty to include everything.
3. **Include CIP 99 residual rollups** — only check if you want
   institution-level total rows.
4. Click **Generate Report**.

**Main area** shows:
- Summary metrics across the top (institution count, CAGR rows, etc.)
- "Download Excel workbook" button — saves to `output/reports/` and offers
  download
- Five tabs: Institutions, CAGR by Institution, Program Growth, Market View,
  Definitions — same content as the corresponding Excel tabs, with the same
  green/red CAGR highlighting

### Stopping the app

Close the Terminal window that opened with the launcher, or press `Ctrl+C` in
it. Your browser tab can stay open.

### Troubleshooting the web app

**"This site can't be reached" in browser** — the launcher didn't finish
starting. Wait 5–10 seconds and refresh. If it still fails, check the
Terminal window for an error.

**"Address already in use" in Terminal** — the app is already running in
another Terminal window. Either use that one, or close it and re-launch.

**Browser doesn't open automatically** — visit
`http://localhost:8501` manually.

---

## Requirements

- Python 3.9 or newer
- ~500 MB of free disk space for IPEDS CSV files
- Packages listed in [requirements.txt](requirements.txt)

```bash
pip install -r requirements.txt
```

---

## Setup

### 1. Download IPEDS data

From [NCES IPEDS Complete Data Files](https://nces.ed.gov/ipeds/use-the-data/download-access-database):

| File | Pattern | Years needed |
|------|---------|--------------|
| Institutional Characteristics | `hd{year}.csv` | one per year in `years.yaml` |
| Completions A — Awards by Program, Race, Gender | `c{year}_a.csv` | one per year in `years.yaml` |
| IPEDS varlist (data dictionary) | `varlist.xlsx` | one (per any year) |
| CIP-SOC crosswalk | `cip_soc_crosswalk.xlsx` | one |

Drop survey files into `data/raw/` and dictionary files into `data/dictionary/`.
The loader accepts `.csv` and `.xlsx` interchangeably.

Default year range is **2020–2024** (configurable in `config/years.yaml`).

### 2. Review configs

- [config/years.yaml](config/years.yaml) — years to process + CAGR endpoints
- [config/cip_filter.yaml](config/cip_filter.yaml) — CIP codes + award levels (leave empty for all)
- [config/institutions.csv](config/institutions.csv) — optional fixed UNITID list

### 3. Smoke test

```bash
python src/main.py --unitids 110635 --cip 51.3801 --awlevel 5
```

Output: `output/reports/IPEDS_Completions_custom_{timestamp}.xlsx`.

---

## CLI reference

The CLI (`src/main.py`) is the same engine the web app uses, exposed for
scripted runs and power-user one-offs. It still works exactly as before — the
web app does not replace it.

### Quick start

```bash
# All institutions in a state, all CIPs / award levels from cip_filter.yaml
python src/main.py --state OR

# State + sub-filters: private nonprofit + 4-year only
python src/main.py --state CA --control 2 --iclevel 1

# Specific institutions, specific programs, specific award levels
python src/main.py --unitids 110635 110662 --cip 51.3801 11.0701 --awlevel 5 7

# Interactive name search (prompts for picks from match list)
python src/main.py --search "Pacific University"

# Combine modes: search results union with state results
python src/main.py --state HI --search "Stanford"

# Verbose logging (per-year row counts at every step)
python src/main.py --state OR --verbose

# Custom output directory
python src/main.py --state OR --output ./my_reports/
```

---

### CLI flags

| Flag | Purpose |
|------|---------|
| `--search NAME` | Substring + fuzzy search against HD `INSTNM`; interactive picker. |
| `--state ABBR` | All institutions in state. Prompts to confirm if > 200 matches. |
| `--control {1,2,3}` | Sub-filter: 1=Public, 2=Private nonprofit, 3=Private for-profit. |
| `--iclevel {1,2,3}` | Sub-filter: 1=4-year, 2=2-year, 3=<2-year. |
| `--unitids ID [ID …]` | Specific UNITIDs (overrides `config/institutions.csv`). |
| `--cip CODE [CODE …]` | Override `cip_filter.yaml` cip_codes. Quote to preserve leading zeros (e.g., `"01.0101"`). |
| `--awlevel N [N …]` | Override `cip_filter.yaml` award_levels. |
| `--include-residual` | Include CIP 99 aggregate-rollup rows (excluded by default). |
| `--output DIR` | Output directory (default: `output/reports/`). |
| `--verbose` | Per-step row count logs for every module. |

Selection modes (`--search`, `--state`, `--unitids`) can combine — the union of
their UNITIDs is processed.

---

## Output workbook

Filename: `output/reports/IPEDS_Completions_{label}_{YYYYMMDD_HHMMSS}.xlsx`

Label is the state code when `--state` is used alone (e.g. `CA`, `OR_ctrl2_lvl1`),
otherwise `custom`. Search-only runs use `search_{query}`.

| # | Sheet | Contents |
|---|-------|----------|
| 1 | `Institutions` | UNITID, Institution, State, Control + label, Level + label, Carnegie, HD source year. |
| 2-6 | `Completions_{year}` | UNITID, Institution, CIPCODE, CIP Title, Award Level, Total/Men/Women completions. Sorted UNITID → CIPCODE → Award Level. |
| 7 | `CAGR_by_Institution` | One row per UNITID × CIPCODE × Award Level. Year columns + CAGR + Flag. Direct fills: green (>0), red (<0), gray (N/A). |
| 8 | `Program_Growth` | One row per UNITID. Distinct (CIPCODE, Award Level) count per year + delta + % change. Color-scale on % change. |
| 9 | `Market_View` | One row per CIPCODE × Award Level. Yearly sums + Market CAGR + institution count per year. Color-scale on Market CAGR. |
| 10 | `Definitions` | IPEDS variable definitions + Award Level codes (including legacy 17–20) + CONTROL / ICLEVEL / CAGR Flag lookups. |

All sheets: bold header (light gray fill), freeze top row + first two columns,
integer formatting on counts, percent formatting (0.0%) on CAGR columns,
auto-fit column widths (capped at 45 chars).

---

## IPEDS data quirks the tool handles

| Quirk | Handling |
|-------|----------|
| **MAJORNUM double-counting** | `loader.py` filters `MAJORNUM == 1` immediately on load. Including second majors over-counts students with double majors. **Never aggregate without this filter.** |
| **CSV encoding (latin-1)** | All CSVs read with `encoding='latin-1'`. |
| **UTF-8 BOM in hd2023+** | BOM (`﻿` / `ï»¿`) stripped from column names. |
| **CIPCODE leading zeros** | Stored as string throughout (`'01.0101'` would otherwise lose its leading zero as float). |
| **Suppressed cells (<3)** | Preserved as NaN; never imputed to zero. Aggregations use `min_count=1` so empty groups stay NaN. |
| **CIP 99 / 2-digit aggregate rollups** | Excluded by default via the `IS_CIP_6DIGIT` flag added at load time. Pass `--include-residual` to keep them. |
| **Carnegie column variance** | HD files split Carnegie across 8 columns. Tool prefers `C21BASIC` (2021 Basic), falls back to `C00CARNEGIE`. hd2020 has neither. |
| **Closed / merged institutions** | UNITIDs absent from the latest HD are sourced from their most recent earlier appearance. `HD Source Year` on the Institutions tab records which HD year supplied the metadata. |
| **Legacy AWLEVEL codes** | Codes 17/18/19/20 (legacy doctorate variants) appear in real data alongside the spec's 1–11. Documented in the Definitions tab. |
| **Multi-sheet xlsx** | `varlist.xlsx` and `cip_soc_crosswalk.xlsx` are multi-sheet; loader targets `Varlist` and `CIP-SOC` sheets specifically. |
| **CIP-SOC row explosion** | Crosswalk has ~5.7 SOCs per CIP. Only `CIPTitle` is joined to C_A. Full crosswalk (with SOC columns) stays loaded for v2 labor market integration. |

---

## Calculations

### 5-year CAGR
```
CAGR = (completions_end / completions_start) ** (1 / (end_year - start_year)) - 1
```
For 2020 → 2024, the exponent is `1/4`.

Edge case flags:

| Condition | Flag | CAGR cell |
|-----------|------|-----------|
| Both endpoints > 0 | `OK` | computed value |
| Start = 0 | `New Program` | blank (N/A) |
| Start > 0, end = 0 | `Program Ended` | -100% |
| Either endpoint is NaN / missing | `Missing Data` | blank (N/A) |

### Program count growth
For each UNITID: count of distinct `(CIPCODE, AWLEVEL)` combinations per year +
absolute delta and percent change from start to end year. Measures breadth of
offerings, not volume of degrees.

### Market view
For each `(CIPCODE, AWLEVEL)` across all selected institutions: sum of
CTOTALT per year, market-level CAGR on the sums, and count of distinct UNITIDs
with CTOTALT > 0 (institution count offering that program).

---

## Troubleshooting

**`config error: cagr_start_year=... must both be in years=[…]`**  
The `cagr_start_year` / `cagr_end_year` values in `config/years.yaml` must
appear in the `years` list.

**`data file missing: Missing HD file for year YYYY in data/raw`**  
Check that every year listed in `config/years.yaml` has both `hd{year}.{csv,xlsx}`
and `c{year}_a.{csv,xlsx}` files in `data/raw/`.

**`warning: CIP codes ['51.38'] do not match the 6-digit format`**  
IPEDS uses 6-digit CIPs (`51.3801`). Pass full 6-digit codes, or use
`--include-residual` if you intended to match a 2-digit aggregate code like `99`.

**`668 institutions matched (>200). Continue? [y/n]`**  
Confirmation prompt to avoid accidentally generating huge workbooks. Type `y`
to proceed, or refine with `--control` / `--iclevel`.

**`⚠ 2021: 1 selected UNITIDs have zero rows after filters`**  
The institution exists in HD but reported no completions matching your CIP /
award-level filters that year. Common for small institutions, suppressed-data
years, or schools that closed mid-window. The Institutions tab's
`HD Source Year` column indicates which HD file the metadata came from.

**`[NOT IN HD — POSSIBLE CLOSURE] sample: [101541]`**  
A UNITID present in older C_A files but absent from the most recent HD —
usually a closure or merger. The resolver's closure cross-reference surfaces
the last year the institution appeared.

---

## Project layout

```
Claude Code IPEDS Lookup Tool/
├── SPEC.md                       Authoritative data-handling specification
├── README.md                     This file
├── requirements.txt
├── Launch IPEDS Tool.command     ★ Double-click to start the web app
├── .streamlit/
│   └── config.toml               APU brand theme for the web app
├── config/
│   ├── years.yaml                Years to process + CAGR endpoints
│   ├── cip_filter.yaml           CIP codes + award levels (leave empty for all)
│   └── institutions.csv          Optional fixed UNITID list
├── data/
│   ├── raw/                      IPEDS HD + C_A survey files
│   └── dictionary/               varlist + CIP-SOC crosswalk
├── src/
│   ├── loader.py                 Reads CSVs, applies MAJORNUM filter, logs counts
│   ├── resolver.py               Search / state filter / fixed-list resolution
│   ├── joiner.py                 Joins C_A → HD metadata + CIPTitle
│   ├── aggregator.py             CAGR, program growth, market view
│   ├── reporter.py               Excel workbook builder
│   ├── main.py                   CLI entry point
│   └── app.py                    Streamlit web app entry point
└── output/
    └── reports/                  Generated workbooks
```

Each module is runnable as `python src/{module}.py` for an isolated smoke test
against the configured data.

---

## Out of scope (v1)

- Demographic breakdowns (by race / gender) — planned for v2
- Automated NCES download
- Web UI
- CIP code crosswalk (2010 → 2020 changes)
- Peer group auto-generation
- SOC / labor market integration (BLS OES wages, employment projections) — SOC
  columns are already retained in the loaded crosswalk to enable this in v2.
