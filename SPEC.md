# IPEDS Completions Analysis Tool — Project Specification

## Purpose
Python CLI tool that ingests IPEDS Completions (C_A) and Institutional Characteristics (HD)
batch CSV downloads, joins institution metadata, and produces an Excel workbook analyzing
degree/award completions by CIP code and award level — including 5-year CAGR, program count
growth, and market-level aggregation — for user-selected institutions or all institutions
in a given state.

---

## Project Structure

```
Claude Code IPEDS Lookup Tool/
├── SPEC.md                      ← this file
├── README.md                    ← generate last
├── requirements.txt
├── config/
│   ├── institutions.csv         ← optional fixed UNITID list (can be empty)
│   ├── cip_filter.yaml          ← CIP codes and award levels to include
│   └── years.yaml               ← academic years to process
├── data/
│   ├── raw/
│   │   ├── hd2019.csv           ← HD survey, one per year
│   │   ├── hd2020.csv
│   │   ├── hd2021.csv
│   │   ├── hd2022.csv
│   │   ├── hd2023.csv
│   │   ├── c2019_a.csv          ← Completions A survey, one per year
│   │   ├── c2020_a.csv
│   │   ├── c2021_a.csv
│   │   ├── c2022_a.csv
│   │   └── c2023_a.csv
│   └── dictionary/
│       ├── varlist.csv          ← IPEDS data dictionary (all surveys)
│       └── cip_soc_crosswalk.csv ← CIP-SOC crosswalk (CIP labels + SOC codes, stored for v2)
├── src/
│   ├── loader.py                ← reads and validates all raw CSVs
│   ├── resolver.py              ← institution search by name or state filter
│   ├── joiner.py                ← joins UNITID → HD metadata
│   ├── aggregator.py            ← computes CAGR, program counts, market totals
│   ├── reporter.py              ← builds Excel workbook
│   └── main.py                  ← CLI entry point
└── output/
    └── reports/
```

---

## Data Inputs

### Survey: HD (Institutional Characteristics)
File pattern: `hd{year}.csv`

Key columns used:
| Column | Description |
|--------|-------------|
| UNITID | Institution identifier (join key) |
| INSTNM | Institution name |
| STABBR | State abbreviation (e.g., CA) |
| CONTROL | 1=Public, 2=Private nonprofit, 3=Private for-profit |
| ICLEVEL | 1=4-year, 2=2-year, 3=Less-than-2-year |
| CARNEGIE | Carnegie classification code |
| CITY | City |

Use the **most recent year's HD file** as the authoritative name/state lookup.
Cross-reference older HD files only to catch UNITIDs that have since closed or merged.

### Survey: Completions A (C_A)
File pattern: `c{year}_a.csv`

Key columns used:
| Column | Description |
|--------|-------------|
| UNITID | Institution identifier |
| CIPCODE | 6-digit CIP code (e.g., 51.3801 = Nursing) |
| MAJORNUM | 1 = first major, 2 = second major |
| AWLEVEL | Award level (see table below) |
| CTOTALT | Total completions, all students |
| CTOTALW | Total completions, women |
| CTOTALM | Total completions, men |

**CRITICAL — MAJORNUM filter:** Always filter to `MAJORNUM == 1` only.
Including second majors (MAJORNUM=2) double-counts completions. This is the single most
common IPEDS completions error. Never aggregate without this filter applied first.

**Award level codes:**
| AWLEVEL | Label |
|---------|-------|
| 1 | Postsecondary award < 1 year |
| 2 | Postsecondary award 1–2 years |
| 3 | Associate's degree |
| 4 | Postsecondary award 2–4 years |
| 5 | Bachelor's degree |
| 6 | Post-baccalaureate certificate |
| 7 | Master's degree |
| 8 | Post-master's certificate |
| 9 | Doctorate — research/scholarship |
| 10 | Doctorate — professional practice |
| 11 | Doctorate — other |

### CIP-SOC Crosswalk (`data/dictionary/cip_soc_crosswalk.csv`)
Download from: https://nces.ed.gov/ipeds/cipcode/resources.aspx (or BLS O*NET crosswalk page)
Contains: `CIPCode`, `CIPTitle`, `SOCCode`, `SOCTitle`

**v1 use:** CIPTitle only — joined to output to display human-readable program names.
**v2 use:** SOCCode and SOCTitle — enables labor market integration (BLS OES wage data,
employment projections, job posting analytics) keyed to the same CIP codes already
in the completions data.

Store the full file as-is. Do not drop SOC columns even though they are unused in v1.
Note in code with `# SOC columns retained for v2 labor market integration`.

**CIP code structure:**
- `51` = 2-digit series (e.g., Health Professions)
- `51.38` = 4-digit family (e.g., Registered Nursing)
- `51.3801` = 6-digit program (e.g., Registered Nursing/Registered Nurse)

The C_A file stores 6-digit codes. Strip trailing zeros carefully:
`51.3801` ≠ `51.38` — always match at the same digit level.

### CIP Filter Config (`config/cip_filter.yaml`)
User specifies which programs and award levels to analyze:
```yaml
# Leave empty to include ALL CIP codes (produces large output)
cip_codes:
  - "51.3801"   # Registered Nursing
  - "52.0201"   # Business Administration
  - "11.0701"   # Computer Science

# Leave empty to include ALL award levels
award_levels:
  - 5    # Bachelor's
  - 7    # Master's
  - 9    # Doctorate - research
  - 10   # Doctorate - professional
```

### Years Config (`config/years.yaml`)
```yaml
years: [2019, 2020, 2021, 2022, 2023]
cagr_start_year: 2019
cagr_end_year: 2023
```

---

## Institution Selection (Three Modes)

### Mode 1 — Name Search (interactive)
```bash
python src/main.py --search "Pacific University"
```
- Searches HD file `INSTNM` column (case-insensitive, partial match)
- Prints numbered list of matches with UNITID, state, control, level
- User selects by number(s) → proceeds with those UNITIDs

### Mode 2 — State Filter
```bash
python src/main.py --state CA
python src/main.py --state CA --control 2    # private nonprofit only
python src/main.py --state CA --iclevel 1    # 4-year only
```
- Pulls all UNITIDs from HD where `STABBR == state`
- Optional sub-filters: `--control` and `--iclevel`
- Prints count of matched institutions before proceeding
- Warn if count > 200 (large output); ask confirmation

### Mode 3 — Fixed List
```bash
python src/main.py --unitids 110644 110635 110662
```
Or populated `config/institutions.csv` used automatically if Modes 1 and 2 not specified.

**Modes can combine:** `--search` result can be added to `--state` results.

---

## Calculations

### 5-Year CAGR
For each institution × CIP code × award level combination:
```
CAGR = (completions_end / completions_start) ^ (1 / (n_years - 1)) - 1
```
Where:
- `completions_end` = CTOTALT for `cagr_end_year`
- `completions_start` = CTOTALT for `cagr_start_year`
- `n_years - 1` = 4 (for 2019→2023)

**Edge cases:**
- `completions_start == 0`: CAGR = undefined → display as `N/A` (new program)
- `completions_end == 0` and `completions_start > 0`: CAGR = -100% (program ended)
- Either year missing from data: CAGR = `N/A` with note in a flag column

### Program Count Growth
For each institution, count distinct `CIPCODE + AWLEVEL` combinations per year.
Report: count per year + absolute change (start→end) + percent change.
This measures breadth of offerings, not volume.

### Market-Level Aggregation
For each `CIPCODE + AWLEVEL` combination across **all selected institutions**:
- Total completions per year (sum of CTOTALT)
- Market CAGR (same formula, applied to summed totals)
- Institution count offering that program (count of distinct UNITIDs with CTOTALT > 0)

---

## Output: Excel Workbook

Filename: `output/reports/IPEDS_Completions_{state_or_custom}_{timestamp}.xlsx`

### Tab Structure

| # | Tab Name | Contents |
|---|----------|----------|
| 1 | `Institutions` | UNITID, name, state, control, level, Carnegie — all selected institutions |
| 2 | `Completions_2019` | UNITID, Institution, CIP, CIP Title, Award Level, CTOTALT, CTOTALM, CTOTALW |
| 3 | `Completions_2020` | Same structure |
| 4 | `Completions_2021` | Same structure |
| 5 | `Completions_2022` | Same structure |
| 6 | `Completions_2023` | Same structure |
| 7 | `CAGR_by_Institution` | One row per UNITID × CIP × AWLEVEL; columns: years + CAGR + flag |
| 8 | `Program_Growth` | One row per UNITID; program count per year + delta |
| 9 | `Market_View` | One row per CIP × AWLEVEL; aggregate totals per year + market CAGR + institution count |
| 10 | `Definitions` | Variable names, survey source, plain-English labels from varlist.csv + AWLEVEL lookup |

### Formatting

**All tabs:**
- Freeze top row + first two columns
- Bold header row, light gray fill (#F2F2F2)
- Auto-fit column widths (max 45 chars)

**Completions tabs (2–6):**
- Integer formatting for all count columns
- Sort: UNITID ascending, then CIPCODE ascending, then AWLEVEL ascending

**CAGR tab (7):**
- CAGR formatted as percentage (1 decimal): `12.3%`
- Conditional formatting: green if CAGR > 0, red if CAGR < 0, gray if N/A
- Flag column values: `OK`, `New Program`, `Program Ended`, `Missing Data`

**Program Growth tab (8):**
- Integer for counts, percentage for % change
- Conditional formatting on % change column: color scale

**Market View tab (9):**
- Bold subtotal rows
- Conditional formatting on Market CAGR: color scale (green high, red low)

---

## CLI Reference

```bash
# Search by institution name (interactive)
python src/main.py --search "Pacific University"

# All institutions in a state
python src/main.py --state CA

# State + institution type filters
python src/main.py --state CA --control 2 --iclevel 1

# Specific UNITIDs
python src/main.py --unitids 110644 110635

# Filter to specific CIP codes at runtime (overrides cip_filter.yaml)
python src/main.py --state CA --cip 51.3801 52.0201

# Filter to specific award levels
python src/main.py --state CA --awlevel 5 7

# Verbose output
python src/main.py --state CA --verbose

# Output path override
python src/main.py --state CA --output ./my_reports/
```

---

## Data Quality Rules

1. **Always filter `MAJORNUM == 1`** before any aggregation — no exceptions
2. **Warn** when a UNITID appears in institutions list but has zero completions rows in C_A for a given year
3. **Warn** when a CIP code in `cip_filter.yaml` returns zero rows across all institutions and years
4. **Never impute** missing values — leave as blank/NaN; note in flag column
5. **Log** for each year: total rows loaded, rows after MAJORNUM filter, rows after institution filter, rows after CIP filter
6. **Validate** CIPCODE format: must match pattern `\d{2}\.\d{4}` — warn and skip malformed codes
7. **Handle mixed file formats** — `loader.py` detects extension and routes accordingly: `pd.read_excel()` for `.xlsx`, `pd.read_csv(encoding='latin-1')` for `.csv`; both may exist across different years for the same survey

---

## Tech Stack

| Package | Purpose |
|---------|---------|
| `pandas` | Data loading, joins, pivots, aggregation |
| `openpyxl` | Excel workbook generation and formatting |
| `pyyaml` | Config file parsing |
| `argparse` | CLI |
| `rich` | Terminal output, progress bars, tables |
| `fuzzywuzzy` or `rapidfuzz` | Institution name fuzzy search |

Python 3.10+

---

## Development Sequence (strict order)

1. `loader.py` — load HD and C_A CSVs for all years; validate columns; print shape summary
2. `resolver.py` — institution name search and state filter; print matched institution table
3. `joiner.py` — join C_A rows to HD metadata; confirm INSTNM appears correctly
4. `aggregator.py` — CAGR, program count, market aggregation; unit test each formula
5. `reporter.py` — build Excel; start with Institutions tab + one Completions tab
6. `reporter.py` continued — add CAGR, Program Growth, Market View, Definitions tabs
7. `main.py` — wire all modules to CLI; test all flag combinations
8. Edge cases — zero-start CAGR, missing years, large state outputs (>200 institutions)
9. `README.md` — generate last

---

## Known IPEDS Quirks

| Quirk | Handling |
|-------|---------|
| **MAJORNUM double-counting** | Filter `MAJORNUM == 1` in `loader.py` immediately on load — before any other operation |
| **CSV encoding** | Open all files with `encoding='latin-1'` |
| **CIP code leading zeros** | Store CIPCODE as string, never float (e.g., `01.0101` loses leading zero as float) |
| **Suppressed small cells** | IPEDS suppresses cells < 3 completions; these appear as blank — treat as NaN, not zero |
| **Variable renamed across years** | Log warning; check varlist.csv; leave column blank if unresolvable |
| **Institution closed/merged** | UNITID absent from HD; log as `[NOT IN HD — POSSIBLE CLOSURE]` |
| **CIPCODE `99` catch-all** | IPEDS uses `99.0000` as a residual code — exclude from analysis by default; add `--include-residual` flag to override |

---

## Out of Scope (v1)
- Demographic breakdowns (by race/gender) — v2
- Automated NCES download
- Web UI
- CIP code crosswalk (2010→2020 changes) — v2
- Peer group auto-generation
- SOC/labor market integration (BLS OES wages, employment projections) — v2
  *(CIP-SOC crosswalk is already stored in `data/dictionary/` to enable this)*
