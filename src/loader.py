"""
loader.py — Reads and validates raw IPEDS CSVs.

Responsibilities:
- Load HD (Institutional Characteristics) and Completions A (C_A) files for each
  configured year.
- Load the varlist dictionary and CIP-SOC crosswalk.
- Dual-extension routing: pd.read_excel() for .xlsx, pd.read_csv(encoding='latin-1')
  for .csv.
- Strip UTF-8 BOM ("\\ufeff" / "ï»¿") and trailing whitespace from column names.
- Enforce CIPCODE as string (preserve leading zeros, e.g., '01.0101').
- CRITICAL — MAJORNUM filter:
    Filter MAJORNUM == 1 on every C_A load BEFORE any other operation.
    Including MAJORNUM == 2 double-counts completions. Never aggregate
    without this filter applied first.
- Flag non-6-digit CIPCODE rows (bare '99', '01', etc.) via IS_CIP_6DIGIT column;
  the aggregator decides whether to exclude based on --include-residual.
- Suppressed count cells (IPEDS suppresses < 3) remain NaN. Never imputed.
- Log before/after row counts at every step (rich console).
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import yaml
from rich.console import Console
from rich.table import Table

console = Console()

# Columns required from each survey for v1 analysis.
HD_REQUIRED_COLS = [
    'UNITID', 'INSTNM', 'STABBR', 'CONTROL', 'ICLEVEL', 'CITY',
]
# C21BASIC is the 2021 Basic Carnegie Classification — chosen as the
# canonical "Carnegie" field for the Institutions tab (spec lists a generic
# CARNEGIE column; the real HD file splits Carnegie across 8 columns).
HD_OPTIONAL_COLS = ['C21BASIC', 'C00CARNEGIE']

CA_REQUIRED_COLS = [
    'UNITID', 'CIPCODE', 'MAJORNUM', 'AWLEVEL',
    'CTOTALT', 'CTOTALM', 'CTOTALW',
]

CIP_6DIGIT_PATTERN = re.compile(r'^\d{2}\.\d{4}$')


# ---------------------------------------------------------------------------
# File-level helpers
# ---------------------------------------------------------------------------

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip BOM and surrounding whitespace from column names."""
    df.columns = [
        str(c).replace('﻿', '').replace('ï»¿', '').strip()
        for c in df.columns
    ]
    return df


def _read_table(
    path: Path,
    dtype: Optional[dict] = None,
    sheet_name: Optional[str] = None,
    keep_cols: Optional[set] = None,
) -> pd.DataFrame:
    """Route by extension. CSVs use latin-1 per IPEDS quirk rules.

    keep_cols, if provided, restricts loading to that set of columns (plus any
    that happen to be absent — callable usecols is tolerant of missing names,
    which matters for HD's optional C21BASIC/C00CARNEGIE columns that exist in
    some years but not others). This is a memory optimization for Streamlit
    Cloud's 1 GB tier: C_A files have ~23 columns of which only 7 are used.
    """
    suffix = path.suffix.lower()
    # BOM-tolerant column matcher: IPEDS CSVs sometimes prefix the first
    # column name with a UTF-8 BOM. _clean_columns strips it later, but
    # usecols runs before that. Strip both the unicode BOM (﻿) and the
    # latin-1 representation of the UTF-8 BOM bytes (ï»¿), since CSVs are
    # decoded as latin-1.
    def _col_filter(col: str) -> bool:
        if keep_cols is None:
            return True
        cleaned = col.lstrip('﻿').lstrip('ï»¿').strip()
        return cleaned in keep_cols

    if suffix in ('.xlsx', '.xls'):
        df = pd.read_excel(path, dtype=dtype, sheet_name=sheet_name or 0)
        if keep_cols is not None:
            df = df.loc[:, [c for c in df.columns if _col_filter(c)]]
    elif suffix == '.csv':
        df = pd.read_csv(
            path,
            encoding='latin-1',
            dtype=dtype,
            low_memory=False,
            usecols=_col_filter if keep_cols is not None else None,
        )
    else:
        raise ValueError(f'Unsupported file extension for {path.name}: {path.suffix}')
    return _clean_columns(df)


def _find_file(directory: Path, stem: str) -> Optional[Path]:
    """Find {stem}.csv or {stem}.xlsx (case-insensitive on stem) in `directory`."""
    if not directory.exists():
        return None
    stem_lower = stem.lower()
    for entry in directory.iterdir():
        if not entry.is_file():
            continue
        if entry.stem.lower() == stem_lower and entry.suffix.lower() in ('.csv', '.xlsx', '.xls'):
            return entry
    return None


# ---------------------------------------------------------------------------
# Config loaders
# ---------------------------------------------------------------------------

def load_years_config(config_path: Path) -> dict:
    with open(config_path) as f:
        cfg = yaml.safe_load(f)
    years = cfg.get('years') or []
    if not years:
        raise ValueError(f'{config_path} has no `years` list')
    start = cfg.get('cagr_start_year', years[0])
    end = cfg.get('cagr_end_year', years[-1])
    if start not in years or end not in years:
        raise ValueError(
            f'cagr_start_year={start} and cagr_end_year={end} must both be in years={years}'
        )
    return {'years': list(years), 'cagr_start_year': start, 'cagr_end_year': end}


def load_cip_filter_config(config_path: Path) -> dict:
    with open(config_path) as f:
        cfg = yaml.safe_load(f) or {}
    cfg.setdefault('cip_codes', [])
    cfg.setdefault('award_levels', [])
    cfg['cip_codes'] = [str(c).strip() for c in (cfg['cip_codes'] or [])]
    cfg['award_levels'] = [int(a) for a in (cfg['award_levels'] or [])]
    return cfg


# ---------------------------------------------------------------------------
# Survey loaders
# ---------------------------------------------------------------------------

def load_hd(year: int, raw_dir: Path) -> pd.DataFrame:
    path = _find_file(raw_dir, f'hd{year}')
    if path is None:
        raise FileNotFoundError(f'Missing HD file for year {year} in {raw_dir}')
    console.log(f'[cyan]HD {year}[/] ← {path.name}')

    hd_keep = set(HD_REQUIRED_COLS) | set(HD_OPTIONAL_COLS)
    df = _read_table(path, keep_cols=hd_keep)
    raw_rows = len(df)

    missing = [c for c in HD_REQUIRED_COLS if c not in df.columns]
    if missing:
        console.log(f'  [yellow]⚠ missing required cols: {missing}[/]')

    optional_present = [c for c in HD_OPTIONAL_COLS if c in df.columns]
    optional_absent = [c for c in HD_OPTIONAL_COLS if c not in df.columns]
    if optional_absent:
        console.log(f'  [dim]optional cols absent: {optional_absent}[/]')

    if 'UNITID' in df.columns:
        df['UNITID'] = pd.to_numeric(df['UNITID'], errors='coerce').astype('Int64')
    if 'STABBR' in df.columns:
        df['STABBR'] = df['STABBR'].astype(str).str.strip().str.upper()
    if 'INSTNM' in df.columns:
        df['INSTNM'] = df['INSTNM'].astype(str).str.strip()
    for c in ('CONTROL', 'ICLEVEL'):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').astype('Int64')

    console.log(f'  rows loaded: {raw_rows:,}  cols: {df.shape[1]}'
                f'  carnegie src: {optional_present or "[none]"}')
    return df


def load_ca(year: int, raw_dir: Path) -> pd.DataFrame:
    """
    Load Completions A for `year`.

    CRITICAL: applies MAJORNUM == 1 filter BEFORE returning. Never aggregate
    completions data that skipped this filter — MAJORNUM == 2 (second majors)
    double-counts every student with a double major.
    """
    path = _find_file(raw_dir, f'c{year}_a')
    if path is None:
        raise FileNotFoundError(f'Missing C_A file for year {year} in {raw_dir}')
    console.log(f'[cyan]C_A {year}[/] ← {path.name}')

    df = _read_table(path, dtype={'CIPCODE': str}, keep_cols=set(CA_REQUIRED_COLS))
    raw_rows = len(df)

    missing = [c for c in CA_REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f'C_A {year}: missing required columns: {missing}')

    df['CIPCODE'] = df['CIPCODE'].astype(str).str.strip()

    # ----- CRITICAL MAJORNUM FILTER -----
    # Filter MAJORNUM == 1 immediately and never aggregate without it.
    # See module docstring + SPEC §Known IPEDS Quirks.
    df['MAJORNUM'] = pd.to_numeric(df['MAJORNUM'], errors='coerce')
    before_majornum = len(df)
    df = df[df['MAJORNUM'] == 1].copy()
    after_majornum = len(df)
    # ------------------------------------

    df['UNITID'] = pd.to_numeric(df['UNITID'], errors='coerce').astype('Int64')
    df['AWLEVEL'] = pd.to_numeric(df['AWLEVEL'], errors='coerce').astype('Int64')
    for col in ('CTOTALT', 'CTOTALM', 'CTOTALW'):
        # Nullable Int64 preserves suppressed cells as <NA> (never imputed).
        df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')

    # Flag 6-digit CIPCODEs vs aggregate-rollup rows (bare '99', '01', etc.).
    # The aggregator decides whether to drop non-6-digit rows based on
    # --include-residual; loader just tags them.
    df['IS_CIP_6DIGIT'] = df['CIPCODE'].apply(
        lambda s: bool(CIP_6DIGIT_PATTERN.match(s)) if isinstance(s, str) else False
    )
    n_6digit = int(df['IS_CIP_6DIGIT'].sum())
    n_aggregate = int((~df['IS_CIP_6DIGIT']).sum())

    console.log(
        f'  rows raw: {raw_rows:,}  →  MAJORNUM==1: {after_majornum:,}'
        f'  ({before_majornum - after_majornum:,} dropped as 2nd major)'
    )
    console.log(
        f'  CIPCODE breakdown after MAJORNUM filter: '
        f'6-digit programs: {n_6digit:,}  |  aggregate rollups (non-6-digit): {n_aggregate:,}'
    )
    return df


def load_varlist(dict_dir: Path) -> pd.DataFrame:
    """
    IPEDS data dictionary. The xlsx export ships with multiple sheets
    (Introduction, Varlist, Description, Frequencies, Statistics,
    Imputation values). The `Varlist` sheet holds the variable definitions;
    fall back to the first sheet for plain CSVs or unexpected layouts.
    """
    path = _find_file(dict_dir, 'varlist')
    if path is None:
        console.log('[yellow]⚠ varlist not found in data/dictionary/[/]')
        return pd.DataFrame()
    console.log(f'[cyan]varlist[/] ← {path.name}')

    sheet = None
    if path.suffix.lower() in ('.xlsx', '.xls'):
        try:
            sheets = pd.ExcelFile(path).sheet_names
        except Exception:
            sheets = []
        if 'Varlist' in sheets:
            sheet = 'Varlist'
        elif sheets:
            console.log(f'  [yellow]⚠ no "Varlist" sheet; using "{sheets[0]}"[/]')
    df = _read_table(path, sheet_name=sheet)
    console.log(f'  rows: {len(df):,}  cols: {df.shape[1]}'
                f'  sheet: {sheet or "default"}')
    return df


# CIP-SOC crosswalk canonical column names (downstream code joins on `CIPCODE`).
# Source workbook uses CIP2020Code/CIP2020Title/SOC2018Code/SOC2018Title.
_CROSSWALK_COLUMN_RENAMES = {
    'CIP2020Code': 'CIPCODE',
    'CIP2020Title': 'CIPTitle',
    'SOC2018Code': 'SOCCode',
    'SOC2018Title': 'SOCTitle',
    # Common alternate headers seen in other crosswalk distributions:
    'CIPCode': 'CIPCODE',
    'CIP Code': 'CIPCODE',
    'CIP Title': 'CIPTitle',
    'SOC Code': 'SOCCode',
    'SOC Title': 'SOCTitle',
}


def load_crosswalk(dict_dir: Path) -> pd.DataFrame:
    """
    CIP-SOC crosswalk.

    The published workbook contains 8 sheets — we want `CIP-SOC` (one row per
    CIP-SOC pair). Sheet 0 is a "File Guide" metadata sheet, not the data.
    SOC columns are retained for v2 labor market integration; do not drop.
    """
    path = _find_file(dict_dir, 'cip_soc_crosswalk')
    if path is None:
        console.log('[yellow]⚠ cip_soc_crosswalk not found in data/dictionary/[/]')
        return pd.DataFrame()
    console.log(f'[cyan]CIP-SOC crosswalk[/] ← {path.name}')

    sheet = None
    if path.suffix.lower() in ('.xlsx', '.xls'):
        try:
            sheets = pd.ExcelFile(path).sheet_names
        except Exception:
            sheets = []
        if 'CIP-SOC' in sheets:
            sheet = 'CIP-SOC'
        elif sheets:
            console.log(f'  [yellow]⚠ no "CIP-SOC" sheet; using "{sheets[0]}"[/]')

    df = _read_table(path, dtype=str, sheet_name=sheet)
    rename_map = {k: v for k, v in _CROSSWALK_COLUMN_RENAMES.items() if k in df.columns}
    if rename_map:
        df = df.rename(columns=rename_map)
    if 'CIPCODE' in df.columns:
        df['CIPCODE'] = df['CIPCODE'].astype(str).str.strip()
    console.log(f'  rows: {len(df):,}  cols: {df.shape[1]}'
                f'  sheet: {sheet or "default"}  renamed: {list(rename_map.values()) or "[none]"}')
    return df


# ---------------------------------------------------------------------------
# Orchestration + summary
# ---------------------------------------------------------------------------

def load_all(years_cfg: dict, raw_dir: Path, dict_dir: Path) -> Dict[str, object]:
    """Load every HD/C_A for configured years plus dictionary files."""
    years = years_cfg['years']
    return {
        'hd': {y: load_hd(y, raw_dir) for y in years},
        'ca': {y: load_ca(y, raw_dir) for y in years},  # MAJORNUM == 1 already applied
        'varlist': load_varlist(dict_dir),
        'crosswalk': load_crosswalk(dict_dir),
    }


def print_shape_summary(loaded: Dict[str, object]) -> None:
    table = Table(title='IPEDS Data Load Summary', show_lines=False)
    table.add_column('File / Source', style='cyan')
    table.add_column('Rows', justify='right')
    table.add_column('Cols', justify='right')
    table.add_column('Notes', style='dim')

    hd_dict = loaded['hd']  # type: ignore[assignment]
    ca_dict = loaded['ca']  # type: ignore[assignment]

    for y, df in sorted(hd_dict.items()):
        table.add_row(f'hd{y}', f'{len(df):,}', f'{df.shape[1]}', 'institutions')
    for y, df in sorted(ca_dict.items()):
        n_6 = int(df['IS_CIP_6DIGIT'].sum()) if 'IS_CIP_6DIGIT' in df.columns else 0
        table.add_row(
            f'c{y}_a',
            f'{len(df):,}',
            f'{df.shape[1]}',
            f'MAJORNUM==1 only · 6-digit rows: {n_6:,}',
        )
    if not loaded['varlist'].empty:  # type: ignore[union-attr]
        df = loaded['varlist']  # type: ignore[assignment]
        table.add_row('varlist', f'{len(df):,}', f'{df.shape[1]}', 'data dictionary')
    if not loaded['crosswalk'].empty:  # type: ignore[union-attr]
        df = loaded['crosswalk']  # type: ignore[assignment]
        table.add_row('cip_soc_crosswalk', f'{len(df):,}', f'{df.shape[1]}',
                      'SOC cols retained for v2')

    console.print(table)


# ---------------------------------------------------------------------------
# Smoke test entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    project_root = Path(__file__).resolve().parent.parent
    years_cfg = load_years_config(project_root / 'config' / 'years.yaml')
    console.rule('[bold]IPEDS Loader — Smoke Test[/]')
    console.print(f'Years configured: {years_cfg["years"]}  '
                  f'(CAGR {years_cfg["cagr_start_year"]} → {years_cfg["cagr_end_year"]})')
    loaded = load_all(
        years_cfg,
        raw_dir=project_root / 'data' / 'raw',
        dict_dir=project_root / 'data' / 'dictionary',
    )
    print_shape_summary(loaded)
