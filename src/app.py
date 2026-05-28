"""
app.py — Streamlit web interface for the IPEDS Completions tool.

Thin UI layer over the verified Python modules:
  loader → joiner → aggregator → reporter

Launch:
  python3 -m streamlit run src/app.py
  (or double-click "Launch IPEDS Tool.command")

The CLI (src/main.py) is unaffected and continues to work as before.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# Bump Pandas Styler's cell-count cap. Default is 262,144 cells; large national
# or all-state slices easily exceed that. Setting to 5M keeps the Styler path
# working for any realistic IPEDS query while staying well under memory limits.
pd.set_option('styler.render.max_elements', 5_000_000)

# Above this row count, the CAGR + Market View tabs skip the per-cell Styler
# (cell coloring / gradient) and render plain DataFrames. The browser-side cost
# of inline-styled HTML for tens of thousands of cells is what freezes the page
# on large queries (e.g., all CIPs × California). The Excel workbook download
# retains full coloring regardless.
STYLER_ROW_THRESHOLD = 2000

# Ensure project root + src/ are importable regardless of cwd.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / 'src'))

from aggregator import (  # noqa: E402
    apply_filters,
    compute_cagr_table,
    compute_market_view,
    compute_national_market_view,
    market_view_to_long,
    merge_selected_and_national_market_view,
)
from joiner import (  # noqa: E402
    build_cip_title_lookup,
    build_institution_metadata,
    join_all_years,
)
from loader import load_all, load_cip_filter_config, load_years_config  # noqa: E402
from reporter import (  # noqa: E402
    AWLEVEL_LABELS,
    CARNEGIE_LABELS,
    CONTROL_LABELS,
    ICLEVEL_LABELS,
    build_workbook,
)

# ── APU brand colors ──
APU_BRICK_RED = '#A8353A'
APU_BRIGHT_RED = '#D34147'
APU_LIGHT_GRAY = '#BCBDC0'
APU_BONE_WHITE = '#EDEBE8'

# ── Conditional formatting kept consistent with Excel output ──
# These are functional indicators (growth=green, decline=red), not APU branding.
# CAGR cell fills match reporter.py so the inline view matches the workbook.
CAGR_POS_FILL = '#C6EFCE'   # light green
CAGR_NEG_FILL = '#FFC7CE'   # light red
CAGR_NA_FILL = APU_LIGHT_GRAY

YEARS_YAML = PROJECT_ROOT / 'config' / 'years.yaml'
CIP_FILTER_YAML = PROJECT_ROOT / 'config' / 'cip_filter.yaml'
DATA_RAW = PROJECT_ROOT / 'data' / 'raw'
DATA_DICT = PROJECT_ROOT / 'data' / 'dictionary'
REPORTS_DIR = PROJECT_ROOT / 'output' / 'reports'


# ---------------------------------------------------------------------------
# Cached data load
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def cached_load_everything() -> dict:
    """
    Run loader + build metadata + CIP titles once per Streamlit session.

    Returns a dict ready to feed downstream:
      hd, ca, varlist, crosswalk, metadata, cip_titles, years_cfg, cip_cfg
    """
    years_cfg = load_years_config(YEARS_YAML)
    cip_cfg = load_cip_filter_config(CIP_FILTER_YAML)
    loaded = load_all(years_cfg, raw_dir=DATA_RAW, dict_dir=DATA_DICT)
    metadata = build_institution_metadata(loaded['hd'])
    cip_titles = build_cip_title_lookup(loaded['crosswalk'])
    return {
        'hd': loaded['hd'],
        'ca': loaded['ca'],
        'varlist': loaded['varlist'],
        'crosswalk': loaded['crosswalk'],
        'metadata': metadata,
        'cip_titles': cip_titles,
        'years_cfg': years_cfg,
        'cip_cfg': cip_cfg,
    }


# ---------------------------------------------------------------------------
# Label helpers — friendly labels for codes
# ---------------------------------------------------------------------------

def institution_label(unitid: int, metadata: pd.DataFrame) -> str:
    """'110635 — University of California-Berkeley (CA)'"""
    row = metadata[metadata['UNITID'] == unitid]
    if row.empty:
        return f'{unitid} — (not in HD)'
    r = row.iloc[0]
    state = r.get('STABBR') or ''
    name = r.get('INSTNM') or ''
    return f'{unitid} — {name} ({state})'


def cip_label(code: str, cip_titles: pd.DataFrame) -> str:
    """'51.3801 — Registered Nursing/Registered Nurse.'"""
    if cip_titles.empty:
        return code
    hit = cip_titles[cip_titles['CIPCODE'] == code]
    if hit.empty:
        return code
    title = hit.iloc[0].get('CIPTitle')
    return f'{code} — {title}' if title else code


def awlevel_label(code: int) -> str:
    label = AWLEVEL_LABELS.get(code)
    return f'{code} — {label}' if label else f'{code}'


def _awlevel_label_or_blank(v) -> str:
    """Map an AWLEVEL value to '5 — Bachelor\\'s degree'. Blank for NaN/missing."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    try:
        code = int(v)
    except (TypeError, ValueError):
        return str(v)
    label = AWLEVEL_LABELS.get(code)
    return f'{code} — {label}' if label else f'{code}'


# ---------------------------------------------------------------------------
# Pandas Styler for CAGR cells (matches Excel direct-fill behavior)
# ---------------------------------------------------------------------------

def _style_cagr_cell(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return f'background-color: {CAGR_NA_FILL};'
    if isinstance(v, (int, float)):
        if v > 0:
            return f'background-color: {CAGR_POS_FILL};'
        if v < 0:
            return f'background-color: {CAGR_NEG_FILL};'
    return ''


def styled_cagr_table(df: pd.DataFrame) -> 'pd.io.formats.style.Styler':
    """
    Expects CAGR pre-multiplied by 100 (i.e., percent value, not decimal).
    Per-cell colored fills are applied via Styler.applymap. The CAGR column's
    percent FORMAT is handled by the caller via st.column_config — Styler.format
    is unreliable for numeric columns in Streamlit's st.dataframe (it gets
    silently overridden by Streamlit's column-type detection).
    """
    sty = df.style
    if 'CAGR' in df.columns:
        sty = sty.applymap(_style_cagr_cell, subset=['CAGR'])
    year_cols = [
        c for c in df.columns
        if c.endswith(' Completions') or c.startswith('CTOTALT_')
    ]
    sty = sty.format({c: _fmt_int_blank for c in year_cols})
    return sty


def _fmt_pct(v) -> str:
    """Format a decimal CAGR (e.g. 0.0148) as '+1.5%'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 'N/A'
    return f'{v * 100:+.1f}%'


def _fmt_int_blank(v) -> str:
    """Format an integer cell; blank for NaN (don't impute to 0)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    try:
        return f'{int(v):,}'
    except (TypeError, ValueError):
        return str(v)


def _render_search_summary(
    mv_wide: pd.DataFrame, start_year: int, end_year: int, n_selected_inst: int,
) -> None:
    """
    Headline panel above the result tabs. For each program in the user's
    filter, show selected vs national side by side with BOTH CAGRs labeled
    explicitly (Completions CAGR and Programs CAGR).
    """
    if mv_wide.empty:
        st.info('No programs match your filters. Try widening the CIP or award-level selection.')
        return

    sel_compl_col = f'SUM_CTOTALT_{end_year}'
    nat_compl_col = f'NAT_SUM_CTOTALT_{end_year}'
    sel_prog_col = f'INST_COUNT_{end_year}'
    nat_prog_col = f'NAT_INST_COUNT_{end_year}'

    sel_total = int(mv_wide.get(sel_compl_col, pd.Series(dtype=float)).fillna(0).sum())
    nat_total = int(mv_wide.get(nat_compl_col, pd.Series(dtype=float)).fillna(0).sum())
    # Number of programs = sum of institution-counts across all (CIP × Award Level)
    # rows in the user's filter. Each (institution, CIP, Award Level) offering = 1 program.
    sel_programs = int(
        mv_wide.get(sel_prog_col, pd.Series(dtype=float)).fillna(0).sum()
    )
    nat_programs = int(
        mv_wide.get(nat_prog_col, pd.Series(dtype=float)).fillna(0).sum()
    )

    cols = st.columns(4)
    cols[0].metric(
        f'Selected — {end_year} completions',
        f'{sel_total:,}',
        help=(
            f'Total completions in {end_year} across your selected institutions and '
            f'the CIP codes / award levels in your filter.'
        ),
    )
    cols[1].metric(
        f'National — {end_year} completions',
        f'{nat_total:,}',
        help=(
            f'Total completions in {end_year} across every U.S. institution reporting '
            f'to IPEDS for the same CIP codes / award levels.'
        ),
    )
    cols[2].metric(
        f'Selected — number of programs offered ({end_year})',
        f'{sel_programs:,}',
        help=(
            f'Total programs offered in {end_year} within your selected region across '
            f'the CIP codes AND award levels in your filter. Each (institution × CIP × '
            f'Award Level) offering counts as 1 program — an institution offering the '
            f'same CIP at multiple of your selected award levels is counted once per '
            f'level. Award levels you did NOT select are excluded.'
        ),
    )
    cols[3].metric(
        f'National — number of programs offered ({end_year})',
        f'{nat_programs:,}',
        help=(
            f'Total programs offered in {end_year} nationally across the same CIP codes '
            f'AND award levels in your filter. Same counting convention as the Selected '
            f'column to its left.'
        ),
    )

    # Per-program callouts — both CAGRs explicitly labeled.
    with st.expander(f'Per-program summary ({start_year} → {end_year})', expanded=True):
        for _, row in mv_wide.iterrows():
            cip = row.get('CIPCODE') or ''
            title = row.get('CIPTitle') or ''
            aw = row.get('AWLEVEL')
            aw_lbl = AWLEVEL_LABELS.get(int(aw), str(aw)) if pd.notna(aw) else ''

            st.markdown(
                f"**{cip} — {title}**  ·  *{aw_lbl}*  \n"
                f"&nbsp;&nbsp;**Selected** in {end_year}: "
                f"**{_fmt_int_blank(row.get(sel_prog_col)) or '—'}** programs (institutions offering), "
                f"**{_fmt_int_blank(row.get(sel_compl_col)) or '—'}** completions  \n"
                f"&nbsp;&nbsp;&nbsp;&nbsp;{start_year}→{end_year} CAGR — "
                f"Completions {_fmt_pct(row.get('MARKET_CAGR'))} · "
                f"Programs {_fmt_pct(row.get('PROGRAMS_CAGR'))}  \n"
                f"&nbsp;&nbsp;**National** in {end_year}: "
                f"**{_fmt_int_blank(row.get(nat_prog_col)) or '—'}** programs (institutions offering), "
                f"**{_fmt_int_blank(row.get(nat_compl_col)) or '—'}** completions  \n"
                f"&nbsp;&nbsp;&nbsp;&nbsp;{start_year}→{end_year} CAGR — "
                f"Completions {_fmt_pct(row.get('NAT_MARKET_CAGR'))} · "
                f"Programs {_fmt_pct(row.get('NAT_PROGRAMS_CAGR'))}",
                unsafe_allow_html=True,
            )


def _render_market_view(mv_long: pd.DataFrame, start_year: int, end_year: int) -> None:
    """
    Render the long-format Market View tab.

    The long-format DataFrame has one row per (CIP × Award Level × Geography × Metric).

    NOTE on CAGR rendering: Streamlit's `st.dataframe` silently overrides
    Styler.format() for numeric columns and shows the raw number. We work
    around this by:
      1. Pre-multiplying CAGR to a percent value so the background gradient
         operates on the same scale as the displayed text.
      2. Using `st.column_config.NumberColumn(format='%+.1f%%')` to force
         Streamlit to format the value as a signed percent at render time.
         column_config wins where Styler.format loses.
    """
    if mv_long.empty:
        st.info('No rows match your filters.')
        return

    display = mv_long.copy()
    year_cols = [c for c in display.columns if c.isdigit()]

    if 'Award Level' in display.columns:
        display['Award Level'] = display['Award Level'].apply(_awlevel_label_or_blank)

    if 'CAGR' in display.columns:
        display['CAGR'] = pd.to_numeric(display['CAGR'], errors='coerce') * 100

    col_config = {}
    if 'CAGR' in display.columns:
        col_config['CAGR'] = st.column_config.NumberColumn(
            'CAGR',
            help=f'{start_year} → {end_year} compound annual growth rate for this row\'s metric.',
            format='%+.1f%%',
        )

    # On very large result sets, skip the Styler — generating an inline-styled
    # HTML cell for every row freezes the browser. Excel download keeps colors.
    if len(display) > STYLER_ROW_THRESHOLD:
        st.caption(
            f'⚡ Rendering {len(display):,} rows without color gradient for performance. '
            f'The Excel download retains full CAGR coloring.'
        )
        st.dataframe(
            display, hide_index=True, use_container_width=True,
            column_config=col_config,
        )
        return

    sty = display.style
    if 'CAGR' in display.columns:
        sty = sty.background_gradient(
            subset=['CAGR'], cmap='RdYlGn', vmin=-30, vmax=30,
        )
    sty = sty.format({c: _fmt_int_blank for c in year_cols})

    st.dataframe(
        sty, hide_index=True, use_container_width=True,
        column_config=col_config,
    )


# ---------------------------------------------------------------------------
# Institution resolution from sidebar widgets
# ---------------------------------------------------------------------------

def resolve_unitids_from_ui(
    selection_mode: str,
    metadata: pd.DataFrame,
    states: List[str],
    control_filters: List[int],
    iclevel_filters: List[int],
    name_query: str,
    chosen_unitids: List[int],
    ca_dict: Optional[Dict[int, pd.DataFrame]] = None,
    cip_filter: Optional[List[str]] = None,
    awlevel_filter: Optional[List[int]] = None,
) -> Tuple[List[int], str]:
    """
    Returns (sorted unique UNITID list, label for filename).

    Modes:
      - 'All institutions nationally' → every UNITID with completions in the
        program filters (CIPs/AWLEVELs). Requires ca_dict. If no program filter
        is set, returns every UNITID with any C_A row.
      - 'By state(s)' → multiselect; union institutions whose latest HD STABBR
        is in `states`.
      - 'By institution name' → substring match on INSTNM.
      - 'Specific UNITIDs' → as-is.
      - 'From institutions.csv config' → loads config/institutions.csv.
    """

    if selection_mode == 'All institutions nationally':
        if ca_dict is None:
            return [], 'national'
        cip_set = {str(c).strip() for c in (cip_filter or [])} or None
        aw_set = {int(a) for a in (awlevel_filter or [])} or None
        uids: set = set()
        for df in ca_dict.values():
            sub = df
            if 'IS_CIP_6DIGIT' in sub.columns:
                sub = sub[sub['IS_CIP_6DIGIT'] == True]  # noqa: E712
            if cip_set:
                sub = sub[sub['CIPCODE'].isin(cip_set)]
            if aw_set:
                sub = sub[sub['AWLEVEL'].isin(aw_set)]
            uids.update(int(u) for u in sub['UNITID'].dropna())
        return sorted(uids), 'national'

    if selection_mode == 'By state(s)':
        if not states:
            return [], 'custom'
        df = metadata[metadata['STABBR'].isin(states)]
        if control_filters:
            df = df[df['CONTROL'].isin(control_filters)]
        if iclevel_filters:
            df = df[df['ICLEVEL'].isin(iclevel_filters)]
        label = '_'.join(sorted(states))
        if control_filters and len(control_filters) == 1:
            label += f'_ctrl{control_filters[0]}'
        if iclevel_filters and len(iclevel_filters) == 1:
            label += f'_lvl{iclevel_filters[0]}'
        return sorted(int(u) for u in df['UNITID'].dropna()), label

    if selection_mode == 'By institution name':
        if not name_query.strip():
            return [], 'custom'
        q = name_query.strip().lower()
        mask = metadata['INSTNM'].astype(str).str.lower().str.contains(q, na=False, regex=False)
        return sorted(int(u) for u in metadata[mask]['UNITID'].dropna()), 'search'

    if selection_mode == 'Specific UNITIDs':
        return sorted(set(int(u) for u in chosen_unitids)), 'custom'

    if selection_mode == 'From institutions.csv config':
        csv_path = PROJECT_ROOT / 'config' / 'institutions.csv'
        if not csv_path.exists():
            return [], 'custom'
        cfg = pd.read_csv(csv_path)
        if 'UNITID' not in cfg.columns:
            return [], 'custom'
        ids = [int(u) for u in cfg['UNITID'].dropna()]
        in_meta = set(metadata['UNITID'].dropna().astype(int))
        return sorted(set(ids) & in_meta), 'custom'

    return [], 'custom'


# ---------------------------------------------------------------------------
# Sidebar UI
# ---------------------------------------------------------------------------

def render_sidebar(everything: dict) -> Optional[dict]:
    """
    Render the sidebar form. Returns a dict of selections when 'Generate Report'
    is clicked, or None on every other render.
    """
    metadata: pd.DataFrame = everything['metadata']
    cip_titles: pd.DataFrame = everything['cip_titles']
    cip_cfg: dict = everything['cip_cfg']
    years_cfg: dict = everything['years_cfg']
    ca_dict: Dict[int, pd.DataFrame] = everything['ca']

    st.sidebar.header('Selection')

    selection_mode = st.sidebar.radio(
        'Mode',
        options=[
            'All institutions nationally',
            'By state(s)',
            'By institution name',
            'Specific UNITIDs',
            'From institutions.csv config',
        ],
        index=1,
        help='"All institutions nationally" picks every institution with completions in the CIP codes you choose below. '
             'Use it to see who offers a given program nationwide.',
    )

    states_selected: List[str] = []
    control_filters: List[int] = []
    iclevel_filters: List[int] = []
    name_query = ''
    chosen_unitids: List[int] = []

    if selection_mode == 'All institutions nationally':
        st.sidebar.info(
            'Institutions will be auto-selected based on the **CIP codes** you choose below — '
            'every institution with at least one completion in those programs across any year. '
            'Leave the CIP filter empty to include every institution with any completions.',
            icon='💡',
        )

    elif selection_mode == 'By state(s)':
        all_states = sorted(metadata['STABBR'].dropna().unique().tolist())
        states_selected = st.sidebar.multiselect(
            'State(s)',
            options=all_states,
            default=['CA'] if 'CA' in all_states else [],
            help='Pick one or more states. Institutions from all selected states are unioned.',
        )
        st.sidebar.markdown('**Control** (leave all unchecked to include all)')
        for code, label in CONTROL_LABELS.items():
            if st.sidebar.checkbox(label, key=f'ctl_{code}'):
                control_filters.append(code)
        st.sidebar.markdown('**Level** (leave all unchecked to include all)')
        for code, label in ICLEVEL_LABELS.items():
            if st.sidebar.checkbox(label, key=f'lvl_{code}'):
                iclevel_filters.append(code)

    elif selection_mode == 'By institution name':
        name_query = st.sidebar.text_input('Search institution name', value='')
        if name_query.strip():
            q = name_query.strip().lower()
            preview = metadata[
                metadata['INSTNM'].astype(str).str.lower().str.contains(q, na=False, regex=False)
            ].head(50)
            st.sidebar.caption(f'{len(preview)} match (showing up to 50)')
            st.sidebar.dataframe(
                preview[['UNITID', 'INSTNM', 'STABBR']].reset_index(drop=True),
                hide_index=True, use_container_width=True, height=200,
            )

    elif selection_mode == 'Specific UNITIDs':
        all_unitids = sorted(int(u) for u in metadata['UNITID'].dropna())
        label_map = {institution_label(u, metadata): u for u in all_unitids}
        picks = st.sidebar.multiselect(
            'Pick institutions (search by name or UNITID)',
            options=list(label_map.keys()),
            default=[],
        )
        chosen_unitids = [label_map[p] for p in picks]

    elif selection_mode == 'From institutions.csv config':
        csv_path = PROJECT_ROOT / 'config' / 'institutions.csv'
        if csv_path.exists():
            cfg_df = pd.read_csv(csv_path)
            n = cfg_df['UNITID'].notna().sum() if 'UNITID' in cfg_df.columns else 0
            st.sidebar.caption(f'{int(n)} UNITIDs in config/institutions.csv')
        else:
            st.sidebar.warning('config/institutions.csv not found.')

    st.sidebar.divider()
    st.sidebar.header('Program filters')

    # CIP multiselect with searchable labels (full set ~2,143 — Streamlit handles).
    cip_label_map = {
        cip_label(c, cip_titles): c
        for c in sorted(cip_titles['CIPCODE'].dropna().unique())
    } if not cip_titles.empty else {}
    default_cips = []
    for c in cip_cfg.get('cip_codes', []):
        for lbl, code in cip_label_map.items():
            if code == c:
                default_cips.append(lbl)
                break
    cip_picks = st.sidebar.multiselect(
        'CIP codes (leave empty for all)',
        options=list(cip_label_map.keys()),
        default=default_cips,
    )
    cip_codes = [cip_label_map[p] for p in cip_picks]

    aw_label_map = {awlevel_label(code): code for code in sorted(AWLEVEL_LABELS.keys())}
    default_aws = [awlevel_label(c) for c in cip_cfg.get('award_levels', []) if c in AWLEVEL_LABELS]
    aw_picks = st.sidebar.multiselect(
        'Award levels (leave empty for all)',
        options=list(aw_label_map.keys()),
        default=default_aws,
    )
    award_levels = [aw_label_map[p] for p in aw_picks]

    include_residual = st.sidebar.checkbox(
        'Include CIP 99 residual rollups', value=False,
        help='IPEDS uses bare 2-digit codes like "99" for institution-level totals. Excluded by default.',
    )

    # Live count — calculated AFTER program filters because national mode depends on them.
    candidate_ids, label_preview = resolve_unitids_from_ui(
        selection_mode, metadata, states_selected, control_filters, iclevel_filters,
        name_query, chosen_unitids,
        ca_dict=ca_dict, cip_filter=cip_codes, awlevel_filter=award_levels,
    )
    st.sidebar.divider()
    st.sidebar.markdown(f'**Selected: {len(candidate_ids):,} institutions**')
    if selection_mode == 'All institutions nationally' and len(candidate_ids) > 200:
        st.sidebar.caption(
            f'Large national selection — Excel + UI rendering may take a few seconds.'
        )

    st.sidebar.caption(
        f'Years: {years_cfg["years"][0]}–{years_cfg["years"][-1]}  '
        f'(CAGR {years_cfg["cagr_start_year"]} → {years_cfg["cagr_end_year"]})'
    )
    st.sidebar.caption('Edit config/years.yaml to change the year range.')

    generate = st.sidebar.button('Generate Report', type='primary', use_container_width=True)

    if not generate:
        return None

    if not candidate_ids:
        st.sidebar.error('No institutions selected. Pick at least one before generating.')
        return None

    return {
        'unitids': candidate_ids,
        'label': label_preview,
        'cip_codes': cip_codes,
        'award_levels': award_levels,
        'include_residual': include_residual,
        'selection_mode': selection_mode,
    }


# ---------------------------------------------------------------------------
# Main area — runs the pipeline and renders result tabs
# ---------------------------------------------------------------------------

def render_results(selections: dict, everything: dict) -> None:
    metadata: pd.DataFrame = everything['metadata']
    cip_titles: pd.DataFrame = everything['cip_titles']
    ca_dict: Dict[int, pd.DataFrame] = everything['ca']
    varlist_df: pd.DataFrame = everything['varlist']
    years_cfg: dict = everything['years_cfg']
    start_year = years_cfg['cagr_start_year']
    end_year = years_cfg['cagr_end_year']

    with st.spinner('Joining + filtering + aggregating…'):
        joined = join_all_years(
            ca_dict, metadata, cip_titles,
            selected_unitids=selections['unitids'], quiet=True,
        )
        filtered = apply_filters(
            joined,
            cip_codes=selections['cip_codes'],
            award_levels=selections['award_levels'],
            include_residual=selections['include_residual'],
            quiet=True,
        )
        cagr_df = compute_cagr_table(filtered, start_year=start_year, end_year=end_year)
        selected_mv = compute_market_view(filtered, start_year=start_year, end_year=end_year)
        national_mv = compute_national_market_view(
            ca_dict, cip_titles,
            cip_codes=selections['cip_codes'],
            award_levels=selections['award_levels'],
            include_residual=selections['include_residual'],
            start_year=start_year, end_year=end_year, quiet=True,
        )
        mv_wide = merge_selected_and_national_market_view(selected_mv, national_mv)
        mv_long = market_view_to_long(mv_wide, start_year=start_year, end_year=end_year)

        # Institutions tab should only show institutions that contributed data —
        # i.e., had at least one matching CIP × Award Level completion in any year
        # after filters. Schools picked in the selection but with no matching rows
        # would be misleading (the user said this was puzzling).
        unitids_with_data: set = set()
        for year_df in filtered.values():
            unitids_with_data.update(
                int(u) for u in year_df['UNITID'].dropna()
            )
        institutions_view = metadata[metadata['UNITID'].isin(unitids_with_data)]
        n_selected = len(selections['unitids'])
        n_with_data = len(institutions_view)

    # ── Search Summary panel — headline numbers per program (selected vs national) ──
    _render_search_summary(mv_wide, start_year, end_year, len(selections['unitids']))

    # ── Excel download ──
    with st.spinner('Building Excel workbook…'):
        out_path = build_workbook(
            output_dir=REPORTS_DIR,
            label=selections['label'],
            institutions_df=institutions_view,
            completions_by_year=filtered,
            cagr_df=cagr_df,
            market_view_df=mv_long,
            varlist_df=varlist_df,
            only_latest_completions=False,
        )
    st.download_button(
        label=f'⬇ Download Excel workbook  ({out_path.name})',
        data=out_path.read_bytes(),
        file_name=out_path.name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
    st.caption(f'Workbook also saved to `{out_path.relative_to(PROJECT_ROOT)}`.')

    # ── Result tabs ──
    tabs = st.tabs(['Institutions', 'CAGR by Institution', 'Market View',
                    'Definitions'])

    with tabs[0]:
        st.subheader('Institutions with completions in your filtered programs')
        no_data_count = n_selected - n_with_data
        if no_data_count > 0:
            st.caption(
                f'You selected **{n_selected:,}** institutions; **{n_with_data:,}** of them '
                f'had at least one completion in your filtered CIP × Award Level across '
                f'{start_year}–{end_year}. {no_data_count:,} selected institutions had no '
                f'matching completions and are not shown.'
            )
        else:
            st.caption(
                f'All **{n_with_data:,}** selected institutions had at least one matching '
                f'completion in {start_year}–{end_year}.'
            )
        view = institutions_view.copy()
        if 'CONTROL' in view.columns:
            view['Control'] = view['CONTROL'].map(CONTROL_LABELS)
        if 'ICLEVEL' in view.columns:
            view['Level'] = view['ICLEVEL'].map(ICLEVEL_LABELS)
        if 'CARNEGIE' in view.columns:
            view['Carnegie Classification'] = view['CARNEGIE'].map(CARNEGIE_LABELS)
        show_cols = [c for c in (
            'UNITID', 'INSTNM', 'STABBR', 'CITY',
            'Control', 'Level', 'CARNEGIE', 'Carnegie Classification',
            'HD_SOURCE_YEAR',
        ) if c in view.columns]
        st.dataframe(
            view[show_cols].rename(columns={
                'INSTNM': 'Institution', 'STABBR': 'State', 'CITY': 'City',
                'CARNEGIE': 'Carnegie Code', 'HD_SOURCE_YEAR': 'HD Source Year',
            }).sort_values('UNITID').reset_index(drop=True),
            hide_index=True, use_container_width=True,
        )

    with tabs[1]:
        st.subheader('CAGR by institution × CIP × award level')
        if cagr_df.empty:
            st.info('No rows match your filters. Try widening the CIP or award-level selection.')
        else:
            flag_counts = cagr_df['Flag'].value_counts().to_dict()
            st.caption('  •  '.join(f'{k}: {v}' for k, v in flag_counts.items()))
            display = cagr_df.copy()
            rename = {f'CTOTALT_{y}': f'{y} Completions' for y in range(start_year, end_year + 1)}
            rename.update({
                'INSTNM': 'Institution',
                'STABBR': 'State',
                'CIPTitle': 'CIP Title',
                'AWLEVEL': 'Award Level',
            })
            display = display.rename(columns=rename)
            if 'Award Level' in display.columns:
                display['Award Level'] = display['Award Level'].apply(_awlevel_label_or_blank)
            # Pre-multiply CAGR to percent value so the gradient + column_config
            # format both operate on the same scale.
            if 'CAGR' in display.columns:
                display['CAGR'] = pd.to_numeric(display['CAGR'], errors='coerce') * 100
            cagr_col_config = {
                'CAGR': st.column_config.NumberColumn(
                    'CAGR',
                    help=f'{start_year} → {end_year} compound annual growth rate of completions.',
                    format='%+.1f%%',
                ),
            }

            # On very large result sets, skip the per-cell Styler colors to
            # keep the browser responsive. Excel download keeps full coloring.
            if len(display) > STYLER_ROW_THRESHOLD:
                st.caption(
                    f'⚡ Rendering {len(display):,} rows without green/red cell coloring '
                    f'for performance. The Excel download retains full CAGR highlighting.'
                )
                st.dataframe(
                    display, hide_index=True, use_container_width=True,
                    column_config=cagr_col_config,
                )
            else:
                st.dataframe(
                    styled_cagr_table(display),
                    hide_index=True, use_container_width=True,
                    column_config=cagr_col_config,
                )

    with tabs[2]:
        st.subheader('Market view — selected institutions vs. national')
        st.caption(
            f'Four rows per program (CIP × Award Level): Selected/Completions, '
            f'Selected/Programs, National/Completions, National/Programs. The Metric column '
            f'states what each row measures, and the CAGR column is the {start_year}→{end_year} '
            f'CAGR for that metric. "Programs" = number of institutions offering this CIP × Award Level.'
        )
        if mv_long.empty:
            st.info('No rows match your filters.')
        else:
            _render_market_view(mv_long, start_year, end_year)

    with tabs[3]:
        st.subheader('Definitions')
        st.caption(
            'IPEDS variable labels + Award Level / Control / ICLEVEL / Carnegie codes + CAGR flag meanings.'
        )
        rows = []
        if not varlist_df.empty:
            for _, r in varlist_df.iterrows():
                rows.append({'Category': 'Variable',
                             'Code': str(r.get('varName') or ''),
                             'Label': str(r.get('varTitle') or ''),
                             'Source': 'IPEDS varlist'})
        for code, label in AWLEVEL_LABELS.items():
            rows.append({'Category': 'Award Level', 'Code': code, 'Label': label,
                         'Source': 'SPEC §AWLEVEL + IPEDS legacy codes'})
        for code, label in CONTROL_LABELS.items():
            rows.append({'Category': 'Control', 'Code': code, 'Label': label, 'Source': 'SPEC §HD'})
        for code, label in ICLEVEL_LABELS.items():
            rows.append({'Category': 'Level (ICLEVEL)', 'Code': code, 'Label': label, 'Source': 'SPEC §HD'})
        for code in sorted(CARNEGIE_LABELS.keys()):
            rows.append({'Category': 'Carnegie Classification', 'Code': code,
                         'Label': CARNEGIE_LABELS[code],
                         'Source': '2021 Carnegie Basic Classification'})
        for code, desc in [
            ('OK', 'Both endpoints > 0; CAGR computed normally.'),
            ('New Program', 'Start completions = 0; CAGR undefined.'),
            ('Program Ended', 'Start > 0 and end = 0; CAGR shown as -100%.'),
            ('Missing Data', 'Either endpoint is suppressed (<3) or absent; CAGR N/A.'),
        ]:
            rows.append({'Category': 'CAGR Flag', 'Code': code, 'Label': desc,
                         'Source': 'SPEC §Calculations'})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ---------------------------------------------------------------------------
# App entry point
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title='APU · IPEDS Completions',
        page_icon='📊',
        layout='wide',
        initial_sidebar_state='expanded',
    )

    st.markdown(
        f"<h1 style='color: {APU_BRICK_RED}; margin-bottom: 0;'>"
        "IPEDS Completions Analysis</h1>"
        "<p style='color: #77787B; margin-top: 0.25rem;'>"
        "Azusa Pacific University · Strategic Planning</p>",
        unsafe_allow_html=True,
    )

    # First-load: long spinner while we read all the IPEDS CSVs.
    with st.spinner('Loading IPEDS data (one-time per session, ~10 seconds)…'):
        try:
            everything = cached_load_everything()
        except FileNotFoundError as e:
            st.error(f'Data file missing: {e}')
            st.info(
                'Check that every year listed in `config/years.yaml` has a matching '
                '`hd{year}` and `c{year}_a` file in `data/raw/`.'
            )
            st.stop()
        except ValueError as e:
            st.error(f'Configuration error: {e}')
            st.stop()

    selections = render_sidebar(everything)
    if selections is None:
        st.info(
            'Pick a selection mode and program filters in the sidebar, then click '
            '**Generate Report**.'
        )
        return

    render_results(selections, everything)


if __name__ == '__main__':
    main()
