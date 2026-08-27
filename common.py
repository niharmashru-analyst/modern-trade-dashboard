"""
COMMON — shared helpers for the Dispatch Tracker dashboard suite.

Used by Home.py and every page in pages/. Streamlit adds the app's
root directory (where Home.py lives) to sys.path for every page, so
`import common` works from anywhere in pages/ without any package
setup.

What lives here (and why it used to live in 4 separate copies):
  - inject_theme()          one visual language for every page
  - get_secret()            st.secrets.get() that never raises
  - fetch_bytes()           SharePoint download + friendly errors
  - read_excel_sheet() /
    read_excel_all_sheets() cheap single-sheet vs. full-workbook read
  - type_columns() /
    build_column_config()   date/currency/percent inference + display
  - normalize_key()         EAN/SKU int-vs-float-string matching
  - resolve_col() /
    resolve_and_rename()    tolerate small header drift (case/spacing)
"""
from __future__ import annotations

import io
from typing import Optional

import pandas as pd
import requests
import streamlit as st

# ================================================================
# THEME — call inject_theme() once near the top of each page, right
# after st.set_page_config(...). Pass extra_css for anything page-
# specific (e.g. a page's own accent color or chart tweaks).
# ================================================================
_FONTS_IMPORT = (
    "https://fonts.googleapis.com/css2?"
    "family=Barlow+Condensed:wght@500;600;700&"
    "family=Inter:wght@400;500;600;700&"
    "family=IBM+Plex+Mono:wght@400;500;600&display=swap"
)

_BASE_CSS = f"""
<style>
@import url('{_FONTS_IMPORT}');

html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
.stApp {{ background: #F4F6FA; }}

h1, h2, h3 {{
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important;
    color: #132238 !important;
}}

div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: #fff; border-radius: 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,.08);
}}
div[data-testid="stVerticalBlockBorderWrapper"] h4 {{ margin-bottom: 4px; }}

div[data-testid="stMetric"] {{
    background: #fff; border: 1px solid #E4E7ED; border-radius: 14px;
    padding: 14px 16px; box-shadow: 0 1px 3px rgba(0,0,0,.08);
}}
div[data-testid="stMetricLabel"] {{
    font-family: 'IBM Plex Mono', monospace !important; font-size: 10.5px !important;
    letter-spacing: 1.2px; text-transform: uppercase; color: #8B93A3 !important;
}}
div[data-testid="stMetricValue"] {{
    font-family: 'Barlow Condensed', sans-serif !important; font-weight: 700 !important;
    color: #132238 !important;
}}

.stCaption, [data-testid="stCaptionContainer"] {{ font-family: 'Inter', sans-serif; color: #8B93A3; }}
[data-testid="stDataFrame"] * {{ font-family: 'Inter', sans-serif; }}
button[data-baseweb="tab"] {{ color: #132238 !important; }}
</style>
"""


def inject_theme(extra_css: str = "") -> None:
    """Apply the shared look (fonts, background, card/metric styling)
    used across Home and every dashboard page. `extra_css` is raw CSS
    appended after the base theme, for page-specific overrides."""
    st.markdown(_BASE_CSS, unsafe_allow_html=True)
    if extra_css:
        st.markdown(f"<style>{extra_css}</style>", unsafe_allow_html=True)


# ================================================================
# SECRETS — st.secrets.get(key, default) raises instead of returning
# the default when no secrets.toml exists at all (only hit in local
# dev before the file's been created). get_secret() makes that case
# behave the same as "not set", so pages can rely on one code path.
# ================================================================
def get_secret(key: str, default: str = "") -> str:
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


# ================================================================
# LIVE FILE FETCH — SharePoint/OneDrive share link -> raw bytes.
# ================================================================
class FetchError(Exception):
    """Raised for any problem getting a live SharePoint file. The
    message is written to be shown directly via st.error(str(e))."""


def sharepoint_download_url(url: str) -> str:
    """Turn a SharePoint/OneDrive 'view' share link into a direct
    download link by appending download=1, so requests gets the raw
    file instead of the web viewer page."""
    if "download=1" in url:
        return url
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}download=1"


def fetch_bytes(url: str, timeout: int = 30) -> bytes:
    """Download a SharePoint/OneDrive file, with a specific, friendly
    message for each failure mode this app actually hits in practice:
    a slow/unreachable link, an expired/wrong link (404/403), or a
    login page coming back instead of the file because the share
    link isn't set to 'Anyone with the link can view'."""
    try:
        resp = requests.get(sharepoint_download_url(url), timeout=timeout)
    except requests.exceptions.Timeout:
        raise FetchError(
            f"Timed out after {timeout}s waiting for the file. "
            "SharePoint or the network may be slow right now — try again."
        )
    except requests.exceptions.ConnectionError:
        raise FetchError(
            "Couldn't reach SharePoint — check your internet connection "
            "and that the link is correct."
        )
    except requests.exceptions.RequestException as e:
        raise FetchError(f"Request failed: {e}")

    if resp.status_code == 404:
        raise FetchError(
            "SharePoint returned 404 (file not found) — the share link "
            "may have expired or been revoked."
        )
    if resp.status_code == 403:
        raise FetchError(
            "SharePoint returned 403 (access denied) — set the share "
            "link to 'Anyone with the link can view'."
        )
    try:
        resp.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise FetchError(f"SharePoint returned an error: {e}")

    if "html" in resp.headers.get("Content-Type", "").lower():
        raise FetchError(
            "Got a login/redirect page instead of the Excel file — the "
            "share link likely needs broader access ('Anyone with the "
            "link can view'), or it has expired."
        )
    return resp.content


def read_excel_sheet(raw: bytes, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Read exactly one sheet from an in-memory workbook. sheet_name
    of None reads only the FIRST sheet (by position) — cheap, unlike
    asking pandas for every sheet just to keep one of them."""
    actual_sheet = 0 if sheet_name is None else sheet_name
    df = pd.read_excel(io.BytesIO(raw), sheet_name=actual_sheet)
    if isinstance(df, dict):  # only if sheet_name was itself a list
        df = next(iter(df.values()))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_excel_all_sheets(raw: bytes) -> dict:
    """Read every sheet in the workbook, e.g. for the Fill Rate page
    where each month lives on its own tab."""
    wb = pd.read_excel(io.BytesIO(raw), sheet_name=None)
    for sdf in wb.values():
        sdf.columns = [str(c).strip() for c in sdf.columns]
    return wb


# ================================================================
# COLUMN TYPING — shared date/currency/percent inference, since the
# source sheets mix text-formatted currency ("₹ 1,827"), plain
# numbers, dates, and percentages all in one flat table.
# ================================================================
DATE_COL_HINTS = ["date"]
CURRENCY_COL_HINTS = ["value", "lacs", "sale loss", "amt"]
PERCENT_COL_HINTS = ["%"]


def looks_like(col: str, hints: list) -> bool:
    c = col.lower()
    return any(h in c for h in hints)


def type_columns(
    df: pd.DataFrame,
    date_hints: list = DATE_COL_HINTS,
    currency_hints: list = CURRENCY_COL_HINTS,
    percent_hints: list = PERCENT_COL_HINTS,
) -> pd.DataFrame:
    """Returns a copy of df with date/currency/percent columns (by
    name-hint match) converted to proper dtypes."""
    df = df.copy()
    for col in df.columns:
        if looks_like(col, date_hints):
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        elif looks_like(col, currency_hints):
            cleaned = (
                df[col].astype(str)
                .str.replace("₹", "", regex=False)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(cleaned, errors="coerce")
        elif looks_like(col, percent_hints):
            cleaned = df[col].astype(str).str.replace("%", "", regex=False).str.strip()
            df[col] = pd.to_numeric(cleaned, errors="coerce")
    return df


def build_column_config(
    columns,
    date_hints: list = DATE_COL_HINTS,
    currency_hints: list = CURRENCY_COL_HINTS,
    percent_hints: list = PERCENT_COL_HINTS,
) -> dict:
    """st.column_config dict for `columns`, using the same hints as
    type_columns() so display formatting always matches the dtype."""
    config = {}
    for c in columns:
        if looks_like(c, date_hints):
            config[c] = st.column_config.DateColumn(c, format="DD-MM-YYYY")
        elif looks_like(c, currency_hints):
            config[c] = st.column_config.NumberColumn(c, format="₹ %.2f")
        elif looks_like(c, percent_hints):
            config[c] = st.column_config.NumberColumn(c, format="%.1f%%")
    return config


# ================================================================
# SKU/EAN KEY NORMALIZATION — Stock Gap Dashboard matches order rows
# to stock rows on this.
# ================================================================
def normalize_key(val) -> str:
    """Turns any EAN/SKU value into a consistent lookup key. Excel
    often stores the same code as an int in one file and a float
    (e.g. 8906121646924.0) in another — usually because one blank
    cell elsewhere in the column forced pandas to upcast the whole
    column to float. Without this, those two would never match."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        try:
            float(s)
            s = s[:-2]
        except ValueError:
            pass
    return s.lower()


# ================================================================
# FLEXIBLE COLUMN RESOLUTION — matches a configured column name to
# whatever the live sheet actually calls it: exact match first, then
# case/whitespace-insensitive. Lets the app survive small header
# drift (e.g. "AWB Number" vs "AWB NUMBER") instead of a hard crash.
# ================================================================
def resolve_col(target: str, columns) -> Optional[str]:
    if target in columns:
        return target
    key = str(target).strip().lower()
    for c in columns:
        if str(c).strip().lower() == key:
            return c
    return None


def resolve_and_rename(df: pd.DataFrame, config: dict, keys: list) -> tuple[pd.DataFrame, list]:
    """For each `key` in `keys`, find the real column in `df` that
    matches `config[key]` (tolerating case/whitespace drift) and
    rename it to the exact `config[key]` string. This means every
    other function can safely index with `df[config[key]]` directly
    instead of re-resolving the column name every time it's used —
    which is what caused a live KeyError before: a column found by a
    tolerant lookup was never actually renamed, so later exact-match
    lookups on the same column could still fail.

    Returns (renamed_df, missing_keys) — missing_keys lists any
    config keys that had no match at all in df.columns.
    """
    rename_map = {}
    missing = []
    for key in keys:
        wanted = config[key]
        actual = resolve_col(wanted, df.columns)
        if actual is None:
            missing.append(key)
        elif actual != wanted:
            rename_map[actual] = wanted
    if rename_map:
        df = df.rename(columns=rename_map)
    return df, missing
