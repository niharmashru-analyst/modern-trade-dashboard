"""
STOCK GAP DASHBOARD
------------------------------------------------------------
Matches an Order sheet against the live Stock workbook and shows
exactly where orders exceed available stock.

The order sheet is upload-only (.xlsx/.xls/.csv). No live order-sheet
link is used on this page.

SETUP: add one link to this app's Secrets —
    STOCK_EXCEL_URL -> the stock workbook (multi-sheet, required)

CONFIG below controls sheet names / column mapping for the STOCK
workbook — edit this block directly if sheet names or columns ever
change; nothing else in the file needs to change. The ORDER sheet's
columns are picked from dropdowns on the page itself (its layout
varies more than the stock workbook does), with a best-guess default.
------------------------------------------------------------
"""

import io

import pandas as pd
import streamlit as st

from common import FetchError, fetch_bytes, get_secret, inject_theme, normalize_key

st.set_page_config(page_title="Stock Gap Dashboard", layout="wide", page_icon="📉")
inject_theme()

# ================================================================
# CONFIG — edit any time the STOCK workbook's sheet names or
# columns change. Nothing else in the script needs to change.
# ================================================================
CONFIG = {
    # Row number (1-indexed) where the real column headers sit in
    # every stock sheet.
    "stock_header_row": 1,

    # Which sheets to read from the stock workbook, and which column
    # LETTER holds EAN / Product Name in each one.
    "stock_sheets": [
        {"name": "Sheet1", "ean_col": "A", "name_col": "B"},
    ],

    # Header TEXT (not column letter) for each stock quantity type —
    # matched by searching each sheet's header row, so it still works
    # even if the column position shifts between sheets.
    "qty_headers": {
        "mwh": "MWH Stock",         # used for Ahmedabad comparison
        "blr": "Direct Shelf BLR",  # used for Bangalore comparison
        "uc": "UC INVENTORY",       # shown alongside for Ahmedabad only
    },

    # When the same EAN appears in more than one sheet above, its
    # quantities are summed together across all of them.
    "order_header_row": 1,
    "order_ean_col_guess": ["ean", "barcode", "upc"],
    "order_name_col_guess": ["product", "name", "description", "item"],
    "order_qty_col_guess": ["order", "qty", "quantity", "demand"],
}

STATUS_LABELS = {
    "stockout": "🔴 Stockout",
    "low": "🟡 Low",
    "ok": "🟢 OK",
    "missing": "❔ Not Found",
}


def _col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _find_header_col(columns, header_text):
    target = header_text.strip().lower()
    for c in columns:
        if str(c).strip().lower() == target:
            return c
    return None


def _guess_col(columns, hints):
    cols_lower = {str(c).lower(): c for c in columns}
    for hint in hints:
        for cl, orig in cols_lower.items():
            if hint in cl:
                return orig
    return columns[0] if len(columns) else None


@st.cache_data(ttl=300, show_spinner="Fetching latest stock…")
def load_stock(url: str) -> pd.DataFrame:
    raw = fetch_bytes(url)
    hdr_idx = CONFIG["stock_header_row"] - 1
    combined = {}
    skipped_sheets = []

    for sheet_cfg in CONFIG["stock_sheets"]:
        try:
            sdf = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_cfg["name"], header=hdr_idx)
        except Exception:
            skipped_sheets.append(sheet_cfg["name"])
            continue
        sdf.columns = [str(c).strip() for c in sdf.columns]

        ean_idx = _col_letter_to_index(sheet_cfg["ean_col"])
        name_idx = _col_letter_to_index(sheet_cfg["name_col"])
        if ean_idx >= len(sdf.columns) or name_idx >= len(sdf.columns):
            skipped_sheets.append(sheet_cfg["name"])
            continue

        mwh_col = _find_header_col(sdf.columns, CONFIG["qty_headers"]["mwh"])
        blr_col = _find_header_col(sdf.columns, CONFIG["qty_headers"]["blr"])
        uc_col = _find_header_col(sdf.columns, CONFIG["qty_headers"]["uc"])

        ean_series = sdf.iloc[:, ean_idx]
        name_series = sdf.iloc[:, name_idx]
        mwh_series = pd.to_numeric(sdf[mwh_col], errors="coerce") if mwh_col else None
        blr_series = pd.to_numeric(sdf[blr_col], errors="coerce") if blr_col else None
        uc_series = pd.to_numeric(sdf[uc_col], errors="coerce") if uc_col else None

        for i in range(len(sdf)):
            if pd.isna(ean_series.iloc[i]):
                continue
            ean_raw = ean_series.iloc[i]
            key = normalize_key(ean_raw)
            if not key:
                continue
            ean = str(ean_raw).strip()
            if ean.endswith(".0") and key == ean[:-2].lower():
                ean = ean[:-2]  # display the clean code too, not "...924.0"
            name = str(name_series.iloc[i]).strip() if pd.notna(name_series.iloc[i]) else ""
            mwh = float(mwh_series.iloc[i]) if mwh_series is not None and pd.notna(mwh_series.iloc[i]) else 0.0
            blr = float(blr_series.iloc[i]) if blr_series is not None and pd.notna(blr_series.iloc[i]) else 0.0
            uc = float(uc_series.iloc[i]) if uc_series is not None and pd.notna(uc_series.iloc[i]) else 0.0

            if key not in combined:
                combined[key] = {"ean": ean, "product": name, "mwh": 0.0, "blr": 0.0, "uc": 0.0}
            row = combined[key]
            if not row["product"] and name:
                row["product"] = name
            row["mwh"] += mwh
            row["blr"] += blr
            row["uc"] += uc

    result = pd.DataFrame(list(combined.values()))
    result.attrs["skipped_sheets"] = skipped_sheets
    return result


def parse_uploaded_orders(file_bytes: bytes, filename: str, header_row: int) -> pd.DataFrame:
    """Parse an uploaded order file (.xlsx/.xls/.csv) using a user-chosen header row."""
    hdr_idx = header_row - 1
    if filename.lower().endswith(".csv"):
        odf = pd.read_csv(io.BytesIO(file_bytes), header=hdr_idx)
    else:
        odf = pd.read_excel(io.BytesIO(file_bytes), header=hdr_idx)
    odf.columns = [str(c).strip() for c in odf.columns]
    return odf


st.title("📉 Stock Gap Dashboard")
st.caption("Matches an uploaded order sheet against the live stock workbook.")

stock_url = get_secret("STOCK_EXCEL_URL")
if not stock_url:
    st.error("Missing link. Add **STOCK_EXCEL_URL** to this app's Secrets.")
    st.stop()

top_l, top_r = st.columns([5, 1])
with top_r:
    if st.button("🔄 Refresh Now", use_container_width=True):
        load_stock.clear()
        st.rerun()

try:
    stock_df = load_stock(stock_url)
except FetchError as e:
    st.error(f"Could not load the live stock file: {e}")
    st.stop()

# ------------------------------------------------------------
# ORDER SHEET SOURCE — upload only
# ------------------------------------------------------------
st.markdown("#### Order Sheet")
st.caption("Upload the order file to compare against the current stock workbook. Live order links are disabled.")

up_col1, up_col2 = st.columns([3, 1])
with up_col1:
    uploaded_order_file = st.file_uploader(
        "Choose order file (.xlsx / .xls / .csv)",
        type=["xlsx", "xls", "csv"],
        key="order_upload",
    )
with up_col2:
    header_row = st.number_input(
        "Header row #",
        min_value=1,
        value=CONFIG["order_header_row"],
        step=1,
        key="order_header_row_upload",
    )

if uploaded_order_file is None:
    st.info("Upload an order sheet to continue.")
    st.stop()

try:
    order_df = parse_uploaded_orders(
        uploaded_order_file.getvalue(),
        uploaded_order_file.name,
        int(header_row),
    )
except Exception as e:
    st.error(f"Could not parse the uploaded order file: {e}")
    st.stop()

order_source_label = f"Uploaded file → `{uploaded_order_file.name}`"

if stock_df.empty:
    st.warning("No stock rows loaded — check CONFIG sheet names/columns still match the live workbook.")
    st.stop()

skipped = stock_df.attrs.get("skipped_sheets", [])
if skipped:
    st.caption(f"⚠️ Skipped sheet(s) not found in the workbook: {', '.join(skipped)} — check CONFIG.")

# ------------------------------------------------------------
# LOADED-DATA PREVIEW — shows exactly what was fetched, so a
# wrong file/header row is obvious immediately.
# ------------------------------------------------------------
with st.expander(f"🔍 Loaded order data preview — {len(order_df):,} rows, {len(order_df.columns)} columns", expanded=False):
    st.caption(f"Source: {order_source_label}")
    st.write("Columns:", list(order_df.columns))
    st.dataframe(order_df.head(5), use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# SETUP CONTROLS
# ------------------------------------------------------------
c1, c2, c3 = st.columns(3)
with c1:
    match_mode = st.radio("Match by", ["EAN", "Product Name"], horizontal=True)
with c2:
    order_cols = list(order_df.columns)
    default_key_col = (
        _guess_col(order_cols, CONFIG["order_ean_col_guess"]) if match_mode == "EAN"
        else _guess_col(order_cols, CONFIG["order_name_col_guess"])
    )
    key_label = "Order sheet's EAN column" if match_mode == "EAN" else "Order sheet's Product Name column"
    key_col = st.selectbox(
        key_label, order_cols,
        index=order_cols.index(default_key_col) if default_key_col in order_cols else 0,
    )
with c3:
    default_qty_col = _guess_col(order_cols, CONFIG["order_qty_col_guess"])
    qty_col = st.selectbox(
        "Order Qty column", order_cols,
        index=order_cols.index(default_qty_col) if default_qty_col in order_cols else 0,
    )

location = st.radio("Warehouse Location", ["Ahmedabad", "Bangalore"], horizontal=True)
st.caption(
    "Ahmedabad compares against **MWH Stock** (UC Inventory shown alongside). "
    "Bangalore compares against **Direct Shelf BLR** only."
)

# ------------------------------------------------------------
# MATCH
# ------------------------------------------------------------
stock_map = {}
for _, r in stock_df.iterrows():
    key = normalize_key(r["ean"]) if match_mode == "EAN" else str(r["product"]).strip().lower()
    stock_map[key] = r


def _compute_row(order_row):
    raw_key = order_row[key_col]
    key = normalize_key(raw_key) if match_mode == "EAN" else (str(raw_key).strip().lower() if pd.notna(raw_key) else "")
    order_qty = pd.to_numeric(order_row[qty_col], errors="coerce")
    order_qty = 0.0 if pd.isna(order_qty) else float(order_qty)

    stock_row = stock_map.get(key)
    if stock_row is None:
        # Cast raw_key to str here (not just left as whatever type the
        # order sheet's cell was) — otherwise unmatched rows leave
        # numeric EANs (e.g. numpy.int64) sitting in the same column as
        # the strings matched rows get from the stock sheet, and pandas/
        # Arrow can't serialize the resulting mixed-type column cleanly.
        return pd.Series({
            "product": str(raw_key) if match_mode != "EAN" else "(unknown)",
            "ean": str(raw_key) if match_mode == "EAN" and pd.notna(raw_key) else "",
            "stock": None, "uc": None, "order": order_qty, "short": None, "status": "missing",
        })

    available = stock_row["mwh"] if location == "Ahmedabad" else stock_row["blr"]
    uc_val = stock_row["uc"] if location == "Ahmedabad" else None
    short = max(order_qty - available, 0)
    if short > 0:
        status = "stockout"
    elif available > 0 and order_qty > 0 and (available - order_qty) <= available * 0.15:
        status = "low"
    else:
        status = "ok"
    return pd.Series({
        "product": stock_row["product"] or raw_key, "ean": stock_row["ean"],
        "stock": available, "uc": uc_val, "order": order_qty, "short": short, "status": status,
    })


if st.button("Match & Build Dashboard", type="primary"):
    st.session_state["gap_results"] = order_df.apply(_compute_row, axis=1)
    st.session_state["gap_location"] = location

results = st.session_state.get("gap_results")

if results is not None and len(results):
    total_skus = len(results)
    short_skus = int((results["status"] == "stockout").sum())
    units_short = float(results["short"].fillna(0).sum())
    missing_skus = int((results["status"] == "missing").sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("SKUs Ordered", f"{total_skus:,}")
    m2.metric("SKUs Short / Out", f"{short_skus:,}")
    m3.metric("Units Short", f"{units_short:,.0f}")

    if missing_skus and missing_skus / total_skus > 0.3:
        st.warning(
            f"⚠️ {missing_skus:,} of {total_skus:,} order rows ({missing_skus/total_skus:.0%}) didn't "
            f"match any stock row — double-check the **{match_mode}** column you picked on the order "
            "sheet actually lines up with the stock file's EAN/product values. Use the **Not Found** "
            "filter below to see them."
        )

    fcol, scol = st.columns([2, 1])
    with fcol:
        status_filter = st.radio("Filter", ["All", "Stockout", "Low", "Not Found"], horizontal=True, key="gap_status_filter")
    with scol:
        gap_search = st.text_input("Search product / EAN", key="gap_search")

    view = results.copy()
    status_map = {"Stockout": "stockout", "Low": "low", "Not Found": "missing"}
    if status_filter != "All":
        view = view[view["status"] == status_map[status_filter]]
    if gap_search:
        s = gap_search.lower()
        view = view[
            view["product"].astype(str).str.lower().str.contains(s, na=False)
            | view["ean"].astype(str).str.lower().str.contains(s, na=False)
        ]

    view = view.sort_values("short", ascending=False, na_position="last")
    view["status"] = view["status"].map(STATUS_LABELS).fillna(view["status"])

    show_cols = ["product", "ean", "stock", "order", "short", "status"]
    if st.session_state.get("gap_location") == "Ahmedabad":
        show_cols.insert(3, "uc")

    label_map = {
        "product": "Product", "ean": "EAN", "stock": "Available", "uc": "UC Inventory",
        "order": "Ordered", "short": "Short", "status": "Status",
    }
    display = view[show_cols].rename(columns=label_map)

    column_config = {
        "Available": st.column_config.NumberColumn(format="localized"),
        "UC Inventory": st.column_config.NumberColumn(format="localized"),
        "Ordered": st.column_config.NumberColumn(format="localized"),
        "Short": st.column_config.NumberColumn(format="localized"),
    }
    st.dataframe(display, use_container_width=True, hide_index=True, height=560, column_config=column_config)

    csv_cols = list(label_map.values())
    csv = view.rename(columns=label_map)[csv_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Export CSV", csv,
        file_name=f"stock_gap_results_{(st.session_state.get('gap_location') or '').lower()}.csv",
        mime="text/csv",
    )
else:
    st.info("Set the columns above and click **Match & Build Dashboard** to see the gap analysis.")
