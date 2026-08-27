"""
ORDER TRACKING — SEARCH & FILTER TOOL
------------------------------------------------------------
Reads LIVE from a SharePoint/OneDrive Excel file (view-only share
link) — no manual upload needed. Whoever updates that file, this
app reflects it (cached 5 min, or hit "Refresh Now" for instantly).

Also reads a SECOND, linked Excel file — a line-item level order
breakdown (GTIN / Description / Order Amt. Exc. GST etc.), keyed by
a "Document No." column that matches this app's "Order Id". When you
open the "Order Details" popup for an order, its line items from
that second file show up automatically underneath.

SETUP (one-time):
  1. In SharePoint/OneDrive, get a share link for the file with
     "Anyone with the link can view" (or your org's equivalent).
  2. Add it to Streamlit secrets:
       Local:  create .streamlit/secrets.toml with:
                   SHAREPOINT_EXCEL_URL = "https://....?e=xxxx"
                   SHAREPOINT_ORDER_ITEMS_URL = "https://....?e=yyyy"
       Streamlit Cloud: App -> Settings -> Secrets -> paste the same lines.
     SHAREPOINT_ORDER_ITEMS_URL is optional — without it, the app
     works as before, just without the line-item breakdown popup.
  3. requirements.txt needs: streamlit, pandas, openpyxl, requests

Shared fetch/caching/typing/theme helpers live in common.py, next to
Home.py, since Order Tracking / Stock Gap / Cancelled Orders / Fill
Rate all repeat the same SharePoint-download + column-typing logic.
------------------------------------------------------------
"""

import pandas as pd
import streamlit as st

from common import (
    FetchError, build_column_config, fetch_bytes, get_secret, inject_theme,
    read_excel_sheet, resolve_col, type_columns,
)

st.set_page_config(page_title="Order Tracking Search", layout="wide", page_icon="📦")
inject_theme()

SHEET_NAME = None  # None = first sheet. Set an exact tab name if the workbook has multiple tabs.

# ------------------------------------------------------------
# Second (linked) Excel source — line-item level order detail, one
# row per GTIN/Description per order, keyed by "Document No." (which
# matches this app's "Order Id"). Powers the extra breakdown table
# inside the "Order Details" popup.
# ------------------------------------------------------------
ITEMS_SHEET_NAME = None  # None = first sheet in that workbook
ITEMS_DOC_NO_COLUMN = "Document No."  # column in the items file that matches "Order Id"
ITEMS_DISPLAY_COLUMNS = [
    "Order Date", "Document No.", "Customer Name", "GTIN", "Description",
    "Order Qty", "Order Amt. Exc. GST", "Invoice Qty",
    "Invoice Amt. Exc. GST", "Invoice No.",
]

ORDER_ID_COLUMN = "Order Id"

# Columns a free-text search box matches against.
SEARCH_COLUMNS = [
    "Order Id", "AWB NUMBER", "InvoiceNumber", "Customer Name",
    "External Document No.", "SRO Number",
]

# Shown by default in the results table — team can change via the
# "Columns to show" picker.
DEFAULT_VISIBLE_COLUMNS = [
    "Order Id", "Customer Name", "Order Received Date", "Order Qty",
    "Order Value", "AWB NUMBER", "COURIER", "Delivery Status", "Standard TAT",
]

# Dropdown filters shown in the Filters expander.
FILTER_COLUMNS = ["DB Code", "Name", "Final Remarks", "InvoiceNumber", "Category"]


@st.cache_data(ttl=300, show_spinner="Fetching latest data…")
def load_data(url: str) -> pd.DataFrame:
    raw = fetch_bytes(url)
    df = read_excel_sheet(raw, SHEET_NAME)
    return type_columns(df)


@st.cache_data(ttl=300, show_spinner="Fetching order item details…")
def load_items_data(url: str) -> pd.DataFrame:
    raw = fetch_bytes(url)
    df = read_excel_sheet(raw, ITEMS_SHEET_NAME)
    return type_columns(df)


st.title("📦 Order Tracking — Search & Filter")

sp_url = get_secret("SHAREPOINT_EXCEL_URL")
items_url = get_secret("SHAREPOINT_ORDER_ITEMS_URL")  # optional — powers the line-item popup table

if not sp_url:
    st.error(
        "No SharePoint link configured. Add SHAREPOINT_EXCEL_URL to this app's "
        "Secrets (see the setup note at the top of this page's source)."
    )
    st.stop()

top_l, top_r = st.columns([5, 1])
with top_r:
    if st.button("🔄 Refresh Now", use_container_width=True):
        load_data.clear()
        load_items_data.clear()
        st.rerun()

try:
    df = load_data(sp_url)
except FetchError as e:
    st.error(f"Could not load the live file: {e}")
    st.stop()

order_id_col = resolve_col(ORDER_ID_COLUMN, df.columns)
if order_id_col is None:
    st.warning(f"Heads up: no '{ORDER_ID_COLUMN}' column found — the full-detail lookup at the bottom needs it to work.")

st.caption(f"{len(df):,} orders loaded · cached up to 5 min, or hit **Refresh Now** for the latest")
if not items_url:
    st.caption(
        "Tip: add SHAREPOINT_ORDER_ITEMS_URL to Secrets to also show line-item "
        "detail (GTIN, Description, Order/Invoice Amt., etc.) inside Order Details."
    )

# ------------------------------------------------------------
# SEARCH + FILTERS
# ------------------------------------------------------------
search_term = st.text_input(
    "🔍 Search",
    placeholder="Order Id, AWB Number, Invoice Number, Customer Name, External Doc No…",
)

with st.expander("Filters"):
    fcols = st.columns(3)
    active_filters = {}
    filter_cols_present = [c for c in FILTER_COLUMNS if c in df.columns]
    for i, col in enumerate(filter_cols_present):
        with fcols[i % 3]:
            options = sorted(df[col].dropna().astype(str).unique().tolist())
            picked = st.multiselect(col, options, key=f"filter_{col}")
            if picked:
                active_filters[col] = picked

filtered = df.copy()
if search_term:
    present_search_cols = [c for c in SEARCH_COLUMNS if c in filtered.columns]
    mask = pd.Series(False, index=filtered.index)
    for c in present_search_cols:
        mask |= filtered[c].astype(str).str.contains(search_term, case=False, na=False)
    filtered = filtered[mask]

for col, vals in active_filters.items():
    filtered = filtered[filtered[col].astype(str).isin(vals)]

# ------------------------------------------------------------
# COLUMN PICKER — team chooses what shows in the results table.
# Collapsible, same as Filters above, so it can be closed once set.
# ------------------------------------------------------------
all_cols = list(df.columns)
default_cols = [c for c in DEFAULT_VISIBLE_COLUMNS if c in all_cols] or all_cols[:8]
with st.expander("Columns to show"):
    visible_cols = st.multiselect("Columns to show", all_cols, default=default_cols,
                                   key="visible_cols", label_visibility="collapsed")
if not visible_cols:
    visible_cols = default_cols

# ------------------------------------------------------------
# RESULTS TABLE
# ------------------------------------------------------------
st.markdown(f"**{len(filtered):,} of {len(df):,} orders**")

display_df = filtered[visible_cols].copy()
st.dataframe(
    display_df, use_container_width=True, hide_index=True, height=560,
    column_config=build_column_config(visible_cols),
)


# ------------------------------------------------------------
# FULL-DETAIL LOOKUP — every column for one order, in a dialog, plus
# the linked line-item breakdown pulled from the second Excel file
# (matched on Document No. == this Order Id).
# ------------------------------------------------------------
@st.dialog("Order Details", width="large")
def show_full_details(order_id):
    row = df[df[order_id_col].astype(str) == str(order_id)]
    if row.empty:
        st.warning("Order not found.")
        return
    row = row.iloc[0]

    detail_rows = []
    for col in df.columns:
        val = row[col]
        if pd.isna(val) or val == "":
            continue
        if isinstance(val, pd.Timestamp):
            val = val.strftime("%d-%m-%Y")
        # Cast every value to str: this column mixes text, numbers, and
        # dates row-to-row (each row is a different field), which
        # otherwise leaves pandas/Arrow guessing a single dtype for a
        # genuinely mixed-type column. It's already shown as a
        # TextColumn below, so nothing about the display changes.
        detail_rows.append({"Field": col, "Value": str(val)})

    details_df = pd.DataFrame(detail_rows)
    st.dataframe(
        details_df, use_container_width=True, hide_index=True,
        height=min(560, 38 * len(details_df) + 38),
        column_config={
            "Field": st.column_config.TextColumn("Field", width="medium"),
            "Value": st.column_config.TextColumn("Value", width="large"),
        },
    )

    st.markdown("---")
    st.markdown("**Order Line Items**")

    if not items_url:
        st.caption(
            "Add SHAREPOINT_ORDER_ITEMS_URL to Secrets to show line-item "
            "detail (Order Date, GTIN, Description, etc.) here."
        )
        return

    try:
        items_df = load_items_data(items_url)
    except FetchError as e:
        st.error(f"Could not load the line-item file: {e}")
        return

    items_doc_col = resolve_col(ITEMS_DOC_NO_COLUMN, items_df.columns)
    if items_doc_col is None:
        st.warning(
            f"'{ITEMS_DOC_NO_COLUMN}' column not found in the line-item file. "
            "Update ITEMS_DOC_NO_COLUMN at the top of this file to match your sheet."
        )
        return

    item_rows = items_df[items_df[items_doc_col].astype(str) == str(order_id)]
    if item_rows.empty:
        st.caption("No line items found for this Order Id.")
        return

    item_cols_present = [c for c in ITEMS_DISPLAY_COLUMNS if c in item_rows.columns]
    st.caption(f"{len(item_rows):,} line item(s) for Document No. {order_id}")
    st.dataframe(
        item_rows[item_cols_present], use_container_width=True, hide_index=True,
        height=min(420, 76 + 35 * len(item_rows)),
        column_config=build_column_config(item_cols_present),
    )


if order_id_col and order_id_col in filtered.columns and len(filtered):
    st.markdown("---")
    lk1, lk2 = st.columns([3, 1])
    with lk1:
        picked_order = st.selectbox(
            "View full details for Order Id",
            options=sorted(filtered[order_id_col].dropna().astype(str).unique().tolist()),
            index=None,
            placeholder="Select an order…",
            key="order_lookup",
        )
    with lk2:
        st.markdown('<div style="padding-top:28px;"></div>', unsafe_allow_html=True)
        if st.button("View", use_container_width=True) and picked_order:
            show_full_details(picked_order)
