"""
FILL RATE EXPLORER — merged Chain / Product / Shop drill-down
------------------------------------------------------------
Merges Order Tracking + Fill Rate into one cross-drill report:

  1. Pick a primary lens: Chain-wise, Product-wise, or Shop-wise.
  2. Once you drill into one item, toggle between the OTHER two
     lenses as a secondary breakdown of just that item.
  3. Click into the secondary breakdown to open the final result —
     the underlying order-level (or line-item-level, for anything
     touching Product) rows, same shape as Order Tracking's detail
     popup.

Data model
----------
Chain and Shop both come straight from the live dispatch tracker
(SHAREPOINT_EXCEL_URL, monthly tabs — same source Fill Rate reads):
  Chain = CONFIG["name"]      (the "Name" column)
  Shop  = CONFIG["customer"]  ("Customer Name")

Product does NOT exist at that granularity — the tracker is one row
per ORDER, not per SKU. Product-wise numbers instead come from the
line-item file (SHAREPOINT_ORDER_ITEMS_URL, the same file Order
Tracking's popup uses), grouped by GTIN. That file already carries
Customer Name on every row, but not Chain — so build_product_frame()
merges in CONFIG["name"] + Month from the order-level table via
Document No. == Order Id, once, up front. Everything downstream can
then treat "the product frame" as self-contained.

This means Product-wise Fill Rate is computed at the SKU/order-line
level and will not always match the order-level Fill Rate shown on
the other pages exactly — a partially-filled order can be 100% on
one line and 0% on another. That is expected, not a bug.

Setup: this page needs BOTH secrets Order Tracking uses. Without
SHAREPOINT_ORDER_ITEMS_URL, Chain-wise and Shop-wise still work;
Product-wise is disabled with an explanation instead of crashing.
Auto-refresh needs one extra package: add `streamlit-autorefresh`
to requirements.txt (`pip install streamlit-autorefresh`).
------------------------------------------------------------
"""

import numpy as np
import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh

from common import (
    FetchError, fetch_bytes, get_secret, inject_theme, normalize_key,
    read_excel_all_sheets, read_excel_sheet, resolve_and_rename, resolve_col, type_columns,
)

st.set_page_config(page_title="Fill Rate Explorer", page_icon="🔀", layout="wide")

# Same column map Fill Rate uses for the order-level dispatch tracker.
CONFIG = {
    "sharepoint_secret": "SHAREPOINT_EXCEL_URL",
    "customer": "Customer Name", "category": "Category", "channel": "Channel", "zone": "Zone", "name": "Name",
    "order_id": "Order Id", "order_qty": "Order Qty", "order_value": "Order Value",
    "invoice_qty": "Invoice Qty", "invoice_value": "Invoice Value", "invoice_number": "InvoiceNumber",
    "sale_loss": "Sale Loss", "db_code": "DB Code", "category": "Category",
}
MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Line-item file — same one Order Tracking's popup reads.
ITEMS_CONFIG = {
    "sharepoint_secret": "SHAREPOINT_ORDER_ITEMS_URL",
    "doc_no": "Document No.", "order_date": "Order Date", "customer": "Customer Name",
    "gtin": "GTIN", "description": "Description",
    "order_qty": "Order Qty", "order_value": "Order Amt. Exc. GST",
    "invoice_qty": "Invoice Qty", "invoice_value": "Invoice Amt. Exc. GST", "invoice_no": "Invoice No.",
}

LENSES = ["chain", "product", "shop"]
LENS_LABEL = {"chain": "Chain", "product": "Product", "shop": "Shop"}

inject_theme(extra_css="""
.section-title { font-size:21px; font-weight:750; color:#132238; margin:20px 0 8px; }
.dashboard-subtitle { color:#6B7280; font-size:14px; margin-bottom:12px; }
div[role="radiogroup"] label { padding:6px 14px !important; }
""")


# ================================================================
# FORMATTERS — identical to Fill Rate's, kept local so this page
# has no cross-page import dependency.
# ================================================================
def fmt_num(v):
    return "—" if pd.isna(v) else f"{float(v):,.0f}"


def fmt_pct(v):
    return "—" if pd.isna(v) else f"{float(v):.1f}%"


def fmt_currency(v):
    if pd.isna(v):
        return "—"
    v = float(v)
    sign = "-" if v < 0 else ""
    v = abs(v)
    if v >= 10_000_000:
        return f"{sign}₹ {v / 10_000_000:.2f} Cr"
    if v >= 100_000:
        return f"{sign}₹ {v / 100_000:.2f} L"
    return f"{sign}₹ {v:,.2f}"


def clean_numeric(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.replace("₹", "", regex=False).str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan, "-": np.nan}),
        errors="coerce",
    )


def month_key(m):
    try:
        return MONTH_ORDER.index(str(m))
    except ValueError:
        return 999


def status_for(fr):
    if pd.isna(fr):
        return "—", "gray"
    if fr >= 95:
        return "Healthy", "green"
    if fr >= 85:
        return "Watch", "orange"
    return "Critical", "red"


# ================================================================
# LOADERS
# ================================================================
@st.cache_data(ttl=300, show_spinner="Fetching dispatch tracker…")
def load_order_df(url):
    raw = fetch_bytes(url, timeout=60)
    wb = read_excel_all_sheets(raw)
    frames = []
    for sname, sdf in wb.items():
        sname = str(sname).strip()
        if sname not in MONTH_ORDER or sdf is None or sdf.empty:
            continue
        d = sdf.copy()
        d["Month"] = sname
        frames.append(d)
    if not frames:
        raise ValueError("No monthly sheets found in the dispatch tracker.")
    df = pd.concat(frames, ignore_index=True, sort=False)
    df, _ = resolve_and_rename(df, CONFIG, [k for k in CONFIG if k != "sharepoint_secret"])
    for key in ["order_qty", "order_value", "invoice_qty", "invoice_value", "sale_loss"]:
        col = CONFIG[key]
        if col in df.columns:
            df[col] = clean_numeric(df[col])
    return df


@st.cache_data(ttl=300, show_spinner="Fetching line items…")
def load_items_df(url):
    raw = fetch_bytes(url, timeout=60)
    df = read_excel_sheet(raw, None)
    df = type_columns(df)
    df, _ = resolve_and_rename(df, ITEMS_CONFIG, [k for k in ITEMS_CONFIG if k != "sharepoint_secret"])
    for key in ["order_qty", "order_value", "invoice_qty", "invoice_value"]:
        col = ITEMS_CONFIG[key]
        if col in df.columns:
            df[col] = clean_numeric(df[col])
    return df


@st.cache_data(ttl=300, show_spinner="Linking line items to chains…")
def build_product_frame(order_df, items_df):
    """Line items + Chain/Month attached from the order-level table,
    joined on Document No. == Order Id. Customer Name is already on
    the items file directly, so Shop-wise off this frame needs no
    join at all — only Chain does."""
    oid = CONFIG["order_id"]
    doc = ITEMS_CONFIG["doc_no"]
    if oid not in order_df.columns or doc not in items_df.columns:
        return items_df.assign(**{CONFIG["name"]: np.nan, "Month": np.nan})
    lookup = order_df[[oid, CONFIG["name"], "Month"]].drop_duplicates(subset=[oid])
    lookup["__key"] = lookup[oid].map(normalize_key)
    items = items_df.copy()
    items["__key"] = items[doc].map(normalize_key)
    merged = items.merge(lookup[["__key", CONFIG["name"], "Month"]], on="__key", how="left")
    merged["__gtin_key"] = merged[ITEMS_CONFIG["gtin"]].map(normalize_key)
    return merged.drop(columns="__key")


# ================================================================
# AGGREGATION — one function serves Chain, Shop AND Product, since
# all three are just "group these rows by this column and sum these
# qty/value columns" once you hand it the right column names.
# ================================================================
def summarize_by(df, group_col, id_col, qty_o, qty_i, val_o, val_i, label_col=None, sale_loss_col=None):
    empty_cols = ["__key", "__label", "orders", "order_qty", "invoice_qty", "fr", "order_value", "invoice_value", "sale_loss", "pending_qty"]
    if df.empty or group_col not in df.columns:
        return pd.DataFrame(columns=empty_cols)
    agg = {"orders": (id_col, "nunique"), "order_qty": (qty_o, "sum"), "invoice_qty": (qty_i, "sum")}
    if val_o in df.columns and val_i in df.columns:
        agg["order_value"] = (val_o, "sum")
        agg["invoice_value"] = (val_i, "sum")
    if sale_loss_col and sale_loss_col in df.columns:
        agg["sale_loss"] = (sale_loss_col, "sum")
    g = df.groupby(group_col, dropna=False).agg(**agg).reset_index()
    g["fr"] = np.where(g.order_qty != 0, g.invoice_qty / g.order_qty * 100, np.nan)
    if "order_value" in g.columns and "sale_loss" not in g.columns:
        g["sale_loss"] = g["order_value"] - g["invoice_value"]  # computed proxy, flagged in the UI
    # __label is a display name derived from the group column (or a
    # separate label_col, e.g. Description for products) — kept under
    # its own name so it can never collide with a real column even
    # when, as with Chain, CONFIG["name"] literally equals "Name".
    if label_col:
        labels = df.groupby(group_col)[label_col].first()
        g["__label"] = g[group_col].map(labels)
    else:
        g["__label"] = g[group_col].astype(str)
    g["__label"] = g["__label"].fillna("Blank / Not Available").astype(str)
    g["pending_qty"] = g.order_qty - g.invoice_qty
    return g.rename(columns={group_col: "__key"})


def lens_source(lens, order_df, product_df):
    """(dataframe, group_col, id_col, qty_o, qty_i, val_o, val_i, label_col, sale_loss_col, granularity)"""
    if lens == "chain":
        return (order_df, CONFIG["name"], CONFIG["order_id"], CONFIG["order_qty"], CONFIG["invoice_qty"],
                CONFIG["order_value"], CONFIG["invoice_value"], None, CONFIG["sale_loss"], "order")
    if lens == "shop":
        return (order_df, CONFIG["customer"], CONFIG["order_id"], CONFIG["order_qty"], CONFIG["invoice_qty"],
                CONFIG["order_value"], CONFIG["invoice_value"], None, CONFIG["sale_loss"], "order")
    return (product_df, "__gtin_key", ITEMS_CONFIG["doc_no"], ITEMS_CONFIG["order_qty"], ITEMS_CONFIG["invoice_qty"],
            ITEMS_CONFIG["order_value"], ITEMS_CONFIG["invoice_value"], ITEMS_CONFIG["description"], None, "item")


def build_summary(lens, order_df, product_df, filter_lens=None, filter_key=None):
    """Top-level summary for `lens`, optionally restricted to rows
    where `filter_lens`'s group column equals `filter_key` (used
    when showing a secondary breakdown inside a drilled-in item)."""
    df, group_col, id_col, qty_o, qty_i, val_o, val_i, label_col, sale_loss_col, _ = lens_source(lens, order_df, product_df)
    if filter_lens and filter_key is not None:
        f_df, f_group_col, *_ = lens_source(filter_lens, order_df, product_df)
        if f_df is df:  # same source table, filter directly
            df = df[df[f_group_col].astype(str) == str(filter_key)]
        else:
            # cross-source filter: need the __gtin_key/Chain/Customer bridge on product_df
            df = df[df[f_group_col].astype(str) == str(filter_key)] if f_group_col in df.columns else df
    return summarize_by(df, group_col, id_col, qty_o, qty_i, val_o, val_i, label_col, sale_loss_col)


def drilldown_rows(order_df, product_df, path):
    """Given the active drill path (list of (lens, key) tuples),
    return the raw underlying rows — order-level unless Product is
    anywhere in the path, in which case item-level."""
    uses_product = any(lens == "product" for lens, _ in path)
    df = product_df if uses_product else order_df
    for lens, key in path:
        _, group_col, *_ = lens_source(lens, order_df, product_df)
        if group_col in df.columns:
            df = df[df[group_col].astype(str) == str(key)]
    return df, uses_product


# ================================================================
# LOAD DATA
# ================================================================
order_url = get_secret(CONFIG["sharepoint_secret"])
items_url = get_secret(ITEMS_CONFIG["sharepoint_secret"])

if not order_url:
    st.error("SHAREPOINT_EXCEL_URL is not configured in Streamlit secrets.")
    st.stop()

top_l, top_r = st.columns([5, 1])
with top_r:
    if st.button("🔄 Refresh Now", width="stretch"):
        load_order_df.clear(); load_items_df.clear(); build_product_frame.clear()
        st.rerun()

try:
    order_df = load_order_df(order_url)
except (FetchError, ValueError) as e:
    st.error(f"Could not load the dispatch tracker: {e}")
    st.stop()

product_df = pd.DataFrame()
product_enabled = bool(items_url)
if product_enabled:
    try:
        items_df = load_items_df(items_url)
        product_df = build_product_frame(order_df, items_df)
    except FetchError as e:
        st.warning(f"Product-wise is unavailable — could not load the line-item file: {e}")
        product_enabled = False

# Silent 5-min auto-refresh: reruns the script (cache TTL also 300s,
# so this is what actually pulls fresh numbers), but every control
# below is session-state-backed, so drill path / lens / columns
# survive the rerun untouched.
st_autorefresh(interval=300_000, key="explorer_autorefresh")

st.title("🔀 Fill Rate Explorer")
st.caption(f"{len(order_df):,} orders loaded · auto-refreshes every 5 min · cached up to 5 min, or hit **Refresh Now**")
if not product_enabled:
    st.caption("Tip: add SHAREPOINT_ORDER_ITEMS_URL to Secrets to enable Product-wise breakdowns.")


# ================================================================
# STATE
# ================================================================
def reset_drill():
    st.session_state["explorer_path"] = []
    st.session_state["explorer_secondary"] = None


if "explorer_primary" not in st.session_state:
    st.session_state["explorer_primary"] = "chain"
if "explorer_path" not in st.session_state:
    st.session_state["explorer_path"] = []  # list of (lens, key, display_name)
if "explorer_secondary" not in st.session_state:
    st.session_state["explorer_secondary"] = None
if "explorer_cols" not in st.session_state:
    st.session_state["explorer_cols"] = {"order_value": True, "invoice_value": False, "pending_qty": True, "sale_loss": False}

lens_options = [l for l in LENSES if l != "product" or product_enabled]


def on_lens_change():
    reset_drill()


st.radio(
    "View by", lens_options, format_func=lambda l: LENS_LABEL[l] + "-wise",
    horizontal=True, key="explorer_primary", on_change=on_lens_change,
)
primary = st.session_state["explorer_primary"]
path = st.session_state["explorer_path"]

# ------------------------------------------------------------
# Breadcrumb
# ------------------------------------------------------------
crumbs = st.columns([1] * (len(path) + 2) + [6])
with crumbs[0]:
    if st.button(f"{LENS_LABEL[primary]}-wise", key="crumb_root", type="primary" if not path else "secondary"):
        reset_drill(); st.rerun()
for i, (lens, key, name) in enumerate(path):
    with crumbs[i + 1]:
        is_last = i == len(path) - 1
        if st.button(name, key=f"crumb_{i}", type="primary" if is_last else "secondary"):
            st.session_state["explorer_path"] = path[: i + 1]
            st.session_state["explorer_secondary"] = None
            st.rerun()

st.markdown("---")

# ================================================================
# TABLE — top-level lens summary, or a secondary breakdown once
# drilled into a specific item.
# ================================================================
OPTIONAL_COLS = {"order_value": "Order Value", "invoice_value": "Invoice Value", "pending_qty": "Pending Qty", "sale_loss": "Sale Loss"}


def render_table(rows_df, key_suffix, is_final_level):
    with st.expander("Columns to show"):
        for k, label in OPTIONAL_COLS.items():
            st.session_state["explorer_cols"][k] = st.checkbox(label, value=st.session_state["explorer_cols"][k], key=f"col_{k}_{key_suffix}")
    cols = st.session_state["explorer_cols"]

    view = rows_df.copy()
    view["Name"] = view["__label"]
    view["Fill Rate"] = view["fr"].map(fmt_pct)
    view["Status"] = view["fr"].map(lambda v: status_for(v)[0])
    view["Order Qty"] = view["order_qty"].map(fmt_num)
    view["Invoice Qty"] = view["invoice_qty"].map(fmt_num)
    view["Orders"] = view["orders"].map(fmt_num)
    show_cols = ["Name", "Orders", "Order Qty", "Invoice Qty"]
    if cols["order_value"] and "order_value" in view.columns:
        view["Order Value"] = view["order_value"].map(fmt_currency); show_cols.append("Order Value")
    if cols["invoice_value"] and "invoice_value" in view.columns:
        view["Invoice Value"] = view["invoice_value"].map(fmt_currency); show_cols.append("Invoice Value")
    if cols["pending_qty"]:
        view["Pending Qty"] = view["pending_qty"].map(fmt_num); show_cols.append("Pending Qty")
    if cols["sale_loss"] and "sale_loss" in view.columns:
        view["Sale Loss"] = view["sale_loss"].map(fmt_currency); show_cols.append("Sale Loss")
    show_cols += ["Fill Rate", "Status"]

    display = view.sort_values("order_qty", ascending=False)[show_cols].reset_index(drop=True)
    keys_in_order = view.sort_values("order_qty", ascending=False)["__key"].reset_index(drop=True)
    names_in_order = view.sort_values("order_qty", ascending=False)["Name"].reset_index(drop=True)

    st.dataframe(display, width="stretch", hide_index=True, height=min(520, 60 + 36 * len(display)))

    verb = "Open order detail for" if is_final_level else "Drill into"
    pick_col, btn_col = st.columns([5, 1])
    with pick_col:
        choice = st.selectbox(
            verb, ["Select…"] + names_in_order.tolist(),
            key=f"pick_{key_suffix}", label_visibility="collapsed",
        )
    with btn_col:
        go = st.button("Go →", key=f"go_{key_suffix}", width="stretch")
    if go and choice != "Select…":
        idx = names_in_order[names_in_order == choice].index[0]
        return keys_in_order.iloc[idx], choice
    return None, None


if not path:
    # Top level — no filter, just the chosen primary lens.
    summary = build_summary(primary, order_df, product_df)
    st.markdown(f'<div class="section-title">{LENS_LABEL[primary]}-wise fill rate</div>', unsafe_allow_html=True)
    st.markdown('<div class="dashboard-subtitle">Click a row to see its breakdown by the other two dimensions.</div>', unsafe_allow_html=True)
    picked_key, picked_name = render_table(summary, "top", is_final_level=False)
    if picked_key is not None:
        st.session_state["explorer_path"] = [(primary, picked_key, picked_name)]
        others = [l for l in lens_options if l != primary]
        st.session_state["explorer_secondary"] = others[0]
        st.rerun()

else:
    drilled_lens, drilled_key, drilled_name = path[-1]
    others = [l for l in lens_options if l != drilled_lens and l not in [p[0] for p in path[:-1]]]
    if not others:
        others = [l for l in lens_options if l != drilled_lens]
    secondary = st.session_state["explorer_secondary"] or others[0]

    st.markdown(f'<div class="section-title">{drilled_name} — by {LENS_LABEL[secondary]}</div>', unsafe_allow_html=True)
    seg_cols = st.columns([1] * len(others) + [6])
    for i, o in enumerate(others):
        with seg_cols[i]:
            if st.button(f"By {LENS_LABEL[o]}", key=f"sec_{o}", type="primary" if o == secondary else "secondary"):
                st.session_state["explorer_secondary"] = o
                st.rerun()

    df_source, group_col, *_ = lens_source(drilled_lens, order_df, product_df)
    breakdown = build_summary(secondary, order_df, product_df, filter_lens=drilled_lens, filter_key=drilled_key)
    if breakdown.empty:
        st.info("No rows for this combination — the two dimensions may not share underlying data (e.g. Product needs the line-item file joined for Chain context).")
    else:
        st.markdown('<div class="dashboard-subtitle">Click a row to open the order-level detail behind it — the final result.</div>', unsafe_allow_html=True)
        picked_key, picked_name = render_table(breakdown, f"sec_{secondary}", is_final_level=True)
        if picked_key is not None:
            full_path = path + [(secondary, picked_key, picked_name)]
            rows, is_item_level = drilldown_rows(order_df, product_df, full_path)

            @st.dialog(f"Order detail — {picked_name}", width="large")
            def show_final():
                st.caption(f"{len(rows):,} {'line item' if is_item_level else 'order'} row(s) for **{' → '.join(n for _, _, n in full_path)}**")
                if is_item_level:
                    wanted = [ITEMS_CONFIG[k] for k in ["order_date", "doc_no", "customer", "gtin", "description", "order_qty", "order_value", "invoice_qty", "invoice_value", "invoice_no"]]
                else:
                    wanted = [CONFIG["order_id"], CONFIG["customer"], CONFIG["name"], CONFIG["category"],
                              CONFIG["order_qty"], CONFIG["order_value"], CONFIG["invoice_qty"], CONFIG["invoice_value"], CONFIG["sale_loss"], "Month"]
                present = [resolve_col(c, rows.columns) or c for c in wanted]
                present = [c for c in present if c in rows.columns]
                display_rows = rows[present].copy()
                for c in display_rows.columns:
                    if display_rows[c].dtype == "object":
                        # Mixed int/text within one column (e.g. an AWB
                        # number stored as text in some rows, a number in
                        # others) breaks Arrow serialization. Stringify
                        # defensively rather than rely on Streamlit's
                        # automatic-fix fallback.
                        display_rows[c] = display_rows[c].apply(lambda v: "" if pd.isna(v) else str(v))
                st.dataframe(display_rows, width="stretch", hide_index=True, height=min(560, 60 + 36 * len(rows)))

            show_final()

st.markdown("---")
st.caption("Prototype note: Product-wise Fill Rate is computed at line-item level and may not exactly match order-level Fill Rate on other pages — this is expected, not a bug.")
