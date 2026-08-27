"""
FILL RATE — month-on-month dashboard
------------------------------------------------------------
Reads the same live dispatch tracker workbook as the other pages
(SHAREPOINT_EXCEL_URL secret), where each month lives on its own
sheet tab (Jan, Feb, Mar, …), and combines them into one table to
compare Fill Rate (Invoice Qty / Order Qty) month-over-month, by
customer, and by category.

CONFIG below maps this page's internal names ("order_qty", "customer",
etc.) to the actual column headers in the workbook — edit it if a
header ever changes. Column lookups tolerate small case/whitespace
differences (e.g. "order qty" vs "Order Qty"): on load, every column
found this way is renamed to its exact CONFIG spelling, so the rest
of the page can index with CONFIG values directly.
------------------------------------------------------------
"""

from datetime import datetime

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

from common import FetchError, fetch_bytes, get_secret, inject_theme, read_excel_all_sheets, resolve_and_rename, resolve_col

st.set_page_config(page_title="Fill Rate MOM Dashboard", page_icon="📊", layout="wide")

CONFIG = {
    "sharepoint_secret": "SHAREPOINT_EXCEL_URL",
    "customer": "Customer Name", "category": "Category", "channel": "Channel", "zone": "Zone", "name": "Name",
    "order_id": "Order Id", "order_qty": "Order Qty", "order_value": "Order Value",
    "invoice_qty": "Invoice Qty", "invoice_value": "Invoice Value", "invoice_number": "InvoiceNumber",
    "fr_qty": "Over all FR % (Qty)", "fr_value": "Over all FR % (Value)", "sale_loss": "Sale Loss",
    "order_received_date": "Order Received Date", "order_upload_date": "Order Upload date", "wh_receiving_date": "Wh Receiving Date",
    "db_code": "DB Code", "external_document": "External Document No.", "invoice_date": "Invoice Date",
    "dispatch_date": "Dispatch Date", "awb": "AWB NUMBER", "courier": "COURIER", "mode": "Mode",
    "delivery_status": "Delivery Status", "delivery_date": "Delivery Date", "standard_tat": "Standard TAT",
    "variance": "Vairance", "order_to_wh": "Order to wh", "oti": "OTI", "otd": "OTD", "otde": "OTDE",
    "dispatch_to_delivery": "Dispatch to Deli TAT", "otd_bucket": "OTD Bucket", "wh_remarks": "Wh. Remarks",
    "wh_remark": "Wh Remark", "logistics_remarks": "Logistics Remarks", "ho_remarks": "HO Remarks", "final_remarks": "Final Remarks",
    "actual_delivery_days": "Actual Deli. Days", "box": "Box", "weight": "Weight", "pin_code": "Pin Code",
}
MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

inject_theme(extra_css="""
.section-title { font-size:21px; font-weight:750; color:#132238; margin:24px 0 8px; }
.dashboard-subtitle { color:#6B7280; font-size:14px; margin-bottom:15px; }
div[data-testid="stMetric"] { min-height:118px; }
div[data-testid="stMetricValue"] { font-size:40px !important; }
""")


def clean_numeric(s):
    return pd.to_numeric(
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "NaN": np.nan, "None": np.nan, "-": np.nan}),
        errors="coerce",
    )


def month_key(m):
    try:
        return MONTH_ORDER.index(str(m))
    except ValueError:
        return 999


def fmt_num(v):
    return "—" if pd.isna(v) else f"{float(v):,.0f}"


def fmt_pct(v):
    return "—" if pd.isna(v) else f"{float(v):.1f}%"


def fmt_sale(v):
    return "—" if pd.isna(v) else f"₹ {float(v):,.2f}"


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


@st.cache_data(ttl=300, show_spinner="Fetching latest Excel file...")
def get_workbook(url):
    raw = fetch_bytes(url, timeout=60)
    return read_excel_all_sheets(raw)


@st.cache_data(ttl=300, show_spinner="Combining monthly sheets...")
def load_data(url):
    wb = get_workbook(url)
    frames, sheets = [], []
    for sname, raw in wb.items():
        sname = str(sname).strip()
        if sname not in MONTH_ORDER or raw is None or raw.empty:
            continue
        d = raw.copy()
        d["Month"] = sname
        frames.append(d)
        sheets.append(sname)
    if not frames:
        raise ValueError("No monthly sheets found. Expected sheets like Apr, May, Jun, Jul, Aug.")

    df = pd.concat(frames, ignore_index=True, sort=False)

    # Tolerate small header drift (case/whitespace) once, up front, by
    # renaming every column CONFIG can find to its exact CONFIG spelling.
    # Every later `df[CONFIG["..."]]` lookup in this file then works
    # directly — without this, a lookup could still KeyError even
    # after the "required columns" check passed, because that check
    # only *confirmed a match exists*, it never renamed the column.
    config_keys = [k for k in CONFIG if k != "sharepoint_secret"]
    df, _ = resolve_and_rename(df, CONFIG, config_keys)

    numeric_keys = [
        "order_qty", "order_value", "invoice_qty", "invoice_value", "fr_qty", "fr_value",
        "sale_loss", "standard_tat", "variance", "order_to_wh", "oti", "otd", "otde",
        "dispatch_to_delivery", "actual_delivery_days", "box", "weight",
    ]
    for key in numeric_keys:
        col = CONFIG.get(key)
        if col in df.columns:
            df[col] = clean_numeric(df[col])

    for key in ["order_received_date", "order_upload_date", "wh_receiving_date", "invoice_date", "dispatch_date", "delivery_date"]:
        col = CONFIG[key]
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    df["__month_sort"] = df["Month"].map(month_key)
    return df.sort_values("__month_sort").reset_index(drop=True), sorted(sheets, key=month_key)


def add_derived(df):
    df = df.copy()
    oq, iq = CONFIG["order_qty"], CONFIG["invoice_qty"]
    ov, iv = CONFIG["order_value"], CONFIG["invoice_value"]
    has_qty = oq in df.columns and iq in df.columns
    has_val = ov in df.columns and iv in df.columns

    df["__FR_Qty"] = np.where(df[oq] != 0, df[iq] / df[oq] * 100, np.nan) if has_qty else np.nan
    df["__FR_Value"] = np.where(df[ov] != 0, df[iv] / df[ov] * 100, np.nan) if has_val else np.nan
    df["__Pending_Qty"] = df[oq].fillna(0) - df[iq].fillna(0) if has_qty else np.nan
    df["__Pending_Value"] = df[ov].fillna(0) - df[iv].fillna(0) if has_val else np.nan

    tat = CONFIG["actual_delivery_days"]
    if tat in df.columns:
        df["__Valid_TAT"] = clean_numeric(df[tat])
        df.loc[(df["__Valid_TAT"] < 0) | (df["__Valid_TAT"] > 365), "__Valid_TAT"] = np.nan
    else:
        df["__Valid_TAT"] = np.nan
    return df


def summarize(df):
    oq = df[CONFIG["order_qty"]].sum()
    iq = df[CONFIG["invoice_qty"]].sum()
    ov = df[CONFIG["order_value"]].sum()
    iv = df[CONFIG["invoice_value"]].sum()
    sl = df[CONFIG["sale_loss"]].sum()
    return {
        "order_qty": oq, "invoice_qty": iq, "fr_qty": iq / oq * 100 if oq else np.nan, "pending_qty": oq - iq,
        "order_value": ov, "invoice_value": iv, "fr_value": iv / ov * 100 if ov else np.nan, "sale_loss": sl,
    }


def change_text(cur, prev, kind):
    d = cur - prev
    if kind == "pct":
        return f"{d:+.1f}%"
    if kind == "currency":
        return f"₹ {d:+,.0f}"
    if kind == "sale":
        return f"₹ {d:+,.2f}"
    return f"{d:+,.0f}"


def comparison_table(cur, prev, ytd):
    rows = [
        ("Order Qty", "number", "order_qty"), ("Invoice Qty", "number", "invoice_qty"),
        ("Fill Rate — Qty", "pct", "fr_qty"), ("Pending Qty", "number", "pending_qty"),
        ("Order Value", "currency", "order_value"), ("Invoice Value", "currency", "invoice_value"),
        ("Fill Rate — Value", "pct", "fr_value"), ("Sale Loss", "sale", "sale_loss"),
    ]
    out = []
    for label, kind, key in rows:
        fmt = fmt_pct if kind == "pct" else fmt_sale if kind == "sale" else fmt_currency if kind == "currency" else fmt_num
        out.append([label, fmt(prev[key]), fmt(cur[key]), change_text(cur[key], prev[key], kind), fmt(ytd[key])])
    return pd.DataFrame(out, columns=["Metric", "Previous Month", "Current Month", "Change", "YTD"])


def monthly_summary(df):
    g = df.groupby("Month", dropna=False).agg(
        order_qty=(CONFIG["order_qty"], "sum"), invoice_qty=(CONFIG["invoice_qty"], "sum"),
        order_value=(CONFIG["order_value"], "sum"), invoice_value=(CONFIG["invoice_value"], "sum"),
        sale_loss=(CONFIG["sale_loss"], "sum"),
    ).reset_index()
    g["fr_qty"] = np.where(g.order_qty != 0, g.invoice_qty / g.order_qty * 100, np.nan)
    g["fr_value"] = np.where(g.order_value != 0, g.invoice_value / g.order_value * 100, np.nan)
    g["pending_qty"] = g.order_qty - g.invoice_qty
    g["pending_value"] = g.order_value - g.invoice_value
    return g.sort_values("Month", key=lambda s: s.map(month_key)).reset_index(drop=True)


def customer_summary(df):
    g = df.groupby(CONFIG["customer"], dropna=False).agg(
        order_count=(CONFIG["order_id"], "nunique"),
        invoice_count=(CONFIG["invoice_number"], lambda x: x.dropna().astype(str).nunique()),
        order_qty=(CONFIG["order_qty"], "sum"), invoice_qty=(CONFIG["invoice_qty"], "sum"),
        sale_loss=(CONFIG["sale_loss"], "sum"), tat_avg=("__Valid_TAT", "mean"),
    ).reset_index()
    g["fill_rate"] = np.where(g.order_qty != 0, g.invoice_qty / g.order_qty * 100, np.nan)
    return g


@st.dialog("Customer Order Details", width="large")
def show_customer_details(customer_name, df):
    rows = df[df[CONFIG["customer"]].astype(str) == str(customer_name)].copy()
    if rows.empty:
        st.warning("No order details found.")
        return
    rows["Fill Rate"] = np.where(rows[CONFIG["order_qty"]] != 0, rows[CONFIG["invoice_qty"]] / rows[CONFIG["order_qty"]] * 100, np.nan)
    rows["TAT"] = rows["__Valid_TAT"]

    wanted = [
        CONFIG["wh_receiving_date"], CONFIG["customer"], CONFIG["db_code"], CONFIG["category"], CONFIG["order_id"],
        CONFIG["external_document"], CONFIG["order_qty"], CONFIG["order_value"], CONFIG["invoice_date"],
        CONFIG["invoice_number"], CONFIG["invoice_qty"], CONFIG["invoice_value"], "Fill Rate", CONFIG["sale_loss"],
        CONFIG["dispatch_date"], CONFIG["awb"], CONFIG["courier"], CONFIG["mode"], CONFIG["delivery_status"],
        CONFIG["delivery_date"], "TAT", CONFIG["standard_tat"], CONFIG["variance"], CONFIG["order_to_wh"],
        CONFIG["oti"], CONFIG["otd"], CONFIG["otde"], CONFIG["dispatch_to_delivery"], CONFIG["otd_bucket"],
        CONFIG["wh_remarks"], CONFIG["wh_remark"], CONFIG["logistics_remarks"], CONFIG["ho_remarks"], CONFIG["final_remarks"],
    ]
    selected = []
    for c in wanted:
        if c in ["Fill Rate", "TAT"]:
            selected.append(c)
        else:
            a = resolve_col(c, rows.columns)
            if a and a not in selected:
                selected.append(a)

    display = rows[selected].copy()
    for c in [CONFIG["order_received_date"], CONFIG["order_upload_date"], CONFIG["wh_receiving_date"],
              CONFIG["invoice_date"], CONFIG["dispatch_date"], CONFIG["delivery_date"]]:
        if c in display.columns:
            display[c] = pd.to_datetime(display[c], errors="coerce").dt.strftime("%d-%m-%Y")

    st.caption(f"{len(display):,} order rows for **{customer_name}**")
    st.dataframe(display, use_container_width=True, hide_index=True, height=550)


url = get_secret(CONFIG["sharepoint_secret"])
if not url:
    st.error("SHAREPOINT_EXCEL_URL is not configured in Streamlit secrets.")
    st.stop()

with st.sidebar:
    if st.button("🔄 Refresh data", use_container_width=True):
        get_workbook.clear()
        load_data.clear()
        st.rerun()

try:
    df, sheets_used = load_data(url)
    df = add_derived(df)
except (FetchError, ValueError) as e:
    st.error(f"Could not load Excel file: {e}")
    st.stop()

required = [
    CONFIG["customer"], CONFIG["order_qty"], CONFIG["invoice_qty"], CONFIG["order_value"],
    CONFIG["invoice_value"], CONFIG["order_id"], CONFIG["category"], CONFIG["sale_loss"],
]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error("Missing required columns: " + ", ".join(missing))
    st.stop()

with st.expander("🔎 Filters", expanded=False):
    # Channel/Zone/Name are optional — some workbooks won't have them,
    # so each filter is only shown (and only ever indexed) when its
    # column actually exists, instead of assuming it's always there.
    f1, f2, f3 = st.columns(3)
    with f1:
        selected_months = st.multiselect(
            "Month (multi-select)",
            sheets_used,
            default=sheets_used,
            key="fill_rate_month_filter",
        )
    with f2:
        selected_categories = st.multiselect("Category", sorted(df[CONFIG["category"]].dropna().astype(str).unique()))
    with f3:
        selected_customers = st.multiselect("Customer", sorted(df[CONFIG["customer"]].dropna().astype(str).unique()))

    f4, f5, f6 = st.columns(3)
    optional_filters = [("channel", "Channel", f4), ("zone", "Zone", f5), ("name", "Name", f6)]
    selected_optional = {}
    for key, label, col in optional_filters:
        cname = CONFIG[key]
        with col:
            if cname in df.columns:
                selected_optional[cname] = st.multiselect(label, sorted(df[cname].dropna().astype(str).unique()))
            else:
                st.caption(f"_{label} column not found — filter unavailable._")

filtered = df.copy()
if selected_months:
    filtered = filtered[filtered.Month.isin(selected_months)]

active_filter_cols = [
    (CONFIG["category"], selected_categories),
    (CONFIG["customer"], selected_customers),
    *selected_optional.items(),
]
for col, vals in active_filter_cols:
    if vals:
        filtered = filtered[filtered[col].astype(str).isin(vals)]
if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

available = sorted(sheets_used, key=month_key)
selected_sorted = sorted(filtered.Month.dropna().unique(), key=month_key)
current_month = selected_sorted[-1]
current_idx = available.index(current_month)
previous_month = available[current_idx - 1] if current_idx > 0 else None
current_df = filtered[filtered.Month == current_month]
previous_df = filtered[filtered.Month == previous_month] if previous_month else pd.DataFrame()

# YTD always starts from the first available month and ends at the current month. Month filter does not truncate YTD.
ytd_months = available[:current_idx + 1]
ytd_df = df[df.Month.isin(ytd_months)].copy()
for col, vals in active_filter_cols:
    if vals:
        ytd_df = ytd_df[ytd_df[col].astype(str).isin(vals)]

cur = summarize(current_df)
prev = summarize(previous_df) if not previous_df.empty else None
ytd = summarize(ytd_df)

st.markdown('<div class="section-title">Executive Overview</div>', unsafe_allow_html=True)
st.caption(f"Current Month: **{current_month}** | Previous Month: **{previous_month or '—'}**")
a, b, c, d = st.columns(4)
a.metric("Order Qty", fmt_num(cur["order_qty"]))
b.metric("Invoice Qty", fmt_num(cur["invoice_qty"]))
c.metric("Fill Rate — Qty", fmt_pct(cur["fr_qty"]))
d.metric("Pending Qty", fmt_num(cur["pending_qty"]))
a, b, c, d = st.columns(4)
a.metric("Order Value", fmt_currency(cur["order_value"]))
b.metric("Invoice Value", fmt_currency(cur["invoice_value"]))
c.metric("Fill Rate — Value", fmt_pct(cur["fr_value"]))
d.metric("Sale Loss", fmt_sale(cur["sale_loss"]))

st.markdown('<div class="section-title">Current Month vs Previous Month — YTD</div>', unsafe_allow_html=True)
st.markdown(f'<div class="dashboard-subtitle"><b>{current_month}</b> vs <b>{previous_month or "—"}</b> | YTD: <b>{ytd_months[0]}–{current_month}</b></div>', unsafe_allow_html=True)
if prev is not None:
    st.dataframe(comparison_table(cur, prev, ytd), use_container_width=True, hide_index=True, height=335)
else:
    st.info("A previous month is required for comparison.")

tab_mom, tab_customer, tab_category = st.tabs(["📈 MOM Overview", "🏪 Customer Wise", "📦 Category Wise"])

with tab_mom:
    monthly = monthly_summary(filtered)
    st.markdown('<div class="section-title">Month-on-Month Fill Rate Comparison</div>', unsafe_allow_html=True)
    st.markdown('<div class="dashboard-subtitle">Monthly Fill Rate trend for the selected data.</div>', unsafe_allow_html=True)

    fr = monthly[["Month", "fr_qty", "fr_value"]].melt(id_vars="Month", var_name="Metric", value_name="Fill Rate")
    fr["Metric"] = fr["Metric"].replace({"fr_qty": "Fill Rate — Qty", "fr_value": "Fill Rate — Value"})
    base = alt.Chart(fr).encode(
        x=alt.X("Month:N", sort=available, title="Month", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Fill Rate:Q", title="Fill Rate (%)", scale=alt.Scale(zero=False)),
    )
    line = base.mark_line(point=True, strokeWidth=3).encode(
        color=alt.Color("Metric:N", title="Metric"),
        tooltip=["Month", "Metric", alt.Tooltip("Fill Rate:Q", format=".1f")],
    ).properties(height=390)
    q = alt.Chart(fr[fr.Metric == "Fill Rate — Qty"]).mark_text(dy=-14, fontWeight="bold").encode(
        x=alt.X("Month:N", sort=available), y="Fill Rate:Q", text=alt.Text("Fill Rate:Q", format=".1f"))
    v = alt.Chart(fr[fr.Metric == "Fill Rate — Value"]).mark_text(dy=16, fontWeight="bold").encode(
        x=alt.X("Month:N", sort=available), y="Fill Rate:Q", text=alt.Text("Fill Rate:Q", format=".1f"))
    st.altair_chart(line + q + v, use_container_width=True)

    st.markdown('<div class="section-title">Month-on-Month Table</div>', unsafe_allow_html=True)
    mom_map = {
        "Order Qty": "order_qty", "Invoice Qty": "invoice_qty", "Fill Rate — Qty": "fr_qty", "Pending Qty": "pending_qty",
        "Order Value": "order_value", "Invoice Value": "invoice_value", "Fill Rate — Value": "fr_value",
        "Pending Value": "pending_value", "Sale Loss": "sale_loss",
    }
    selected_cols = st.multiselect("Choose columns", list(mom_map), default=list(mom_map), key="mom_cols") or list(mom_map)
    md = monthly[["Month"] + [mom_map[x] for x in selected_cols]].rename(columns={v: k for k, v in mom_map.items()})
    for col in md.columns:
        if col == "Month":
            continue
        if "Fill Rate" in col:
            md[col] = md[col].map(fmt_pct)
        elif col == "Sale Loss":
            md[col] = md[col].map(fmt_sale)
        elif "Value" in col:
            md[col] = md[col].map(fmt_currency)
        else:
            md[col] = md[col].map(fmt_num)
    st.dataframe(md, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Fill Rate — Qty vs Value</div>', unsafe_allow_html=True)
    st.altair_chart(
        alt.Chart(fr).mark_bar().encode(
            x=alt.X("Month:N", sort=available, title="Month", axis=alt.Axis(labelAngle=0)),
            xOffset="Metric:N", y=alt.Y("Fill Rate:Q", title="Fill Rate (%)"),
            color=alt.Color("Metric:N", title="Metric"),
            tooltip=["Month", "Metric", alt.Tooltip("Fill Rate:Q", format=".1f")],
        ).properties(height=350),
        use_container_width=True,
    )

    st.markdown('<div class="section-title">Sale Loss MOM</div>', unsafe_allow_html=True)
    sl = alt.Chart(monthly).mark_bar().encode(
        x=alt.X("Month:N", sort=available, title="Month", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("sale_loss:Q", title="Sale Loss (₹)"),
        tooltip=["Month", alt.Tooltip("sale_loss:Q", format=",.2f")],
    ).properties(height=350)
    sll = alt.Chart(monthly).mark_text(dy=-8, fontWeight="bold").encode(
        x=alt.X("Month:N", sort=available), y="sale_loss:Q", text=alt.Text("sale_loss:Q", format=",.2f"))
    st.altair_chart(sl + sll, use_container_width=True)

with tab_customer:
    st.markdown('<div class="section-title">Customer Wise Performance</div>', unsafe_allow_html=True)
    st.markdown('<div class="dashboard-subtitle">Customer-level performance for the selected month(s). Select a customer below to open order-level details.</div>', unsafe_allow_html=True)

    cs = customer_summary(filtered)
    search = st.text_input("🔍 Search Customer", key="customer_search")
    if search:
        cs = cs[cs[CONFIG["customer"]].astype(str).str.contains(search, case=False, na=False)]

    customer_choice = st.selectbox(
        "Customer detail", ["Select a customer"] + cs[CONFIG["customer"]].dropna().astype(str).tolist(),
        key="customer_detail_choice",
    )
    if customer_choice != "Select a customer":
        show_customer_details(customer_choice, filtered)

    cd = cs.copy()
    cd["Order (count)"] = cd.order_count.map(fmt_num)
    cd["Invoice (count)"] = cd.invoice_count.map(fmt_num)
    cd["Order Qty"] = cd.order_qty.map(fmt_num)
    cd["Invoice Qty"] = cd.invoice_qty.map(fmt_num)
    cd["Fill Rate"] = cd.fill_rate.map(fmt_pct)
    cd["Sale Loss (In Lacs)"] = cd.sale_loss.map(fmt_sale)
    cd["TAT (avg)"] = cd.tat_avg.map(lambda x: "—" if pd.isna(x) or x < 0 or x > 365 else f"{x:.1f}")
    st.dataframe(
        cd[[CONFIG["customer"], "Order (count)", "Invoice (count)", "Order Qty", "Invoice Qty",
            "Fill Rate", "Sale Loss (In Lacs)", "TAT (avg)"]],
        use_container_width=True, hide_index=True, height=500,
    )

with tab_category:
    st.markdown('<div class="section-title">Category Wise Performance</div>', unsafe_allow_html=True)
    st.markdown('<div class="dashboard-subtitle">Category-level order, invoice, fill rate and sale loss performance.</div>', unsafe_allow_html=True)

    cat = filtered.groupby(CONFIG["category"], dropna=False).agg(
        orders=(CONFIG["order_id"], "nunique"), order_qty=(CONFIG["order_qty"], "sum"),
        invoice_qty=(CONFIG["invoice_qty"], "sum"), order_value=(CONFIG["order_value"], "sum"),
        invoice_value=(CONFIG["invoice_value"], "sum"), sale_loss=(CONFIG["sale_loss"], "sum"),
    ).reset_index()
    cat["fr_qty"] = np.where(cat.order_qty != 0, cat.invoice_qty / cat.order_qty * 100, np.nan)
    cat["fr_value"] = np.where(cat.order_value != 0, cat.invoice_value / cat.order_value * 100, np.nan)
    cat["pending_qty"] = cat.order_qty - cat.invoice_qty
    cat["pending_value"] = cat.order_value - cat.invoice_value
    cat[CONFIG["category"]] = (
        cat[CONFIG["category"]].fillna("Blank / Not Available").astype(str)
        .replace({"None": "Blank / Not Available", "nan": "Blank / Not Available", "": "Blank / Not Available"})
    )

    cd = cat[[CONFIG["category"], "orders", "order_qty", "invoice_qty", "fr_qty", "fr_value",
              "pending_qty", "order_value", "invoice_value", "pending_value", "sale_loss"]].rename(columns={
        CONFIG["category"]: "Category", "orders": "Orders", "order_qty": "Order Qty", "invoice_qty": "Invoice Qty",
        "fr_qty": "FR % Qty", "fr_value": "FR % Value", "pending_qty": "Pending Qty", "order_value": "Order Value",
        "invoice_value": "Invoice Value", "pending_value": "Pending Value", "sale_loss": "Sale Loss",
    })
    for c in ["Orders", "Order Qty", "Invoice Qty", "Pending Qty"]:
        cd[c] = cd[c].map(fmt_num)
    for c in ["FR % Qty", "FR % Value"]:
        cd[c] = cd[c].map(fmt_pct)
    for c in ["Order Value", "Invoice Value", "Pending Value"]:
        cd[c] = cd[c].map(fmt_currency)
    cd["Sale Loss"] = cd["Sale Loss"].map(fmt_sale)
    st.dataframe(cd, use_container_width=True, hide_index=True, height=430)

    cc = alt.Chart(cat).mark_bar().encode(
        x=alt.X(f"{CONFIG['category']}:N", title="Category", sort="-y", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("fr_qty:Q", title="Fill Rate (%)"),
        tooltip=[alt.Tooltip(f"{CONFIG['category']}:N", title="Category"), alt.Tooltip("fr_qty:Q", title="Fill Rate", format=".1f")],
    ).properties(height=350)
    ccl = alt.Chart(cat).mark_text(dy=-8, fontWeight="bold").encode(
        x=alt.X(f"{CONFIG['category']}:N", sort="-y"), y="fr_qty:Q", text=alt.Text("fr_qty:Q", format=".1f"))
    st.altair_chart(cc + ccl, use_container_width=True)

st.markdown("---")
st.caption(f"Sheets: {', '.join(sheets_used)} | Rows: {len(filtered):,} | Updated: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
