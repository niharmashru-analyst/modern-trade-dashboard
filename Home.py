"""
HOME — entry point for the multi-page dashboard project.
Streamlit auto-builds the sidebar page-switcher from every file inside
the pages/ folder next to this one — no manual sidebar code needed.
Deploy with this file (Home.py) set as the app's main file.

This page DISCOVERS the actual .py files inside pages/ at runtime
(instead of hardcoding filenames), so st.page_link can never point at
a file that doesn't exist. It picks an icon/title/description for
each discovered file by matching a keyword in the filename — see
KEYWORD_PRIORITY / CARD_INFO below. Any page whose filename doesn't
match a known keyword still gets a card (generic icon, title guessed
from the filename), it just won't have a description.
"""
import os

import streamlit as st

from common import inject_theme

st.set_page_config(page_title="Dashboards Home", layout="wide", page_icon="🏠")
inject_theme()

PAGES_DIR = "pages"

# ================================================================
# CARD LOOKUP — matched by keyword found in the page's filename.
# Add/edit entries here if a page's title/description should change,
# or if a new page needs its own keyword.
# ================================================================
CARD_INFO = {
    "cancel": {
        "icon": "🚫", "title": "Cancelled Orders",
        "description": "Filtered view of the dispatch tracker showing only rows "
                        "matching configured cancel-reason terms (e.g. \"Order "
                        "below 7k\") — editable on the page itself.",
    },
    "fill": {
        "icon": "📶", "title": "Fill Rate",
        "description": "Compares Fill Rate (Invoice Qty / Order Qty) across chains "
                        "with a bar chart, then a shop-by-shop breakdown — click a "
                        "shop to see every order behind its numbers.",
    },
    "stock": {
        "icon": "📉", "title": "Stock Gap Dashboard",
        "description": "Matches a live order sheet against live stock and shows "
                        "exactly where orders exceed available stock — Ahmedabad "
                        "(MWH) or Bangalore (Direct Shelf BLR).",
    },
    "dispatch": {
        "icon": "📦", "title": "Order Tracking",
        "description": "Search and filter live order/delivery data (AWB, courier, "
                        "delivery status, TAT) — reads straight from the live "
                        "SharePoint file, always up to date.",
    },
    "order": {
        "icon": "📦", "title": "Order Tracking",
        "description": "Search and filter live order/delivery data (AWB, courier, "
                        "delivery status, TAT) — reads straight from the live "
                        "SharePoint file, always up to date.",
    },
}

# Checked in this order, first hit wins — more specific keywords
# (cancel/fill/stock/dispatch) are listed before the generic "order"
# one so e.g. "Cancelled_Orders.py" doesn't get mis-tagged as Order
# Tracking just because it contains "order".
KEYWORD_PRIORITY = ["cancel", "fill", "stock", "dispatch", "order"]


def _card_for_filename(filename: str) -> dict:
    key = filename.lower()
    for kw in KEYWORD_PRIORITY:
        if kw in key:
            return CARD_INFO[kw]
    return {
        "icon": "📄",
        "title": filename[:-3].replace("_", " ").replace("-", " ").title(),
        "description": "",
    }


def _discover_pages() -> list:
    if not os.path.isdir(PAGES_DIR):
        return []
    files = sorted(f for f in os.listdir(PAGES_DIR) if f.endswith(".py"))
    return [{"path": f"{PAGES_DIR}/{f}", **_card_for_filename(f)} for f in files]


st.title("🏠 Dashboards")
st.caption("Click a card to open a dashboard.")
st.write("")

pages = _discover_pages()

if not pages:
    st.warning(f"No pages found in the '{PAGES_DIR}/' folder yet.")
else:
    cols = st.columns(2)
    for i, page in enumerate(pages):
        with cols[i % 2]:
            with st.container(border=True):
                st.markdown(f"#### {page['icon']} {page['title']}")
                if page["description"]:
                    st.caption(page["description"])
                st.page_link(page["path"], label="Open →", use_container_width=True)
