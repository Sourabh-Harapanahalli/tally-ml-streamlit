"""
Look and feel for the Tally AI app.

Everything visual that is shared between pages lives here: the global style
sheet, the sidebar navigation, the home screen with its tool cards, the page
header, numbered step headings and the Excel-style file preview. ``app.py``
keeps the per-tool logic and calls into this module for presentation.
"""

import html
import io
import re

import pandas as pd
import streamlit as st

HOME = "Home"

# --------------------------------------------------------------------------
# Global style sheet. Selectors use Streamlit's data-testid attributes and the
# ``st-key-<key>`` class Streamlit puts on any element created with a key.
# --------------------------------------------------------------------------
_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
  --tm-ink: #1c2033;
  --tm-text: #3d4358;
  --tm-muted: #6f768a;
  --tm-line: #e7e9f0;
  --tm-card: #ffffff;
  --tm-accent: #5b5fd6;
  --tm-accent-2: #7a6ee6;
  --tm-accent-ink: #4347b8;
  --tm-accent-soft: #f1f2fe;
  --tm-accent-line: #dcdefa;
  --tm-excel: #1d7a4a;
  --tm-radius: 16px;
  --tm-shadow: 0 1px 2px rgba(28, 32, 51, .04), 0 6px 20px rgba(28, 32, 51, .05);
  --tm-shadow-lift: 0 2px 6px rgba(28, 32, 51, .05), 0 16px 36px rgba(91, 95, 214, .12);
}

html, body, .stApp, [data-testid="stAppViewContainer"],
[data-testid="stMarkdownContainer"], button, input, textarea, select {
  font-family: 'Inter', -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}

/* ---- Page canvas ---------------------------------------------------- */
[data-testid="stAppViewContainer"] {
  background:
    radial-gradient(900px 460px at 92% -10%, rgba(196, 203, 255, .28), transparent 62%),
    radial-gradient(700px 400px at 30% -14%, rgba(214, 240, 232, .30), transparent 60%),
    #f8f9fb;
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stMainBlockContainer"], .main .block-container {
  max-width: 1120px;
  padding-top: 2.4rem;
  padding-bottom: 5rem;
}
h1, h2, h3 { color: var(--tm-ink); letter-spacing: -.02em; }
[data-testid="stMain"] [data-testid="stMarkdownContainer"] p, [data-testid="stMain"] li { color: var(--tm-text); }
h3 { font-weight: 700; }

/* ---- Sidebar: light, soft, same family as the page ------------------- */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #ffffff 0%, #fafafd 100%);
  border-right: 1px solid var(--tm-line);
}
[data-testid="stSidebar"] [data-testid="stSidebarContent"] { padding-top: .5rem; }
.tm-brand { display: flex; align-items: center; gap: .7rem; padding: .2rem .2rem 1rem; }
.tm-brand-mark {
  width: 40px; height: 40px; border-radius: 12px; display: grid; place-items: center;
  font-size: 1.25rem; background: var(--tm-accent-soft); border: 1px solid var(--tm-accent-line);
}
.tm-brand-name { font-weight: 800; font-size: 1.08rem; color: var(--tm-ink); letter-spacing: -.01em; }
.tm-brand-sub { font-size: .74rem; color: var(--tm-muted); margin-top: -2px; }
[data-testid="stSidebar"] .stButton { margin-bottom: -.72rem; }
[data-testid="stSidebar"] .stButton button {
  background: transparent; border: 0; box-shadow: none; border-radius: 10px;
  justify-content: flex-start; text-align: left; padding: .42rem .6rem; min-height: 0;
  color: var(--tm-text); font-weight: 500; font-size: .88rem;
  transition: background .18s ease, color .18s ease, transform .18s ease;
}
[data-testid="stSidebar"] .stButton button > div { justify-content: flex-start; width: 100%; }
[data-testid="stSidebar"] .stButton button p { font-size: .88rem; text-align: left; color: var(--tm-text); }
[data-testid="stSidebar"] .stButton button:hover { background: #f3f4f8; transform: translateX(2px); }
[data-testid="stSidebar"] .stButton button:hover p { color: var(--tm-ink); }
[data-testid="stSidebar"] [class*="st-key-navA_"] button {
  background: var(--tm-accent-soft); box-shadow: inset 3px 0 0 var(--tm-accent);
}
[data-testid="stSidebar"] [class*="st-key-navA_"] button:hover { background: var(--tm-accent-soft); }
[data-testid="stSidebar"] [class*="st-key-navA_"] button p { color: var(--tm-accent-ink); font-weight: 650; }

/* Section groups: highlighted, expandable headers */
[data-testid="stSidebar"] [class*="st-key-navsec"] { margin-top: .35rem; }
[data-testid="stSidebar"] [data-testid="stExpander"] details {
  background: #fff; border: 1px solid var(--tm-line); border-radius: 14px;
  box-shadow: 0 1px 2px rgba(28, 32, 51, .03); overflow: hidden;
  transition: border-color .2s ease, box-shadow .2s ease;
}
[data-testid="stSidebar"] [data-testid="stExpander"] details:hover {
  border-color: var(--tm-accent-line); box-shadow: 0 4px 14px rgba(28, 32, 51, .06);
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
  padding: .62rem .75rem; border-radius: 13px; background: transparent; transition: background .2s ease;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover { background: #f7f8fb; }
[data-testid="stSidebar"] [data-testid="stExpander"] summary p {
  font-size: .92rem; font-weight: 700; color: var(--tm-ink); letter-spacing: -.005em;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary svg { color: var(--tm-muted); fill: var(--tm-muted); }
[data-testid="stSidebar"] [data-testid="stExpander"] details[open] summary {
  border-bottom: 1px solid var(--tm-line); border-radius: 13px 13px 0 0;
}
[data-testid="stSidebar"] [class*="st-key-navsecA_"] [data-testid="stExpander"] details {
  border-color: var(--tm-accent-line); box-shadow: 0 6px 18px rgba(91, 95, 214, .09);
}
[data-testid="stSidebar"] [class*="st-key-navsecA_"] [data-testid="stExpander"] summary {
  background: linear-gradient(90deg, var(--tm-accent-soft), #fafaff);
}
[data-testid="stSidebar"] [class*="st-key-navsecA_"] [data-testid="stExpander"] summary p { color: var(--tm-accent-ink); }
[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stExpanderDetails"] { padding: .45rem .4rem .75rem; }
[data-testid="stSidebar"] [class*="st-key-navsec"] [data-testid="stVerticalBlock"],
[data-testid="stSidebar"] [class*="st-key-navsec"] [data-testid="stElementContainer"] { width: 100% !important; max-width: 100%; }

.tm-side-foot { font-size: .72rem; color: #9aa0b2; padding: 1.6rem .55rem .5rem; line-height: 1.5; }

/* ---- Buttons -------------------------------------------------------- */
[data-testid="stMain"] .stButton button,
[data-testid="stMain"] .stDownloadButton button,
[data-testid="stMain"] .stFormSubmitButton button {
  border-radius: 12px; font-weight: 600; border: 1px solid var(--tm-line);
  transition: transform .16s ease, box-shadow .16s ease, border-color .16s ease, background .16s ease;
}
[data-testid="stMain"] .stButton button:hover,
[data-testid="stMain"] .stDownloadButton button:hover,
[data-testid="stMain"] .stFormSubmitButton button:hover {
  transform: translateY(-1px); border-color: var(--tm-accent-line);
  box-shadow: 0 6px 16px rgba(28, 32, 51, .08);
}
[data-testid="stMain"] .stButton button:active,
[data-testid="stMain"] .stDownloadButton button:active { transform: translateY(0); }
[data-testid="stMain"] button[kind="primary"],
[data-testid="stMain"] button[data-testid$="-primary"] {
  background: var(--tm-accent); border: 0; color: #fff;
  box-shadow: 0 1px 2px rgba(28, 32, 51, .10), 0 6px 16px rgba(91, 95, 214, .20);
}
[data-testid="stMain"] button[kind="primary"]:hover,
[data-testid="stMain"] button[data-testid$="-primary"]:hover {
  background: #4f53c8; box-shadow: 0 2px 4px rgba(28, 32, 51, .10), 0 10px 22px rgba(91, 95, 214, .26);
}
[data-testid="stMain"] .stDownloadButton button:not([kind="primary"]):not([data-testid$="-primary"]) {
  background: #fff; border-color: var(--tm-line);
}
[data-testid="stMain"] button[kind="primary"] p,
[data-testid="stMain"] button[data-testid$="-primary"] p { color: #fff; }

/* ---- Inputs, uploads, tables ---------------------------------------- */
[data-testid="stFileUploaderDropzone"] {
  border: 1.5px dashed #d3d6e6; border-radius: var(--tm-radius);
  background: #fff;
  padding: 1.4rem 1.2rem; transition: border-color .2s ease, background .2s ease, transform .2s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: var(--tm-accent); background: var(--tm-accent-soft); transform: translateY(-1px);
}
[data-testid="stTextInput"] input, [data-testid="stNumberInput"] input,
[data-testid="stTextArea"] textarea, [data-baseweb="select"] > div {
  border-radius: 12px !important;
}
[data-testid="stDataFrame"], [data-testid="stDataEditor"], [data-testid="stTable"] {
  border-radius: 14px; overflow: hidden; box-shadow: var(--tm-shadow);
  border: 1px solid var(--tm-line);
}
[data-testid="stAlert"], [data-testid="stAlertContainer"] { border-radius: 14px; }
[data-testid="stAlert"] { box-shadow: 0 1px 2px rgba(15, 23, 42, .04); }
[data-testid="stExpander"] details {
  border-radius: 14px; border: 1px solid var(--tm-line); background: #fff;
  box-shadow: 0 1px 2px rgba(15, 23, 42, .03);
}
[data-testid="stMetric"] {
  background: #fff; border: 1px solid var(--tm-line); border-radius: 14px;
  padding: .9rem 1rem; box-shadow: var(--tm-shadow);
}
[data-testid="stMetricValue"] { font-weight: 700; letter-spacing: -.02em; font-size: 1.7rem; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: var(--tm-radius); }
[data-testid="stProgress"] > div > div > div > div {
  background: var(--tm-accent);
}
hr { border-color: var(--tm-line) !important; margin: 1.6rem 0 !important; }

/* ---- Tabs (selectors cover old and new Streamlit markup) ------------ */
[data-testid="stTabs"] [role="tablist"] { gap: .4rem; border-bottom: 0 !important; box-shadow: none !important; padding: .15rem 0 .5rem; }
[data-testid="stTabs"] [role="tab"] {
  border-radius: 999px !important; padding: .35rem .95rem !important; height: auto !important; background: #fff;
  border: 1px solid var(--tm-line) !important; transition: border-color .16s ease, background .16s ease;
  margin: 0 !important;
}
[data-testid="stTabs"] [role="tab"]:hover { border-color: var(--tm-accent-line) !important; background: #fafaff; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
  background: var(--tm-accent-soft); border-color: var(--tm-accent-line) !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p { color: var(--tm-accent-ink); font-weight: 600; }
[data-testid="stTabs"] [role="tab"]::before, [data-testid="stTabs"] [role="tab"]::after,
[data-testid="stTabs"] [role="tablist"]::before, [data-testid="stTabs"] [role="tablist"]::after { display: none !important; }
[data-testid="stTabs"] [data-baseweb="tab-highlight"],
[data-testid="stTabs"] [data-baseweb="tab-border"] { display: none !important; }

/* ---- Hero (home): soft pastel mesh, dark text ---------------------- */
.tm-hero {
  position: relative; overflow: hidden; border-radius: 24px; padding: 2.5rem 2.6rem 2.3rem;
  background:
    radial-gradient(560px 320px at 6% -6%, #dde2ff 0%, transparent 64%),
    radial-gradient(520px 320px at 98% 4%, #fbe4ee 0%, transparent 62%),
    radial-gradient(620px 340px at 60% 122%, #d4f2e4 0%, transparent 64%),
    #ffffff;
  border: 1px solid var(--tm-line); box-shadow: var(--tm-shadow);
  animation: tm-rise .55s cubic-bezier(.2, .7, .2, 1) both;
}
.tm-hero::before, .tm-hero::after {
  content: ""; position: absolute; border-radius: 50%; filter: blur(28px); opacity: .55;
  animation: tm-float 14s ease-in-out infinite; pointer-events: none;
}
.tm-hero::before { width: 260px; height: 260px; right: -40px; top: -110px; background: #dfe3ff; }
.tm-hero::after {
  width: 220px; height: 220px; right: 210px; bottom: -150px; background: #d8f3e7; animation-delay: -6s;
}
.tm-hero-tag {
  position: relative; display: inline-block; font-size: .74rem; font-weight: 600; letter-spacing: .03em;
  color: var(--tm-accent-ink); background: rgba(255, 255, 255, .75); border: 1px solid var(--tm-accent-line);
  padding: .26rem .7rem; border-radius: 999px; backdrop-filter: blur(6px);
}
.tm-hero h1 {
  position: relative; color: var(--tm-ink); font-size: 2.25rem; font-weight: 800; line-height: 1.14;
  margin: .9rem 0 .55rem; padding: 0; letter-spacing: -.03em;
}
.tm-hero p { position: relative; color: var(--tm-text); font-size: 1.04rem; max-width: 620px; margin: 0; }
.tm-flow { display: grid; grid-template-columns: repeat(3, 1fr); gap: .9rem; margin: 1.1rem 0 .4rem; }
.tm-flow-item {
  display: flex; align-items: center; gap: .8rem; background: #fff; border: 1px solid var(--tm-line);
  border-radius: 16px; padding: .85rem 1rem; box-shadow: var(--tm-shadow);
  animation: tm-rise .55s cubic-bezier(.2, .7, .2, 1) both;
}
.tm-flow-item:nth-child(2) { animation-delay: .07s; }
.tm-flow-item:nth-child(3) { animation-delay: .14s; }
.tm-flow-num {
  flex: none; width: 34px; height: 34px; border-radius: 11px; display: grid; place-items: center;
  font-weight: 700; color: var(--tm-accent-ink); background: var(--tm-accent-soft);
}
.tm-flow-title { font-weight: 650; color: var(--tm-ink); font-size: .95rem; }
.tm-flow-sub { color: var(--tm-muted); font-size: .82rem; margin-top: 1px; }

.tm-section { margin: 2rem 0 .2rem; }
.tm-section-title { font-size: 1.18rem; font-weight: 750; color: var(--tm-ink); letter-spacing: -.02em; }
.tm-section-sub { color: var(--tm-muted); font-size: .92rem; margin-top: .1rem; }

.st-key-home_search { margin: .9rem 0 .2rem; }
.st-key-home_search input {
  background: #fff; border-radius: 14px !important; padding: .85rem 1rem; font-size: 1rem;
  box-shadow: var(--tm-shadow);
}
.st-key-home_search [data-baseweb="input"], .st-key-home_search [data-baseweb="base-input"] {
  border-radius: 14px !important; background: #fff;
}

/* ---- Tool cards (home) ---------------------------------------------- */
[class*="st-key-card_"] {
  background: var(--tm-card); border: 1px solid var(--tm-line) !important; border-radius: 18px !important;
  padding: 1.05rem 1.05rem .95rem; box-shadow: var(--tm-shadow); height: 100%;
  transition: transform .2s cubic-bezier(.2, .7, .2, 1), box-shadow .2s ease, border-color .2s ease;
  animation: tm-rise .5s cubic-bezier(.2, .7, .2, 1) both;
}
[class*="st-key-card_"]:hover {
  transform: translateY(-3px); box-shadow: var(--tm-shadow-lift); border-color: var(--tm-accent-line) !important;
}
.tm-card-icon {
  width: 46px; height: 46px; border-radius: 14px; display: grid; place-items: center; font-size: 1.4rem;
  background: #f5f6fa; border: 1px solid var(--tm-line); margin-bottom: .7rem;
  transition: transform .25s ease;
}
[class*="st-key-card_"]:hover .tm-card-icon { transform: scale(1.08) rotate(-4deg); }
.tm-card-title { font-weight: 700; color: var(--tm-ink); font-size: 1rem; line-height: 1.3; }
.tm-card-desc { color: var(--tm-muted); font-size: .86rem; line-height: 1.45; margin-top: .3rem; min-height: 4.4em; }
.tm-card-title { min-height: 2.6em; display: flex; align-items: flex-end; }
[class*="st-key-card_"] .stButton button {
  background: #f5f6fa; border: 1px solid var(--tm-line); color: var(--tm-text); font-weight: 600;
}
[class*="st-key-card_"] .stButton button:hover {
  background: var(--tm-accent-soft); border-color: var(--tm-accent-line); box-shadow: none;
}
[class*="st-key-card_"] .stButton button:hover p { color: var(--tm-accent-ink); }

/* ---- Page header and steps ------------------------------------------ */
.tm-crumb { font-size: .8rem; color: var(--tm-muted); font-weight: 500; margin-bottom: .7rem; }
.tm-crumb b { color: var(--tm-accent-ink); font-weight: 600; }
.tm-head {
  display: flex; align-items: center; gap: 1rem; margin-bottom: 1.1rem;
  animation: tm-rise .45s cubic-bezier(.2, .7, .2, 1) both;
}
.tm-head-icon {
  flex: none; width: 60px; height: 60px; border-radius: 18px; display: grid; place-items: center;
  font-size: 1.8rem; background: #fff; border: 1px solid var(--tm-line);
  box-shadow: var(--tm-shadow);
}
.tm-head-title { font-size: 1.85rem; font-weight: 800; color: var(--tm-ink); letter-spacing: -.03em; line-height: 1.15; }
.tm-head-sub { color: var(--tm-muted); font-size: 1rem; margin-top: .2rem; }
.tm-step { display: flex; align-items: center; gap: .75rem; margin: .5rem 0 .7rem; }
.tm-step-num {
  flex: none; min-width: 34px; height: 34px; padding: 0 .55rem; border-radius: 11px; display: grid;
  place-items: center; font-weight: 700; font-size: .95rem; color: var(--tm-accent-ink);
  background: var(--tm-accent-soft); border: 1px solid var(--tm-accent-line);
}
.tm-step-title { font-size: 1.2rem; font-weight: 700; color: var(--tm-ink); letter-spacing: -.02em; }

/* ---- Excel-style preview -------------------------------------------- */
[class*="st-key-xl_"] {
  border: 1px solid #cfe3d7 !important; border-radius: 14px !important; overflow: hidden;
  background: #fff; box-shadow: var(--tm-shadow); padding: 0 0 .5rem; gap: .5rem;
}
.tm-xl-bar {
  display: flex; align-items: center; gap: .65rem; padding: .6rem .9rem;
  background: #f1f8f4; color: #14532d; border-bottom: 1px solid #d5eadd;
}
.tm-xl-logo {
  width: 26px; height: 26px; border-radius: 6px; background: var(--tm-excel); color: #fff;
  display: grid; place-items: center; font-weight: 800; font-size: .9rem;
}
.tm-xl-name { font-weight: 600; font-size: .92rem; color: #14532d; }
.tm-xl-meta { margin-left: auto; font-size: .78rem; color: #4d7c63; }
[class*="st-key-xl_"] > div:not(:first-child) { padding: 0 .8rem; }
[class*="st-key-xl_"] [data-testid="stTabs"] [role="tab"] { border-radius: 8px !important; }
[class*="st-key-xl_"] [data-testid="stTabs"] [role="tab"][aria-selected="true"] {
  background: #e7f4ec; border-color: var(--tm-excel) !important;
}
[class*="st-key-xl_"] [data-testid="stTabs"] [role="tab"][aria-selected="true"] p { color: #0e6b38; }
[class*="st-key-xl_"] [data-testid="stTabs"] [role="tab"]:hover { border-color: var(--tm-excel) !important; }

/* ---- Result cards (error log) --------------------------------------- */
[class*="st-key-logerr_"], [class*="st-key-logok_"] {
  background: #fff; border: 1px solid var(--tm-line) !important; border-radius: 16px !important;
  padding: 1rem 1.15rem; box-shadow: var(--tm-shadow); border-left: 4px solid #6cc79a !important;
  transition: transform .18s ease, box-shadow .18s ease;
}
[class*="st-key-logerr_"] { border-left-color: #e58a8a !important; }
[class*="st-key-logerr_"]:hover, [class*="st-key-logok_"]:hover { transform: translateY(-2px); box-shadow: var(--tm-shadow-lift); }

/* Older Streamlit gives blocks a fixed pixel width; keep padded cards and
   their contents inside the column. */
[class*="st-key-card_"], [class*="st-key-logerr_"], [class*="st-key-logok_"], [class*="st-key-xl_"] {
  box-sizing: border-box; max-width: 100%;
}
[class*="st-key-card_"] > div, [class*="st-key-logerr_"] > div, [class*="st-key-logok_"] > div {
  width: 100% !important; max-width: 100%;
}
[class*="st-key-card_"] .stButton, [class*="st-key-card_"] .stButton button { width: 100% !important; }
[class*="st-key-card_"] [data-testid="stVerticalBlock"], [class*="st-key-card_"] [data-testid="stElementContainer"],
[class*="st-key-logerr_"] [data-testid="stVerticalBlock"], [class*="st-key-logerr_"] [data-testid="stElementContainer"],
[class*="st-key-logok_"] [data-testid="stVerticalBlock"], [class*="st-key-logok_"] [data-testid="stElementContainer"],
[class*="st-key-xl_"] [data-testid="stElementContainer"] {
  width: 100% !important; max-width: 100%;
}
[class*="st-key-card_"] .stMarkdown, [class*="st-key-card_"] [data-testid="stMarkdownContainer"] {
  width: 100% !important; max-width: 100%;
}
[class*="st-key-card_"] { overflow: hidden; }

@keyframes tm-rise { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: none; } }
@keyframes tm-float { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(-18px, 16px) scale(1.07); } }
@media (prefers-reduced-motion: reduce) { * { animation: none !important; transition: none !important; } }
@media (max-width: 760px) {
  .tm-flow { grid-template-columns: 1fr; }
  .tm-hero { padding: 1.8rem 1.4rem; }
  .tm-hero h1 { font-size: 1.7rem; }
}
"""


def _html(markup):
    st.markdown(markup, unsafe_allow_html=True)


def inject_style():
    """Add the app's style sheet. Call once, right after set_page_config."""
    _html(f"<style>{_CSS}</style>")


def slug(name):
    """URL-safe form of a tool name, used for the ?tool= deep link."""
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


# --------------------------------------------------------------------------
# Navigation
# --------------------------------------------------------------------------
def go_to(tool):
    """Button callback: open a tool (or Home) and reflect it in the URL."""
    st.session_state["active_tool"] = tool
    if tool == HOME:
        st.query_params.pop("tool", None)
    else:
        st.query_params["tool"] = slug(tool)


def initial_tool(all_tools):
    """The tool named by ?tool=… in the URL, or Home."""
    wanted = st.query_params.get("tool", "")
    return next((t for t in all_tools if slug(t) == wanted), HOME)


def sidebar_nav(sections, meta, active, section_icons=None):
    """Brand block plus one expandable group of tool buttons per section.

    The section holding the open tool starts expanded and is highlighted;
    the others start collapsed so the four section names stand out.
    """
    section_icons = section_icons or {}
    with st.sidebar:
        _html('<div class="tm-brand"><div class="tm-brand-mark">📒</div>'
              '<div><div class="tm-brand-name">Tally AI</div>'
              '<div class="tm-brand-sub">Excel ⇄ Tally made simple</div></div></div>')
        st.button("🏠  Home", key="navA_home" if active == HOME else "nav_home",
                  on_click=go_to, args=(HOME,), use_container_width=True)
        for s, (section, tools) in enumerate(sections.items()):
            here = active in tools
            label = f"{section_icons.get(section, '📁')}  {section}"
            with st.container(key=f"{'navsecA' if here else 'navsec'}_{s}"):
                with st.expander(label, expanded=here):
                    for t, tool in enumerate(tools):
                        info = meta.get(tool, {})
                        st.button(
                            f"{info.get('icon', '•')}  {info.get('title', tool)}",
                            key=f"{'navA' if tool == active else 'nav'}_{s}_{t}",
                            on_click=go_to, args=(tool,),
                            use_container_width=True, help=info.get("desc"))
        _html('<div class="tm-side-foot">Your files are processed on this '
              'computer and are not stored.</div>')


# --------------------------------------------------------------------------
# Home screen
# --------------------------------------------------------------------------
def home(sections, meta, section_notes):
    """Hero, a three-step explainer and a searchable grid of tool cards."""
    _html('<div class="tm-hero"><span class="tm-hero-tag">Tally AI · Excel ⇄ Tally</span>'
          '<h1>What would you like to do today?</h1>'
          '<p>Pick a task below. Each one walks you through it step by step — '
          'upload your file, check it, and download the result.</p></div>')
    _html('<div class="tm-flow">'
          '<div class="tm-flow-item"><div class="tm-flow-num">1</div><div>'
          '<div class="tm-flow-title">Pick a task</div>'
          '<div class="tm-flow-sub">Choose a card below or from the menu</div></div></div>'
          '<div class="tm-flow-item"><div class="tm-flow-num">2</div><div>'
          '<div class="tm-flow-title">Upload your file</div>'
          '<div class="tm-flow-sub">Excel, XML, PDF, JSON or DBF</div></div></div>'
          '<div class="tm-flow-item"><div class="tm-flow-num">3</div><div>'
          '<div class="tm-flow-title">Download the result</div>'
          '<div class="tm-flow-sub">Ready for Tally or Excel</div></div></div></div>')

    query = st.text_input(
        "Search", key="home_search", label_visibility="collapsed",
        placeholder="🔎  Search — try “ledger”, “bank statement”, “GST”, “Marg”…")
    words = query.lower().split()

    def matches(tool):
        info = meta.get(tool, {})
        text = " ".join([tool, info.get("title", ""), info.get("desc", ""),
                         info.get("tags", "")]).lower()
        return all(w in text for w in words)

    shown = 0
    for s, (section, tools) in enumerate(sections.items()):
        tools = [t for t in tools if matches(t)]
        if not tools:
            continue
        shown += len(tools)
        _html(f'<div class="tm-section"><div class="tm-section-title">{html.escape(section)}</div>'
              f'<div class="tm-section-sub">{html.escape(section_notes.get(section, ""))}</div></div>')
        for row in range(0, len(tools), 3):
            cols = st.columns(3)
            for col, tool in zip(cols, tools[row:row + 3]):
                info = meta.get(tool, {})
                key = f"{s}_{slug(tool)}"
                with col, st.container(key=f"card_{key}"):
                    _html(f'<div class="tm-card-icon">{info.get("icon", "•")}</div>'
                          f'<div class="tm-card-title">{html.escape(info.get("title", tool))}</div>'
                          f'<div class="tm-card-desc">{html.escape(info.get("desc", ""))}</div>')
                    st.button("Open  →", key=f"open_{key}", on_click=go_to,
                              args=(tool,), use_container_width=True)
    if not shown:
        st.info("Nothing matches that search. Try a shorter word, such as "
                "“ledger” or “excel”.")


# --------------------------------------------------------------------------
# Page furniture
# --------------------------------------------------------------------------
def page_header(tool, meta, section=""):
    """Breadcrumb, icon tile, title and one-line description for a tool."""
    info = meta.get(tool, {})
    crumb = f'Home  ›  {html.escape(section)}  ›  ' if section else "Home  ›  "
    _html(f'<div class="tm-crumb">{crumb}<b>{html.escape(info.get("title", tool))}</b></div>'
          f'<div class="tm-head"><div class="tm-head-icon">{info.get("icon", "•")}</div><div>'
          f'<div class="tm-head-title">{html.escape(info.get("title", tool))}</div>'
          f'<div class="tm-head-sub">{html.escape(info.get("desc", ""))}</div></div></div>')


def step(text):
    """A step heading: 'Step 2a — Verify names' -> numbered badge + title."""
    m = re.match(r"\s*Step\s+(\w+)\s*[—–-]\s*(.*)", text)
    if not m:
        st.subheader(text)
        return
    _html(f'<div class="tm-step"><div class="tm-step-num">{html.escape(m.group(1))}</div>'
          f'<div class="tm-step-title">{html.escape(m.group(2))}</div></div>')


# --------------------------------------------------------------------------
# Excel-style preview
# --------------------------------------------------------------------------
_PREVIEW_ROWS = 500


@st.cache_data(show_spinner=False)
def _read_sheets(file_bytes):
    """{sheet name: (DataFrame of the first rows, total rows)} for a workbook."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, dtype=object)
    return {name: (df.head(_PREVIEW_ROWS), len(df)) for name, df in sheets.items()}


def excel_preview(file_bytes, name, key):
    """Show a workbook the way Excel does: title bar, grid, sheet tabs.

    Row numbers match Excel's (the header is row 1, so data starts at row 2).
    """
    try:
        sheets = _read_sheets(file_bytes)
    except Exception as exc:  # noqa: BLE001 - a preview must never break a page
        st.caption(f"Preview not available for this file ({exc}).")
        return
    if not sheets:
        return
    total = sum(rows for _, rows in sheets.values())
    with st.container(key=f"xl_{key}"):
        _html(f'<div class="tm-xl-bar"><div class="tm-xl-logo">X</div>'
              f'<div class="tm-xl-name">{html.escape(name)}</div>'
              f'<div class="tm-xl-meta">{len(sheets)} sheet{"s" if len(sheets) != 1 else ""}'
              f' · {total:,} row{"s" if total != 1 else ""}</div></div>')
        for tab, (sheet, (df, rows)) in zip(st.tabs(list(sheets)), sheets.items()):
            with tab:
                if df.empty:
                    st.caption("This sheet has no data rows.")
                    continue
                # Text, so blank cells show empty (not "None") and long
                # numbers such as pincodes appear exactly as typed.
                grid = df.fillna("").astype(str)
                grid.index = range(2, len(grid) + 2)
                st.dataframe(grid, use_container_width=True,
                             height=min(420, 38 + 35 * len(grid)))
                note = f"{rows:,} row{'s' if rows != 1 else ''} × {len(df.columns)} columns"
                if rows > len(df):
                    note += f" — showing the first {len(df):,}"
                st.caption(note)
