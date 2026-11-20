"""
Look and feel for the TALLY ML app.

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
  --tm-ink: #0f172a;
  --tm-muted: #64748b;
  --tm-line: #e5e8f2;
  --tm-card: #ffffff;
  --tm-accent: #6366f1;
  --tm-accent-2: #8b5cf6;
  --tm-accent-soft: #eef0ff;
  --tm-excel: #107c41;
  --tm-radius: 16px;
  --tm-shadow: 0 1px 2px rgba(15, 23, 42, .04), 0 8px 24px rgba(15, 23, 42, .06);
  --tm-shadow-lift: 0 2px 4px rgba(15, 23, 42, .06), 0 18px 40px rgba(79, 70, 229, .16);
}

html, body, .stApp, [data-testid="stAppViewContainer"],
[data-testid="stMarkdownContainer"], button, input, textarea, select {
  font-family: 'Inter', -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}

/* ---- Page canvas ---------------------------------------------------- */
[data-testid="stAppViewContainer"] {
  background:
    radial-gradient(900px 420px at 88% -8%, rgba(139, 92, 246, .13), transparent 62%),
    radial-gradient(760px 380px at -6% 4%, rgba(99, 102, 241, .11), transparent 60%),
    #f6f7fb;
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stMainBlockContainer"], .main .block-container {
  max-width: 1120px;
  padding-top: 2.4rem;
  padding-bottom: 5rem;
}
h1, h2, h3 { color: var(--tm-ink); letter-spacing: -.02em; }
h3 { font-weight: 700; }

/* ---- Sidebar -------------------------------------------------------- */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #0b1020 0%, #151a3a 55%, #1e1b4b 100%);
  border-right: 0;
}
[data-testid="stSidebar"] * { color: #cbd5e1; }
[data-testid="stSidebar"] [data-testid="stSidebarContent"] { padding-top: .5rem; }
.tm-brand { display: flex; align-items: center; gap: .7rem; padding: .2rem .2rem 1rem; }
.tm-brand-mark {
  width: 40px; height: 40px; border-radius: 12px; display: grid; place-items: center;
  font-size: 1.25rem;
  background: linear-gradient(135deg, var(--tm-accent), var(--tm-accent-2));
  box-shadow: 0 8px 20px rgba(99, 102, 241, .45);
}
.tm-brand-name { font-weight: 800; font-size: 1.08rem; color: #fff !important; letter-spacing: -.01em; }
.tm-brand-sub { font-size: .74rem; color: #94a3b8 !important; margin-top: -2px; }
.tm-nav-section {
  font-size: .68rem; font-weight: 700; letter-spacing: .09em; text-transform: uppercase;
  color: #7c86a8 !important; padding: 1rem .55rem .3rem;
}
[data-testid="stSidebar"] .stButton { margin-bottom: -.72rem; }
[data-testid="stSidebar"] .stButton button {
  background: transparent; border: 0; box-shadow: none; border-radius: 10px;
  justify-content: flex-start; text-align: left; padding: .42rem .6rem; min-height: 0;
  color: #cbd5e1; font-weight: 500; font-size: .88rem;
  transition: background .18s ease, color .18s ease, transform .18s ease;
}
[data-testid="stSidebar"] .stButton button > div { justify-content: flex-start; width: 100%; }
[data-testid="stSidebar"] .stButton button p { font-size: .88rem; text-align: left; }
[data-testid="stSidebar"] .stButton button:hover {
  background: rgba(255, 255, 255, .08); color: #fff; transform: translateX(2px);
}
[data-testid="stSidebar"] .stButton button:hover * { color: #fff; }
[data-testid="stSidebar"] [class*="st-key-navA_"] button {
  background: linear-gradient(90deg, rgba(99, 102, 241, .95), rgba(139, 92, 246, .9));
  box-shadow: 0 6px 16px rgba(99, 102, 241, .35);
}
[data-testid="stSidebar"] [class*="st-key-navA_"] button * { color: #fff; font-weight: 600; }
.tm-side-foot { font-size: .72rem; color: #64748b !important; padding: 1.6rem .55rem .5rem; line-height: 1.5; }

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
  transform: translateY(-1px); border-color: var(--tm-accent);
  box-shadow: 0 8px 20px rgba(99, 102, 241, .16);
}
[data-testid="stMain"] .stButton button:active,
[data-testid="stMain"] .stDownloadButton button:active { transform: translateY(0); }
[data-testid="stMain"] button[kind="primary"],
[data-testid="stMain"] button[data-testid$="-primary"] {
  background: linear-gradient(135deg, var(--tm-accent), var(--tm-accent-2));
  border: 0; color: #fff; box-shadow: 0 8px 20px rgba(99, 102, 241, .28);
}
[data-testid="stMain"] button[kind="primary"]:hover,
[data-testid="stMain"] button[data-testid$="-primary"]:hover {
  box-shadow: 0 12px 28px rgba(99, 102, 241, .38); filter: brightness(1.05);
}
[data-testid="stMain"] .stDownloadButton button {
  background: #fff; border-color: #cfd5ee;
}

/* ---- Inputs, uploads, tables ---------------------------------------- */
[data-testid="stFileUploaderDropzone"] {
  border: 2px dashed #c7cdf0; border-radius: var(--tm-radius);
  background: linear-gradient(180deg, #ffffff, #f8f9ff);
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
  background: linear-gradient(90deg, var(--tm-accent), var(--tm-accent-2));
}
hr { border-color: var(--tm-line) !important; margin: 1.6rem 0 !important; }

/* ---- Tabs (selectors cover old and new Streamlit markup) ------------ */
[data-testid="stTabs"] [role="tablist"] { gap: .4rem; border-bottom: 0 !important; box-shadow: none !important; padding: .15rem 0 .5rem; }
[data-testid="stTabs"] [role="tab"] {
  border-radius: 999px !important; padding: .35rem .95rem !important; height: auto !important; background: #fff;
  border: 1px solid var(--tm-line) !important; transition: border-color .16s ease, background .16s ease;
  margin: 0 !important;
}
[data-testid="stTabs"] [role="tab"]:hover { border-color: var(--tm-accent) !important; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
  background: var(--tm-accent-soft); border-color: var(--tm-accent) !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p { color: #4338ca; font-weight: 600; }
[data-testid="stTabs"] [role="tab"]::before, [data-testid="stTabs"] [role="tab"]::after,
[data-testid="stTabs"] [role="tablist"]::before, [data-testid="stTabs"] [role="tablist"]::after { display: none !important; }
[data-testid="stTabs"] [data-baseweb="tab-highlight"],
[data-testid="stTabs"] [data-baseweb="tab-border"] { display: none !important; }

/* ---- Hero (home) ---------------------------------------------------- */
.tm-hero {
  position: relative; overflow: hidden; border-radius: 26px; padding: 2.6rem 2.6rem 2.4rem;
  background: linear-gradient(125deg, #312e81 0%, #4f46e5 46%, #8b5cf6 100%);
  box-shadow: 0 24px 60px rgba(79, 70, 229, .28);
  animation: tm-rise .55s cubic-bezier(.2, .7, .2, 1) both;
}
.tm-hero::before, .tm-hero::after {
  content: ""; position: absolute; border-radius: 50%; filter: blur(8px); opacity: .5;
  animation: tm-float 11s ease-in-out infinite;
}
.tm-hero::before { width: 320px; height: 320px; right: -70px; top: -120px; background: rgba(255, 255, 255, .16); }
.tm-hero::after {
  width: 230px; height: 230px; right: 190px; bottom: -150px; background: rgba(196, 181, 253, .32);
  animation-delay: -5s;
}
.tm-hero-tag {
  display: inline-block; font-size: .74rem; font-weight: 600; letter-spacing: .04em;
  color: #e0e7ff; background: rgba(255, 255, 255, .14); border: 1px solid rgba(255, 255, 255, .22);
  padding: .26rem .7rem; border-radius: 999px; backdrop-filter: blur(6px);
}
.tm-hero h1 {
  position: relative; color: #fff; font-size: 2.35rem; font-weight: 800; line-height: 1.12;
  margin: .9rem 0 .55rem; padding: 0; letter-spacing: -.03em;
}
.tm-hero p { position: relative; color: #dfe3ff; font-size: 1.04rem; max-width: 620px; margin: 0; }
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
  font-weight: 700; color: #4338ca; background: var(--tm-accent-soft);
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
  transform: translateY(-4px); box-shadow: var(--tm-shadow-lift); border-color: #c7cbfb !important;
}
.tm-card-icon {
  width: 46px; height: 46px; border-radius: 14px; display: grid; place-items: center; font-size: 1.4rem;
  background: linear-gradient(135deg, #eef0ff, #f5edff); margin-bottom: .7rem;
  transition: transform .25s ease;
}
[class*="st-key-card_"]:hover .tm-card-icon { transform: scale(1.08) rotate(-4deg); }
.tm-card-title { font-weight: 700; color: var(--tm-ink); font-size: 1rem; line-height: 1.3; }
.tm-card-desc { color: var(--tm-muted); font-size: .86rem; line-height: 1.45; margin-top: .3rem; min-height: 4.4em; }
.tm-card-title { min-height: 2.6em; display: flex; align-items: flex-end; }
[class*="st-key-card_"] .stButton button {
  background: var(--tm-accent-soft); border: 0; color: #4338ca; font-weight: 600;
}
[class*="st-key-card_"] .stButton button:hover {
  background: linear-gradient(135deg, var(--tm-accent), var(--tm-accent-2)); color: #fff;
}
[class*="st-key-card_"] .stButton button:hover * { color: #fff; }

/* ---- Page header and steps ------------------------------------------ */
.tm-crumb { font-size: .8rem; color: var(--tm-muted); font-weight: 500; margin-bottom: .7rem; }
.tm-crumb b { color: #4338ca; font-weight: 600; }
.tm-head {
  display: flex; align-items: center; gap: 1rem; margin-bottom: 1.1rem;
  animation: tm-rise .45s cubic-bezier(.2, .7, .2, 1) both;
}
.tm-head-icon {
  flex: none; width: 60px; height: 60px; border-radius: 18px; display: grid; place-items: center;
  font-size: 1.8rem; background: #fff; border: 1px solid #dfe2fb;
  box-shadow: 0 12px 28px rgba(99, 102, 241, .22);
}
.tm-head-title { font-size: 1.85rem; font-weight: 800; color: var(--tm-ink); letter-spacing: -.03em; line-height: 1.15; }
.tm-head-sub { color: var(--tm-muted); font-size: 1rem; margin-top: .2rem; }
.tm-step { display: flex; align-items: center; gap: .75rem; margin: .5rem 0 .7rem; }
.tm-step-num {
  flex: none; min-width: 34px; height: 34px; padding: 0 .55rem; border-radius: 11px; display: grid;
  place-items: center; font-weight: 700; font-size: .95rem; color: #fff;
  background: linear-gradient(135deg, var(--tm-accent), var(--tm-accent-2));
  box-shadow: 0 6px 14px rgba(99, 102, 241, .3);
}
.tm-step-title { font-size: 1.2rem; font-weight: 700; color: var(--tm-ink); letter-spacing: -.02em; }

/* ---- Excel-style preview -------------------------------------------- */
[class*="st-key-xl_"] {
  border: 1px solid #cfe3d7 !important; border-radius: 14px !important; overflow: hidden;
  background: #fff; box-shadow: var(--tm-shadow); padding: 0 0 .5rem; gap: .5rem;
}
.tm-xl-bar {
  display: flex; align-items: center; gap: .65rem; padding: .6rem .9rem;
  background: linear-gradient(90deg, #0e6b38, var(--tm-excel)); color: #fff;
}
.tm-xl-logo {
  width: 26px; height: 26px; border-radius: 6px; background: #fff; color: var(--tm-excel);
  display: grid; place-items: center; font-weight: 800; font-size: .9rem;
}
.tm-xl-name { font-weight: 600; font-size: .92rem; color: #fff; }
.tm-xl-meta { margin-left: auto; font-size: .78rem; color: #d7f0e1; }
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
  padding: 1rem 1.15rem; box-shadow: var(--tm-shadow); border-left: 5px solid #22c55e !important;
  transition: transform .18s ease, box-shadow .18s ease;
}
[class*="st-key-logerr_"] { border-left-color: #ef4444 !important; }
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


def sidebar_nav(sections, meta, active):
    """Brand block plus one button per tool, grouped by section."""
    with st.sidebar:
        _html('<div class="tm-brand"><div class="tm-brand-mark">📒</div>'
              '<div><div class="tm-brand-name">TALLY ML</div>'
              '<div class="tm-brand-sub">Excel ⇄ Tally made simple</div></div></div>')
        st.button("🏠  Home", key="navA_home" if active == HOME else "nav_home",
                  on_click=go_to, args=(HOME,), use_container_width=True)
        for s, (section, tools) in enumerate(sections.items()):
            _html(f'<div class="tm-nav-section">{html.escape(section)}</div>')
            for t, tool in enumerate(tools):
                info = meta.get(tool, {})
                st.button(
                    f"{info.get('icon', '•')}  {info.get('title', tool)}",
                    key=f"{'navA' if tool == active else 'nav'}_{s}_{t}",
                    on_click=go_to, args=(tool,), use_container_width=True,
                    help=info.get("desc"))
        _html('<div class="tm-side-foot">Your files are processed on this '
              'computer and are not stored.</div>')


# --------------------------------------------------------------------------
# Home screen
# --------------------------------------------------------------------------
def home(sections, meta, section_notes):
    """Hero, a three-step explainer and a searchable grid of tool cards."""
    _html('<div class="tm-hero"><span class="tm-hero-tag">TALLY ML · Excel ⇄ Tally</span>'
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
