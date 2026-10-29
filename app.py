# ═════════════════════════════════════════════════════════════════════════════
#  M.E. COMMAND CENTER  v4.0
#  Marine diesel performance intelligence  ·  TEC-005 Performance Data Sheets
#
#  • Pure Python  : olefile + pandas + plotly + pydantic  (no OS packages, no external APIs)
#  • Input        : legacy Word 97-2003 binary .doc, drag & drop - no conversion needed
#  • Parser       : geometry-based - reads the Word table structure (piece table + PAPX cell
#                   boundaries) and matches header cells to data cells by x-position, so
#                   merged cells, shifted columns, typed-in header cells and blank rows
#                   cannot corrupt the extraction.
#  • Trust model  : every value carries provenance, is cross-checked against the sheet's own
#                   "Average" cells, passes a Pydantic data-integrity shield and is confirmed
#                   by a human before analysis.
# ═════════════════════════════════════════════════════════════════════════════

import bisect
import hashlib
import html as _html
import inspect
import io
import json
import math
import re
import struct
import time
from datetime import date as _date
from io import BytesIO
from typing import Optional

import olefile
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from pydantic import BaseModel, Field, ValidationError, model_validator

APP_VERSION = "4.0"

st.set_page_config(page_title="M.E. Command Center", page_icon="⚙️", layout="wide",
                   initial_sidebar_state="expanded")


# ── small helpers ────────────────────────────────────────────────────────────
def esc(x) -> str:
    return _html.escape("" if x is None else str(x), quote=True)


def H(s: str) -> str:
    """Flatten an HTML snippet: no blank lines / indentation, so Markdown never treats it as code."""
    return "\n".join(l.strip() for l in s.strip().splitlines() if l.strip())


def md(s: str) -> None:
    st.markdown(H(s), unsafe_allow_html=True)


def stretch(fn) -> dict:
    """Version-proof 'full width' kwargs (Streamlit renamed use_container_width -> width='stretch')."""
    try:
        params = inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return {}
    if "width" in params:
        return {"width": "stretch"}
    if "use_container_width" in params:
        return {"use_container_width": True}
    return {}


def fnum(v, nd=1, dash="—") -> str:
    if v is None or (isinstance(v, float) and v != v):
        return dash
    return f"{v:,.{nd}f}"


def fdate(d) -> str:
    return d.strftime("%d %b %Y") if isinstance(d, _date) else "—"


# ═════════════════════════════════════════════════════════════════════════════
#  1. DESIGN SYSTEM
# ═════════════════════════════════════════════════════════════════════════════
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Condensed:wght@500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
:root{
  --bg:#09090D; --bg2:#0D0D13; --card:#121219; --card2:#181822; --line:rgba(255,255,255,.07);
  --gold-line:rgba(196,163,90,.30); --text:#ECE8DD; --muted:#8E8B80; --faint:#5C5A52;
  --gold:#C4A35A; --gold2:#E2C888; --gold-soft:rgba(196,163,90,.10);
  --ok:#5BB98C; --warn:#E5B54A; --orange:#EA8B4A; --crit:#E8645E; --info:#7BAFD9;
  --ok-s:rgba(91,185,140,.10); --warn-s:rgba(229,181,74,.10); --orange-s:rgba(234,139,74,.10);
  --crit-s:rgba(232,100,94,.10); --info-s:rgba(123,175,217,.10);
  --disp:'Barlow Condensed','Arial Narrow',sans-serif; --body:'Barlow',system-ui,sans-serif;
  --mono:'JetBrains Mono',ui-monospace,monospace;
}
html, body, .stApp, [data-testid="stAppViewContainer"]{ background:var(--bg) !important; color:var(--text); font-family:var(--body); }
.stApp{ background:radial-gradient(1200px 520px at 78% -8%, rgba(196,163,90,.07), transparent 60%), var(--bg) !important; }
[data-testid="stHeader"]{ background:transparent !important; height:0 !important; }
[data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer, .stDeployButton{ display:none !important; visibility:hidden !important; }
.block-container{ max-width:1320px !important; padding:1.3rem 2.2rem 4rem !important; }
.stMarkdown p, .stMarkdown li, label, .stCaption, [data-testid="stCaptionContainer"]{ color:var(--text); font-family:var(--body); }
[data-testid="stCaptionContainer"] p{ color:var(--muted) !important; font-size:.82rem; }
h1,h2,h3,h4{ font-family:var(--disp) !important; color:var(--text) !important; letter-spacing:.04em; text-transform:uppercase; font-weight:600 !important; }
hr{ border-color:var(--line) !important; margin:1.1rem 0 !important; }
::-webkit-scrollbar{ width:9px; height:9px } ::-webkit-scrollbar-thumb{ background:#2A2A33; border-radius:8px } ::-webkit-scrollbar-track{ background:transparent }

/* sidebar */
[data-testid="stSidebar"]{ background:var(--bg2) !important; border-right:1px solid var(--line) !important; }
[data-testid="stSidebar"] .block-container, [data-testid="stSidebarContent"]{ padding-top:1rem !important; }
[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] *{ font-family:"Material Symbols Rounded","Material Symbols Outlined","Material Icons" !important; letter-spacing:normal !important; text-transform:none !important; }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{ color:var(--muted) !important; font-size:.72rem !important; letter-spacing:.14em; text-transform:uppercase; font-family:var(--disp); font-weight:600; }

/* widgets */
[data-testid="stWidgetLabel"] p{ color:var(--muted) !important; font-size:.78rem !important; }
[data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="select"] > div{ background:#0B0B10 !important; border-color:var(--line) !important; border-radius:9px !important; }
[data-baseweb="input"]:focus-within, [data-baseweb="base-input"]:focus-within{ border-color:var(--gold) !important; box-shadow:0 0 0 1px var(--gold-line) !important; }
input, textarea{ color:var(--text) !important; font-family:var(--mono) !important; font-size:.92rem !important; -webkit-text-fill-color:var(--text) !important; }
input::placeholder{ color:var(--faint) !important; -webkit-text-fill-color:var(--faint) !important; }
[data-testid="stNumberInputContainer"]{ background:#0B0B10 !important; border:1px solid var(--line) !important; border-radius:9px !important; overflow:hidden; }
[data-testid="stNumberInputContainer"]:focus-within{ border-color:var(--gold) !important; box-shadow:0 0 0 1px var(--gold-line) !important; }
[data-testid="stNumberInput"] button{ background:#14141B !important; color:var(--muted) !important; border:none !important; }
[data-testid="stNumberInput"] button:hover{ color:var(--gold) !important; background:#1B1B24 !important; }
[data-baseweb="popover"] > div, [data-baseweb="menu"], ul[role="listbox"]{ background:#14141B !important; color:var(--text) !important; border:1px solid var(--line) !important; }
li[role="option"]{ color:var(--text) !important; } li[role="option"]:hover{ background:var(--gold-soft) !important; }
[data-testid="stRadio"] label p{ color:var(--text) !important; font-size:.88rem !important; }
[data-testid="stCheckbox"] label p, [data-testid="stToggle"] label p{ color:var(--text) !important; font-size:.86rem !important; }
[data-baseweb="slider"] div[role="slider"]{ background:var(--gold) !important; }

.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button{
  background:linear-gradient(180deg,#D3B46B,#B08F44) !important; color:#14110A !important; border:1px solid #E2C888 !important;
  border-radius:10px !important; font-family:var(--disp) !important; font-weight:700 !important; letter-spacing:.13em; text-transform:uppercase;
  font-size:.95rem !important; padding:.55rem 1.3rem !important; box-shadow:0 6px 22px rgba(196,163,90,.16) !important; transition:all .18s ease; }
.stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover{ transform:translateY(-1px); filter:brightness(1.06); box-shadow:0 10px 28px rgba(196,163,90,.26) !important; }
.stButton > button p, .stDownloadButton > button p, [data-testid="stFormSubmitButton"] > button p{ color:#14110A !important; font-family:var(--disp) !important; font-weight:700 !important; letter-spacing:.13em; }
.stButton > button[kind="secondary"], .stDownloadButton > button[kind="secondary"]{ }

[data-testid="stForm"]{ background:linear-gradient(180deg,var(--card),var(--bg2)) !important; border:1px solid var(--line) !important; border-radius:16px !important; padding:1.3rem 1.4rem !important; }
[data-testid="stExpander"]{ background:var(--card) !important; border:1px solid var(--line) !important; border-radius:14px !important; overflow:hidden; }
[data-testid="stExpander"] summary{ background:transparent !important; }
[data-testid="stExpander"] summary p{ font-family:var(--disp); letter-spacing:.1em; text-transform:uppercase; font-weight:600; color:var(--text) !important; font-size:.95rem; }
[data-testid="stExpander"] summary:hover p{ color:var(--gold2) !important; }
[data-testid="stExpander"] svg{ color:var(--muted) !important; }

[data-baseweb="tab-list"]{ gap:6px !important; background:var(--bg2) !important; padding:5px !important; border:1px solid var(--line); border-radius:12px; width:fit-content; }
button[data-baseweb="tab"]{ background:transparent !important; border-radius:8px !important; padding:.42rem 1.15rem !important; height:auto !important; }
button[data-baseweb="tab"] p{ color:var(--muted) !important; font-family:var(--disp); font-weight:600; letter-spacing:.14em; text-transform:uppercase; font-size:.9rem !important; }
button[data-baseweb="tab"][aria-selected="true"]{ background:var(--gold-soft) !important; box-shadow:inset 0 0 0 1px var(--gold-line); }
button[data-baseweb="tab"][aria-selected="true"] p{ color:var(--gold2) !important; }
[data-baseweb="tab-highlight"], [data-baseweb="tab-border"], .react-aria-SelectionIndicator{ display:none !important; }
[role="tablist"]{ gap:6px !important; background:var(--bg2) !important; padding:5px !important; border:1px solid var(--line) !important; border-radius:12px !important; width:fit-content; box-shadow:none !important; }
[data-testid="stTab"]{ background:transparent !important; border-radius:8px !important; padding:.42rem 1.15rem !important; border:none !important; }
[data-testid="stTab"] p{ color:var(--muted) !important; font-family:var(--disp) !important; font-weight:600; letter-spacing:.14em; text-transform:uppercase; font-size:.9rem !important; }
[data-testid="stTab"][aria-selected="true"], [data-testid="stTab"][data-selected="true"]{ background:var(--gold-soft) !important; box-shadow:inset 0 0 0 1px var(--gold-line) !important; }
[data-testid="stTab"][aria-selected="true"] p, [data-testid="stTab"][data-selected="true"] p{ color:var(--gold2) !important; }

[data-testid="stFileUploader"] section, [data-testid="stFileUploaderDropzone"]{ background:linear-gradient(180deg,#101017,#0B0B10) !important; border:1.5px dashed var(--gold-line) !important; border-radius:16px !important; padding:1.4rem !important; }
[data-testid="stFileUploaderDropzone"]:hover{ border-color:var(--gold) !important; background:#12121A !important; }
[data-testid="stFileUploaderDropzone"] *{ color:var(--muted) !important; }
[data-testid="stFileUploaderDropzone"] button{ background:transparent !important; color:var(--gold2) !important; border:1px solid var(--gold-line) !important; box-shadow:none !important; font-size:.8rem !important; letter-spacing:.1em; }
[data-testid="stFileChip"]{ background:#14141B !important; border:1px solid var(--gold-line) !important; border-radius:10px !important; }
[data-testid="stFileChipName"]{ color:var(--text) !important; font-family:var(--mono) !important; font-size:.8rem !important; }
[data-testid="stFileChip"] *{ color:var(--muted) !important; }
[data-testid="stFileChipDeleteBtn"] button{ background:transparent !important; box-shadow:none !important; border:none !important; }
[data-testid="stAlert"]{ background:var(--card) !important; border:1px solid var(--line) !important; border-radius:12px !important; color:var(--text) !important; }
[data-testid="stAlert"] *{ color:var(--text) !important; }
[data-testid="stSpinner"] *{ color:var(--gold2) !important; }
[data-testid="stPlotlyChart"]{ border-radius:14px; }

/* components */
.hdr{ display:flex; align-items:center; justify-content:space-between; gap:16px; padding:.2rem 0 1.2rem; border-bottom:1px solid var(--line); margin-bottom:1.3rem; flex-wrap:wrap; }
.hdr-l{ display:flex; align-items:center; gap:16px; }
.logo{ width:50px; height:50px; border-radius:13px; display:flex; align-items:center; justify-content:center; font-family:var(--disp); font-weight:700; font-size:1.55rem; letter-spacing:.03em; color:#16120A;
  background:linear-gradient(145deg,#E7CD8D,#B3913F); box-shadow:0 8px 30px rgba(196,163,90,.28), inset 0 1px 0 rgba(255,255,255,.35); }
.brand{ font-family:var(--disp); font-size:1.85rem; font-weight:700; letter-spacing:.11em; line-height:1; color:var(--text); }
.tag{ font-size:.72rem; letter-spacing:.2em; text-transform:uppercase; color:var(--muted); margin-top:5px; font-weight:500; }
.chip{ display:inline-flex; align-items:center; gap:6px; padding:3px 11px; border-radius:999px; font-family:var(--disp); font-weight:600; letter-spacing:.13em; text-transform:uppercase; font-size:.74rem;
  border:1px solid var(--line); color:var(--muted); background:rgba(255,255,255,.02); white-space:nowrap; }
.chip.gold{ color:var(--gold2); border-color:var(--gold-line); background:var(--gold-soft); }
.chip.ok{ color:var(--ok); border-color:rgba(91,185,140,.35); background:var(--ok-s); }
.chip.warn{ color:var(--warn); border-color:rgba(229,181,74,.35); background:var(--warn-s); }
.chip.orange{ color:var(--orange); border-color:rgba(234,139,74,.35); background:var(--orange-s); }
.chip.crit{ color:var(--crit); border-color:rgba(232,100,94,.38); background:var(--crit-s); }
.chip.info{ color:var(--info); border-color:rgba(123,175,217,.35); background:var(--info-s); }

.steps{ display:flex; gap:10px; margin:.2rem 0 1.2rem; flex-wrap:wrap; }
.step{ display:flex; align-items:center; gap:10px; padding:.55rem 1rem; border-radius:12px; border:1px solid var(--line); background:var(--card); flex:1; min-width:180px; }
.step .n{ width:26px; height:26px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-family:var(--mono); font-size:.78rem; border:1px solid var(--line); color:var(--muted); }
.step .t{ font-family:var(--disp); letter-spacing:.14em; text-transform:uppercase; font-size:.86rem; font-weight:600; color:var(--muted); }
.step.on{ border-color:var(--gold-line); background:var(--gold-soft); } .step.on .n{ background:var(--gold); color:#14110A; border-color:var(--gold); } .step.on .t{ color:var(--gold2); }
.step.done .n{ background:var(--ok-s); color:var(--ok); border-color:rgba(91,185,140,.4); } .step.done .t{ color:var(--text); }

.card{ background:linear-gradient(180deg,var(--card),var(--bg2)); border:1px solid var(--line); border-radius:16px; padding:1.15rem 1.3rem; }
.card.gold{ border-color:var(--gold-line); box-shadow:0 14px 50px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.03); }
.sec{ display:flex; align-items:baseline; gap:12px; margin:1.7rem 0 .75rem; }
.sec .no{ font-family:var(--mono); color:var(--gold); font-size:.8rem; letter-spacing:.1em; }
.sec .ti{ font-family:var(--disp); font-size:1.38rem; letter-spacing:.14em; text-transform:uppercase; font-weight:600; color:var(--text); }
.sec .su{ color:var(--muted); font-size:.82rem; margin-left:auto; text-align:right; }
.lbl{ font-family:var(--disp); font-size:.74rem; letter-spacing:.15em; text-transform:uppercase; color:var(--muted); font-weight:600; margin:0; }
.val{ font-family:var(--body); font-size:.98rem; color:var(--text); margin:2px 0 10px; font-weight:500; word-break:break-word; }
.mono{ font-family:var(--mono); }

.vs{ display:grid; grid-template-columns:1.25fr 1fr; gap:22px; align-items:center; }
@media(max-width:900px){ .vs{ grid-template-columns:1fr; } }
.vname{ font-family:var(--disp); font-size:2.45rem; font-weight:700; letter-spacing:.06em; line-height:1.02; color:var(--text); text-transform:uppercase; }
.vsub{ color:var(--muted); font-size:.9rem; margin:6px 0 14px; }
.facts{ display:grid; grid-template-columns:repeat(auto-fill,minmax(172px,1fr)); gap:2px 18px; }
.ring-wrap{ display:flex; gap:18px; align-items:center; justify-content:flex-end; flex-wrap:wrap; }
.bars{ min-width:230px; flex:1; }
.bar-row{ margin:7px 0; } .bar-row .bl{ display:flex; justify-content:space-between; font-size:.73rem; color:var(--muted); letter-spacing:.06em; margin-bottom:3px; }
.bar-row .bl b{ color:var(--text); font-family:var(--mono); font-weight:500; }
.track{ height:5px; border-radius:6px; background:rgba(255,255,255,.07); overflow:hidden; } .fill{ height:100%; border-radius:6px; background:linear-gradient(90deg,#8F7336,var(--gold2)); }

.kpis{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:14px; }
.kpi{ background:linear-gradient(180deg,var(--card),var(--bg2)); border:1px solid var(--line); border-radius:15px; padding:1rem 1.15rem; position:relative; overflow:hidden; }
.kpi::before{ content:''; position:absolute; left:0; top:0; bottom:0; width:3px; background:var(--gold); opacity:.8; }
.kpi.ok::before{ background:var(--ok); } .kpi.warn::before{ background:var(--warn); } .kpi.crit::before{ background:var(--crit); } .kpi.orange::before{ background:var(--orange); }
.kpi .k{ font-family:var(--disp); font-size:.74rem; letter-spacing:.2em; text-transform:uppercase; color:var(--muted); font-weight:600; }
.kpi .v{ font-family:var(--mono); font-size:1.85rem; font-weight:500; color:var(--text); line-height:1.15; margin-top:4px; }
.kpi .v small{ font-size:.85rem; color:var(--muted); margin-left:4px; }
.kpi .s{ font-size:.76rem; color:var(--muted); margin-top:2px; }

.banner{ display:flex; gap:14px; align-items:flex-start; padding:.95rem 1.2rem; border-radius:14px; border:1px solid var(--line); margin:.9rem 0; }
.banner .ic{ font-size:1.25rem; line-height:1.25; } .banner .bt{ font-family:var(--disp); letter-spacing:.12em; text-transform:uppercase; font-weight:700; font-size:1rem; } .banner .bd{ color:var(--muted); font-size:.88rem; margin-top:2px; }
.banner.ok{ background:var(--ok-s); border-color:rgba(91,185,140,.32); } .banner.ok .bt{ color:var(--ok); }
.banner.high{ background:var(--info-s); border-color:rgba(123,175,217,.30); } .banner.high .bt{ color:var(--info); }
.banner.review{ background:var(--warn-s); border-color:rgba(229,181,74,.34); } .banner.review .bt{ color:var(--warn); }
.banner.low, .banner.err{ background:var(--crit-s); border-color:rgba(232,100,94,.36); } .banner.low .bt, .banner.err .bt{ color:var(--crit); }
.banner.blank{ background:var(--gold-soft); border-color:var(--gold-line); } .banner.blank .bt{ color:var(--gold2); }

.fcard{ border-radius:14px; padding:1rem 1.2rem; margin-bottom:.7rem; border:1px solid var(--line); border-left:4px solid var(--faint); background:var(--card); }
.fcard.critical{ border-left-color:var(--crit); background:linear-gradient(90deg,var(--crit-s),var(--card) 40%); }
.fcard.warning{ border-left-color:var(--warn); background:linear-gradient(90deg,var(--warn-s),var(--card) 40%); }
.fcard.mechanical{ border-left-color:var(--orange); background:linear-gradient(90deg,var(--orange-s),var(--card) 40%); }
.fcard.advisory{ border-left-color:var(--info); background:linear-gradient(90deg,var(--info-s),var(--card) 40%); }
.fcard.nominal{ border-left-color:var(--ok); background:linear-gradient(90deg,var(--ok-s),var(--card) 40%); }
.fcard .ft{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:5px; }
.fcard .fh{ font-family:var(--disp); font-size:1.08rem; font-weight:700; letter-spacing:.09em; text-transform:uppercase; }
.fcard .ff{ color:var(--text); font-size:.9rem; margin:2px 0 6px; }
.fcard .fa{ color:var(--muted); font-size:.84rem; border-top:1px dashed var(--line); padding-top:6px; }
.fcard .fa b{ color:var(--gold2); font-family:var(--disp); letter-spacing:.14em; font-size:.74rem; text-transform:uppercase; margin-right:6px; }

.tbl{ width:100%; border-collapse:separate; border-spacing:0; font-size:.88rem; }
.tbl th{ font-family:var(--disp); font-size:.76rem; letter-spacing:.1em; text-transform:uppercase; color:var(--muted); text-align:right; padding:.55rem .6rem; border-bottom:1px solid var(--gold-line); font-weight:600; white-space:nowrap; }
.tbl th:first-child, .tbl td:first-child{ text-align:left; }
.tbl td{ padding:.55rem .6rem; text-align:right; border-bottom:1px solid var(--line); color:var(--text); font-family:var(--mono); font-size:.86rem; white-space:nowrap; }
.tbl tr:hover td{ background:rgba(196,163,90,.04); }
.tbl td.t{ font-family:var(--body); text-align:left; }
.tbl .dim{ color:var(--faint); }
.tbl-wrap{ overflow-x:auto; border:1px solid var(--line); border-radius:14px; background:var(--card); padding:.2rem .4rem .4rem; }
.cyl-chip{ display:inline-flex; align-items:center; justify-content:center; min-width:34px; height:34px; border-radius:9px; font-family:var(--mono); font-weight:600; font-size:.95rem; color:var(--gold2); background:var(--gold-soft); border:1px solid var(--gold-line); }
.fh-row{ font-family:var(--disp); font-size:.74rem; letter-spacing:.17em; text-transform:uppercase; color:var(--muted); font-weight:600; padding-bottom:2px; }
.note{ color:var(--muted); font-size:.82rem; }
.legend{ display:flex; gap:16px; flex-wrap:wrap; margin:.3rem 0 .9rem; color:var(--muted); font-size:.78rem; }
.dot{ display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:6px; vertical-align:middle; }
.empty{ text-align:center; padding:2.6rem 1rem 1.2rem; }
.empty .big{ font-family:var(--disp); font-size:2.3rem; letter-spacing:.12em; text-transform:uppercase; font-weight:700; color:var(--text); }
.empty .sm{ color:var(--muted); max-width:640px; margin:.5rem auto 0; font-size:.98rem; }
.feat{ display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:14px; margin-top:1.6rem; }
.feat .card .fi{ font-size:1.3rem; } .feat .card .fh{ font-family:var(--disp); letter-spacing:.14em; text-transform:uppercase; font-size:1rem; font-weight:700; margin:.3rem 0 .2rem; color:var(--gold2); } .feat .card .fd{ color:var(--muted); font-size:.86rem; }
"""


def inject_css() -> None:
    st.markdown("<style>" + CSS + "</style>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
#  2. PARSER  -  Word 97-2003 binary table reader + geometry-based TEC-005 extraction
# ═════════════════════════════════════════════════════════════════════════════
# ═══ BEGIN PARSER ═══
class ParseError(Exception):
    """Raised when the file cannot be read as a Word 97-2003 binary document."""


# ─────────────────────────────────────────────────────────────────────────────
# 1. LOW-LEVEL WORD (.doc) READER  —  piece table + PAPX table geometry
# ─────────────────────────────────────────────────────────────────────────────
_SPRA = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}


def _walk_sprms(grpprl):
    i, n = 0, len(grpprl)
    while i + 2 <= n:
        op = struct.unpack_from("<H", grpprl, i)[0]
        i += 2
        spra = op >> 13
        if spra == 6:
            if i >= n:
                break
            if op in (0xD608, 0xC615):                     # sprmTDefTable: 2-byte size
                if i + 2 > n:
                    break
                sz = struct.unpack_from("<H", grpprl, i)[0]
                yield op, grpprl[i + 2:i + 2 + sz - 1]
                i += 2 + sz - 1
                continue
            sz = grpprl[i]
            yield op, grpprl[i + 1:i + 1 + sz]
            i += 1 + sz
        else:
            sz = _SPRA.get(spra, 1)
            yield op, grpprl[i:i + sz]
            i += sz


def _expand(grpprl, data):
    """Resolve sprmPHugePapx (0x6646) into the sprms stored in the Data stream."""
    out = []
    for op, operand in _walk_sprms(grpprl):
        if op == 0x6646 and len(operand) >= 4:
            fc = struct.unpack_from("<I", operand, 0)[0]
            if fc + 2 <= len(data):
                cb = struct.unpack_from("<H", data, fc)[0]
                out.extend(_walk_sprms(data[fc + 2:fc + 2 + cb]))
        else:
            out.append((op, operand))
    return out


class Row:
    """One table row: cells = [(x0, x1, text)] positioned in twips."""
    __slots__ = ("cells", "degraded", "idx")

    def __init__(self, cells, tdef):
        self.degraded = False
        self.idx = -1
        xs = None
        if tdef and len(tdef) >= 1:
            n = tdef[0]
            if n == len(cells) and len(tdef) >= 1 + 2 * (n + 1):
                xs = list(struct.unpack_from("<%dh" % (n + 1), tdef, 1))
        if xs is None:                                       # synthetic geometry (rare)
            self.degraded = True
            xs = [k * 1000 for k in range(len(cells) + 1)]
        self.cells = [(xs[i], xs[i + 1], cells[i].replace("\xa0", " ")) for i in range(len(cells))]


def read_doc_grid(file_bytes):
    """Return list[Row] for every table row of the document (nested rows included)."""
    try:
        return _read_doc_grid(file_bytes)
    except ParseError:
        raise
    except (struct.error, IndexError, ValueError, KeyError, OverflowError, MemoryError) as e:
        raise ParseError("Document structure is damaged or unsupported") from e


def _read_doc_grid(file_bytes):
    try:
        ole = olefile.OleFileIO(BytesIO(file_bytes))
    except Exception as e:                                    # noqa: BLE001
        raise ParseError("Not an OLE2 compound file") from e
    try:
        if not ole.exists("WordDocument"):
            raise ParseError("No WordDocument stream - not a Word document")
        wd = ole.openstream("WordDocument").read()
        if len(wd) < 0x1A2 + 8 or struct.unpack_from("<H", wd, 0)[0] != 0xA5EC:
            raise ParseError("Not a Word 97-2003 binary document")
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        tname = "1Table" if flags & 0x0200 else "0Table"
        if not ole.exists(tname):
            raise ParseError("Table stream missing")
        tbl = ole.openstream(tname).read()
        data = ole.openstream("Data").read() if ole.exists("Data") else b""
    finally:
        ole.close()

    fcClx, lcbClx = struct.unpack_from("<II", wd, 0x01A2)
    ccpText = struct.unpack_from("<i", wd, 0x4C)[0]
    clx = tbl[fcClx:fcClx + lcbClx]
    i = 0
    while i < len(clx) and clx[i] == 1:
        i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
    if i >= len(clx) or clx[i] != 2:
        raise ParseError("Piece table (CLX) not found")
    lcb = struct.unpack_from("<I", clx, i + 1)[0]
    plc = clx[i + 5:i + 5 + lcb]
    n = (lcb - 4) // 12
    cps = struct.unpack_from("<%di" % (n + 1), plc, 0)

    chars, fcs = [], []
    for k in range(n):
        fcraw = struct.unpack_from("<I", plc, 4 * (n + 1) + 8 * k + 2)[0]
        comp = bool(fcraw & 0x40000000)
        fc = (fcraw & 0x3FFFFFFF) // 2 if comp else fcraw
        cnt = cps[k + 1] - cps[k]
        if comp:
            txt = wd[fc:fc + cnt].decode("cp1252", errors="replace")
            step = 1
        else:
            txt = wd[fc:fc + 2 * cnt].decode("utf-16-le", errors="replace")
            step = 2
        for j, ch in enumerate(txt):
            chars.append(ch)
            fcs.append(fc + j * step)
    chars, fcs = chars[:ccpText], fcs[:ccpText]

    fcB, lcbB = struct.unpack_from("<II", wd, 0x0102)
    plcb = tbl[fcB:fcB + lcbB]
    nb = (lcbB - 4) // 8
    pns = struct.unpack_from("<%dI" % nb, plcb, 4 * (nb + 1))
    runs = []
    for pn in pns:
        page = wd[pn * 512:pn * 512 + 512]
        if len(page) < 512:
            continue
        crun = page[511]
        rgfc = struct.unpack_from("<%dI" % (crun + 1), page, 0)
        for r in range(crun):
            bo = page[4 * (crun + 1) + 13 * r]
            if bo == 0:
                runs.append((rgfc[r], rgfc[r + 1], b""))
                continue
            off = bo * 2
            cb = page[off]
            if cb == 0:
                content = page[off + 2:off + 2 + page[off + 1] * 2]
            else:
                content = page[off + 1:off + 1 + cb * 2 - 1]
            runs.append((rgfc[r], rgfc[r + 1], content[2:]))
    starts = [r[0] for r in runs]

    def run_for(fc):
        k = bisect.bisect_right(starts, fc) - 1
        return runs[k] if k >= 0 and runs[k][0] <= fc < runs[k][1] else None

    out_rows, cur, icur, buf = [], [], [], []
    for idx, ch in enumerate(chars):
        if ch != "\x07":
            buf.append(ch)
            continue
        run = run_for(fcs[idx])
        ttp = ittp = icell = False
        tdef = None
        if run:
            for op, operand in _expand(run[2], data):
                if op == 0x2417 and operand and operand[0] == 1:
                    ttp = True
                elif op == 0x244C and operand and operand[0] == 1:
                    ittp = True
                elif op == 0x244B and operand and operand[0] == 1:
                    icell = True
                elif op == 0xD608:
                    tdef = operand
        if ittp:
            out_rows.append(Row(icur, tdef))
            icur = []
        elif ttp:
            out_rows.append(Row(cur, tdef))
            cur = []
        elif icell:
            icur.append("".join(buf))
            buf = []
        else:
            cur.append("".join(buf))
            buf = []
            continue
        buf = []
    if not out_rows:
        raise ParseError("No table found in document")
    for k, r in enumerate(out_rows):
        r.idx = k
    return out_rows


# ─────────────────────────────────────────────────────────────────────────────
# 2. VALUE PARSING
# ─────────────────────────────────────────────────────────────────────────────
def cl(t):
    """Collapse all whitespace/control characters."""
    return " ".join(str(t).replace("\xa0", " ").replace("\r", " ").replace("\x0b", " ")
                    .replace("\x07", " ").replace("\x0c", " ").split())


_NUM_RE = re.compile(r"(?<!\w)[-+]?(?:\d+(?:[.,]\d+)*|[.,]\d+)")
_MISSING = {"n/a", "na", "x", "-", "--", "—", "–", "n.a.", "nil", "?", "xx", "xxx"}
_UNIT_RE = re.compile(r"(?i)\b(?:bar|kg/cm2?|deg\.?\s*c|°\s*c|c|rpm|r/min|mm\s*hg|mmwc|mmh2o|mm\s*h2o|kw|bhp|%|t/h|knots?|kn|mbar|hrs?|m/s|g/kwh|g/bhph)\b\.?")


def _num_from_token(s, comma="decimal"):
    s = s.strip()
    sign = -1.0 if s.startswith("-") else 1.0
    s = s.lstrip("+-")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        if comma == "thousands" and re.fullmatch(r"[1-9]\d{0,2}(,\d{3})+", s):
            s = s.replace(",", "")
        elif s.count(",") == 1:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(".") > 1:
        s = s.replace(".", "") if re.fullmatch(r"\d{1,3}(\.\d{3})+", s) else s
    try:
        return sign * float(s)
    except ValueError:
        return None


def to_num(text, comma="decimal"):
    """First number in the text (None if absent)."""
    m = _NUM_RE.search(cl(text))
    return _num_from_token(m.group(0), comma) if m else None


def numbers_in(text, comma="decimal"):
    return [v for v in (_num_from_token(m.group(0), comma) for m in _NUM_RE.finditer(cl(text))) if v is not None]


def cell_value(text, comma="decimal"):
    """Strict reading of a data cell.
    Returns (value, kind): kind in {'num','missing','empty','text','multi'}."""
    t = cl(text)
    if not t:
        return None, "empty"
    if t.lower() in _MISSING:
        return None, "missing"
    t2 = _UNIT_RE.sub(" ", t.replace("%", " ")).strip(" :=")
    t2 = re.sub(r"(?i)^(?:ave(?:rage)?\.?|avg\.?)\s*[:=]?\s*", "", t2).strip()
    if re.fullmatch(r"[-+]?(?:\d+(?:[.,]\d+)*|[.,]\d+)", t2):
        return to_num(t2, comma), "num"
    nums = numbers_in(t2, comma)
    letters = re.sub(r"[\d.,+\-\s/:=()%°]", "", t2)
    if len(nums) == 1 and len(letters) <= 2:
        return nums[0], "num"
    if len(nums) >= 2 and len(letters) <= 2:
        return nums, "multi"
    return None, "text"


_MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


def parse_date(text):
    """Parse many date notations (day-first). Returns datetime.date or None."""
    t = cl(text)
    if not t:
        return None
    pats = [
        (r"(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)", "ymd"),
        (r"(?<!\d)(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4}|\d{2})(?!\d)", "dmy"),
        (r"(?<!\d)(\d{1,2})(?:st|nd|rd|th)?[\s\-/.]*([A-Za-z]{3,9})\.?,?[\s\-/.'‘’`]*(\d{4}|\d{2})(?!\d)", "dMy"),
        (r"([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", "Mdy"),
    ]
    for rx, kind in pats:
        m = re.search(rx, t)
        if not m:
            continue
        try:
            if kind == "ymd":
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            elif kind == "dmy":
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if mo > 12 >= d:
                    d, mo = mo, d
                if y < 100:
                    y += 2000
            elif kind == "dMy":
                d, y = int(m.group(1)), int(m.group(3))
                mo = _MONTHS.get(m.group(2)[:3].lower())
                if y < 100:
                    y += 2000
            else:
                mo = _MONTHS.get(m.group(1)[:3].lower())
                d, y = int(m.group(2)), int(m.group(3))
            if mo and 1990 <= y <= 2100:
                return _date(y, mo, d)
        except ValueError:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# 3. GEOMETRY HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _ctr(c):
    return (c[0] + c[1]) / 2.0


def cells_in(row, xa, xb):
    """Non-empty cells whose centre lies inside [xa, xb)."""
    return [c for c in row.cells if cl(c[2]) and xa <= _ctr(c) < xb]


def in_col(c, xa, xb):
    """Data cell belongs to a column if it overlaps it substantially."""
    ov = min(c[1], xb) - max(c[0], xa)
    if ov <= 0:
        return False
    return ov >= 0.4 * min(c[1] - c[0], xb - xa) or xa <= _ctr(c) < xb


_CLAIMS = {"s": None}          # set of (row, x0, x1) keys of cells the extractors have consumed


def claim(row, c):
    s = _CLAIMS["s"]
    if s is not None and c is not None:
        s.add((row.idx, c[0], c[1]))


def find_cell(rows, i0, i1, rx, xa=-10 ** 9, xb=10 ** 9):
    for i in range(max(i0, 0), min(i1, len(rows))):
        for c in rows[i].cells:
            if re.search(rx, cl(c[2]), re.I) and xa <= _ctr(c) < xb:
                claim(rows[i], c)
                return i, c
    return None


def find_all(rows, i0, i1, rx, xa=-10 ** 9, xb=10 ** 9):
    out = []
    for i in range(max(i0, 0), min(i1, len(rows))):
        for c in rows[i].cells:
            if re.search(rx, cl(c[2]), re.I) and xa <= _ctr(c) < xb:
                claim(rows[i], c)
                out.append((i, c))
    return out


def scan_below(rows, i, xa, xb, max_rows=5, comma="decimal"):
    """First non-empty cell below row i inside the column. Returns number or None."""
    for j in range(i + 1, min(i + 1 + max_rows, len(rows))):
        for c in rows[j].cells:
            if cl(c[2]) and in_col(c, xa, xb):
                v, kind = cell_value(c[2], comma)
                if kind == "num":
                    claim(rows[j], c)
                return (v if kind == "num" else None), j
    return None, None


def numbered_values(rows, i0, i1, xa, xb, maxn=4):
    """Single-column readings numbered 1,2,.. (T/C 1, T/C 2): {n: value}."""
    seq = []
    for j in range(i0, min(i1, len(rows))):
        for cc in rows[j].cells:
            if cl(cc[2]) and in_col(cc, xa, xb):
                seq.append((rows[j], cc, cl(cc[2])))
    vals, cur, expect = {}, None, 1
    for k, (rw, cc, t) in enumerate(seq):
        nxt = seq[k + 1][2] if k + 1 < len(seq) else None
        if t == str(expect) and expect <= maxn and (cur is None or cur in vals or nxt != str(expect)):
            cur, expect = expect, expect + 1
            claim(rw, cc)
            continue
        mm = re.match(r"^%d\s+(\S.*)$" % expect, t) if expect <= maxn else None
        if mm:                                     # number and reading typed into the same cell ("1 359")
            v, kind = cell_value(mm.group(1))
            if kind == "num":
                cur, expect = expect, expect + 1
                vals.setdefault(cur, v)
                claim(rw, cc)
                continue
        v, kind = cell_value(t)
        if kind != "num":
            continue
        if cur is None:
            cur, expect = 1, 2
        vals.setdefault(cur, v)
        claim(rw, cc)
    return vals


# ─────────────────────────────────────────────────────────────────────────────
# 4. CYLINDER GROUPS  (numbering row + data row pairs, matched by x-geometry)
# ─────────────────────────────────────────────────────────────────────────────
def extract_trio(rows, r0, r1, xa, xb):
    """Per-cylinder values from a group laid out as [1 2 3 / data / 4 5 6 / data / 7 8 9 / data].
    Returns ({cyl: dict(value, src, row)}, issues, n_numbering_rows)."""
    out, issues = {}, []
    nrows = []
    for i in range(r0, r1):
        qual = []
        for c in cells_in(rows[i], xa, xb):
            toks = cl(c[2]).split()
            if toks and re.fullmatch(r"[1-9]", toks[0]):
                qual.append((c, int(toks[0]), toks[1:]))
        if len(qual) >= 2:
            ds = [q[1] for q in qual]
            if all(ds[k] < ds[k + 1] for k in range(len(ds) - 1)) and len({(d - 1) // 3 for d in ds}) == 1:
                nrows.append((i, qual))

    def record(d, v, src, row):
        cur = out.get(d)
        if cur and cur["value"] is not None and v is not None:
            if abs(cur["value"] - v) > 1e-9:
                issues.append(("warn", f"Cyl {d}: conflicting values {cur['value']:g} ({cur['src']}) vs {v:g} ({src}) - using the cell value"))
                if src == "cell" and cur["src"] != "cell":
                    out[d] = dict(value=v, src=src, row=row)
            return
        if cur and cur["value"] is not None:
            return
        out[d] = dict(value=v, src=src, row=row)

    for k, (i, qual) in enumerate(nrows):
        end = nrows[k + 1][0] if k + 1 < len(nrows) else r1
        numcells = [(c[0], c[1], d) for c, d, _ in qual]
        for c, d, rest in qual:
            claim(rows[i], c)
            out.setdefault(d, dict(value=None, src=None, row=i))
            if rest:
                v, kind = cell_value(rest[0])
                if kind == "num":
                    record(d, v, "embedded", i)
                    if len(rest) > 1:
                        issues.append(("info", f"Cyl {d}: extra text next to the number in its header cell"))
        for j in range(i + 1, end):
            for c in cells_in(rows[j], xa, xb):
                spanned = [d for (n0, n1, d) in numcells if c[0] <= (n0 + n1) / 2 <= c[1]]
                v, kind = cell_value(c[2])
                if kind in ("empty", "missing"):
                    continue
                if kind == "text":
                    issues.append(("info", f"Ignored text '{cl(c[2])[:24]}' in a data row"))
                    continue
                if len(spanned) >= 2:
                    toks = numbers_in(c[2])
                    if len(toks) == len(spanned):
                        for d, tv in zip(spanned, toks):
                            record(d, tv, "cell", j)
                        claim(rows[j], c)
                    else:
                        issues.append(("warn", f"A merged cell covers cylinders {spanned} and cannot be split"))
                    continue
                if kind == "multi":
                    issues.append(("warn", f"A single cell holds several numbers ('{cl(c[2])[:24]}')"))
                    continue
                best = max(numcells, key=lambda n: min(c[1], n[1]) - max(c[0], n[0]))
                ov = min(c[1], best[1]) - max(c[0], best[0])
                if ov <= 0.3 * min(c[1] - c[0], best[1] - best[0]):
                    issues.append(("info", f"Orphan value '{cl(c[2])[:12]}' not aligned to any cylinder"))
                    continue
                record(best[2], v, "cell", j)
                claim(rows[j], c)
    return out, issues, len(nrows)


def _avg_row(rows, i0, i1, xa=-10 ** 9, xb=10 ** 9):
    for i in range(i0, min(i1, len(rows))):
        for c in rows[i].cells:
            if re.match(r"(?i)^\s*(ave|avg|mean)", cl(c[2])) and xa <= _ctr(c) < xb + 10 ** 8:
                return i
    return None


def _avg_value(row, xa, xb):
    """Document 'Average' value from cells in column range; returns (value, decimals) or (None, 0)."""
    for c in row.cells:
        if cl(c[2]) and xa <= _ctr(c) < xb:
            claim(row, c)
    txt = " ".join(cl(c[2]) for c in row.cells if cl(c[2]) and xa <= _ctr(c) < xb)
    txt = re.sub(r"(?i)\b(ave(?:rage)?|avg|mean)\b\.?\s*[:=]?", " ", txt)
    m = _NUM_RE.search(txt)
    if not m:
        return None, 0
    tok = m.group(0)
    v = _num_from_token(tok)
    dec = len(re.split(r"[.,]", tok)[-1]) if re.search(r"[.,]\d", tok) else 0
    return v, dec


# ─────────────────────────────────────────────────────────────────────────────
# 5. TOP-LEVEL DOCUMENT PARSER
# ─────────────────────────────────────────────────────────────────────────────
PLAUSIBLE = {  # key: (low, high)
    "p_max": (25, 220), "p_comp": (15, 190), "fuel_index": (5, 150), "exh_temp": (120, 650),
    "pmax_adj": (0, 200), "cw_out": (30, 110), "piston_out": (25, 90), "pi": (0.5, 60),
}
GROUPS = [("p_max", r"^p\s*[-.]?\s*max"), ("p_comp", r"^p\s*[-.]?\s*comp"),
          ("fuel_index", r"^(fuel\s*pump|fuel\s*index|f\.?\s*p\.?\s*i)")]


def _engine_cyl_count(engine_type):
    t = cl(engine_type).upper()
    m = re.search(r"(?<!\d)(\d{1,2})\s*([A-Z]{1,2})\s*([0-9O]{2})", t)
    if m and 3 <= int(m.group(1)) <= 14:
        return int(m.group(1))
    return None


def parse_tec_005(file_bytes):
    """Parse a TEC-005 Main Engine Performance Data Sheet (.doc). Never raises for content problems;
    raises ParseError only if the file is not a readable Word binary document."""
    rows = read_doc_grid(file_bytes)
    _CLAIMS["s"] = set()
    issues = []          # (severity, message)   severity: info | warn | error
    res = dict(identity={}, op={}, cylinders=[], avgs_doc={}, avgs_calc={}, tc={}, singles={}, header_extra={}, tc_info={}, remarks_series=[],
               remarks="", issues=issues, n_cylinders=0, status="failed", layout={})

    if any(r.degraded for r in rows):
        issues.append(("warn", "Some table rows had no readable geometry; positional matching is degraded"))

    # ── locate anchors ──
    hit = find_cell(rows, 0, len(rows), r"^p\s*[-.]?\s*max\s*bar") or find_cell(rows, 0, len(rows), r"^p\s*[-.]?\s*max\b(?!.*adjust)")
    if not hit:
        issues.append(("error", "Performance table header ('P max bar') not found - this does not look like a TEC-005 sheet"))
        return res
    H = hit[0]
    A = _avg_row(rows, H + 1, len(rows))
    sp = find_cell(rows, H, len(rows), r"special\s*points")
    exh = find_cell(rows, H + 1, len(rows), r"exhaust\s*gas")
    res["layout"] = dict(header_row=H, avg_row=A, exhaust_row=exh[0] if exh else None, special_row=sp[0] if sp else None)
    perf_end = A if A is not None else (exh[0] if exh else min(H + 12, len(rows)))

    # ── group ranges from the header row ──
    ranges = {}
    for key, rx in GROUPS:
        for c in rows[H].cells:
            if re.search(rx, cl(c[2]), re.I):
                ranges[key] = (c[0], c[1])
                claim(rows[H], c)
                break
    structure_notes = 0
    if len(ranges) < 3:                               # fallback: chunk the first long numbering row into triplets
        issues.append(("warn", "Group headers incomplete - groups inferred from numbering geometry"))
        structure_notes += 1
        best = None
        for i in range(H + 1, perf_end):
            nc = [c for c in rows[i].cells if re.fullmatch(r"[1-9]", cl(c[2]).split(" ")[0] if cl(c[2]) else "")]
            if len(nc) >= 9 and (best is None or len(nc) > len(best)):
                best = nc
        if best:
            chunks = [best[k:k + 3] for k in range(0, len(best) - 2, 3)]
            keys = ["pi", "p_max", "p_comp", "fuel_index"][-len(chunks):] if len(chunks) <= 4 else None
            if keys:
                for k, ch in zip(keys, chunks):
                    ranges.setdefault(k, (ch[0][0], ch[-1][1]))

    per = {}      # key -> {cyl: {...}}
    nnum = {}
    for key, _ in GROUPS:
        if key not in ranges:
            continue
        xa, xb = ranges[key]
        got, iss, nn = extract_trio(rows, H + 1, perf_end, xa, xb)
        per[key], nnum[key] = got, nn
        issues.extend(iss)
        if A is not None:
            v, dec = _avg_value(rows[A], xa, xb)
            if v is not None:
                res["avgs_doc"][key] = (v, dec)

    # ── Pi (mean indicated pressure) ──
    for cc in rows[H].cells:
        if re.match(r"(?i)^pi\b", cl(cc[2])):
            claim(rows[H], cc)
            got, iss, nn = extract_trio(rows, H + 1, perf_end, cc[0], cc[1])
            per["pi"], nnum["pi"] = got, nn
            issues.extend(i for i in iss if i[0] != "info")
            if A is not None:
                v, dec = _avg_value(rows[A], cc[0], cc[1])
                if v is not None:
                    res["avgs_doc"]["pi"] = (v, dec)
            break

    # ── exhaust block ──
    exh_rng = adj_rng = None
    E_h = E_a = None
    if exh:
        E_h = exh[0]
        sub = find_cell(rows, E_h, E_h + 4, r"exhaust\s*valve")
        exh_rng = (sub[1][0], sub[1][1]) if sub else (exh[1][0], exh[1][1])
        adj = find_cell(rows, E_h, E_h + 1, r"adjust")
        if adj:
            adj_rng = (adj[1][0] - 50, exh_rng[0] - 1)
        E_a = _avg_row(rows, E_h + 1, len(rows))
        end = E_a if E_a is not None else (sp[0] if sp else len(rows))
        for key, rng in (("exh_temp", exh_rng), ("pmax_adj", adj_rng)):
            if not rng:
                continue
            got, iss, nn = extract_trio(rows, E_h + 1, end, rng[0], rng[1])
            per[key], nnum[key] = got, nn
            issues.extend(iss)
            if E_a is not None:
                v, dec = _avg_value(rows[E_a], rng[0], rng[1])
                if v is not None:
                    res["avgs_doc"][key] = (v, dec)
    else:
        issues.append(("warn", "Exhaust gas temperature block not found"))

    # ── special points: jacket CW outlet / piston oil outlet per cylinder ──
    S_a = None
    if sp:
        S = sp[0]
        S_a = _avg_row(rows, S + 1, len(rows))
        send = S_a if S_a is not None else len(rows)
        for key, rx in (("cw_out", r"outlet\s*cyl"), ("piston_out", r"outlet\s*pist")):
            h = find_cell(rows, S, S + 4, rx)
            if not h:
                continue
            xa, xb = h[1][0], h[1][1]
            got, iss, nn = extract_trio(rows, h[0] + 1, send, xa, xb)
            per[key], nnum[key] = got, nn
            issues.extend(i for i in iss if i[0] != "info")
            if S_a is not None:
                v, dec = _avg_value(rows[S_a], xa, xb)
                if v is not None:
                    res["avgs_doc"][key] = (v, dec)
    res["layout"].update(special_avg_row=S_a, exhaust_avg_row=E_a)

    # ── header, operating conditions, T/C, singles, remarks (best effort, never fatal) ──
    for fn, label in ((lambda: _read_identity(rows, H, res), "Header fields"),
                      (lambda: _harvest_header(rows, H, res), "Header details"),
                      (lambda: _read_operating(rows, H, perf_end, ranges, res), "Operating conditions"),
                      (lambda: _read_extras(rows, E_h, E_a, exh_rng, sp, S_a, res), "Turbocharger / auxiliary readings"),
                      (lambda: _read_remarks(rows, res), "Remarks"),
                      (lambda: _mine_remarks(res), "Remark values")):
        try:
            fn()
        except Exception as e:                                           # noqa: BLE001
            issues.append(("info", f"{label}: partly unreadable ({type(e).__name__})"))

    # ── report date: majority of all dated fields ──
    cands = res["identity"].pop("_date_candidates", [])
    ds = [d for _, d in cands if d]
    if ds:
        best = max(set(ds), key=ds.count)
        res["identity"]["date"] = best
        if len(set(ds)) > 1:
            issues.append(("warn", "Dates on the sheet disagree: " + ", ".join(sorted({d.strftime('%d.%m.%Y') for d in ds})) +
                           f" - using {best.strftime('%d.%m.%Y')}"))
    else:
        res["identity"]["date"] = None

    # ── assemble per-cylinder table & reconcile cylinder count ──
    keys = ("p_max", "p_comp", "fuel_index", "exh_temp", "pmax_adj", "cw_out", "piston_out", "pi")
    n_data = 0
    for key in ("p_max", "p_comp", "fuel_index", "exh_temp"):
        for d, rec in per.get(key, {}).items():
            if rec["value"] is not None:
                n_data = max(n_data, d)
    n_type = _engine_cyl_count(res["identity"].get("engine_type", ""))
    res["n_type"], res["n_data"] = n_type, n_data
    if n_data == 0:
        n = n_type or max([len(v) for v in per.values()] + [0]) or 6
        n = min(max(n, 3), 9)
    elif n_type and n_type >= n_data:
        n = n_type
        if n_type > n_data:
            issues.append(("warn", f"Engine type says {n_type} cylinders but data only reaches cylinder {n_data} - "
                                   f"cylinder(s) {', '.join(str(i) for i in range(n_data + 1, n_type + 1))} left empty for manual entry"))
    elif n_type and n_type < n_data:
        n = n_data
        issues.append(("warn", f"Data found for {n_data} cylinders but engine type says {n_type} - trusting the data; please verify"))
    else:
        n = n_data
        issues.append(("info", "Cylinder count taken from data (engine type unreadable)"))
    n = min(n, 9)
    res["n_cylinders"] = n

    cyls = []
    for d in range(1, n + 1):
        row = {"id": d}
        for key in keys:
            rec = per.get(key, {}).get(d)
            row[key] = rec["value"] if rec else None
            row["src_" + key] = rec["src"] if rec else None
        cyls.append(row)
    res["cylinders"] = cyls
    res["perf_ranges"] = ranges

    # ── plausibility sanitising (never silently drop - flag) ──
    nbad = 0
    for row in cyls:
        for key, (lo, hi) in PLAUSIBLE.items():
            v = row.get(key)
            if v is not None and not (lo <= v <= hi):
                issues.append(("warn", f"Cyl {row['id']} {key}: {v:g} is outside the physical range {lo}-{hi}"))
                nbad += 1
    res["n_implausible"] = nbad

    # ── averages: calculate & cross-validate against the document's own Average cells ──
    xv = {}
    for key in keys:
        vals = [r[key] for r in cyls if r[key] is not None]
        if vals:
            res["avgs_calc"][key] = sum(vals) / len(vals)
        doc = res["avgs_doc"].get(key)
        if doc and vals and doc[0] and doc[0] > 0:
            tier, label = _avg_tier(res["avgs_calc"][key], doc[0], doc[1])
            xv[key] = dict(tier=tier, label=label, doc=doc[0], calc=res["avgs_calc"][key], dec=doc[1])
            if label in ("weak", "mismatch"):
                issues.append(("warn", f"{_KEYNAME[key]}: printed Average {doc[0]:g} differs from the mean of the cylinder readings "
                                       f"{res['avgs_calc'][key]:.1f} - the sheet is internally inconsistent (or a reading was mis-entered)"))
            elif label == "near":
                issues.append(("info", f"{_KEYNAME[key]}: printed Average {doc[0]:g} vs calculated {res['avgs_calc'][key]:.2f} (minor rounding difference)"))
        elif doc and doc[0] == 0 and not vals:
            res["avgs_doc"].pop(key, None)
    res["xval"] = xv

    # ── status ──
    core_present = sum(1 for r in cyls for k in ("p_max", "p_comp", "fuel_index") if r[k] is not None)
    if core_present == 0 and not any(r["exh_temp"] is not None for r in cyls):
        res["status"] = "blank"
        issues.append(("info", "Blank form: no per-cylinder measurements were entered on this sheet"))
    else:
        res["status"] = "ok"
    res["confidence"] = _confidence(res, ranges, nnum, A is not None, E_h is not None)
    try:
        _coverage(rows, res)
    except Exception as e:                                                   # noqa: BLE001
        issues.append(("info", f"Coverage audit unavailable ({type(e).__name__})"))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# 6. HEADER / OPERATING / EXTRAS
# ─────────────────────────────────────────────────────────────────────────────
_LABEL_START = re.compile(r"(?i)^(bhp|kw|r\s*/\s*min|builder|yard|no|built\s*year|engine\s*(type|no|layout)|checked\s*by|date|m\s*/\s*v|"
                          r"maker|type|governor|cylinder\s*constant|max\s*(rpm|temp)|oil\s*brand|density|sulphur|sulfur|kjoule|heat)\b")


def _is_label(t):
    t = cl(t)
    return (not t) or t.endswith(":") or bool(_LABEL_START.match(t))


def _label_value(rows, i, ci, below=True, max_below=1, rx=None):
    """Value belonging to label cell rows[i].cells[ci]: embedded in the cell / right neighbour / cell below."""
    c = rows[i].cells[ci]
    claim(rows[i], c)
    txt = cl(c[2])
    rest = ""
    if rx:
        rest = cl(re.sub(rx + r"\s*:?", "", txt, count=1, flags=re.I))
    elif ":" in txt:
        rest = cl(txt.split(":", 1)[1])
    if rest:
        return rest
    for k in range(ci + 1, len(rows[i].cells)):
        nc = rows[i].cells[k]
        if cl(nc[2]):
            if not _is_label(nc[2]):
                claim(rows[i], nc)
                return cl(nc[2])
            break
    if below:
        for j in range(i + 1, min(i + 1 + max_below, len(rows))):
            for nc in rows[j].cells:
                if cl(nc[2]) and in_col(nc, c[0], c[1]) and not _is_label(nc[2]):
                    claim(rows[j], nc)
                    return cl(nc[2])
    return ""


def _read_identity(rows, H, res):
    ident = res["identity"]
    cands = []
    end = min(H, 14)
    spec = [
        ("vessel", r"^m\s*/\s*v\b", True), ("engine_type", r"^engine\s*type", False),
        ("builder", r"^builder", False), ("yard", r"^yard", False), ("engine_no", r"^engine\s*no\.?", True),
        ("built_year", r"^built\s*year", True), ("checked_by", r"^checked\s*by", False),
    ]
    for key, rx, below in spec:
        for i in range(0, end):
            hit = [(ci, c) for ci, c in enumerate(rows[i].cells) if re.search(rx, cl(c[2]), re.I)]
            if not hit:
                continue
            v = _label_value(rows, i, hit[0][0], below=below, rx=rx)
            if v:
                ident[key] = v
            break
    if ident.get("built_year"):
        d = parse_date(ident["built_year"])
        m = re.search(r"(?<!\d)(19|20)\d{2}(?!\d)", ident["built_year"])
        ident["built_year"] = str(d.year) if d else (m.group(0) if m else ident["built_year"])
    # rated power & speed
    for i in range(0, min(H, 6)):
        for ci, c in enumerate(rows[i].cells):
            t = cl(c[2])
            if re.match(r"(?i)^(bhp|kw)\b", t) and "rated_power" not in ident:
                v = _label_value(rows, i, ci, below=False, rx=r"^(bhp|kw)\b")
                n = to_num(v, "thousands")
                if n:
                    ident["rated_power"] = n
                    ident["power_unit"] = "kW" if (t.lower().startswith("kw") or "kw" in (v or "").lower()) else "BHP"
            elif re.match(r"(?i)^r\s*/\s*min", t) and "rated_rpm" not in ident:
                n = to_num(_label_value(rows, i, ci, below=False, rx=r"^r\s*/\s*min"), "thousands")
                if n:
                    ident["rated_rpm"] = n
    # fuel
    for key, rx in (("fuel_density", r"^density(\s*at\s*15\s*c)?"), ("fuel_sulphur", r"^sul(ph|f)ur(\s*%)?")):
        for i in range(8, H):
            hit = [(ci, c) for ci, c in enumerate(rows[i].cells) if re.search(rx, cl(c[2]), re.I)]
            if hit:
                n = to_num(_label_value(rows, i, hit[0][0], below=False, rx=rx))
                if n is not None:
                    ident[key] = n
                break
    # date candidates
    for i in range(0, min(H, 6)):
        for ci, c in enumerate(rows[i].cells):
            if re.match(r"(?i)^date\s*:?", cl(c[2])):
                raw = _label_value(rows, i, ci, below=False, rx=r"^date")
                cands.append(("header", parse_date(raw) if raw else None))
    h_first = rows[H].cells[0] if rows[H].cells else None
    if h_first and re.match(r"(?i)^date\b", cl(h_first[2])):
        claim(rows[H], h_first)
        cands.append(("table-head", parse_date(cl(h_first[2]))))
        for j in range(H + 1, H + 3):
            for c in rows[j].cells:
                if cl(c[2]) and in_col(c, h_first[0], h_first[1]):
                    d = parse_date(c[2])
                    if d:
                        claim(rows[j], c)
                        cands.append(("table-row", d))
    ident["_date_candidates"] = [(s, d) for s, d in cands if d]


_OP_SPEC = [
    ("draught_fore", r"^draught\s*fore(?:\s*m\b)?", "decimal"), ("draught_aft", r"^draught\s*aft(?:\s*m\b)?", "decimal"),
    ("run_hours", r"^total\s*run\.?\s*(?:hrs?\.?)?", "thousands"), ("rpm", r"^rpm\b", "decimal"),
    ("wind", r"^wind\s*(?:speed)?\s*(?:\(?\d*\s*m\s*/\s*s\)?)?", "decimal"),
    ("wave", r"^wave\s*height", "decimal"), ("speed_setting", r"^speed\s*setting", "decimal"),
    ("kw_indic", r"^kw\s*indic\w*", "thousands"), ("kw_eff", r"^kw\s*effec\w*", "thousands"),
    ("load_pct", r"^load\s*%?", "decimal"), ("log_knots", r"^log\s*knots", "decimal"),
    ("obs_knots", r"^obs\s*knots", "decimal"), ("governor_index", r"^governor\s*index", "decimal"),
    ("bhp_op", r"^bhp\b", "thousands"), ("barometer", r"^barom\w*\.?\s*(?:mbar)?", "thousands"),
    ("sfoc_kwh", r"^g\s*/\s*kwh\s*(?:effective)?", "decimal"),
]


def _read_operating(rows, H, end, ranges, res):
    op = res["op"]
    xmax = min([r[0] for r in ranges.values()] + [10 ** 9])
    for key, rx, mode in _OP_SPEC:
        for i in range(H, end):
            found = False
            for c in rows[i].cells:
                t = cl(c[2])
                m = re.match(rx, t, re.I)
                if not m or _ctr(c) >= xmax:
                    continue
                found = True
                claim(rows[i], c)
                rest = t[m.end():].strip()
                v = kind = None
                if rest:
                    v, kind = cell_value(rest, mode)
                    if kind == "multi":
                        v = v[0]
                        kind = "num"
                if kind != "num":
                    v, _j = scan_below(rows, i, c[0], c[1], max_rows=2, comma=mode)
                    if v is None:        # a "value row" cell may hold text like 'SSE 14' (wind)
                        for j in range(i + 1, min(i + 3, len(rows))):
                            for nc in rows[j].cells:
                                if cl(nc[2]) and in_col(nc, c[0], c[1]):
                                    nv = to_num(nc[2], mode)
                                    if nv is not None and key in ("wind", "wave", "load_pct"):
                                        v = nv
                                        claim(rows[j], nc)
                                    break
                            if v is not None:
                                break
                if v is not None:
                    op[key] = v
                break
            if found and key in op:
                break
    # directions (first 'Direct' = wind, second = waves) and the hour of the test
    dirs = find_all(rows, H, end, r"^direct", -10 ** 9, xmax)
    for n_, (i, cc) in enumerate(dirs[:2]):
        t = cl(cc[2])[6:].strip()
        if not t:
            for j in range(i + 1, min(i + 3, len(rows))):
                for nc in rows[j].cells:
                    if cl(nc[2]) and in_col(nc, cc[0], cc[1]):
                        t = "" if re.match(r"(?i)^direct", cl(nc[2])) else cl(nc[2])
                        if t:
                            claim(rows[j], nc)
                        break
                if t:
                    break
        if t and re.fullmatch(r"(?i)[NSEW]{1,3}(?:[\s/-]+[NSEW]{1,3})*|variable|calm|var\.?", t):
            op["wind_dir" if n_ == 0 else "wave_dir"] = t
    hr = find_cell(rows, H, end, r"^hour", -10 ** 9, xmax)
    if hr:
        t = cl(hr[1][2])[4:].strip(" :")
        if not t:
            for j in range(hr[0] + 1, min(hr[0] + 3, len(rows))):
                for nc in rows[j].cells:
                    if cl(nc[2]) and in_col(nc, hr[1][0], hr[1][1]):
                        t = cl(nc[2])
                        claim(rows[j], nc)
                        break
                if t:
                    break
        if t:
            op["hour"] = t


def _read_extras(rows, E_h, E_a, exh_rng, sp, S_a, res):
    tc = {}
    singles = res["singles"]
    if E_h is not None and exh_rng:
        end = E_a if E_a is not None else (sp[0] if sp else len(rows))
        x0 = exh_rng[1] - 5
        spec = [
            ("turb_in", r"^inlet$"), ("turb_out", r"^outlet$"), ("exh_press", r"^receiver$"),
            ("turb_out_press", r"^turbine\s*outlet"), ("tc_rpm", r"t\s*/\s*c\s*rpm"),
            ("dp_filter", r"filter"), ("dp_cooler", r"p\s*cooler"), ("inl_blower", r"(inl|in\.|inlet)\w*\.?\s*blower"),
            ("bef_cooler", r"^bef"), ("aft_cooler", r"^aft"),
        ]
        for key, rx in spec:
            hit = find_cell(rows, E_h, E_h + 4, rx, x0, 10 ** 9)
            if not hit:
                continue
            vals = numbered_values(rows, hit[0] + 1, end, hit[1][0], hit[1][1])
            for n, v in vals.items():
                tc.setdefault(n, {})[key] = v
        # scavenge receiver: unit-labelled rows (Bar / mmHg / deg C)
        rec = find_all(rows, E_h, E_h + 4, r"^receiver$", x0, 10 ** 9)
        if len(rec) >= 2 or (rec and "exh_press" in tc.get(1, {}) is False):
            pass
        if len(rec) >= 2:
            i, c = rec[1]
            unit = None
            for j in range(i + 1, end):
                for nc in rows[j].cells:
                    if cl(nc[2]) and in_col(nc, c[0], c[1]):
                        claim(rows[j], nc)
                        t = cl(nc[2])
                        low = t.lower()
                        v, kind = cell_value(t)
                        if kind == "num":
                            if unit:
                                singles.setdefault(unit, v)
                        elif "bar" in low:
                            unit = "scav_bar"
                        elif "mmh2o" in low.replace(" ", "") or "mmwc" in low.replace(" ", ""):
                            unit = "scav_mmwc"
                        elif "mmhg" in low.replace(" ", ""):
                            unit = "scav_mmhg"
                        elif "deg" in low or "temp" in low:
                            unit = "scav_temp"
                        break
    for n_, d_ in tc.items():
        if d_.get("tc_rpm") is not None and 0 < d_["tc_rpm"] < 100:
            res["issues"].append(("info", f"T/C {n_} rpm {d_['tc_rpm']:g} looks like thousands - read as {d_['tc_rpm'] * 1000:,.0f}"))
            d_["tc_rpm"] *= 1000
    res["tc"] = tc
    if sp:
        S = sp[0]
        send = S_a if S_a is not None else len(rows)
        send_all = len(rows)

        def lv(rx, i0, i1, key, xlo=-10 ** 9, xhi=10 ** 9, maxrows=5):
            hit = find_cell(rows, i0, i1, rx, xlo, xhi)
            if hit:
                v, _ = scan_below(rows, hit[0], hit[1][0], hit[1][1], max_rows=maxrows)
                if v is not None:
                    singles[key] = v

        lv(r"^system\s*oil", S, send_all, "lo_system_p")
        lv(r"^cooling\s*oil", S, send_all, "lo_cooling_p")
        lv(r"^ex\s*valve\s*oil", S, send_all, "lo_exvalve_p")
        lv(r"^turbine\s*oil", S, send_all, "lo_turbine_p")
        lv(r"^inlet\s*eng", S, send_all, "lo_inlet_t")
        lv(r"^before\s*filter", S, send_all, "fo_before_filter")
        lv(r"^after\s*filter", S, send_all, "fo_after_filter")
        lv(r"^sea\s*water", S, send_all, "sea_water_t", maxrows=6)
        # FO temperature sits on the Average line in the same column
        ft = find_cell(rows, S, send_all, r"^fo\s*temp")
        if ft:
            v, _ = scan_below(rows, ft[0], ft[1][0], ft[1][1], max_rows=3)
            if v is not None:
                singles["fo_temp"] = v
        ac = find_cell(rows, S, S + 3, r"^air\s*cooler")
        if ac:
            ins = find_cell(rows, ac[0] + 1, ac[0] + 3, r"^inlet", ac[1][0], ac[1][1])
            outs = find_cell(rows, ac[0] + 1, ac[0] + 3, r"^outlet", ac[1][0], ac[1][1])
            for key, h in (("cw_air_in", ins), ("cw_air_out", outs)):
                if h:
                    vals = numbered_values(rows, h[0] + 1, send_all, h[1][0], h[1][1], maxn=2)
                    if 1 in vals:
                        singles[key] = vals[1]
        me = find_cell(rows, S, S + 3, r"^main\s*engine")
        if me:
            inl = find_cell(rows, me[0] + 1, me[0] + 3, r"^inlet", me[1][0], me[1][1])
            if inl:
                v, _ = scan_below(rows, inl[0], inl[1][0], inl[1][1], max_rows=3)
                if v is not None:
                    singles["cw_me_in"] = v
        for key, rx in (("spiral_p", r"pspiral"), ("dp_inout", r"inn\s*/\s*out")):
            hh = find_cell(rows, S, S + 4, rx)
            if hh:
                for n_, v in numbered_values(rows, hh[0] + 1, send_all, hh[1][0], hh[1][1]).items():
                    tc.setdefault(n_, {})[key] = v
        tt = find_cell(rows, S, S + 3, r"^turbine$")
        if tt:
            oo = find_cell(rows, tt[0] + 1, tt[0] + 3, r"^outlet", tt[1][0] - 20, tt[1][1] + 20)
            if oo:
                for n_, v in numbered_values(rows, oo[0] + 1, send_all, oo[1][0], oo[1][1], maxn=4).items():
                    if v >= 10:
                        tc.setdefault(n_, {})["cw_turb_out"] = v
        tb = find_cell(rows, S, S + 4, r"^turbochargers")
        if tb:
            for key, rx in (("tc_oil_in", r"^inlet"), ("tc_oil_out", r"^outlet")):
                h = find_cell(rows, tb[0] + 1, tb[0] + 3, rx, tb[1][0] - 20, 10 ** 9)
                if h:
                    vals = numbered_values(rows, h[0] + 1, send_all, h[1][0], h[1][1], maxn=2)
                    if 1 in vals:
                        singles[key] = vals[1]


def _read_remarks(rows, res):
    for i, r in enumerate(rows):
        for ci, c in enumerate(r.cells):
            if re.match(r"(?i)^remarks", cl(c[2])):
                claim(r, c)
                for x in r.cells[ci + 1:]:
                    claim(r, x)
                parts = [cl(x[2]) for x in r.cells[ci + 1:] if cl(x[2])]
                parts.insert(0, cl(c[2])[7:].strip(" :"))
                for j in range(i + 1, min(i + 4, len(rows))):
                    parts.extend(cl(x[2]) for x in rows[j].cells if cl(x[2]))
                    for x in rows[j].cells:
                        claim(rows[j], x)
                res["remarks"] = " ".join(p for p in parts if p).strip()
                return


_LABEL_VOCAB = re.compile(r"(?i)^(maker|type|max\s*rpm|max\s*temp|cylinder\s*constant|governor|tc\s*specification|bunker\s*station|oil\s*brand|"
                          r"fuel\s*oil\s*viscosity|at\s*deg|kjoule|heat\s*value|sulphur|sulfur|density|cylinder\s*oil|circulating\s*oil|"
                          r"turbo\s*oil|mt\s*/\s*day|engine\s*layout|lub\.?\s*oil\s*system)")


def _labelish(t):
    t = cl(t)
    return t.endswith(":") or bool(_LABEL_VOCAB.match(t)) or bool(re.match(r"(?i)^(serial\s*no|no)\.?$", t))


def _harvest_header(rows, H, res):
    """Generic label -> value harvest of every still-unclaimed label in the header block."""
    claimed = _CLAIMS["s"]
    T = find_cell(rows, 0, H, r"^turbocharger")
    L = find_cell(rows, 0, H, r"^lub\.?\s*oil\s*system")
    t_row = T[0] if T else 4
    l_row = L[0] if L else 9
    cc_ = find_cell(rows, t_row, l_row + 1, r"^cylinder\s*constant")
    split = cc_[1][0] - 5 if cc_ else 4300
    out = res["header_extra"]
    info = res["tc_info"]
    serials = {}
    SER = re.compile(r"(?i)^(?:[A-Z]{1,3}\s?-?)?\d{4,}[A-Z]?$")
    free = lambda r, c_: (r.idx, c_[0], c_[1]) not in claimed

    def put(key, val):
        k, n = key, 2
        while k in out:
            k, n = f"{key} ({n})", n + 1
        out[k] = val

    for i in range(1, H):
        r = rows[i]
        in_tc = t_row <= i < l_row
        for k, c_ in enumerate(r.cells):
            t = cl(c_[2])
            if not t or not _labelish(t):
                continue
            if not free(r, c_):
                nx = next((z for z in r.cells[k + 1:] if cl(z[2])), None)
                if nx is None or not free(r, nx) or _labelish(cl(nx[2])):
                    continue
            if re.match(r"(?i)^(lub\.?\s*oil\s*system|brand$|type$|serial\s*no|no\.?$)", t) and (i >= l_row or re.match(r"(?i)^(lub|serial|no)", t)):
                if re.match(r"(?i)^lub", t):
                    claim(r, c_)
                continue
            label, emb = (t.split(":", 1) + [""])[:2] if ":" in t else (t, "")
            label, emb = label.strip(" :"), emb.strip()
            vals = [emb] if emb else []
            m = k + 1
            while m < len(r.cells) and len(vals) < 3:
                n_ = r.cells[m]
                nt = cl(n_[2])
                m += 1
                if not nt:
                    continue
                if not free(r, n_) or _labelish(nt) or re.match(r"(?i)^(brand|type)$", nt):
                    break
                if re.fullmatch(r"[1-4]", nt) and c_[0] < split:
                    continue
                if in_tc and c_[0] < split and SER.match(nt) and not re.match(r"(?i)^(max)", label):
                    serials.setdefault(len(serials) + 1, nt)
                    claim(r, n_)
                    continue
                vals.append(nt)
                claim(r, n_)
            if not vals:
                for j in range(i + 1, min(i + 2, len(rows))):
                    for n_ in rows[j].cells:
                        nt = cl(n_[2])
                        if nt and free(rows[j], n_) and in_col(n_, c_[0], c_[1]) and not _labelish(nt) \
                                and not any(_labelish(z[2]) and abs(z[0] - c_[0]) < 60 for z in rows[j].cells):
                            vals.append(nt)
                            claim(rows[j], n_)
                            break
            claim(r, c_)
            if not vals:
                continue
            key = label
            if in_tc:
                if c_[0] < split:
                    key = "Turbocharger " + label[0].lower() + label[1:] if label.lower() in ("maker", "type", "max rpm", "max temp. c", "max temp") else label
                elif label.lower() == "type":
                    key = "Governor type"
            put(key, " · ".join(vals))

    for r in rows[H:]:
        for c_ in r.cells:
            mm = re.match(r"(?i)^fo\s*press\.?\s*(?:bar)?\s*(\d+(?:[.,]\d+)?)\s*(?:kg/cm.*)?$", cl(c_[2]))
            if mm and free(r, c_):
                put("FO pressure (bar)", mm.group(1).replace(",", "."))
                claim(r, c_)
    # turbocharger serial numbers: digit cell followed by a text cell, or text typed on the 'SERIAL No.' line
    for i in range(t_row, l_row):
        r = rows[i]
        for k, c_ in enumerate(r.cells):
            t = cl(c_[2])
            if re.fullmatch(r"[1-4]", t) and free(r, c_) and c_[0] < split:
                for n_ in r.cells[k + 1:]:
                    nt = cl(n_[2])
                    if nt:
                        if free(r, n_) and not _labelish(nt) and n_[0] < split + 400 and not re.fullmatch(r"[1-4]", nt):
                            serials[int(t)] = nt
                            claim(r, n_)
                        break
                claim(r, c_)
    ser = find_cell(rows, t_row, t_row + 2, r"^serial")
    if ser:
        r = rows[ser[0]]
        for n_ in r.cells:
            nt = cl(n_[2])
            if nt and n_[0] > ser[1][0] and free(r, n_) and not _labelish(nt):
                serials.setdefault(1, nt)
                claim(r, n_)
    m_ = re.fullmatch(r"(?i)(h\s*t\s*\d{5,})", (out.get("TC specification") or "").replace(" ", ""))
    if m_:
        serials.setdefault(1, out.pop("TC specification"))
    if serials:
        info["serials"] = serials
    # numeric readings of the T/C spec with unit normalisation
    raw = out.get("Turbocharger max RPM") or out.get("Max RPM")
    if raw:
        n = to_num(raw, "thousands")
        if n:
            if re.search(r"1\s*/\s*s\b", raw, re.I):
                n *= 60
                res["issues"].append(("info", f"T/C max speed given per second ('{raw}') - converted to {n:,.0f} rpm"))
            info["max_rpm"] = n
    raw = out.get("Turbocharger max temp. C") or out.get("Turbocharger max temp")
    if raw and to_num(raw):
        info["max_temp"] = to_num(raw)
    for key, name in (("Turbocharger maker", "maker"), ("Turbocharger type", "type")):
        if out.get(key):
            info[name] = out[key]
    if "Cylinder constant (HP, bar)" in out or any(k.lower().startswith("cylinder constant") for k in out):
        kk = next(k for k in out if k.lower().startswith("cylinder constant"))
        info["cyl_constant"] = out[kk]
    for kk in out:
        if kk.lower().startswith("governor") and kk != "Governor type":
            info.setdefault("governor", out[kk])
    if out.get("Governor type"):
        info["governor_type"] = out["Governor type"]


def _mine_remarks(res):
    """Pull 'No.1-59, No.2-56 ...' style per-cylinder series out of the free-text remarks."""
    txt = res.get("remarks") or ""
    rx = re.compile(r"(?:\bno\.?|\bcyl(?:inder)?\.?|#)\s*#?\s*(\d)\s*[-–:=(]*\s*\(?\s*(\d+(?:[.,]\d+)?)", re.I)
    ms = list(rx.finditer(txt))
    series, cur = [], []
    for m in ms:
        if cur and m.start() - cur[-1].end() > 30:
            series.append(cur)
            cur = []
        cur.append(m)
    if cur:
        series.append(cur)
    for s in series:
        if len(s) < 2:
            continue
        pre = txt[max(0, s[0].start() - 45):s[0].start()]
        words = re.findall(r"[A-Za-z][A-Za-z.]*", pre)
        ctx = " ".join(words[-4:]).strip(". ") or "values in remarks"
        vals = {}
        for m in s:
            v = _num_from_token(m.group(2))
            if v is not None:
                vals.setdefault(int(m.group(1)), v)
        if len(vals) >= 2:
            res["remarks_series"].append(dict(context=ctx, values=vals))


# ─────────────────────────────────────────────────────────────────────────────
# 7. CROSS-VALIDATION & CONFIDENCE
# ─────────────────────────────────────────────────────────────────────────────
_KEYNAME = {"p_max": "Pmax", "p_comp": "Pcomp", "fuel_index": "Fuel index", "exh_temp": "Exhaust temp",
            "pmax_adj": "Pmax adjustment", "cw_out": "Jacket CW outlet", "piston_out": "Piston oil outlet", "pi": "Pi"}


def _avg_tier(calc, doc, dec):
    """Compare our mean with the sheet's printed Average (allow rounding / truncation)."""
    step = 10.0 ** (-dec)
    d = abs(calc - doc)
    if d <= 0.5 * step + 1e-9:
        return 1.0, "match"
    if abs((int(calc / step) * step) - doc) < 1e-9:           # crew truncated instead of rounding
        return 1.0, "match"
    rel = d / max(abs(doc), 1e-9)
    if d <= step + 1e-9 or rel <= 0.01:
        return 0.75, "near"
    if rel <= 0.03:
        return 0.4, "weak"
    return 0.0, "mismatch"


def _confidence(res, ranges, nnum, has_avg, has_exh):
    cyls = res["cylinders"]
    n = max(len(cyls), 1)
    if res["status"] == "blank":
        return dict(total=0.0, structure=0.0, completeness=0.0, plausibility=0.0, crossval=0.0, level="blank",
                    label="Blank form", verified_by_sheet=False)
    groups_ok = sum(1 for k in ("p_max", "p_comp", "fuel_index") if nnum.get(k, 0) > 0)
    structure = 0.6 * groups_ok / 3 + 0.2 * (1 if has_avg else 0) + 0.2 * (1 if has_exh else 0)
    if any(sev == "warn" and ("merged" in m or "geometry" in m) for sev, m in res["issues"]):
        structure *= 0.8
    exp = n * 4
    present = sum(1 for r in cyls for k in ("p_max", "p_comp", "fuel_index", "exh_temp") if r[k] is not None)
    completeness = present / exp
    vals = [(r[k], k, r) for r in cyls for k in ("p_max", "p_comp", "fuel_index", "exh_temp") if r[k] is not None]
    good = sum(1 for v, k, _ in vals if PLAUSIBLE[k][0] <= v <= PLAUSIBLE[k][1])
    both = [r for r in cyls if r["p_max"] is not None and r["p_comp"] is not None]
    good += sum(1 for r in both if r["p_comp"] < r["p_max"])
    plaus = good / max(len(vals) + len(both), 1)
    xv = res.get("xval", {})
    w = {"p_max": 2, "p_comp": 2, "fuel_index": 2, "exh_temp": 1}
    num = sum(w[k] * xv[k]["tier"] for k in w if k in xv)
    den = sum(w[k] for k in w if k in xv)
    crossval = num / den if den else 0.6
    total = 100 * (0.20 * structure + 0.30 * completeness + 0.15 * plaus + 0.35 * crossval)
    warns = sum(1 for sev, _ in res["issues"] if sev == "warn")
    total -= min(8.0, 2.0 * warns)
    total = max(0.0, min(100.0, total))
    if total >= 97:
        level, label = "verified", "Verified"
    elif total >= 90:
        level, label = "high", "High confidence"
    elif total >= 75:
        level, label = "review", "Review recommended"
    else:
        level, label = "low", "Low confidence"
    return dict(total=round(total, 1), structure=round(structure, 3), completeness=round(completeness, 3),
                plausibility=round(plaus, 3), crossval=round(crossval, 3), level=level, label=label,
                verified_by_sheet=bool(den))


# ─────────────────────────────────────────────────────────────────────────────
# 8. COVERAGE AUDIT  -  every non-empty cell is claimed, template text, or listed as "not mapped"
# ─────────────────────────────────────────────────────────────────────────────
FURNITURE = [
    r"^\d*this form must be", r"^n/?a$", r"^mm ?h2o$", r"^external (system|from)", r"^(mmhg|mmwc|bar|deg\.? ?c|on|off|brand|type|turbine|receiver|inlet|outlet|average)$",
    r"^g/(bhp|kw)", r"^exh\. pressure", r"^scavenge air (pressure|temp)", r"^auxil", r"^lubricating oil$", r"^cooling water temp",
    r"^temperature deg", r"^p bar$", r"^fo p bar$", r"^fo press\.?\s*(bar)?(\s*kg/cm.*)?$", r"^special points", r"^sea water temp", r"^cooling oil$", r"^system oil$",
    r"^(before|after) filter$", r"^exhaust gas temp", r"^pmax adjustment", r"^(p )?(max|comp) bar$", r"^fuel pump index$",
    r"^lub\.?\s*oil system", r"^turbocharger\(s\)", r"^serial no", r"^checked by", r"^[1-9]$",
]


def _is_furniture(t):
    if len(t) <= 1 and not t.isdigit():
        return True
    return any(re.match(p, t, re.I) for p in FURNITURE)


def _coverage(rows, res):
    claimed = _CLAIMS["s"] or set()
    tot = cl_n = fur = 0
    left = []
    for r in rows:
        for ci, c in enumerate(r.cells):
            t = cl(c[2])
            if not t:
                continue
            tot += 1
            if (r.idx, c[0], c[1]) in claimed:
                cl_n += 1
            elif _is_furniture(t):
                fur += 1
            else:
                prev = next((cl(x[2]) for x in reversed(r.cells[:ci]) if cl(x[2])), "")
                above = ""
                for j in range(r.idx - 1, max(r.idx - 6, -1), -1):
                    hit = [cl(x[2]) for x in rows[j].cells if cl(x[2]) and in_col(x, c[0], c[1])]
                    if hit:
                        above = hit[0]
                        break
                left.append(dict(row=r.idx, col=ci + 1, text=t[:160], left=prev[:60], above=above[:60]))
    res["unclaimed"] = left
    res["coverage"] = dict(total=tot, claimed=cl_n, template=fur, unmapped=len(left),
                           pct=round(100.0 * (cl_n + fur) / tot, 1) if tot else 100.0)
# ═══ END PARSER ═══


# ═════════════════════════════════════════════════════════════════════════════
#  3. DATA-INTEGRITY SHIELD  (Pydantic)  +  DIAGNOSTICS ENGINE  (Pandas)
# ═════════════════════════════════════════════════════════════════════════════
class CylinderIn(BaseModel):
    """One cylinder's confirmed readings. Anything that violates physics is rejected."""
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0, le=250)
    p_comp: float = Field(..., gt=0, le=220)
    fuel_index: float = Field(..., gt=0, lt=200)
    exh_temp: Optional[float] = Field(default=None, gt=0, le=700)
    cw_out: Optional[float] = Field(default=None, gt=0, le=130)
    piston_out: Optional[float] = Field(default=None, gt=0, le=130)

    @model_validator(mode="after")
    def _compression_below_firing(self):
        if self.p_comp >= self.p_max:
            raise ValueError("P comp must be lower than P max")
        return self


FIELDS = [  # key, header, unit, required
    ("p_max", "P max", "bar", True), ("p_comp", "P comp", "bar", True), ("fuel_index", "Fuel index", "", True),
    ("exh_temp", "Exhaust", "°C", False), ("cw_out", "Jacket CW", "°C", False), ("piston_out", "Piston oil", "°C", False),
]
FIELD_NAME = {k: h for k, h, _, _ in FIELDS}
CORE = ("p_max", "p_comp", "fuel_index")

TH_DEFAULT = dict(ratio_low=1.30, ratio_high=1.55, exh_spread=25.0, exh_abs=380.0, fuel_dev=4.0,
                  pmax_dev=5.0, pcomp_dev=5.0, cw_dev=4.0, piston_dev=4.0, cw_abs=90.0)


def validate_cylinders(vals: dict, n: int):
    """vals: {cyl_id: {field: float|None}}. Returns (list[CylinderIn], errors[list[str]], skipped[list[int]])."""
    ok, errs, skipped = [], [], []
    for cid in range(1, n + 1):
        v = vals.get(cid, {}) or {}
        if all(v.get(k) is None for k in CORE) and all(v.get(k) is None for k in ("exh_temp", "cw_out", "piston_out")):
            skipped.append(cid)
            continue
        missing = [FIELD_NAME[k] for k in CORE if v.get(k) is None]
        if missing:
            errs.append(f"Cylinder {cid}: missing {', '.join(missing)} - enter it, or clear the whole row to exclude this cylinder")
            continue
        try:
            ok.append(CylinderIn(id=cid, **{k: v.get(k) for k, *_ in FIELDS}))
        except ValidationError as e:
            for er in e.errors():
                loc = er["loc"][-1] if er.get("loc") else ""
                name = FIELD_NAME.get(str(loc), "")
                msg = er["msg"].replace("Value error, ", "")
                errs.append(f"Cylinder {cid}{' · ' + name if name else ''}: {msg}")
    if not ok and not errs:
        errs.append("No readings entered - provide P max, P comp and fuel index for at least one cylinder")
    return ok, errs, skipped


SEV_RANK = {"critical": 5, "mechanical": 4, "warning": 3, "advisory": 2, "nominal": 1}
SEV_LABEL = {"critical": "Critical", "mechanical": "Mechanical", "warning": "Warning", "advisory": "Advisory", "nominal": "Nominal"}
SEV_CHIP = {"critical": "crit", "mechanical": "orange", "warning": "warn", "advisory": "info", "nominal": "ok"}


def analyze(cyls, th: dict) -> dict:
    """Per-cylinder diagnostics + engine health score."""
    df = pd.DataFrame([c.model_dump() for c in cyls]).sort_values("id").reset_index(drop=True)
    for k in ("p_max", "p_comp", "fuel_index", "exh_temp", "cw_out", "piston_out"):
        df[k] = pd.to_numeric(df[k], errors="coerce")
    df["ratio"] = (df["p_max"] / df["p_comp"]).round(3)
    n = len(df)
    mean = {k: (df[k].mean() if df[k].notna().any() else None) for k in
            ("p_max", "p_comp", "fuel_index", "exh_temp", "cw_out", "piston_out", "ratio")}
    findings = []

    def add(sev, cyl, title, fault, action, rule):
        findings.append(dict(sev=sev, cyl=cyl, title=title, fault=fault, action=action, rule=rule))

    for _, r in df.iterrows():
        c, ratio = int(r["id"]), r["ratio"]
        # R1 / R2 - combustion ratio
        if ratio < th["ratio_low"]:
            add("critical", c, f"Cyl {c} · Poor combustion · ratio {ratio:.3f}",
                f"Pmax/Pcomp = {ratio:.3f} is below the {th['ratio_low']:.2f} limit (Pmax {r['p_max']:.0f} bar vs Pcomp {r['p_comp']:.0f} bar).",
                "Test and overhaul the fuel valve / atomizer (needle lift, seat, nozzle holes). Check fuel cam timing. If the fuel index is normal "
                "relative to sister cylinders suspect a fouled nozzle rather than a pump fault.", "R1")
        elif ratio > th["ratio_high"]:
            add("warning", c, f"Cyl {c} · Harsh combustion · ratio {ratio:.3f}",
                f"Pmax/Pcomp = {ratio:.3f} exceeds the {th['ratio_high']:.2f} limit.",
                "Check injection timing for early injection (VIT / fuel cam follower wear). Verify fuel quality and ignition quality; look for "
                "high-pressure pipe or fuel-valve dribbling.", "R2")
        if n < 3:
            continue
        # R3 - worn plunger signature
        fuel_hi = r["fuel_index"] > mean["fuel_index"] * (1 + th["fuel_dev"] / 100)
        pmax_lo = r["p_max"] < mean["p_max"] * (1 - th["pmax_dev"] / 100)
        worn = fuel_hi and pmax_lo
        if worn:
            add("mechanical", c, f"Cyl {c} · Worn fuel pump plunger suspected",
                f"Fuel index {r['fuel_index']:.1f} is {r['fuel_index'] / mean['fuel_index'] * 100 - 100:+.1f}% vs engine mean {mean['fuel_index']:.1f}, "
                f"while Pmax {r['p_max']:.0f} bar is {r['p_max'] / mean['p_max'] * 100 - 100:+.1f}% vs mean {mean['p_max']:.1f}: more rack for less pressure - the classic worn-plunger signature.",
                "Measure the plunger/barrel helix and clearance, compare with the last overhaul record, and schedule pump replacement. Check the suction "
                "valve and the fuel-valve opening pressure before condemning the pump.", "R3")
        # R4 - exhaust
        if r["exh_temp"] == r["exh_temp"] and r["exh_temp"] is not None and mean["exh_temp"] is not None:
            d = r["exh_temp"] - mean["exh_temp"]
            if abs(d) > th["exh_spread"]:
                hi = d > 0
                add("warning", c, f"Cyl {c} · Exhaust temperature {d:+.0f} °C vs mean",
                    f"{r['exh_temp']:.0f} °C against an engine mean of {mean['exh_temp']:.0f} °C (limit ±{th['exh_spread']:.0f} °C).",
                    ("Check exhaust valve seating / burn-through, injection timing and cylinder oil feed rate; confirm the thermocouple." if hi else
                     "Inspect the fuel valve for partial blockage; check for exhaust-valve or ring leakage (blow-by cools the exhaust); confirm the thermocouple."), "R4")
            if r["exh_temp"] > th["exh_abs"]:
                add("warning", c, f"Cyl {c} · High absolute exhaust temperature {r['exh_temp']:.0f} °C",
                    f"Exceeds the {th['exh_abs']:.0f} °C advisory limit.",
                    "Monitor closely; check exhaust valve condition, scavenge air temperature and cylinder lubrication.", "R4b")
        # R5 - Pmax balance (skip when plunger rule already explains it)
        dp = r["p_max"] / mean["p_max"] * 100 - 100
        if not worn and abs(dp) > th["pmax_dev"]:
            add("advisory", c, f"Cyl {c} · Pmax {dp:+.1f}% vs engine mean",
                f"{r['p_max']:.0f} bar against a mean of {mean['p_max']:.1f} bar (limit ±{th['pmax_dev']:.0f}%).",
                "Re-balance the cylinders (VIT / fuel index) at the next suitable load; verify the indicator reading with a second measurement.", "R5")
        # R6 - compression
        dc = r["p_comp"] / mean["p_comp"] * 100 - 100
        if dc < -th["pcomp_dev"]:
            add("warning", c, f"Cyl {c} · Low compression {dc:+.1f}% vs mean",
                f"Pcomp {r['p_comp']:.0f} bar against a mean of {mean['p_comp']:.1f} bar.",
                "Suspect exhaust-valve leakage or ring/liner blow-by; inspect rings through the scavenge ports and check the exhaust valve (leak test).", "R6")
        elif dc > th["pcomp_dev"]:
            add("advisory", c, f"Cyl {c} · High compression {dc:+.1f}% vs mean",
                f"Pcomp {r['p_comp']:.0f} bar against a mean of {mean['p_comp']:.1f} bar.",
                "Verify the indicator calibration; check for deposits or a changed compression volume (bottom-end bearing clearance).", "R6")
        # R7 - cooling
        if r["cw_out"] == r["cw_out"] and r["cw_out"] is not None and mean["cw_out"] is not None:
            d = r["cw_out"] - mean["cw_out"]
            if abs(d) > th["cw_dev"]:
                add("advisory", c, f"Cyl {c} · Jacket CW outlet {d:+.1f} °C vs mean",
                    f"{r['cw_out']:.0f} °C against a mean of {mean['cw_out']:.1f} °C.",
                    "Check the cylinder-jacket cooling flow (orifice / strainer / air lock) and the thermometer.", "R7")
            if r["cw_out"] > th["cw_abs"]:
                add("warning", c, f"Cyl {c} · Jacket CW outlet {r['cw_out']:.0f} °C is high", f"Above the {th['cw_abs']:.0f} °C limit.",
                    "Increase cooling flow / check jacket-water temperature control.", "R7b")
        if r["piston_out"] == r["piston_out"] and r["piston_out"] is not None and mean["piston_out"] is not None:
            d = r["piston_out"] - mean["piston_out"]
            if abs(d) > th["piston_dev"]:
                add("advisory", c, f"Cyl {c} · Piston cooling-oil outlet {d:+.1f} °C vs mean",
                    f"{r['piston_out']:.0f} °C against a mean of {mean['piston_out']:.1f} °C.",
                    "Check the piston cooling-oil flow (telescopic pipe / sight glass) and look for coke build-up.", "R7")

    findings.sort(key=lambda f: (-SEV_RANK[f["sev"]], f["cyl"]))
    # health score
    pen = {"critical": 15, "mechanical": 10, "warning": 6, "advisory": 2}
    penalty = sum(pen.get(f["sev"], 0) for f in findings)
    if mean["exh_temp"] is not None and df["exh_temp"].notna().sum() > 1:
        if df["exh_temp"].max() - df["exh_temp"].min() > 2 * th["exh_spread"]:
            penalty += 4
    score = 100.0 / (1.0 + penalty / 40.0)                      # saturating: many small findings never read as "dead engine"
    sevs = {f["sev"] for f in findings}
    if "critical" in sevs:
        score = min(score, 74.0)
    elif "mechanical" in sevs:
        score = min(score, 84.0)
    score = int(round(max(0, min(100, score))))
    grade = "Excellent" if score >= 85 else "Good" if score >= 70 else "Needs attention" if score >= 50 else "Critical"
    worst = {int(c): "nominal" for c in df["id"]}
    for f in findings:
        if f["cyl"] in worst and SEV_RANK[f["sev"]] > SEV_RANK[worst[f["cyl"]]]:
            worst[f["cyl"]] = f["sev"]
    df["status"] = df["id"].map(worst)
    if not findings:
        findings.append(dict(sev="nominal", cyl=0, title="All cylinders nominal",
                             fault="Combustion ratio, pressure balance, exhaust temperatures and cooling are within the configured limits.",
                             action="Continue routine monitoring; repeat the performance check per the planned-maintenance schedule.", rule="-"))
    spread = {}
    for k in ("p_max", "p_comp", "fuel_index", "exh_temp"):
        if df[k].notna().sum() > 1 and mean[k]:
            spread[k] = (df[k].max() - df[k].min())
    return dict(df=df, mean=mean, findings=findings, score=score, grade=grade, worst=worst, spread=spread, n=n)


def health_color(score: int) -> str:
    return "#5BB98C" if score >= 85 else "#C4A35A" if score >= 70 else "#E5B54A" if score >= 50 else "#E8645E"


# ═════════════════════════════════════════════════════════════════════════════
#  4. CHARTS
# ═════════════════════════════════════════════════════════════════════════════
C_GOLD, C_GOLD2, C_OK, C_WARN, C_ORANGE, C_CRIT, C_INFO, C_MUTED = (
    "#C4A35A", "#E2C888", "#5BB98C", "#E5B54A", "#EA8B4A", "#E8645E", "#7BAFD9", "#8E8B80")
STATUS_COLOR = {"nominal": C_OK, "advisory": C_INFO, "warning": C_WARN, "mechanical": C_ORANGE, "critical": C_CRIT}


def _style(fig, height=360, legend=True):
    fig.update_layout(
        height=height, margin=dict(l=12, r=12, t=36, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Barlow, sans-serif", color="#9C998E", size=12), hovermode="x unified",
        hoverlabel=dict(bgcolor="#14141B", bordercolor="rgba(196,163,90,.4)", font=dict(family="JetBrains Mono, monospace", color="#ECE8DD", size=12)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, bgcolor="rgba(0,0,0,0)", font=dict(size=11, color="#B8B4A8")) if legend else None,
        showlegend=legend, bargap=0.38)
    fig.update_xaxes(gridcolor="rgba(255,255,255,.04)", linecolor="rgba(255,255,255,.10)", zeroline=False, tickfont=dict(family="JetBrains Mono, monospace", size=11))
    fig.update_yaxes(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.10)", zeroline=False, tickfont=dict(family="JetBrains Mono, monospace", size=11))
    return fig


def plot(fig, key: str) -> None:
    st.plotly_chart(fig, key=key, config={"displayModeBar": False, "responsive": True}, **stretch(st.plotly_chart))


def _cats(df):
    return [str(int(i)) for i in df["id"]]


def chart_pressure(df, th):
    x = _cats(df)
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.08, row_heights=[0.68, 0.32],
                        subplot_titles=("Pressure profile · bar", "Fuel pump index"))
    lo, hi = (df["p_comp"] * th["ratio_low"]).tolist(), (df["p_comp"] * th["ratio_high"]).tolist()
    fig.add_trace(go.Scatter(x=x + x[::-1], y=hi + lo[::-1], fill="toself", fillcolor="rgba(91,185,140,.10)",
                             line=dict(color="rgba(91,185,140,.35)", width=1, dash="dot"), name=f"Optimal zone (ratio {th['ratio_low']:.2f}–{th['ratio_high']:.2f})",
                             hoverinfo="skip"), 1, 1)
    fig.add_trace(go.Scatter(x=x, y=df["p_comp"], name="P comp", mode="lines+markers", line=dict(color=C_INFO, width=2.4, dash="dot"),
                             marker=dict(size=9, color=C_INFO, line=dict(color="#09090D", width=2))), 1, 1)
    fig.add_trace(go.Scatter(x=x, y=df["p_max"], name="P max", mode="lines+markers", line=dict(color=C_GOLD2, width=3),
                             marker=dict(size=11, symbol="diamond", color=C_GOLD2, line=dict(color="#09090D", width=2))), 1, 1)
    pm = df["p_max"].mean()
    fig.add_hline(y=pm, line=dict(color=C_GOLD, width=1, dash="dash"), annotation_text=f"mean Pmax {pm:.1f}",
                  annotation_font=dict(color=C_GOLD, size=10), annotation_position="top left", row=1, col=1)
    fm = df["fuel_index"].mean()
    colors = [C_ORANGE if v > fm * (1 + th["fuel_dev"] / 100) else C_GOLD for v in df["fuel_index"]]
    fig.add_trace(go.Bar(x=x, y=df["fuel_index"], name="Fuel index", marker=dict(color=colors, opacity=.9), showlegend=False,
                         text=[f"{v:.1f}" for v in df["fuel_index"]], textposition="outside", textfont=dict(size=10, color="#9C998E")), 2, 1)
    fig.add_hline(y=fm, line=dict(color=C_MUTED, width=1, dash="dash"), row=2, col=1)
    _style(fig, 470)
    fig.update_xaxes(title_text="Cylinder", row=2, col=1)
    fig.update_yaxes(range=[max(0, min(df["fuel_index"]) * .7), max(df["fuel_index"]) * 1.12], row=2, col=1)
    fig.update_annotations(font=dict(family="Barlow Condensed, sans-serif", size=13, color="#8E8B80"), xanchor="left", x=0.0)
    return fig


def chart_ratio(df, th):
    x = _cats(df)
    cols = [C_CRIT if r < th["ratio_low"] else C_WARN if r > th["ratio_high"] else C_OK for r in df["ratio"]]
    fig = go.Figure()
    fig.add_hrect(y0=th["ratio_low"], y1=th["ratio_high"], fillcolor="rgba(91,185,140,.08)", line=dict(color="rgba(91,185,140,.3)", width=1, dash="dot"))
    fig.add_trace(go.Bar(x=x, y=df["ratio"], marker=dict(color=cols, opacity=.92), text=[f"{v:.3f}" for v in df["ratio"]],
                         textposition="outside", textfont=dict(size=11, color="#9C998E"), name="Pmax / Pcomp"))
    _style(fig, 330, legend=False)
    fig.update_yaxes(range=[min(1.0, df["ratio"].min() - .06), max(df["ratio"].max(), th["ratio_high"]) + .1], title_text="Pmax / Pcomp")
    fig.update_xaxes(title_text="Cylinder")
    return fig


def chart_exhaust(df, th):
    d = df.dropna(subset=["exh_temp"])
    if d.empty:
        return None
    x = _cats(d)
    mu = d["exh_temp"].mean()
    cols = [C_CRIT if abs(v - mu) > th["exh_spread"] else C_WARN if abs(v - mu) > th["exh_spread"] * .6 else C_OK for v in d["exh_temp"]]
    fig = go.Figure()
    fig.add_hrect(y0=mu - th["exh_spread"], y1=mu + th["exh_spread"], fillcolor="rgba(196,163,90,.06)", line=dict(width=0))
    fig.add_trace(go.Bar(x=x, y=d["exh_temp"], marker=dict(color=cols, opacity=.92), text=[f"{v:.0f}" for v in d["exh_temp"]],
                         textposition="outside", textfont=dict(size=11, color="#9C998E"), name="Exhaust °C"))
    fig.add_hline(y=mu, line=dict(color=C_GOLD, width=1.4, dash="dash"), annotation_text=f"mean {mu:.1f} °C", annotation_font=dict(color=C_GOLD, size=10))
    _style(fig, 330, legend=False)
    fig.update_yaxes(range=[d["exh_temp"].min() - 45, d["exh_temp"].max() + 35], title_text="°C")
    fig.update_xaxes(title_text="Cylinder")
    return fig


def chart_heatmap(df):
    rows = [("p_max", "P max"), ("p_comp", "P comp"), ("ratio", "Ratio"), ("fuel_index", "Fuel index"),
            ("exh_temp", "Exhaust"), ("cw_out", "Jacket CW"), ("piston_out", "Piston oil")]
    ylab, z, txt = [], [], []
    for k, lab in rows:
        s = df[k]
        if s.notna().sum() < 2:
            continue
        mu = s.mean()
        ylab.append(lab)
        z.append([(v / mu * 100 - 100) if v == v else None for v in s])
        txt.append([(f"{v:.3f}" if k == "ratio" else f"{v:.1f}") if v == v else "" for v in s])
    if not z:
        return None
    fig = go.Figure(go.Heatmap(z=z, x=_cats(df), y=ylab, text=txt, texttemplate="%{text}", zmid=0, zmin=-8, zmax=8,
                               colorscale=[[0, "#7BAFD9"], [.5, "#1B1B24"], [1, "#EA8B4A"]], xgap=4, ygap=4,
                               colorbar=dict(title=dict(text="% vs mean", font=dict(size=10)), thickness=10, len=.8, tickfont=dict(size=10)),
                               hovertemplate="Cyl %{x} · %{y}<br>%{z:+.1f}% vs engine mean<extra></extra>", textfont=dict(family="JetBrains Mono, monospace", size=11, color="#ECE8DD")))
    _style(fig, 60 + 44 * len(ylab), legend=False)
    fig.update_layout(hovermode="closest", margin=dict(l=10, r=10, t=10, b=10))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showgrid=False, title_text="Cylinder")
    return fig


def chart_cooling(df):
    has_cw, has_po = df["cw_out"].notna().any(), df["piston_out"].notna().any()
    if not (has_cw or has_po):
        return None
    x = _cats(df)
    fig = go.Figure()
    if has_cw:
        fig.add_trace(go.Bar(x=x, y=df["cw_out"], name="Jacket CW outlet", marker=dict(color=C_INFO, opacity=.9)))
        fig.add_hline(y=df["cw_out"].mean(), line=dict(color=C_INFO, width=1, dash="dot"))
    if has_po:
        fig.add_trace(go.Bar(x=x, y=df["piston_out"], name="Piston oil outlet", marker=dict(color=C_ORANGE, opacity=.9)))
        fig.add_hline(y=df["piston_out"].mean(), line=dict(color=C_ORANGE, width=1, dash="dot"))
    _style(fig, 320)
    fig.update_layout(barmode="group", bargap=0.3)
    fig.update_yaxes(title_text="°C")
    fig.update_xaxes(title_text="Cylinder")
    return fig


def chart_fleet_health(rows):
    names = [r["label"] for r in rows]
    scores = [r["score"] for r in rows]
    fig = go.Figure(go.Bar(x=scores, y=names, orientation="h", marker=dict(color=[health_color(s) for s in scores], opacity=.92),
                           text=[f"{s}" for s in scores], textposition="outside", textfont=dict(size=11, color="#9C998E")))
    _style(fig, 90 + 46 * len(rows), legend=False)
    fig.update_layout(hovermode="closest")
    fig.update_xaxes(range=[0, 108], title_text="Health score")
    fig.update_yaxes(autorange="reversed")
    return fig


def chart_trend(series):
    """series: {metric: [(date, value)]} for one vessel."""
    fig = make_subplots(rows=2, cols=2, subplot_titles=("Mean Pmax · bar", "Mean Pcomp · bar", "Mean fuel index", "Mean exhaust · °C"),
                        vertical_spacing=0.16, horizontal_spacing=0.07)
    pos = {"p_max": (1, 1), "p_comp": (1, 2), "fuel_index": (2, 1), "exh_temp": (2, 2)}
    for k, (r, c) in pos.items():
        pts = series.get(k, [])
        if pts:
            fig.add_trace(go.Scatter(x=[p[0] for p in pts], y=[p[1] for p in pts], mode="lines+markers", showlegend=False,
                                     line=dict(color=C_GOLD2, width=2.6), marker=dict(size=9, color=C_GOLD2, line=dict(color="#09090D", width=2))), r, c)
    _style(fig, 430, legend=False)
    fig.update_annotations(font=dict(family="Barlow Condensed", size=13, color="#8E8B80"), xanchor="left")
    for ann in fig.layout.annotations:
        ann.x = ann.x - 0.0
    return fig


# ═════════════════════════════════════════════════════════════════════════════
#  5. SESSION STATE  ("smart memory flush")
# ═════════════════════════════════════════════════════════════════════════════
def init_state() -> None:
    ss = st.session_state
    ss.setdefault("reports", {})
    ss.setdefault("order", [])
    ss.setdefault("uploader_gen", 0)
    ss.setdefault("auto_confirm", False)
    ss.setdefault("th", dict(TH_DEFAULT))


@st.cache_data(show_spinner=False, max_entries=64)
def parse_cached(data: bytes):
    try:
        return parse_tec_005(data), None
    except ParseError as e:
        return None, f"{e}. The file does not appear to be a Word 97-2003 (.doc) document."
    except Exception as e:                                                    # noqa: BLE001
        return None, f"Unexpected problem while reading the document ({type(e).__name__}). The file may be damaged."


def new_report(name: str, data: bytes, rid: str) -> dict:
    res, err = parse_cached(data)
    vals = {i: {k: None for k, *_ in FIELDS} for i in range(1, 10)}
    n = 0
    if res:
        n = int(res["n_cylinders"])
        for c in res["cylinders"]:
            vals[c["id"]] = {k: c.get(k) for k, *_ in FIELDS}
    return dict(id=rid, name=name, size=len(data), parsed=res, error=err, n=n, vals=vals,
                confirmed=False, errors=[], skipped=[], auto_tried=False, added=time.time())


def flush_report(rid: str) -> None:
    """Drop a report and every widget key that belongs to it."""
    ss = st.session_state
    ss["reports"].pop(rid, None)
    for k in [k for k in ss.keys() if isinstance(k, str) and k.startswith(f"w_{rid}_")]:
        del ss[k]
    if rid in ss["order"]:
        ss["order"].remove(rid)


def sync_reports(files) -> None:
    ss = st.session_state
    seen = []
    for f in files or []:
        data = f.getvalue()
        rid = hashlib.sha1(f.name.encode("utf-8", "ignore") + data).hexdigest()[:12]
        if rid in seen:
            continue
        seen.append(rid)
        if rid not in ss["reports"]:
            ss["reports"][rid] = new_report(f.name, data, rid)
    for rid in list(ss["reports"].keys()):
        if rid not in seen:
            flush_report(rid)                       # file removed from the uploader -> memory flush
    ss["order"] = seen


def clear_all() -> None:
    ss = st.session_state
    for rid in list(ss["reports"].keys()):
        flush_report(rid)
    ss["order"] = []
    ss["uploader_gen"] += 1


def reset_thresholds() -> None:
    for k, v in TH_DEFAULT.items():
        st.session_state[f"th_{k}"] = v


def rep_label(rep: dict) -> str:
    p = rep.get("parsed")
    if not p:
        return rep["name"]
    v = p["identity"].get("vessel") or rep["name"]
    d = p["identity"].get("date")
    return f"{v} · {d.strftime('%d %b %y')}" if d else v


def get_analysis(rep: dict, th: dict):
    """Analysis on confirmed values; otherwise (for the fleet view) on raw extracted values."""
    if not rep.get("parsed"):
        return None, False
    cyls, errs, _ = validate_cylinders(rep["vals"], rep["n"])
    if not cyls or (rep["confirmed"] and errs):
        return None, rep["confirmed"]
    try:
        return analyze(cyls, th), rep["confirmed"]
    except Exception:                                                         # noqa: BLE001
        return None, rep["confirmed"]


# ═════════════════════════════════════════════════════════════════════════════
#  6. UI BUILDING BLOCKS
# ═════════════════════════════════════════════════════════════════════════════

def ring_svg(pct: float, color: str, big: str, small: str, size: int = 124) -> str:
    r = 48
    c = 2 * math.pi * r
    off = c * (1 - max(0.0, min(100.0, pct)) / 100)
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 120 120">'
            f'<circle cx="60" cy="60" r="{r}" fill="none" stroke="rgba(255,255,255,.07)" stroke-width="9"/>'
            f'<circle cx="60" cy="60" r="{r}" fill="none" stroke="{color}" stroke-width="9" stroke-linecap="round" '
            f'stroke-dasharray="{c:.1f}" stroke-dashoffset="{off:.1f}" transform="rotate(-90 60 60)"/>'
            f'<text x="60" y="61" text-anchor="middle" fill="#ECE8DD" font-family="JetBrains Mono, monospace" font-size="25" font-weight="500">{esc(big)}</text>'
            f'<text x="60" y="79" text-anchor="middle" fill="#8E8B80" font-family="Barlow Condensed, sans-serif" font-size="9" letter-spacing="1.4">{esc(small)}</text></svg>')


def chip(text, tone="") -> str:
    return f'<span class="chip {tone}">{esc(text)}</span>'


def kpi(label, value, unit="", sub="", tone="") -> str:
    u = f"<small>{esc(unit)}</small>" if unit else ""
    s = f'<div class="s">{esc(sub)}</div>' if sub else ""
    return f'<div class="kpi {tone}"><div class="k">{esc(label)}</div><div class="v">{esc(value)}{u}</div>{s}</div>'


def section(no: str, title: str, sub: str = "") -> None:
    md(f'<div class="sec"><span class="no">{esc(no)}</span><span class="ti">{esc(title)}</span><span class="su">{esc(sub)}</span></div>')


def fact(label, value) -> str:
    return f'<div><p class="lbl">{esc(label)}</p><p class="val">{esc(value) if value not in (None, "") else "—"}</p></div>'


def table_html(headers, rows) -> str:
    th = "".join(f"<th>{h}</th>" for h in headers)
    body = "".join("<tr>" + "".join(f'<td class="{"t" if i == 0 else ""}">{c}</td>' for i, c in enumerate(r)) + "</tr>" for r in rows)
    return f'<div class="tbl-wrap"><table class="tbl"><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table></div>'


def render_header() -> None:
    n = len(st.session_state["reports"])
    right = chip(f"v{APP_VERSION}", "gold") + " " + (chip(f"{n} report{'s' if n != 1 else ''} loaded", "info") if n else chip("awaiting data"))
    md(f'<div class="hdr"><div class="hdr-l"><div class="logo">ME</div><div><div class="brand">M.E. COMMAND CENTER</div>'
       f'<div class="tag">Marine diesel performance intelligence · TEC-005</div></div></div><div>{right}</div></div>')


def render_steps(step: int) -> None:
    names = ["Upload", "Verify", "Analyze"]
    out = []
    for i, nm in enumerate(names, 1):
        cls = "done" if i < step else "on" if i == step else ""
        out.append(f'<div class="step {cls}"><span class="n">{"✓" if i < step else i}</span><span class="t">{nm}</span></div>')
    md('<div class="steps">' + "".join(out) + "</div>")


def render_empty() -> None:
    md("""<div class="empty"><div class="big">Drop your performance sheets</div>
    <div class="sm">Drag one or many legacy TEC-005 <b>.doc</b> files into the zone above. No conversion, no cloud, no external services - the sheets are read in memory, checked against their own printed averages and analysed in seconds.</div></div>
    <div class="feat">
    <div class="card"><div class="fi">🧭</div><div class="fh">Layout-agnostic reading</div><div class="fd">Understands the Word table geometry, so shifted columns, merged cells and values typed into header cells still land in the right cylinder.</div></div>
    <div class="card"><div class="fi">🛡️</div><div class="fh">Integrity shield</div><div class="fd">Every value carries its provenance, is reconciled with the sheet's Average cells and must pass physical-limit validation before analysis.</div></div>
    <div class="card"><div class="fi">🔧</div><div class="fh">5 · 6 · 7+ cylinders</div><div class="fd">Cylinder count is reconciled between the engine type and the data - blank cylinders are never silently turned into zeros.</div></div>
    <div class="card"><div class="fi">📊</div><div class="fh">Fleet intelligence</div><div class="fd">Load many vessels at once: health scores, alert counts and month-to-month trends in one overview.</div></div>
    </div>""")


def render_sidebar() -> None:
    ss = st.session_state
    with st.sidebar:
        md('<div style="display:flex;align-items:center;gap:11px;margin-bottom:.6rem"><div class="logo" style="width:36px;height:36px;font-size:1.05rem;border-radius:10px">ME</div>'
           '<div><div class="brand" style="font-size:1.15rem">COMMAND CENTER</div><div class="tag" style="margin-top:2px;font-size:.62rem">TEC-005 · v' + APP_VERSION + '</div></div></div>')
        ids = ss["order"]
        if ids:
            if ss.get("active") not in ids:
                ss["active"] = ids[-1]
            if len(ids) > 1:
                st.radio("Active report", ids, format_func=lambda i: rep_label(ss["reports"][i]), key="active")
            else:
                md(f'<p class="lbl">Active report</p><p class="val">{esc(rep_label(ss["reports"][ids[0]]))}</p>')
                ss["active"] = ids[0]
        st.toggle("Auto-confirm verified sheets", key="auto_confirm",
                  help="Sheets whose extraction reconciles fully with their own printed averages skip the manual confirmation step. "
                       "You can still edit and re-confirm any report.")
        with st.expander("Diagnostic thresholds"):
            th = {}
            specs = [("ratio_low", "Ratio · poor combustion below", .01, "%.2f"), ("ratio_high", "Ratio · harsh combustion above", .01, "%.2f"),
                     ("exh_spread", "Exhaust deviation limit (±°C)", 1.0, "%.0f"), ("exh_abs", "Exhaust absolute limit (°C)", 5.0, "%.0f"),
                     ("fuel_dev", "Fuel index high (% over mean)", .5, "%.1f"), ("pmax_dev", "Pmax deviation (± %)", .5, "%.1f"),
                     ("pcomp_dev", "Pcomp deviation (± %)", .5, "%.1f"), ("cw_dev", "Jacket CW deviation (± °C)", .5, "%.1f"),
                     ("piston_dev", "Piston oil deviation (± °C)", .5, "%.1f"), ("cw_abs", "Jacket CW absolute (°C)", 1.0, "%.0f")]
            for k, lab, step, fmt in specs:
                th[k] = st.number_input(lab, value=float(TH_DEFAULT[k]), step=step, format=fmt, key=f"th_{k}")
            ss["th"] = th
            st.button("Reset to defaults", on_click=reset_thresholds, key="reset_th")
            st.caption("Limits are fixed, not load-aware. Judge part-load sheets (< 50 %) against their own history.")
        if ids:
            st.button("Clear all reports", on_click=clear_all, key="clear_all")
        md('<p class="note" style="margin-top:1.2rem">All processing happens in memory in this session. Nothing is stored or sent anywhere.</p>')


# ═════════════════════════════════════════════════════════════════════════════
#  7. REPORT VIEW
# ═════════════════════════════════════════════════════════════════════════════
OP_LABELS = [("draught_fore", "Draught fore", "m", 2), ("draught_aft", "Draught aft", "m", 2), ("run_hours", "Total running hours", "h", 0),
             ("rpm", "Engine RPM", "", 1), ("load_pct", "Load", "%", 1), ("kw_indic", "kW indicated", "", 0), ("kw_eff", "kW effective", "", 0),
             ("bhp_op", "BHP", "", 0), ("governor_index", "Governor index", "", 1), ("speed_setting", "Speed setting", "", 1),
             ("log_knots", "Speed (log)", "kn", 1), ("obs_knots", "Speed (obs)", "kn", 1), ("wind", "Wind", "", 1), ("wave", "Wave height", "m", 1),
             ("barometer", "Barometer", "mbar", 0), ("sfoc_kwh", "SFOC", "g/kWh", 1)]
SINGLE_LABELS = [("scav_bar", "Scavenge pressure", "bar", 2), ("scav_temp", "Scavenge air temp", "°C", 0), ("sea_water_t", "Sea water temp", "°C", 0),
                 ("cw_me_in", "Jacket CW inlet", "°C", 0), ("cw_air_in", "Air cooler CW in", "°C", 0), ("cw_air_out", "Air cooler CW out", "°C", 0),
                 ("lo_system_p", "System oil pressure", "bar", 2), ("lo_inlet_t", "Lube oil inlet temp", "°C", 0), ("lo_cooling_p", "Cooling oil pressure", "bar", 2),
                 ("lo_exvalve_p", "Ex.valve oil pressure", "bar", 2), ("lo_turbine_p", "T/C oil pressure", "bar", 2), ("tc_oil_in", "T/C oil inlet temp", "°C", 0),
                 ("tc_oil_out", "T/C oil outlet temp", "°C", 0), ("fo_before_filter", "FO pressure before filter", "bar", 2),
                 ("fo_after_filter", "FO pressure after filter", "bar", 2), ("fo_temp", "FO temperature", "°C", 0)]
TC_LABELS = [("tc_rpm", "T/C RPM", 0), ("turb_in", "Turbine inlet °C", 0), ("turb_out", "Turbine outlet °C", 0), ("exh_press", "Exh. receiver pressure", 2),
             ("dp_filter", "Δp air filter", 1), ("dp_cooler", "Δp air cooler", 2), ("inl_blower", "Blower inlet °C", 0),
             ("bef_cooler", "Air before cooler °C", 0), ("aft_cooler", "Air after cooler °C", 0),
             ("spiral_p", "Spiral housing pressure", 2), ("dp_inout", "T/C Δp in/out", 1), ("cw_turb_out", "CW turbine outlet °C", 0)]


def render_report_card(rep: dict) -> None:
    p, idn, conf = rep["parsed"], rep["parsed"]["identity"], rep["parsed"]["confidence"]
    op = p["op"]
    level = conf["level"]
    tone = {"verified": "#5BB98C", "high": "#7BAFD9", "review": "#E5B54A", "low": "#E8645E", "blank": "#C4A35A"}[level]
    total = conf["total"]
    ring = ring_svg(total, tone, f"{total:.0f}" if level != "blank" else "—", "CONFIDENCE")
    xv_note = "" if conf.get("verified_by_sheet") else " (no sheet averages to check)"

    def bar(label, frac, extra=""):
        return (f'<div class="bar-row"><div class="bl"><span>{label}</span><b>{frac * 100:.0f}%{extra}</b></div>'
                f'<div class="track"><div class="fill" style="width:{frac * 100:.0f}%"></div></div></div>')

    bars = (bar("Structure located", conf["structure"]) + bar("Completeness", conf["completeness"]) + bar("Physical plausibility", conf["plausibility"])
            + bar("Sheet cross-check", conf["crossval"], xv_note))
    n_note = f'{p["n_cylinders"]}' + (f' (type says {p["n_type"]})' if p.get("n_type") and p["n_type"] != p["n_cylinders"] else "")
    power = f'{idn["rated_power"]:,.0f} {idn.get("power_unit", "")}' + (f' @ {idn["rated_rpm"]:.0f} rpm' if idn.get("rated_rpm") else "") if idn.get("rated_power") else None
    fuel = " · ".join(x for x in [f'ρ {idn["fuel_density"]}' if idn.get("fuel_density") else "", f'S {idn["fuel_sulphur"]} %' if idn.get("fuel_sulphur") else ""] if x)
    test = " · ".join(x for x in [f'{op["rpm"]:.0f} rpm' if op.get("rpm") else "", f'{op["load_pct"]:.0f} % load' if op.get("load_pct") else ""] if x)
    facts = (fact("Report date", fdate(idn.get("date"))) + fact("Cylinders", n_note) + fact("Rated output", power) + fact("Test condition", test)
             + fact("Engine no.", idn.get("engine_no")) + fact("Running hours", f'{op["run_hours"]:,.0f}' if op.get("run_hours") else None)
             + fact("Fuel", fuel) + fact("Checked by", idn.get("checked_by"))
             + fact("Sheet coverage", (f'{p["coverage"]["pct"]:.0f} % of cells mapped' + (f' · {p["coverage"]["unmapped"]} not mapped' if p["coverage"]["unmapped"] else "")) if p.get("coverage") else None))
    sub = " · ".join(x for x in [idn.get("engine_type"), idn.get("builder"), (f'yard {idn["yard"]}' if idn.get("yard") else ""),
                                  (f'built {idn["built_year"]}' if idn.get("built_year") else "")] if x)
    md(f'<div class="card gold"><div class="vs"><div><p class="lbl">Vessel</p><div class="vname">{esc(idn.get("vessel") or rep["name"])}</div>'
       f'<div class="vsub">{esc(sub)}</div><div class="facts">{facts}</div></div>'
       f'<div class="ring-wrap"><div>{ring}</div><div class="bars">{bars}</div></div></div></div>')


def render_banner(rep: dict) -> None:
    p = rep["parsed"]
    conf, issues = p["confidence"], p["issues"]
    warns = [m for s, m in issues if s == "warn"]
    n_ok = sum(1 for v in p.get("xval", {}).values() if v["label"] == "match")
    n_x = len(p.get("xval", {}))
    lvl = conf["level"]
    if lvl == "blank":
        rem = p.get("remarks")
        md(f'<div class="banner blank"><div class="ic">📄</div><div><div class="bt">Blank form - no measurements entered</div>'
           f'<div class="bd">This sheet contains no per-cylinder performance data{(" · crew remark: " + esc(rem)) if rem else ""}. '
           f'You can still type readings manually below.</div></div></div>')
    elif lvl == "verified":
        md(f'<div class="banner ok"><div class="ic">✅</div><div><div class="bt">Extraction verified</div><div class="bd">'
           f'Every cylinder reading reconciles with the averages printed on the sheet ({n_ok} of {n_x} cross-checks matched exactly). Please confirm below.</div></div></div>')
    elif lvl == "high":
        extra = f" · {esc(warns[0])}" if warns else ""
        md(f'<div class="banner high"><div class="ic">ℹ️</div><div><div class="bt">High confidence</div><div class="bd">'
           f'Values were extracted cleanly; {n_ok} of {n_x} cross-checks matched exactly{extra}. See the extraction audit for details.</div></div></div>')
    elif lvl == "review":
        items = "".join(f"<br>• {esc(w)}" for w in warns[:3]) or "<br>• Some values could not be reconciled with the sheet."
        md(f'<div class="banner review"><div class="ic">⚠️</div><div><div class="bt">Review recommended</div><div class="bd">Please compare the pre-filled values with the sheet.{items}</div></div></div>')
    else:
        items = "".join(f"<br>• {esc(w)}" for w in warns[:3])
        md(f'<div class="banner low"><div class="ic">⛔</div><div><div class="bt">Low confidence - manual verification required</div><div class="bd">'
           f'The extraction could not be reconciled with the sheet.{items}</div></div></div>')
    infos_blank = [m for s, m in issues if s == "error"]
    for m in infos_blank:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Structure error</div><div class="bd">{esc(m)}</div></div></div>')


def _prov_note(p: dict, cid: int) -> str:
    cy = p["cylinders"][cid - 1] if cid <= len(p["cylinders"]) else None
    if cy is None:
        return '<span class="note">manual entry</span>'
    srcs = [cy.get("src_" + k) for k in CORE]
    vals = [cy.get(k) for k in CORE]
    if all(v is None for v in vals):
        return '<span class="note" style="color:#E5B54A">no readings on sheet</span>'
    notes = []
    if any(s == "embedded" for s in srcs):
        notes.append('<span class="note" style="color:#E5B54A">typed in header</span>')
    miss = [FIELD_NAME[k] for k, v in zip(CORE, vals) if v is None]
    if miss:
        notes.append(f'<span class="note" style="color:#E8645E">missing {esc(", ".join(miss))}</span>')
    return " · ".join(notes)


def render_form(rep: dict) -> None:
    ss, p, rid = st.session_state, rep["parsed"], rep["id"]
    c1, c2 = st.columns([1, 3])
    with c1:
        n = st.number_input("Cylinders", min_value=3, max_value=9, value=int(rep["n"]), step=1, key=f"w_{rid}_n")
    rep["n"] = int(n)
    with c2:
        md('<div class="legend" style="margin-top:1.55rem"><span><i class="dot" style="background:#C4A35A"></i>read from its own cell</span>'
           '<span><i class="dot" style="background:#E5B54A"></i>typed in a header cell</span><span><i class="dot" style="background:#E8645E"></i>missing / not on sheet</span>'
           '<span class="note">Clear a whole row to exclude that cylinder.</span></div>')
    widths = [0.8] + [1] * len(FIELDS)
    new_vals = {}
    with st.form(f"form_{rid}"):
        hc = st.columns(widths)
        hc[0].markdown(H('<div class="fh-row">Cyl</div>'), unsafe_allow_html=True)
        for j, (k, head, unit, req) in enumerate(FIELDS):
            hc[j + 1].markdown(H(f'<div class="fh-row">{esc(head)}{" · " + esc(unit) if unit else ""}</div>'), unsafe_allow_html=True)
        for cid in range(1, rep["n"] + 1):
            cols = st.columns(widths)
            cols[0].markdown(H(f'<div style="display:flex;flex-direction:column;gap:3px"><div class="cyl-chip">{cid}</div>{_prov_note(p, cid)}</div>'), unsafe_allow_html=True)
            row = {}
            for j, (k, head, unit, req) in enumerate(FIELDS):
                cur = rep["vals"].get(cid, {}).get(k)
                row[k] = cols[j + 1].number_input(f"{head} cyl {cid}", value=(float(cur) if cur is not None else None), min_value=0.0, max_value=1000.0,
                                                   step=1.0, format="%.1f", key=f"w_{rid}_{k}_{cid}", label_visibility="collapsed", placeholder="—")
            new_vals[cid] = row
        ad = p.get("avgs_doc", {})
        ac = st.columns(widths)
        ac[0].markdown(H('<div class="fh-row" style="padding-top:.55rem;color:#C4A35A">Sheet avg</div>'), unsafe_allow_html=True)
        for j, (k, *_r) in enumerate(FIELDS):
            v = ad.get(k)
            ac[j + 1].markdown(H(f'<div class="mono" style="color:#8E8B80;padding-top:.45rem;font-size:.9rem">{(f"{v[0]:.{v[1]}f}" if v else "—")}</div>'), unsafe_allow_html=True)
        md('<div style="height:.5rem"></div>')
        submitted = st.form_submit_button("Confirm & run analysis", **stretch(st.form_submit_button))
    if submitted:
        for cid, row in new_vals.items():
            rep["vals"][cid] = row
        _ok, errs, skipped = validate_cylinders(rep["vals"], rep["n"])
        rep["errors"], rep["skipped"] = errs, skipped
        rep["confirmed"] = not errs
    if rep["errors"]:
        items = "".join(f"<br>• {esc(e)}" for e in rep["errors"])
        md(f'<div class="banner err"><div class="ic">🛡️</div><div><div class="bt">Data-integrity shield - analysis blocked</div><div class="bd">Fix the following and confirm again:{items}</div></div></div>')


def render_findings(an: dict) -> None:
    counts = {}
    for f in an["findings"]:
        counts[f["sev"]] = counts.get(f["sev"], 0) + 1
    chips = " ".join(chip(f"{counts[s]} {SEV_LABEL[s]}", SEV_CHIP[s]) for s in ("critical", "mechanical", "warning", "advisory", "nominal") if s in counts)
    md(f'<div style="margin-bottom:.7rem">{chips}</div>')
    for f in an["findings"]:
        md(f'<div class="fcard {f["sev"]}"><div class="ft"><span class="fh">{esc(f["title"])}</span>{chip(SEV_LABEL[f["sev"]], SEV_CHIP[f["sev"]])}</div>'
           f'<div class="ff">{esc(f["fault"])}</div><div class="fa"><b>Action</b>{esc(f["action"])}</div></div>')


def cylinder_table(an: dict) -> str:
    df = an["df"]

    def c(v, nd=1):
        return f"{v:.{nd}f}" if v == v and v is not None else '<span class="dim">—</span>'

    rows = []
    for _, r in df.iterrows():
        st_ = r["status"]
        rows.append([f'<span class="cyl-chip" style="min-width:30px;height:28px;font-size:.85rem">{int(r["id"])}</span>', c(r["p_max"], 1), c(r["p_comp"], 1),
                     f'{r["ratio"]:.3f}', c(r["fuel_index"], 1), c(r["exh_temp"], 0), c(r["cw_out"], 1), c(r["piston_out"], 1), chip(SEV_LABEL[st_], SEV_CHIP[st_])])
    m = an["mean"]
    rows.append(['<b style="color:#C4A35A;font-family:var(--disp);letter-spacing:.14em">MEAN</b>', c(m["p_max"]), c(m["p_comp"]), c(m["ratio"], 3), c(m["fuel_index"]),
                 c(m["exh_temp"], 1), c(m["cw_out"]), c(m["piston_out"]), ""])
    return table_html(["Cyl", "P max", "P comp", "Ratio", "Fuel idx", "Exhaust °C", "Jacket CW °C", "Piston oil °C", "Status"], rows)


def render_results(rep: dict, an: dict) -> None:
    rid, df, m = rep["id"], an["df"], an["mean"]
    sc = an["score"]
    col = health_color(sc)
    tone = "ok" if sc >= 85 else "" if sc >= 70 else "warn" if sc >= 50 else "crit"
    spread_e = an["spread"].get("exh_temp")
    cards = [
        kpi("Engine health", f"{sc}", "/100", an["grade"], tone),
        kpi("Mean Pmax", fnum(m["p_max"]), "bar", f"spread {fnum(an['spread'].get('p_max'))} bar" if "p_max" in an["spread"] else ""),
        kpi("Mean Pcomp", fnum(m["p_comp"]), "bar", f"spread {fnum(an['spread'].get('p_comp'))} bar" if "p_comp" in an["spread"] else ""),
        kpi("Mean ratio", fnum(m["ratio"], 3), "", f"window {st.session_state['th']['ratio_low']:.2f}–{st.session_state['th']['ratio_high']:.2f}",
            "ok" if st.session_state["th"]["ratio_low"] <= m["ratio"] <= st.session_state["th"]["ratio_high"] else "warn"),
        kpi("Mean fuel index", fnum(m["fuel_index"]), "", f"spread {fnum(an['spread'].get('fuel_index'))}" if "fuel_index" in an["spread"] else ""),
        kpi("Mean exhaust", fnum(m["exh_temp"], 0) if m["exh_temp"] is not None else "—", "°C" if m["exh_temp"] is not None else "",
            f"spread {fnum(spread_e, 0)} °C" if spread_e is not None else "no exhaust data"),
    ]
    md('<div class="kpis">' + "".join(cards) + "</div>")
    section("02", "Diagnostic findings", f"{an['n']} cylinders analysed")
    render_findings(an)
    section("03", "Charts", "optimal zone = Pcomp × ratio window")
    plot(chart_pressure(df, st.session_state["th"]), f"{rid}_pressure")
    a, b = st.columns(2)
    with a:
        md('<p class="lbl">Combustion ratio</p>')
        plot(chart_ratio(df, st.session_state["th"]), f"{rid}_ratio")
    with b:
        md('<p class="lbl">Exhaust gas temperature</p>')
        fe = chart_exhaust(df, st.session_state["th"])
        if fe:
            plot(fe, f"{rid}_exh")
        else:
            md('<div class="card"><p class="note">No exhaust temperatures available for this report.</p></div>')
    fh = chart_heatmap(df)
    if fh:
        md('<p class="lbl" style="margin-top:.8rem">Deviation heat-map · every parameter versus the engine mean</p>')
        plot(fh, f"{rid}_heat")
    fc = chart_cooling(df)
    if fc:
        md('<p class="lbl" style="margin-top:.8rem">Cooling outlets per cylinder</p>')
        plot(fc, f"{rid}_cool")
    section("04", "Cylinder summary")
    md(cylinder_table(an))


def export_csv(an: dict) -> bytes:
    d = an["df"].rename(columns={"id": "cylinder", "p_max": "p_max_bar", "p_comp": "p_comp_bar", "exh_temp": "exhaust_c", "cw_out": "jacket_cw_out_c",
                                 "piston_out": "piston_oil_out_c"})
    return d.to_csv(index=False).encode("utf-8")


def export_json(rep: dict, an: dict) -> bytes:
    p = rep["parsed"]
    doc = dict(generator=f"M.E. Command Center v{APP_VERSION}", file=rep["name"], generated=time.strftime("%Y-%m-%d %H:%M:%S"),
               identity=p["identity"], operating=p["op"], confirmed=rep["confirmed"], confidence=p["confidence"], cross_checks=p.get("xval", {}),
               sheet_averages={k: v[0] for k, v in p.get("avgs_doc", {}).items()}, cylinders=an["df"].where(an["df"].notna(), None).to_dict(orient="records"),
               health_score=an["score"], findings=an["findings"], turbocharger=p.get("tc", {}), auxiliary=p.get("singles", {}), remarks=p.get("remarks"), header_details=p.get("header_extra"), tc_specification=p.get("tc_info"), remarks_series=p.get("remarks_series"),
               coverage=p.get("coverage"), unmapped_cells=p.get("unclaimed"), cylinders_raw=p.get("cylinders"),
               notes=[dict(severity=s, message=m) for s, m in p["issues"]])
    return json.dumps(doc, indent=2, default=str, ensure_ascii=False).encode("utf-8")


def export_html(rep: dict, an: dict) -> bytes:
    idn = rep["parsed"]["identity"]
    rows = "".join(
        f"<tr><td>{int(r['id'])}</td><td>{fnum(r['p_max'])}</td><td>{fnum(r['p_comp'])}</td><td>{r['ratio']:.3f}</td><td>{fnum(r['fuel_index'])}</td>"
        f"<td>{fnum(r['exh_temp'], 0)}</td><td>{fnum(r['cw_out'])}</td><td>{fnum(r['piston_out'])}</td><td>{SEV_LABEL[r['status']]}</td></tr>" for _, r in an["df"].iterrows())
    fnd = "".join(f"<div class='f {f['sev']}'><b>{esc(f['title'])}</b><br>{esc(f['fault'])}<br><i>Action: {esc(f['action'])}</i></div>" for f in an["findings"])
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>Performance report - {esc(idn.get('vessel'))}</title><style>
body{{font-family:Arial,Helvetica,sans-serif;color:#1b1b1b;margin:32px auto;max-width:900px}}h1{{margin:0;letter-spacing:.04em}}.sub{{color:#666;margin-bottom:18px}}
table{{border-collapse:collapse;width:100%;margin:12px 0}}th,td{{border:1px solid #ccc;padding:6px 9px;text-align:right}}th{{background:#f3efe2}}td:first-child,th:first-child{{text-align:center}}
.f{{border-left:5px solid #999;background:#fafafa;padding:8px 12px;margin:8px 0}}.f.critical{{border-color:#d33}}.f.warning{{border-color:#d9a500}}.f.mechanical{{border-color:#e07b2a}}.f.advisory{{border-color:#4a8ac4}}.f.nominal{{border-color:#2f9e6b}}
.k{{display:inline-block;margin-right:22px}}</style></head><body>
<h1>{esc(idn.get('vessel') or rep['name'])} - main engine performance</h1>
<div class="sub">{esc(idn.get('engine_type'))} · report date {fdate(idn.get('date'))} · source file {esc(rep['name'])} · generated by M.E. Command Center v{APP_VERSION}</div>
<div><span class="k"><b>Health score</b> {an['score']}/100 ({an['grade']})</span><span class="k"><b>Mean Pmax</b> {fnum(an['mean']['p_max'])} bar</span><span class="k"><b>Mean Pcomp</b> {fnum(an['mean']['p_comp'])} bar</span><span class="k"><b>Mean ratio</b> {fnum(an['mean']['ratio'], 3)}</span></div>
<table><tr><th>Cyl</th><th>Pmax</th><th>Pcomp</th><th>Ratio</th><th>Fuel idx</th><th>Exhaust °C</th><th>Jacket CW °C</th><th>Piston oil °C</th><th>Status</th></tr>{rows}</table>
<h3>Findings</h3>{fnd}
<p style="color:#888;font-size:12px">Limits are fixed and not load-aware. Values were confirmed by the operator before analysis.</p></body></html>"""
    return page.encode("utf-8")


def render_audit(rep: dict) -> None:
    p, rid = rep["parsed"], rep["id"]
    conf = p["confidence"]
    # cross-checks
    names = {"p_max": "P max", "p_comp": "P comp", "fuel_index": "Fuel index", "exh_temp": "Exhaust temp", "pmax_adj": "Pmax adjustment",
             "cw_out": "Jacket CW outlet", "piston_out": "Piston oil outlet", "pi": "Pi"}
    tone = {"match": "ok", "near": "info", "weak": "warn", "mismatch": "crit"}
    rows = []
    for k, v in p.get("xval", {}).items():
        dec = max(v["dec"], 1)
        rows.append([names[k], f'{v["doc"]:.{v["dec"]}f}', f'{v["calc"]:.{dec + 1}f}', f'{v["calc"] - v["doc"]:+.{dec + 1}f}', chip(v["label"], tone[v["label"]])])
    md('<p class="lbl">Reconciliation with the averages printed on the sheet</p>')
    if rows:
        md(table_html(["Parameter", "Sheet average", "Calculated mean", "Δ", "Verdict"], rows))
    else:
        md('<div class="card"><p class="note">The sheet contains no Average values to reconcile against.</p></div>')
    md(f'<p class="note" style="margin-top:.5rem">Confidence {conf["total"]:.1f}% = 20% structure · 30% completeness · 15% plausibility · 35% sheet cross-check, minus 2 points per warning. '
       f'A mismatch can mean a mis-read value <i>or</i> an arithmetic slip by the crew on the sheet itself - compare with the document when in doubt.</p>')
    # provenance
    prow = []
    for cy in p["cylinders"]:
        cells = [f'<span class="cyl-chip" style="min-width:28px;height:26px;font-size:.8rem">{cy["id"]}</span>']
        for k in CORE + ("exh_temp",):
            s = cy.get("src_" + k)
            v = cy.get(k)
            if v is None:
                cells.append('<span class="dim">—</span>')
            else:
                cells.append(f'{v:g} <span class="dim" style="font-size:.7rem">{"hdr" if s == "embedded" else "cell"}</span>')
        prow.append(cells)
    md('<p class="lbl" style="margin-top:1.1rem">Value provenance (extracted)</p>')
    md(table_html(["Cyl", "P max", "P comp", "Fuel idx", "Exhaust"], prow))
    # notes
    md('<p class="lbl" style="margin-top:1.1rem">Extraction notes</p>')
    if p["issues"]:
        sevtone = {"info": "info", "warn": "warn", "error": "crit"}
        md("".join(f'<div style="margin:5px 0">{chip(s.upper(), sevtone[s])} <span style="font-size:.88rem">{esc(m)}</span></div>' for s, m in p["issues"]))
    else:
        md('<div class="note">No extraction notes - nothing unusual was found.</div>')
    # operating
    items = [(lab, f"{p['op'][k]:,.{nd}f} {u}".strip()) for k, lab, u, nd in OP_LABELS if k in p["op"]]
    if p["op"].get("wind_dir"):
        items.append(("Wind direction", p["op"]["wind_dir"]))
    md('<p class="lbl" style="margin-top:1.1rem">Operating conditions at the time of the test</p>')
    md('<div class="card"><div class="facts">' + "".join(fact(a, b) for a, b in items) + "</div></div>" if items else '<div class="note">Not filled in on the sheet.</div>')
    # singles & TCs
    sitems = [(lab, f"{p['singles'][k]:,.{nd}f} {u}".strip()) for k, lab, u, nd in SINGLE_LABELS if k in p["singles"]]
    md('<p class="lbl" style="margin-top:1.1rem">Auxiliary readings (cooling water · lube oil · fuel oil · scavenge)</p>')
    md('<div class="card"><div class="facts">' + "".join(fact(a, b) for a, b in sitems) + "</div></div>" if sitems else '<div class="note">Not filled in on the sheet.</div>')
    if p.get("tc"):
        trows = []
        for k, lab, nd in TC_LABELS:
            if any(k in v for v in p["tc"].values()):
                trows.append([lab] + [(f'{p["tc"][t][k]:,.{nd}f}' if t in p["tc"] and k in p["tc"][t] else '<span class="dim">—</span>') for t in sorted(p["tc"])])
        if trows:
            md('<p class="lbl" style="margin-top:1.1rem">Turbocharger readings</p>')
            md(table_html(["Reading"] + [f"T/C {t}" for t in sorted(p["tc"])], trows))
    ops = [(lab, p["op"][k]) for k, lab in (("hour", "Hour of test"), ("wave_dir", "Wave direction")) if p["op"].get(k)]
    if ops:
        md('<div class="card"><div class="facts">' + "".join(fact(a, b) for a, b in ops) + "</div></div>")
    ti = p.get("tc_info") or {}
    if ti:
        tl = [("Maker", ti.get("maker")), ("Type", ti.get("type")), ("Max speed", f'{ti["max_rpm"]:,.0f} rpm' if ti.get("max_rpm") else None),
              ("Max temperature", f'{ti["max_temp"]:,.0f} °C' if ti.get("max_temp") else None), ("Serial no.", ", ".join(f"T/C {k}: {v}" for k, v in sorted(ti.get("serials", {}).items())) or None),
              ("Governor", ti.get("governor")), ("Governor type", ti.get("governor_type")), ("Cylinder constant", ti.get("cyl_constant"))]
        md('<p class="lbl" style="margin-top:1.1rem">Turbocharger & governor specification</p><div class="card"><div class="facts">' + "".join(fact(a, b) for a, b in tl if b) + "</div></div>")
    hx = {k: v for k, v in (p.get("header_extra") or {}).items() if not k.startswith(("Turbocharger", "Governor", "Cylinder constant"))}
    if hx:
        md('<p class="lbl" style="margin-top:1.1rem">Fuel, lubricants & other header details</p><div class="card"><div class="facts">' + "".join(fact(k, v) for k, v in hx.items()) + "</div></div>")
    for s_ in p.get("remarks_series") or []:
        md(f'<p class="lbl" style="margin-top:1.1rem">Values read from the remarks · {esc(s_["context"])}</p>'
           + table_html(["Cyl"] + [str(k) for k in sorted(s_["values"])], [["Value"] + [f"{s_['values'][k]:g}" for k in sorted(s_["values"])]])
           + '<p class="note">Parsed from free text - not part of the sheet columns, shown for reference only.</p>')
    cov = p.get("coverage")
    if cov:
        md(f'<p class="lbl" style="margin-top:1.1rem">Sheet coverage</p><div class="note">{cov["claimed"]} of {cov["total"]} cells mapped to a field · {cov["template"]} printed template labels · '
           f'<b>{cov["unmapped"]} not mapped</b>.</div>')
        if p.get("unclaimed"):
            md(table_html(["Row", "Text on the sheet", "Next to", "Above"], [[str(u["row"]), esc(u["text"][:120]), esc(u["left"][:40]), esc(u["above"][:40])] for u in p["unclaimed"]])
               + '<p class="note">These cells carry text the reader could not assign to a field. Nothing is hidden - check them against the document.</p>')
    if p.get("remarks"):
        md(f'<p class="lbl" style="margin-top:1.1rem">Remarks written on the sheet</p><div class="card"><p style="margin:0;font-size:.9rem">{esc(p["remarks"])}</p></div>')


def render_report(rep: dict) -> None:
    ss = st.session_state
    if rep["error"] or not rep["parsed"]:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Could not read {esc(rep["name"])}</div><div class="bd">{esc(rep["error"] or "Unknown error")}</div></div></div>')
        return
    p = rep["parsed"]
    if any(s == "error" for s, _ in p["issues"]) and not p["cylinders"]:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Not a TEC-005 performance sheet</div><div class="bd">{esc(p["issues"][0][1])}</div></div></div>')
        return
    # optional auto-confirm for fully verified sheets
    if ss["auto_confirm"] and not rep["confirmed"] and not rep["auto_tried"] and p["confidence"]["level"] == "verified":
        rep["auto_tried"] = True
        _ok, errs, skipped = validate_cylinders(rep["vals"], rep["n"])
        rep["errors"], rep["skipped"], rep["confirmed"] = errs, skipped, not errs
    render_steps(3 if rep["confirmed"] else 2)
    render_report_card(rep)
    render_banner(rep)
    section("01", "Verify extracted readings", "pre-filled from the sheet · edit anything that differs")
    render_form(rep)
    if rep["confirmed"]:
        an, _ = get_analysis(rep, ss["th"])
        if an is None:
            md('<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Analysis unavailable</div><div class="bd">The confirmed values could not be analysed.</div></div></div>')
        else:
            if rep["skipped"]:
                md(f'<p class="note">Cylinder(s) {", ".join(map(str, rep["skipped"]))} left empty and excluded from the analysis.</p>')
            render_results(rep, an)
            section("05", "Export")
            a, b, c = st.columns(3)
            base = re.sub(r"[^A-Za-z0-9_-]+", "_", (p["identity"].get("vessel") or rep["name"])) + "_" + (p["identity"]["date"].strftime("%Y%m%d") if p["identity"].get("date") else "report")
            a.download_button("Cylinder data · CSV", export_csv(an), f"{base}.csv", "text/csv", key=f"w_{rep['id']}_dl_csv", **stretch(st.download_button))
            b.download_button("Full report · JSON", export_json(rep, an), f"{base}.json", "application/json", key=f"w_{rep['id']}_dl_json", **stretch(st.download_button))
            c.download_button("Printable report · HTML", export_html(rep, an), f"{base}.html", "text/html", key=f"w_{rep['id']}_dl_html", **stretch(st.download_button))
    else:
        md('<p class="note" style="margin-top:.6rem">The analysis appears here as soon as the readings are confirmed.</p>')
    with st.expander("Extraction audit · how every number was obtained"):
        render_audit(rep)


# ═════════════════════════════════════════════════════════════════════════════
#  8. FLEET VIEW
# ═════════════════════════════════════════════════════════════════════════════
def render_fleet() -> None:
    ss = st.session_state
    rows, ok_rows, by_vessel = [], [], {}
    for rid in ss["order"]:
        rep = ss["reports"][rid]
        p = rep.get("parsed")
        if not p:
            rows.append([esc(rep["name"]), "—", "—", "—", chip("unreadable", "crit"), "—", "—", "—", "—"])
            continue
        an, confirmed = get_analysis(rep, ss["th"])
        idn, conf = p["identity"], p["confidence"]
        vessel = idn.get("vessel") or rep["name"]
        counts = {s: 0 for s in SEV_RANK}
        if an:
            for f in an["findings"]:
                counts[f["sev"]] += 1
            ok_rows.append(dict(label=f"{vessel} · {idn['date'].strftime('%d %b') if idn.get('date') else ''}", score=an["score"]))
            if idn.get("date"):
                by_vessel.setdefault(vessel, []).append((idn["date"], an))
        lvl_tone = {"verified": "ok", "high": "info", "review": "warn", "low": "crit", "blank": "gold"}[conf["level"]]
        status = chip("confirmed", "ok") if confirmed else chip("unverified", "warn") if an else chip(conf["label"].lower(), lvl_tone)
        crit = counts["critical"] + counts["mechanical"]
        rows.append([f'<b>{esc(vessel)}</b>', fdate(idn.get("date")), str(p["n_cylinders"]), f'{conf["total"]:.0f}%' if conf["level"] != "blank" else "blank",
                     status, str(an["score"]) if an else "—", (chip(str(crit), "crit") if crit else '<span class="dim">0</span>') if an else "—",
                     (chip(str(counts["warning"]), "warn") if counts["warning"] else '<span class="dim">0</span>') if an else "—",
                     fnum(an["mean"]["ratio"], 3) if an else "—"])
    section("F1", "Fleet overview", f"{len(ss['order'])} sheet(s) loaded")
    md(table_html(["Vessel", "Report date", "Cyl", "Confidence", "Status", "Health", "Critical / mech.", "Warnings", "Mean ratio"], rows))
    md('<p class="note" style="margin-top:.5rem">“Unverified” rows are analysed on the raw extracted values - open the report and confirm to make them official.</p>')
    if ok_rows:
        section("F2", "Health ranking")
        plot(chart_fleet_health(sorted(ok_rows, key=lambda r: -r["score"])), "fleet_health")
    trends = {v: sorted(lst, key=lambda t: t[0]) for v, lst in by_vessel.items() if len({d for d, _ in lst}) >= 2}
    if trends:
        section("F3", "Trends", "same vessel · several dates")
        for v, lst in trends.items():
            md(f'<p class="lbl">{esc(v)}</p>')
            series = {k: [(d, a["mean"][k]) for d, a in lst if a["mean"][k] is not None] for k in ("p_max", "p_comp", "fuel_index", "exh_temp")}
            plot(chart_trend(series), f"trend_{v}")
    else:
        md('<p class="note" style="margin-top:1rem">Upload several months for the same vessel to unlock month-to-month trend charts.</p>')


# ═════════════════════════════════════════════════════════════════════════════
#  9. MAIN
# ═════════════════════════════════════════════════════════════════════════════
def main() -> None:
    inject_css()
    init_state()
    header_slot = st.container()
    files = st.file_uploader("Drop TEC-005 performance sheets (.doc) here", type=["doc"], accept_multiple_files=True,
                             key=f"uploader_{st.session_state['uploader_gen']}", label_visibility="collapsed",
                             help="Legacy Word 97-2003 binary files. No conversion needed - you can drop several vessels at once.")
    sync_reports(files)
    with header_slot:
        render_header()
    render_sidebar()
    ss = st.session_state
    if not ss["order"]:
        render_steps(1)
        render_empty()
        return
    if len(ss["order"]) == 1:
        render_report(ss["reports"][ss["order"][0]])
        return
    t1, t2 = st.tabs(["Report", "Fleet overview"])
    with t1:
        render_report(ss["reports"][ss.get("active", ss["order"][-1])])
    with t2:
        render_fleet()


main()
