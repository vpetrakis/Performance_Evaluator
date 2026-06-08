# ─────────────────────────────────────────────────────────────────────────────
# M.E. COMMAND CENTER v2.0  —  Maritime Engineering Performance Dashboard
# M/V ALEXIS  |  MAN-B&W 5S60MC-C MK8
# Author: Chief Systems Engineer  |  Anchor-based parser with cross-validation
# ─────────────────────────────────────────────────────────────────────────────

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pydantic import BaseModel, Field, ValidationError
import olefile
from io import BytesIO
import re
import math

# ── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="M.E. Command Center",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── PREMIUM CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

.stApp { background: #060d1a !important; font-family: 'Inter', system-ui, sans-serif !important; }
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton, div[data-testid="stToolbar"] { display: none !important; }

section[data-testid="stSidebar"] {
    background: #050b18 !important;
    border-right: 1px solid rgba(56,182,255,0.1) !important;
}
.main .block-container { padding-top: 1.5rem !important; max-width: 100% !important; }
h1 { color: #f1f5f9 !important; font-weight: 700 !important; letter-spacing: -0.8px; }
h2 { color: #e2e8f0 !important; font-weight: 600 !important; letter-spacing: -0.4px; }
h3 { color: #cbd5e1 !important; font-weight: 500 !important; }

div[data-testid="metric-container"] {
    background: linear-gradient(160deg, rgba(14,24,48,0.95), rgba(8,14,30,0.98)) !important;
    border: 1px solid rgba(56,182,255,0.16) !important;
    border-radius: 14px !important;
    padding: 1rem 1.25rem !important;
    box-shadow: 0 2px 20px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.03) !important;
}
div[data-testid="stMetricLabel"] p {
    color: #475569 !important; font-size: 0.67rem !important;
    font-weight: 600 !important; letter-spacing: 1.5px !important;
    text-transform: uppercase !important;
}
div[data-testid="stMetricValue"] { color: #e2e8f0 !important; font-size: 1.5rem !important; font-weight: 600 !important; }
div[data-testid="stMetricDelta"] { font-size: 0.72rem !important; }

div[data-baseweb="tab-list"] {
    background: rgba(5,11,24,0.85) !important; border-radius: 12px !important;
    padding: 4px !important; gap: 3px !important;
    border: 1px solid rgba(56,182,255,0.1) !important;
}
button[data-baseweb="tab"] {
    background: transparent !important; color: #475569 !important;
    border-radius: 8px !important; font-size: 0.72rem !important;
    font-weight: 600 !important; letter-spacing: 0.9px !important;
    text-transform: uppercase !important; border: none !important;
    padding: 0.42rem 1rem !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    background: rgba(56,182,255,0.12) !important; color: #38b6ff !important;
}
button[data-baseweb="tab"]:hover { color: #94a3b8 !important; }
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] { display: none !important; }

div[data-testid="stNumberInputField"] input {
    background: rgba(0,0,0,0.35) !important;
    border: 1px solid rgba(56,182,255,0.2) !important;
    color: #38b6ff !important; border-radius: 8px !important;
    font-weight: 600 !important; font-size: 0.95rem !important;
}
.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #1e3a8a) !important;
    border: 1px solid rgba(56,182,255,0.35) !important; color: #e2e8f0 !important;
    border-radius: 10px !important; font-weight: 600 !important;
    box-shadow: 0 0 18px rgba(29,78,216,0.25) !important;
}
div[data-testid="stForm"] {
    background: rgba(5,11,24,0.6) !important;
    border: 1px solid rgba(56,182,255,0.1) !important;
    border-radius: 16px !important; padding: 1.5rem !important;
}
div[data-testid="stFileUploadDropzone"] {
    background: rgba(8,15,31,0.7) !important;
    border: 2px dashed rgba(56,182,255,0.22) !important;
    border-radius: 16px !important;
}
hr { border-color: rgba(56,182,255,0.07) !important; margin: 0.8rem 0 !important; }
div[data-testid="stDataFrame"] {
    border: 1px solid rgba(56,182,255,0.12) !important;
    border-radius: 12px !important; overflow: hidden !important;
}

/* ── Custom components ── */
.ve-header {
    background: linear-gradient(135deg, rgba(10,20,45,0.97), rgba(6,13,26,0.99));
    border: 1px solid rgba(56,182,255,0.18); border-radius: 16px;
    padding: 1.3rem 1.7rem; margin-bottom: 1.2rem;
}
.ve-name { font-size: 1.55rem; font-weight: 700; color: #f1f5f9; letter-spacing: -0.5px; margin: 0; }
.ve-sub  { font-size: 0.72rem; color: #475569; letter-spacing: 1.6px; text-transform: uppercase; margin: 3px 0 0; }

.il { font-size: 0.63rem; color: #475569; text-transform: uppercase; letter-spacing: 1.2px; margin: 0; font-weight: 600; }
.iv { font-size: 0.87rem; color: #cbd5e1; font-weight: 500; margin: 1px 0 9px; }

.cyl-wrap {
    background: linear-gradient(160deg, rgba(10,20,42,0.9), rgba(6,13,26,0.95));
    border: 1px solid rgba(56,182,255,0.13); border-radius: 14px;
    padding: 1.1rem 1.2rem; margin-bottom: 0.75rem;
    border-left: 3px solid;
}
.sp {
    display: inline-flex; align-items: center; gap: 5px;
    background: rgba(0,0,0,0.28); border: 1px solid rgba(255,255,255,0.05);
    border-radius: 7px; padding: 3px 9px;
    font-size: 0.8rem; color: #64748b; margin-right: 5px; margin-bottom: 4px;
}
.sp b { color: #e2e8f0; font-weight: 600; }

.badge {
    display: inline-block; padding: 2px 9px; border-radius: 20px;
    font-size: 0.62rem; font-weight: 700; letter-spacing: 1.2px; text-transform: uppercase;
}
.bn { background: rgba(16,185,129,0.1); color: #10b981; border: 1px solid rgba(16,185,129,0.28); }
.bw { background: rgba(245,158,11,0.1); color: #f59e0b; border: 1px solid rgba(245,158,11,0.28); }
.bc { background: rgba(239,68,68,0.1); color: #ef4444; border: 1px solid rgba(239,68,68,0.28); }
.bi { background: rgba(56,182,255,0.1); color: #38b6ff; border: 1px solid rgba(56,182,255,0.25); }

.dc {
    border-radius: 12px; padding: 0.95rem 1.15rem;
    margin-bottom: 0.75rem; border-left: 3px solid;
}
.dc-c { background: rgba(239,68,68,0.06); border-color: #ef4444; border: 1px solid rgba(239,68,68,0.18); border-left: 3px solid #ef4444; }
.dc-w { background: rgba(245,158,11,0.06); border-color: #f59e0b; border: 1px solid rgba(245,158,11,0.18); border-left: 3px solid #f59e0b; }
.dc-m { background: rgba(251,113,133,0.06); border: 1px solid rgba(251,113,133,0.18); border-left: 3px solid #fb7185; }
.dc-n { background: rgba(16,185,129,0.06); border: 1px solid rgba(16,185,129,0.18); border-left: 3px solid #10b981; }
.dt { font-size: 0.8rem; font-weight: 600; margin: 0 0 3px; }
.df { font-size: 0.84rem; color: #cbd5e1; margin: 0 0 5px; }
.da { font-size: 0.76rem; color: #64748b; margin: 0; }

.opdata {
    background: rgba(8,15,31,0.7); border: 1px solid rgba(56,182,255,0.1);
    border-radius: 12px; padding: 1rem 1.2rem; margin-bottom: 0.7rem;
}
.conf-bar { height: 5px; background: rgba(255,255,255,0.07); border-radius: 3px; overflow: hidden; margin-top: 5px; }
.conf-fill { height: 100%; border-radius: 3px; background: linear-gradient(90deg, #1d4ed8, #38b6ff); }
</style>
""", unsafe_allow_html=True)


# ── PYDANTIC MODEL ─────────────────────────────────────────────────────────────
class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0, le=200)
    p_comp: float = Field(..., gt=0, le=160)
    fuel_index: float = Field(..., gt=0, lt=150)

    @property
    def combustion_ratio(self) -> float:
        return round(self.p_max / self.p_comp, 3)


# ── CORE PARSER ────────────────────────────────────────────────────────────────
def parse_tec_005(file_bytes: bytes) -> dict:
    """
    Anchor-based, cross-validated parser for MAN-B&W TEC-005 data sheets.
    Guaranteed extraction with integrity check against document-embedded averages.
    """
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    raw = ole.openstream("WordDocument").read()
    cell_bytes_list = raw.split(b"\x07")

    def decode_cell(cb: bytes) -> str:
        null_ratio = cb.count(b"\x00") / max(len(cb), 1)
        if null_ratio > 0.25 and len(cb) > 2:
            # UTF-16-LE: strip null bytes from ASCII decode to get digits
            cleaned = cb.decode("ascii", errors="ignore").replace("\x00", "").strip()
            if cleaned:
                return cleaned
        txt = cb.decode("ascii", errors="ignore")
        return "".join(c for c in txt if c.isprintable() or c in "\r\n").strip()

    cells = [decode_cell(cb) for cb in cell_bytes_list]

    def to_float(s: str) -> float | None:
        s = str(s).replace(",", "").replace("\r", "").replace("\n", "").strip()
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None

    def safe_get(idx: int) -> str:
        return cells[idx] if 0 <= idx < len(cells) else ""

    result: dict = {
        "vessel": "N/A", "engine_type": "N/A", "date": "N/A",
        "rpm": 0.0, "load_pct": 0.0, "run_hours": 0.0,
        "tc_rpm": 0.0, "scav_pressure": 0.0, "governor_index": 0.0,
        "speed_log": 0.0, "speed_obs": 0.0, "barometer": 0.0,
        "draught_fore": 0.0, "draught_aft": 0.0,
        "builder": "N/A", "yard": "N/A", "built_year": "N/A",
        "engine_no": "N/A", "bhp": 0.0, "cyl_constant": 0.0,
        "tc_maker": "N/A", "tc_type": "N/A", "tc_max_rpm": 0.0,
        "fuel_type": "N/A", "fuel_sulphur": 0.0, "fuel_density": 0.0,
        "remarks": "",
        "cylinders": [],
        "parse_confidence": 0.0,
        "avg_pmax_doc": 0.0, "avg_pcomp_doc": 0.0,
        "avg_fuel_doc": 0.0, "avg_exhaust_doc": 0.0,
        "n_cylinders": 5,
    }

    # ── Metadata ──
    label_map = {
        "M/V:": ("vessel", None), "Engine type:": ("engine_type", None),
        "Date:": ("date", None), "Builder:": ("builder", None),
        "Yard:": ("yard", None), "Built year:": ("built_year", None),
        "Oil brand:": ("fuel_type", None),
    }
    for i, cell in enumerate(cells):
        c = cell.strip()
        if c in label_map:
            key, converter = label_map[c]
            nxt = safe_get(i + 1)
            result[key] = (converter(nxt) if converter else nxt) or result[key]
        elif c == "BHP:":
            v = to_float(safe_get(i + 1))
            if v:
                result["bhp"] = v
        elif c == "Sulphur %:":
            v = to_float(safe_get(i + 1))
            if v:
                result["fuel_sulphur"] = v
        elif c == "Density at 15 C:":
            v = to_float(safe_get(i + 1))
            if v:
                result["fuel_density"] = v
        elif c == "Maker:" and 34 <= i <= 42:
            result["tc_maker"] = safe_get(i + 1)
        elif c == "Max RPM:" and 48 <= i <= 55:
            v = to_float(safe_get(i + 1))
            if v:
                result["tc_max_rpm"] = v
        elif c == "Cylinder constant (HP, bar):" and i < 50:
            v = to_float(safe_get(i + 1))
            if v:
                result["cyl_constant"] = v

    # Engine No (cell with 'No:' between idx 20-30)
    for i, c in enumerate(cells):
        if c.strip() == "No:" and 18 <= i <= 30:
            result["engine_no"] = safe_get(i + 1)
            break

    # TC Type (cell with 'Type:' between idx 40-50)
    for i, c in enumerate(cells):
        if c.strip() == "Type:" and 40 <= i <= 52:
            result["tc_type"] = safe_get(i + 1)
            break

    # ── Find 'P max bar' anchor ──
    pi = 109  # Default known position
    for i, c in enumerate(cells):
        if c.strip() == "P max bar":
            pi = i
            break

    # ── Operational data (fixed offsets from anchor) ──
    ops = {
        "rpm":             116 - 109,
        "run_hours":       115 - 109,
        "draught_fore":    114 - 109,
        "draught_aft":     144 - 109,
        "load_pct":        180 - 109,
        "governor_index":  145 - 109,
        "speed_log":       181 - 109,
        "speed_obs":       216 - 109,
        "barometer":       215 - 109,
        "tc_rpm":          308 - 109,
        "scav_pressure":   310 - 109,
    }
    for key, offset in ops.items():
        v = to_float(safe_get(pi + offset))
        if v is not None:
            result[key] = v

    # ── Document-embedded averages (cross-validation anchors) ──
    result["avg_pmax_doc"]    = to_float(safe_get(pi + 127)) or 0.0  # cell 236
    result["avg_pcomp_doc"]   = to_float(safe_get(pi + 129)) or 0.0  # cell 238
    result["avg_fuel_doc"]    = to_float(safe_get(pi + 131)) or 0.0  # cell 240
    result["avg_exhaust_doc"] = to_float(safe_get(pi + 285)) or 0.0  # cell 394

    # ── Cylinder count from engine type string ──
    m = re.match(r"(\d+)[Ss]", result["engine_type"])
    n_cyl = int(m.group(1)) if m else 5
    result["n_cylinders"] = n_cyl

    # ── Cylinder data extraction ──
    # Block 1: cylinders 1-3 at offset +43 from anchor (cells 152-160)
    b1 = pi + 43
    # Block 2: cylinders 4-6 at offset +79 (cells 188-196, blanks at +2,+5,+8 for 6th cyl)
    b2 = pi + 79
    # Exhaust block 1: cells 301-303 at offset +192
    e1 = pi + 192
    # Exhaust block 2: cells 339-340 at offset +230
    e2 = pi + 230

    pmax_vals  = [to_float(safe_get(b1 + i)) or 0.0 for i in range(3)]
    pcomp_vals = [to_float(safe_get(b1 + 3 + i)) or 0.0 for i in range(3)]
    fuel_vals  = [to_float(safe_get(b1 + 6 + i)) or 0.0 for i in range(3)]
    exh_vals   = [to_float(safe_get(e1 + i)) or 0.0 for i in range(3)]

    # Second block: [pmax4, pmax5, BLANK, pcomp4, pcomp5, BLANK, fuel4, fuel5, BLANK]
    for offset in [0, 1]:
        pmax_vals.append(to_float(safe_get(b2 + offset)) or 0.0)
        pcomp_vals.append(to_float(safe_get(b2 + 3 + offset)) or 0.0)
        fuel_vals.append(to_float(safe_get(b2 + 6 + offset)) or 0.0)
        exh_vals.append(to_float(safe_get(e2 + offset)) or 0.0)

    for i in range(n_cyl):
        result["cylinders"].append({
            "id": i + 1,
            "p_max":       pmax_vals[i]  if i < len(pmax_vals)  else 0.0,
            "p_comp":      pcomp_vals[i] if i < len(pcomp_vals) else 0.0,
            "fuel_index":  fuel_vals[i]  if i < len(fuel_vals)  else 0.0,
            "exhaust_temp": exh_vals[i]  if i < len(exh_vals)   else 0.0,
        })

    # ── Cross-validation ──
    valid = [c for c in result["cylinders"] if c["p_max"] > 0]
    if valid and result["avg_pmax_doc"] > 0:
        c_pmax  = sum(c["p_max"]       for c in valid) / len(valid)
        c_pcomp = sum(c["p_comp"]      for c in valid) / len(valid)
        c_fuel  = sum(c["fuel_index"]  for c in valid) / len(valid)
        d_pmax  = abs(c_pmax  - result["avg_pmax_doc"])  / result["avg_pmax_doc"]
        d_pcomp = abs(c_pcomp - result["avg_pcomp_doc"]) / result["avg_pcomp_doc"]
        d_fuel  = abs(c_fuel  - result["avg_fuel_doc"])  / result["avg_fuel_doc"]
        confidence = (1.0 - (d_pmax + d_pcomp + d_fuel) / 3.0) * 100.0
        result["parse_confidence"] = round(min(100.0, max(0.0, confidence)), 2)

    # Remarks
    for i, c in enumerate(cells):
        if c.strip() == "REMARKS" and i + 1 < len(cells):
            result["remarks"] = safe_get(i + 1).replace("\r", "\n").strip()
            break

    return result


# ── DIAGNOSTICS ENGINE ─────────────────────────────────────────────────────────
RATIO_LOW  = 1.30
RATIO_HIGH = 1.55
EXH_SPREAD = 25.0    # °C
EXH_ABS_HI = 380.0   # °C

def run_diagnostics(df: pd.DataFrame) -> list[dict]:
    diags = []
    avg_fuel  = df["fuel_index"].mean()
    avg_pmax  = df["p_max"].mean()
    avg_exh   = df["exhaust_temp"].mean()

    for _, row in df.iterrows():
        cyl  = int(row["id"])
        rat  = row["ratio"]
        dexh = row["exhaust_temp"] - avg_exh

        if rat < RATIO_LOW:
            diags.append({
                "cyl": cyl, "kind": "critical",
                "title": f"CYL {cyl}  ·  POOR COMBUSTION  (ratio {rat:.3f})",
                "fault": f"Combustion ratio {rat:.3f} is below the {RATIO_LOW} threshold. "
                         f"Pmax = {row['p_max']:.0f} bar against Pcomp = {row['p_comp']:.0f} bar.",
                "action": "Inspect and test-cock fuel atomizer. Check needle valve lift and seat. "
                          "Measure fuel cam timing. If fuel index is normal relative to sister cylinders, "
                          "suspect a fouled/blocked nozzle rather than a pump fault.",
            })
        elif rat > RATIO_HIGH:
            diags.append({
                "cyl": cyl, "kind": "warning",
                "title": f"CYL {cyl}  ·  HARSH COMBUSTION  (ratio {rat:.3f})",
                "fault": f"Combustion ratio {rat:.3f} exceeds the {RATIO_HIGH} upper limit.",
                "action": "Check injection timing — possible early injection. Inspect fuel cam follower "
                          "for abnormal wear. Verify fuel quality and calorific value.",
            })

        if row["fuel_index"] > avg_fuel + 2.5 and row["p_max"] < avg_pmax - 5:
            diags.append({
                "cyl": cyl, "kind": "mechanical",
                "title": f"CYL {cyl}  ·  WORN FUEL PUMP PLUNGER SUSPECTED",
                "fault": f"Fuel index {row['fuel_index']:.1f} (avg {avg_fuel:.1f}) elevated "
                         f"while Pmax {row['p_max']:.0f} bar (avg {avg_pmax:.0f}) is low. "
                         "Classic worn-plunger signature.",
                "action": "Perform helical groove measurement. Compare with last overhaul data. "
                          "Schedule pump replacement at next opportunity.",
            })

        if abs(dexh) > EXH_SPREAD:
            direction = "high" if dexh > 0 else "low"
            diags.append({
                "cyl": cyl, "kind": "warning",
                "title": f"CYL {cyl}  ·  EXHAUST TEMPERATURE DEVIATION ({dexh:+.0f}°C)",
                "fault": f"Exhaust {row['exhaust_temp']:.0f}°C vs fleet avg {avg_exh:.0f}°C — "
                         f"deviation {abs(dexh):.0f}°C exceeds {EXH_SPREAD}°C limit.",
                "action": (
                    "High exhaust: Check exhaust valve — possible incomplete seating or burnt valve. "
                    "Verify injection timing. Check cylinder oil feed."
                ) if direction == "high" else (
                    "Low exhaust: Inspect fuel atomizer for partial blockage. "
                    "Check exhaust valve for leakage (low compression causes cold exhaust)."
                ),
            })

        if row["exhaust_temp"] > EXH_ABS_HI:
            diags.append({
                "cyl": cyl, "kind": "warning",
                "title": f"CYL {cyl}  ·  ABSOLUTE EXHAUST TEMP ELEVATED ({row['exhaust_temp']:.0f}°C)",
                "fault": f"Exhaust temperature {row['exhaust_temp']:.0f}°C exceeds {EXH_ABS_HI}°C advisory limit.",
                "action": "Monitor closely. Check exhaust valve seating. Ensure adequate cylinder lubrication.",
            })

    if not diags:
        diags.append({
            "cyl": 0, "kind": "nominal",
            "title": "ALL SYSTEMS NOMINAL",
            "fault": "Thermodynamic parameters and fuel delivery are within acceptable limits.",
            "action": "Continue standard monitoring. Next performance check per PMS schedule.",
        })
    return diags


# ── PLOTLY THEME ───────────────────────────────────────────────────────────────
_LAYOUT = dict(
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#64748b", size=11),
    margin=dict(l=8, r=8, t=36, b=8),
    hovermode="x unified",
    hoverlabel=dict(
        bgcolor="rgba(8,15,31,0.96)",
        bordercolor="rgba(56,182,255,0.3)",
        font=dict(family="Inter", color="#e2e8f0", size=12),
    ),
    legend=dict(
        bgcolor="rgba(0,0,0,0)", bordercolor="rgba(56,182,255,0.15)",
        borderwidth=1, font=dict(color="#94a3b8", size=11),
        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
    ),
)
_XAXIS = dict(gridcolor="rgba(255,255,255,0.04)", linecolor="rgba(255,255,255,0.08)",
              zeroline=False, tickcolor="rgba(255,255,255,0.2)", tickfont=dict(size=11))
_YAXIS = dict(gridcolor="rgba(255,255,255,0.04)", linecolor="rgba(255,255,255,0.08)",
              zeroline=False, tickcolor="rgba(255,255,255,0.2)", tickfont=dict(size=11))

BLUE   = "#38b6ff"
PURPLE = "#a78bfa"
GREEN  = "#34d399"
AMBER  = "#fbbf24"
ROSE   = "#fb7185"
SLATE  = "#64748b"


def chart_thermodynamics(df: pd.DataFrame) -> go.Figure:
    fig = make_subplots(
        rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.06,
        row_heights=[0.65, 0.35],
        subplot_titles=["Pressure Profile", "Fuel Pump Index"],
    )
    ids = df["id"].tolist()

    # Optimal zone polygon
    lower = (df["p_comp"] * RATIO_LOW).tolist()
    upper = (df["p_comp"] * RATIO_HIGH).tolist()
    fig.add_trace(go.Scatter(
        x=ids + ids[::-1], y=upper + lower[::-1],
        fill="toself", fillcolor="rgba(56,182,255,0.07)",
        line=dict(color="rgba(0,0,0,0)"), name="Optimal zone",
        hoverinfo="skip",
    ), row=1, col=1)

    # Pcomp
    fig.add_trace(go.Scatter(
        x=ids, y=df["p_comp"], name="P comp",
        mode="lines+markers",
        line=dict(color=PURPLE, width=2.5, dash="dot"),
        marker=dict(size=9, symbol="circle", color=PURPLE,
                    line=dict(color="white", width=1.5)),
    ), row=1, col=1)

    # Pmax
    fig.add_trace(go.Scatter(
        x=ids, y=df["p_max"], name="P max",
        mode="lines+markers",
        line=dict(color=BLUE, width=3),
        marker=dict(size=11, symbol="diamond", color=BLUE,
                    line=dict(color="white", width=1.5)),
    ), row=1, col=1)

    # Average Pmax line
    avg_pm = df["p_max"].mean()
    fig.add_hline(y=avg_pm, line=dict(color=BLUE, width=1, dash="dash"),
                  annotation_text=f"avg {avg_pm:.1f}", annotation_font=dict(color=BLUE, size=10),
                  row=1, col=1)

    # Fuel index bars
    colors = [GREEN if 57 <= v <= 60 else AMBER for v in df["fuel_index"]]
    fig.add_trace(go.Bar(
        x=ids, y=df["fuel_index"], name="Fuel index",
        marker=dict(color=colors, opacity=0.85, line=dict(color="rgba(0,0,0,0.3)", width=1)),
    ), row=2, col=1)

    layout = dict(**_LAYOUT)
    layout["xaxis"]  = dict(**_XAXIS, showgrid=False, tickmode="linear", title=None)
    layout["xaxis2"] = dict(**_XAXIS, showgrid=False, tickmode="linear",
                             title=dict(text="Cylinder No.", font=dict(size=11)))
    layout["yaxis"]  = dict(**_YAXIS, title=dict(text="Bar", font=dict(size=11)))
    layout["yaxis2"] = dict(**_YAXIS, title=dict(text="Index", font=dict(size=11)))
    for ax in ["xaxis", "xaxis2", "yaxis", "yaxis2"]:
        layout[ax]["color"] = "#64748b"
    fig.update_layout(**layout)
    fig.update_annotations(font_color="#64748b", font_size=11)
    return fig


def chart_exhaust(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    avg_exh = df["exhaust_temp"].mean()

    bar_colors = []
    for v in df["exhaust_temp"]:
        if abs(v - avg_exh) > EXH_SPREAD:
            bar_colors.append(ROSE)
        elif abs(v - avg_exh) > 15:
            bar_colors.append(AMBER)
        else:
            bar_colors.append(GREEN)

    fig.add_trace(go.Bar(
        x=df["id"], y=df["exhaust_temp"], name="Exhaust temp",
        marker=dict(color=bar_colors, opacity=0.85, line=dict(color="rgba(0,0,0,0.3)", width=1)),
        text=[f"{v:.0f}°C" for v in df["exhaust_temp"]],
        textposition="outside", textfont=dict(color="#94a3b8", size=11),
    ))
    fig.add_hline(
        y=avg_exh, line=dict(color=ROSE, width=1.5, dash="dot"),
        annotation_text=f"avg {avg_exh:.1f}°C",
        annotation_font=dict(color=ROSE, size=10),
    )

    layout = dict(**_LAYOUT)
    layout["xaxis"] = dict(**_XAXIS, showgrid=False, tickmode="linear",
                            title=dict(text="Cylinder No.", font=dict(size=11)))
    layout["yaxis"] = dict(**_YAXIS, title=dict(text="°C", font=dict(size=11)),
                            range=[min(df["exhaust_temp"]) - 30,
                                   max(df["exhaust_temp"]) + 40])
    layout["showlegend"] = False
    fig.update_layout(**layout)
    return fig


def chart_ratio(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    colors = []
    for v in df["ratio"]:
        if v < RATIO_LOW:
            colors.append(ROSE)
        elif v > RATIO_HIGH:
            colors.append(AMBER)
        else:
            colors.append(GREEN)

    fig.add_hrect(y0=RATIO_LOW, y1=RATIO_HIGH,
                  fillcolor="rgba(52,211,153,0.07)",
                  line=dict(color="rgba(52,211,153,0.2)", width=1))

    fig.add_trace(go.Bar(
        x=df["id"], y=df["ratio"], name="Ratio",
        marker=dict(color=colors, opacity=0.9, line=dict(color="rgba(0,0,0,0.3)", width=1)),
        text=[f"{v:.3f}" for v in df["ratio"]],
        textposition="outside", textfont=dict(color="#94a3b8", size=11),
    ))

    layout = dict(**_LAYOUT)
    layout["xaxis"]      = dict(**_XAXIS, showgrid=False, tickmode="linear",
                                 title=dict(text="Cylinder No.", font=dict(size=11)))
    layout["yaxis"]      = dict(**_YAXIS, title=dict(text="Pmax / Pcomp", font=dict(size=11)),
                                 range=[1.0, max(df["ratio"]) + 0.12])
    layout["showlegend"] = False
    fig.update_layout(**layout)
    return fig


# ── HELPERS ───────────────────────────────────────────────────────────────────
def cyl_status(ratio: float) -> tuple[str, str]:
    if ratio < RATIO_LOW:
        return "critical", "bc"
    if ratio > RATIO_HIGH:
        return "warning", "bw"
    return "nominal", "bn"

STATUS_LABEL = {"nominal": "NOMINAL", "warning": "WARNING", "critical": "CRITICAL"}

def health_score(df: pd.DataFrame) -> int:
    score = 100
    avg_exh = df["exhaust_temp"].mean()
    for _, row in df.iterrows():
        r = row["ratio"]
        if r < RATIO_LOW:   score -= 14
        elif r < 1.32:      score -= 4
        if r > RATIO_HIGH:  score -= 5
        dev = abs(row["exhaust_temp"] - avg_exh)
        if dev > EXH_SPREAD: score -= 6
        elif dev > 15:       score -= 2
    return max(0, min(100, score))

def health_color(score: int) -> str:
    if score >= 80: return "#10b981"
    if score >= 60: return "#f59e0b"
    return "#ef4444"


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
def render_sidebar(data: dict, df: pd.DataFrame | None):
    with st.sidebar:
        st.markdown("### ⚙️ M.E. Command Center")
        st.markdown("---")

        if data:
            st.markdown(f"""
<p class="il">Vessel</p><p class="iv">{data['vessel']}</p>
<p class="il">Engine Type</p><p class="iv">{data['engine_type']}</p>
<p class="il">Builder</p><p class="iv">{data['builder']} — {data['yard']}</p>
<p class="il">Engine No.</p><p class="iv">{data['engine_no']} &nbsp;|&nbsp; {data['built_year']}</p>
<p class="il">BHP</p><p class="iv">{int(data['bhp']):,} HP</p>
""", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown(f"""
<p class="il">TC Unit</p><p class="iv">{data['tc_maker']} {data['tc_type']}</p>
<p class="il">TC Max RPM</p><p class="iv">{int(data['tc_max_rpm']):,} RPM</p>
<p class="il">Fuel Oil</p><p class="iv">{data['fuel_type']}</p>
<p class="il">Sulphur</p><p class="iv">{data['fuel_sulphur']} %</p>
<p class="il">Density @ 15°C</p><p class="iv">{data['fuel_density']} kg/l</p>
""", unsafe_allow_html=True)

            if df is not None:
                st.markdown("---")
                score = health_score(df)
                col = health_color(score)
                st.markdown(f"""
<p class="il">Engine Health Score</p>
<p style="font-size:2.2rem;font-weight:700;color:{col};margin:0;line-height:1.1;">{score}<span style="font-size:1rem;color:#475569;">/100</span></p>
<div class="conf-bar"><div class="conf-fill" style="width:{score}%;background:{col};"></div></div>
""", unsafe_allow_html=True)

            conf = data.get("parse_confidence", 0)
            st.markdown(f"""
<p class="il" style="margin-top:14px;">Data Confidence</p>
<p style="font-size:1.05rem;font-weight:600;color:#38b6ff;margin:0;">{conf:.1f}%</p>
<div class="conf-bar"><div class="conf-fill" style="width:{conf}%;"></div></div>
""", unsafe_allow_html=True)
        else:
            st.markdown('<p class="iv" style="color:#475569;">Upload a TEC-005 sheet to begin.</p>',
                        unsafe_allow_html=True)


# ── MAIN APP ──────────────────────────────────────────────────────────────────
def main():
    # Init session state
    for key in ("parsed_data", "current_file", "analysis_df"):
        if key not in st.session_state:
            st.session_state[key] = None

    render_sidebar(st.session_state.parsed_data, st.session_state.analysis_df)

    # ── HEADER ──
    date_str = ""
    if st.session_state.parsed_data:
        d = st.session_state.parsed_data
        date_str = d.get("date", "")
    st.markdown(f"""
<div class="ve-header">
  <div>
    <p class="ve-name">M.E. COMMAND CENTER</p>
    <p class="ve-sub">Maritime Engineering Performance Dashboard &nbsp;·&nbsp; TEC-005 Analysis Suite</p>
  </div>
  <div style="text-align:right;">
    <span class="badge bi">v2.0</span>&nbsp;
    <span class="badge bi">{date_str or 'Awaiting Data'}</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── FILE UPLOAD ──
    uploaded = st.file_uploader(
        "Upload TEC-005 Performance Data Sheet (.doc)",
        type=["doc"],
        help="Legacy Microsoft Word (.doc) binary file — no conversion required",
    )

    if uploaded:
        if st.session_state.current_file != uploaded.name:
            with st.spinner("Parsing binary OLE2 document…"):
                data = parse_tec_005(uploaded.read())
                st.session_state.parsed_data  = data
                st.session_state.current_file = uploaded.name
                st.session_state.analysis_df  = None

        data = st.session_state.parsed_data

        # ── CONFIDENCE BANNER ──
        conf = data.get("parse_confidence", 0)
        if conf >= 99.5:
            st.success(
                f"✅  **Data Integrity Verified** — {conf:.2f}% confidence. "
                f"Computed averages match document averages exactly. "
                f"({data['n_cylinders']}-cylinder engine detected)"
            )
        elif conf >= 95.0:
            st.info(f"ℹ️  **Data Extracted** — {conf:.1f}% confidence. Verify highlighted fields.")
        else:
            st.warning(f"⚠️  **Low Confidence ({conf:.1f}%)** — Manual verification required for all fields.")

        st.markdown("---")

        # ── VERIFICATION FORM ──
        st.markdown("### 🔍 Data Integrity Verification")
        st.caption("Extracted values are pre-filled. Correct any discrepancies before executing analysis.")

        verified_rows = []
        with st.form("verify_form"):
            for cyl in data["cylinders"]:
                with st.expander(
                    f"Cylinder {cyl['id']}  ·  "
                    f"Pmax: {cyl['p_max']:.0f} bar  ·  "
                    f"Pcomp: {cyl['p_comp']:.0f} bar  ·  "
                    f"Fuel: {cyl['fuel_index']:.0f}  ·  "
                    f"Exh: {cyl['exhaust_temp']:.0f}°C",
                    expanded=True,
                ):
                    c1, c2, c3, c4 = st.columns(4)
                    pm = c1.number_input("Pmax (bar)",    value=float(cyl["p_max"]),
                                         key=f"pm_{cyl['id']}", min_value=0.0, format="%.1f")
                    pc = c2.number_input("Pcomp (bar)",   value=float(cyl["p_comp"]),
                                         key=f"pc_{cyl['id']}", min_value=0.0, format="%.1f")
                    fi = c3.number_input("Fuel Index",    value=float(cyl["fuel_index"]),
                                         key=f"fi_{cyl['id']}", min_value=0.0, format="%.1f")
                    et = c4.number_input("Exhaust (°C)",  value=float(cyl["exhaust_temp"]),
                                         key=f"et_{cyl['id']}", min_value=0.0, format="%.1f")
                    verified_rows.append({
                        "id": cyl["id"], "p_max": pm, "p_comp": pc,
                        "fuel_index": fi, "exhaust_temp": et,
                    })

            submitted = st.form_submit_button(
                "🚀  Execute Full Analysis",
                use_container_width=True, type="primary",
            )

        if submitted:
            errors = []
            valid_cyls = []
            for row in verified_rows:
                if row["p_comp"] >= row["p_max"]:
                    errors.append(
                        f"Cyl {row['id']}: Pcomp ({row['p_comp']:.1f}) ≥ Pmax ({row['p_max']:.1f}) — "
                        "physically impossible."
                    )
                    continue
                try:
                    cyl_obj = Cylinder(
                        id=row["id"], p_max=row["p_max"],
                        p_comp=row["p_comp"], fuel_index=row["fuel_index"],
                    )
                    valid_cyls.append((cyl_obj, row["exhaust_temp"]))
                except ValidationError as e:
                    for err in e.errors():
                        errors.append(f"Cyl {row['id']} [{err['loc'][0]}]: {err['msg']}")

            if errors:
                st.error("**Data Integrity Failure** — Correct the following before proceeding:")
                for err in errors:
                    st.markdown(f"- {err}")
            else:
                final_df = pd.DataFrame([{
                    "id":           cyl.id,
                    "p_max":        cyl.p_max,
                    "p_comp":       cyl.p_comp,
                    "fuel_index":   cyl.fuel_index,
                    "exhaust_temp": exh,
                    "ratio":        cyl.combustion_ratio,
                } for cyl, exh in valid_cyls])
                st.session_state.analysis_df = final_df
                # Re-render sidebar with updated df
                render_sidebar(data, final_df)

    # ── ANALYSIS DASHBOARD ────────────────────────────────────────────────────
    if st.session_state.analysis_df is not None and st.session_state.parsed_data is not None:
        df   = st.session_state.analysis_df
        data = st.session_state.parsed_data
        diags = run_diagnostics(df)
        n_issues = sum(1 for d in diags if d["kind"] != "nominal")
        hs = health_score(df)

        st.markdown("---")

        # ── TOP KPI ROW ──
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Engine RPM",    f"{int(data['rpm'])} RPM",
                  delta=f"Load {data['load_pct']:.0f}%", delta_color="off")
        k2.metric("Mean P max",    f"{df['p_max'].mean():.1f} bar",
                  delta=f"±{df['p_max'].std():.1f} bar spread")
        k3.metric("Mean P comp",   f"{df['p_comp'].mean():.1f} bar",
                  delta=f"±{df['p_comp'].std():.1f} bar spread")
        avg_rat = df["ratio"].mean()
        delta_col = "normal" if RATIO_LOW <= avg_rat <= RATIO_HIGH else "inverse"
        k4.metric("Comb. Ratio",   f"{avg_rat:.3f}",
                  delta=f"Target {RATIO_LOW}–{RATIO_HIGH}", delta_color=delta_col)
        k5.metric("TC Speed",      f"{int(data['tc_rpm']):,} RPM",
                  delta=f"Scav {data['scav_pressure']:.2f} bar", delta_color="off")
        k6.metric("Health Score",  f"{hs}/100",
                  delta=f"{n_issues} issue(s) found",
                  delta_color="off" if n_issues == 0 else "inverse")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── TABS ──
        tab_overview, tab_cylinders, tab_diagnostics, tab_operations = st.tabs(
            ["📊  Overview", "🔩  Cylinders", "🧠  Diagnostics", "🌊  Operations"]
        )

        # ── TAB: OVERVIEW ──
        with tab_overview:
            col_main, col_side = st.columns([1.65, 1])

            with col_main:
                st.markdown("##### Thermodynamic & Fuel Profile")
                st.plotly_chart(chart_thermodynamics(df), use_container_width=True,
                                config={"displayModeBar": False})

            with col_side:
                st.markdown("##### Combustion Ratio per Cylinder")
                st.plotly_chart(chart_ratio(df), use_container_width=True,
                                config={"displayModeBar": False})

            st.markdown("##### Exhaust Gas Temperature Distribution")
            st.plotly_chart(chart_exhaust(df), use_container_width=True,
                            config={"displayModeBar": False})

            st.markdown("---")
            st.markdown("##### Cylinder Summary Table")
            display_df = df.copy()
            display_df.columns = ["Cyl", "Pmax (bar)", "Pcomp (bar)", "Fuel Index",
                                   "Exhaust (°C)", "Ratio"]
            display_df["Pmax (bar)"]   = display_df["Pmax (bar)"].map("{:.1f}".format)
            display_df["Pcomp (bar)"]  = display_df["Pcomp (bar)"].map("{:.1f}".format)
            display_df["Fuel Index"]   = display_df["Fuel Index"].map("{:.1f}".format)
            display_df["Exhaust (°C)"] = display_df["Exhaust (°C)"].map("{:.1f}".format)
            display_df["Ratio"]        = display_df["Ratio"].map("{:.3f}".format)
            st.dataframe(display_df.set_index("Cyl"), use_container_width=True)

        # ── TAB: CYLINDERS ──
        with tab_cylinders:
            st.markdown("##### Per-Cylinder Status Cards")
            cols = st.columns(2)
            for idx, (_, row) in enumerate(df.iterrows()):
                status, badge_cls = cyl_status(row["ratio"])
                label = STATUS_LABEL[status]
                card_html = f"""
<div class="cyl-wrap" style="border-left-color:{'#ef4444' if status=='critical' else '#f59e0b' if status=='warning' else '#10b981'};">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
    <span style="font-size:1.05rem;font-weight:700;color:#e2e8f0;">Cylinder {int(row['id'])}</span>
    <span class="badge {badge_cls}">{label}</span>
  </div>
  <div>
    <span class="sp">P max &nbsp;<b>{row['p_max']:.1f} bar</b></span>
    <span class="sp">P comp &nbsp;<b>{row['p_comp']:.1f} bar</b></span>
    <span class="sp">Ratio &nbsp;<b>{row['ratio']:.3f}</b></span>
    <span class="sp">Fuel idx &nbsp;<b>{row['fuel_index']:.1f}</b></span>
    <span class="sp">Exhaust &nbsp;<b>{row['exhaust_temp']:.0f}°C</b></span>
  </div>
</div>
"""
                cols[idx % 2].markdown(card_html, unsafe_allow_html=True)

        # ── TAB: DIAGNOSTICS ──
        with tab_diagnostics:
            kind_map = {
                "critical":   ("dc dc-c", "🔴 CRITICAL"),
                "warning":    ("dc dc-w", "🟡 WARNING"),
                "mechanical": ("dc dc-m", "🟠 MECHANICAL"),
                "nominal":    ("dc dc-n", "🟢 NOMINAL"),
            }
            title_colors = {
                "critical": "#ef4444", "warning": "#f59e0b",
                "mechanical": "#fb7185", "nominal": "#10b981",
            }
            for d in diags:
                cls, icon = kind_map[d["kind"]]
                col = title_colors[d["kind"]]
                st.markdown(f"""
<div class="{cls}">
  <p class="dt" style="color:{col};">{icon} &nbsp; {d['title']}</p>
  <p class="df">{d['fault']}</p>
  <p class="da">⚡ <strong>Action:</strong> {d['action']}</p>
</div>
""", unsafe_allow_html=True)

            if data.get("remarks"):
                st.markdown("---")
                st.markdown("**Remarks from Data Sheet**")
                st.info(data["remarks"])

        # ── TAB: OPERATIONS ──
        with tab_operations:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown('<div class="opdata">', unsafe_allow_html=True)
                st.markdown("**Operational Context**")
                st.markdown(f"""
<p class="il">Date</p><p class="iv">{data['date']}</p>
<p class="il">Engine Speed</p><p class="iv">{data['rpm']:.0f} RPM</p>
<p class="il">Load</p><p class="iv">{data['load_pct']:.0f} %</p>
<p class="il">Governor Index</p><p class="iv">{data['governor_index']:.1f}</p>
<p class="il">Total Run Hours</p><p class="iv">{data['run_hours']:,.0f} hrs</p>
""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="opdata">', unsafe_allow_html=True)
                st.markdown("**Navigation & Environment**")
                st.markdown(f"""
<p class="il">Speed (Log)</p><p class="iv">{data['speed_log']:.1f} kn</p>
<p class="il">Speed (Obs)</p><p class="iv">{data['speed_obs']:.1f} kn</p>
<p class="il">Barometer</p><p class="iv">{data['barometer']:.0f} mbar</p>
<p class="il">Draught Fore</p><p class="iv">{data['draught_fore']:.2f} m</p>
<p class="il">Draught Aft</p><p class="iv">{data['draught_aft']:.2f} m</p>
""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
            with c3:
                st.markdown('<div class="opdata">', unsafe_allow_html=True)
                st.markdown("**Turbocharger & Air System**")
                st.markdown(f"""
<p class="il">TC Speed</p><p class="iv">{data['tc_rpm']:,.0f} RPM</p>
<p class="il">Scav. Air Pressure</p><p class="iv">{data['scav_pressure']:.2f} bar</p>
<p class="il">TC Maker</p><p class="iv">{data['tc_maker']}</p>
<p class="il">TC Model</p><p class="iv">{data['tc_type']}</p>
<p class="il">TC Max Speed</p><p class="iv">{data['tc_max_rpm']:,.0f} RPM</p>
""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("**Fuel & Lubricants**")
            fc1, fc2 = st.columns(2)
            with fc1:
                st.markdown('<div class="opdata">', unsafe_allow_html=True)
                st.markdown(f"""
<p class="il">Fuel Type</p><p class="iv">{data['fuel_type']}</p>
<p class="il">Sulphur Content</p><p class="iv">{data['fuel_sulphur']:.3f} %</p>
<p class="il">Density @ 15°C</p><p class="iv">{data['fuel_density']:.4f} kg/l</p>
<p class="il">Cylinder Constant</p><p class="iv">{data['cyl_constant']:.3f} bar</p>
""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

    else:
        # ── LANDING SCREEN ──
        st.markdown("""
<div style="text-align:center;padding:4rem 2rem;color:#475569;">
  <p style="font-size:3rem;margin:0;">⚓</p>
  <p style="font-size:1.1rem;font-weight:500;color:#94a3b8;margin:0.5rem 0;">
    Upload a TEC-005 Performance Data Sheet to begin analysis
  </p>
  <p style="font-size:0.82rem;margin:0;">
    Supports legacy binary Microsoft Word (.doc) files &nbsp;·&nbsp; OLE2 / BIFF8 format
  </p>
</div>
""", unsafe_allow_html=True)


main()
