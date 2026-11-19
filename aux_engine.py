"""
Auxiliary engines (TEC-014)  -  reader, analysis and Streamlit tab in one module.

Sections:  1. PARSER     read the workbook by label (pure Python: openpyxl / xlrd, no Streamlit)
           2. DIAGNOSIS  input checks -> shop-trial reference -> engine findings -> probable causes -> verdict -> crew action list
           3. FLEET      comparisons between generators and between months
           4. UI         the Streamlit tab (re-uses the main-engine design system; every key is prefixed `dg_`)

The sheet is read by LABEL, never by cell address: crews insert / delete rows, so the same field sits on
a different row in different files.  Layout assumed from the printed form:

    col A  section title (merged downwards)      col D:E  actual test values
    col B  row label ("No. 3", "JACKET INLET")   col F:H  shop-trial reference, load point A
    col C  unit                                  col I:K  shop-trial reference, load point B

Every non-empty cell ends up either claimed by a field, recognised as printed template text, or listed in
`unclaimed` - nothing is lost silently.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import re
import time
import zipfile
from datetime import date as _date
from typing import Any, NamedTuple, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

from me_engine import (APP_VERSION, C_GOLD, C_GOLD2, C_INFO, C_MUTED, C_OK, C_ORANGE, C_WARN, _style, chip, esc, fact, fnum,
                       kpi, md, plot, ring_svg, section, stretch, table_html)

# ====================================================================================================
# 1. PARSER
# ====================================================================================================

PARSER_VERSION = "1.0"
MAX_BYTES = 25 * 1024 * 1024
MAX_UNZIPPED = 120 * 1024 * 1024
MAX_ROWS, MAX_COLS = 400, 40


class AuxParseError(Exception):
    """The file is not a readable TEC-014 workbook."""


# ───────────────────────────── parameter catalogue ─────────────────────────────
# key -> (label, unit)
PARAMS: dict[str, tuple[str, str]] = {
    "cw_jacket_in": ("Jacket CW inlet", "°C"), "cw_engine_out": ("Engine (jacket) CW outlet", "°C"),
    "cw_ac_in": ("Air cooler CW inlet", "°C"), "cw_ac_out": ("Air cooler CW outlet", "°C"),
    "cw_lo_in": ("L.O. cooler CW inlet", "°C"), "cw_lo_out": ("L.O. cooler CW outlet", "°C"),
    "lo_t_in": ("L.O. cooler inlet", "°C"), "lo_t_out": ("L.O. cooler outlet", "°C"),
    "ba_tc_in": ("Boost air · T/C inlet", "°C"), "ba_ac_in": ("Boost air · air cooler inlet", "°C"),
    "ba_cyl_in": ("Boost air · cylinder inlet", "°C"),
    "tc_in_upper": ("T/C inlet (cyl. 4-6, upper)", "°C"), "tc_in_lower": ("T/C inlet (cyl. 1-3, lower)", "°C"),
    "tc_in": ("T/C inlet", "°C"), "tc_out": ("T/C outlet", "°C"),
    "tc_pair_in": ("T/C in (from ΔT row)", "°C"), "tc_pair_out": ("T/C out (from ΔT row)", "°C"),
    "lo_p_engine": ("L.O. pressure · engine", "kgf/cm²"), "lo_p_tc": ("L.O. pressure · T/C", "kgf/cm²"),
    "fo_p_in": ("FO pressure, engine inlet", "bar"), "fo_t_in": ("FO temperature, engine inlet", "°C"),
    "fo_visc_heater_out": ("Viscosity control · heater outlet", "°C"), "fo_visc_set": ("Viscosity control · set point", "cSt"),
    "hfo_tank_t": ("HFO service tank heating", "°C"),
    "cw_p_jacket": ("Cooling water pressure · jacket", "kgf/cm²"), "cw_p_cooler": ("Cooling water pressure · cooler", "kgf/cm²"),
    "boost_p": ("Boost air pressure", "kgf/cm²"), "exh_p": ("Exhaust pressure, T/C outlet", "bar"),
    "gen_brg_de": ("Generator bearing · engine side", "°C"), "gen_brg_nde": ("Generator bearing · opposite side", "°C"),
}
PARAM_LAYOUT: list[tuple[str, list[str]]] = [
    ("Cooling water temperatures", ["cw_jacket_in", "cw_engine_out", "cw_ac_in", "cw_ac_out", "cw_lo_in", "cw_lo_out"]),
    ("Lube oil", ["lo_t_in", "lo_t_out", "lo_p_engine", "lo_p_tc"]),
    ("Boost air & turbocharger", ["boost_p", "ba_tc_in", "ba_ac_in", "ba_cyl_in", "tc_in_upper", "tc_in_lower", "tc_in", "tc_out",
                                  "tc_pair_in", "tc_pair_out", "exh_p"]),
    ("Fuel oil", ["fo_p_in", "fo_t_in", "fo_visc_heater_out", "fo_visc_set", "hfo_tank_t"]),
    ("Cooling water pressure", ["cw_p_jacket", "cw_p_cooler"]),
    ("Generator", ["gen_brg_de", "gen_brg_nde"]),
]
CYL_METRICS = {"inj": ("FO injection timing", "° BTDC"), "exh": ("Exhaust gas temperature", "°C"),
               "pmax": ("Pmax", ""), "rack": ("Fuel rack position", "mm")}
AVG_METRICS = ("exh", "pmax", "rack")
GROUPS = ("act", "shopA", "shopB")

# plausibility windows (value must lie inside, after unit normalisation)
PLAUSIBLE: dict[str, tuple[float, float]] = {k: (-5.0, 130.0) for k in PARAMS}
PLAUSIBLE.update({
    "cw_jacket_in": (20, 105), "cw_engine_out": (25, 110), "lo_t_in": (20, 110), "lo_t_out": (20, 120),
    "ba_tc_in": (0, 300), "ba_ac_in": (20, 300), "ba_cyl_in": (10, 120), "tc_in_upper": (100, 750), "tc_in_lower": (100, 750),
    "tc_in": (100, 750), "tc_out": (80, 700), "tc_pair_in": (100, 750), "tc_pair_out": (80, 700),
    "lo_p_engine": (0.3, 12), "lo_p_tc": (0.3, 12), "fo_p_in": (0.5, 25), "fo_t_in": (10, 170), "fo_visc_heater_out": (10, 170),
    "fo_visc_set": (2, 40), "hfo_tank_t": (10, 120), "cw_p_jacket": (0.2, 8), "cw_p_cooler": (0.2, 8), "boost_p": (0.1, 6),
    "exh_p": (0, 3), "gen_brg_de": (10, 150), "gen_brg_nde": (10, 150),
})
CYL_PLAUSIBLE = {"inj": (0.0, 45.0), "exh": (100.0, 750.0), "pmax_bar": (20.0, 260.0), "rack": (0.0, 100.0)}

NA_TOKENS = {"N/A", "NA", "N.A", "N.A.", "NIL", "-", "--", "—", "/", "NONE", "N / A", "NOT APPLICABLE", "NOT AVAILABLE", "N/A."}
IN_ = r"(?:INLET|\bIN\b)"
OUT_ = r"(?:OUTLET|\bOUT\b)"

SECTION_RX = [
    ("inj", r"INJECTION"), ("cw_t", r"COOLING WATER.*TEMP"), ("exh_p", r"^EXH\w*\.?\s*PRESS"), ("exh", r"^EXH\w*\.?\s*GAS"),
    ("lo_t", r"^(?:LUB\w*|L\.?\s?O\.?)(?: OIL)?\s*TEMP"), ("lo_p", r"^(?:LUB\w*|L\.?\s?O\.?)(?: OIL)?\s*PRESS"),
    ("ba_t", r"BOOST AIR.*TEMP"), ("boost_p", r"BOOST AIR.*PRESS"), ("cw_p", r"COOLING WATER.*PRESS"),
    ("fuel", r"^MDO|HFO$|^FUEL (?:OIL )?SYSTEM"), ("hfo_tank", r"SERVICE TANK"),
    ("pmax", r"COMBUSTION PRESS|\bPMAX\b|\bP MAX\b"), ("rack", r"RACK"), ("gen_brg", r"(?:GENERATOR|ALTERNATOR) BEARING"),
    ("smoke", r"SMOKE"),
]
LABEL_RX: dict[str, list[tuple[str, str]]] = {
    "cw_t": [("cw_jacket_in", rf"JACKET.*{IN_}"), ("cw_engine_out", rf"(?:ENGINE|JACKET).*{OUT_}"),
             ("cw_ac_in", rf"AIR COOLER.*{IN_}"), ("cw_ac_out", rf"AIR COOLER.*{OUT_}"),
             ("cw_lo_in", rf"(?:L\.?\s?O|LUB).*{IN_}"), ("cw_lo_out", rf"(?:L\.?\s?O|LUB).*{OUT_}")],
    "lo_t": [("lo_t_in", rf"(?:L\.?\s?O|LUB|COOLER).*{IN_}"), ("lo_t_out", rf"(?:L\.?\s?O|LUB|COOLER).*{OUT_}")],
    "ba_t": [("ba_tc_in", r"T/?C"), ("ba_ac_in", r"AIR COOLER"), ("ba_cyl_in", r"CYLINDER")],
    "lo_p": [("lo_p_engine", r"ENGINE"), ("lo_p_tc", r"T/?C|TURBO")],
    "cw_p": [("cw_p_jacket", r"JACKET|ENGINE"), ("cw_p_cooler", r"COOLER")],
    "hfo_tank": [("hfo_tank_t", r"HEAT|TEMP")],
    "gen_brg": [("gen_brg_de", r"ENGINE|DRIVE|\bDE\b"), ("gen_brg_nde", r"OPPOSITE|\bNDE\b|NON")],
    "fuel": [("fuel_mode", r"MDO\s*/?\s*HFO|^FUEL"), ("fo_p_in", r"PRESS"), ("fo_t_in", r"TEMP.*ENG|ENG.*TEMP|^TEMP"),
             ("fo_visc_heater_out", r"VIS.*(?:HEATER|OUT)"), ("fo_visc_set", r"VIS.*SET|VIS.*CST")],
}
SINGLE_SECTIONS = {"boost_p": "boost_p", "exh_p": "exh_p", "smoke": "smoke"}
CYL_ROW = re.compile(r"^(?:NO|CYL(?:INDER)?|#)\.?\s*(\d{1,2})$")

HDR_LEFT = [("vessel", r"^SHIP'?S NAME"), ("maker", r"^D/?G MAKER|^ENGINE MAKER|^MAKER"), ("type", r"^D/?G TYPE|^ENGINE TYPE"),
            ("output", r"^OUTPUT"), ("rpm", r"^RPM"), ("run_hours", r"^RUN(?:NING)?\.? HOURS|^R/?H\b"), ("datetime", r"^DATE")]
HDR_RIGHT = [("fuel_used", r"^FUEL USED"), ("density", r"^DENSITY"), ("viscosity", r"^VISCOSITY"),
             ("fo_cons", r"^FUEL OIL CONS|^F\.?O\.? CONS"), ("lo_cons", r"^LUB\w* OIL CONS|^L\.?O\.? CONS"),
             ("ambient_t", r"^AMBIENT"), ("sea_t", r"^SEA WATER")]
DG_NO = re.compile(r"^D/?G\s*(?:#|NO\.?|NUMBER)?\s*(\d+)\b")
FURNITURE = [r"^MONTHLY AUX", r"^TO BE SENT", r"^NOTE\s*\d", r"^SHOP TRIALS?$", r"^PREPARED BY", r"^CHECKED BY", r"^APPROVED BY",
             r"^(?:O?C|°C|DEG\.? ?C|MPA|BAR|CST|MM|KW|A|%|KGR?/CM2|KG/CM2)$", r"^KW\s*/\s*A\s*/\s*%$", r"^KW\s*/\s*A$"]
PUNCT_ONLY = re.compile(r"^[\W_]+$")


# ───────────────────────────── small helpers ─────────────────────────────
def nt(v: Any) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip().upper()


def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def addr(r: int, c: int) -> str:
    return f"{col_letter(c)}{r}"


def _clean(v: Any):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return None if f != f or f in (float("inf"), float("-inf")) else f
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return v
    s = str(v).replace("\xa0", " ").replace("​", "").strip()
    return s or None


def decimals(v: float) -> int:
    s = repr(float(v))
    if "e" in s or "E" in s:
        return 4
    if "." not in s:
        return 0
    return min(len(s.split(".")[1].rstrip("0")), 4)


def to_num(v: Any) -> tuple[Optional[float], str]:
    """-> (number | None, kind) with kind in num / na / text / blank."""
    if v is None:
        return None, "blank"
    if isinstance(v, float):
        return v, "num"
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return None, "text"
    s = str(v).strip()
    if not s:
        return None, "blank"
    if nt(s) in NA_TOKENS:
        return None, "na"
    t = s.replace(" ", "")
    if re.fullmatch(r"[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", t):
        return float(t.replace(",", "")), "num"
    if re.fullmatch(r"[-+]?\d+,\d+", t):
        return float(t.replace(",", ".")), "num"
    if re.fullmatch(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", t):
        return float(t), "num"
    m = re.match(r"^\s*([-+]?\d+(?:[.,]\d+)?)\s*(\D*)$", s)
    if m and not re.search(r"\d", m.group(2)):
        rest = m.group(2).strip()
        if len(rest) <= 12:
            return float(m.group(1).replace(",", ".")), "num"
    return None, "text"


def _tier(calc: float, doc: float, dec: int) -> tuple[float, str]:
    """Compare our mean with the printed Average (allow rounding / truncation)."""
    step = 10.0 ** (-dec)
    d = abs(calc - doc)
    if d <= 0.5 * step + 1e-9:
        return 1.0, "match"
    if abs((int(calc / step) * step) - doc) < 1e-9 or d < step - 1e-9:      # truncated instead of rounded
        return 1.0, "match"
    rel = d / max(abs(doc), 1e-9)
    if d <= 1.5 * step + 1e-9 or rel <= 0.01:
        return 0.75, "near"
    if rel <= 0.03:
        return 0.4, "weak"
    return 0.0, "mismatch"


_MONTHS = {m: i for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1)}


def parse_datetime(v: Any) -> tuple[Optional[dt.date], Optional[str]]:
    """'MAY 29, 2026/1300H', '29/05/2026 13:00', datetime cells ... -> (date, 'HH:MM')"""
    if isinstance(v, dt.datetime):
        return v.date(), v.strftime("%H:%M") if (v.hour or v.minute) else None
    if isinstance(v, dt.date):
        return v, None
    if v is None:
        return None, None
    s = str(v).strip()
    d = None
    try:
        m = re.search(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", s)
        if m and m.group(1)[:3].upper() in _MONTHS:
            d = dt.date(int(m.group(3)), _MONTHS[m.group(1)[:3].upper()], int(m.group(2)))
        if d is None:
            m = re.search(r"(\d{1,2})(?:st|nd|rd|th)?[\s./-]+([A-Za-z]{3,9})\.?[\s./,-]+(\d{2,4})", s)
            if m and m.group(2)[:3].upper() in _MONTHS:
                y = int(m.group(3))
                d = dt.date(y + 2000 if y < 100 else y, _MONTHS[m.group(2)[:3].upper()], int(m.group(1)))
        if d is None:
            m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
            if m:
                y = int(m.group(3))
                d = dt.date(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(1)))     # day first (European sheets)
    except ValueError:
        d = None
    tm = None
    m = re.search(r"(\d{1,2})\s*[:.]\s*(\d{2})", s) or re.search(r"(?<!\d)(\d{2})(\d{2})\s*(?:H|HRS?|HOURS)\b", s, re.I)
    if m and int(m.group(1)) < 24 and int(m.group(2)) < 60:
        tm = f"{int(m.group(1)):02d}:{m.group(2)}"
    return d, tm


# ───────────────────────────── file → grid ─────────────────────────────
def _grid_xlsx(data: bytes) -> list[dict]:
    import openpyxl
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if sum(i.file_size for i in z.infolist()) > MAX_UNZIPPED:
                raise AuxParseError("The workbook expands to an unreasonable size.")
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        wf = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    except AuxParseError:
        raise
    except (zipfile.BadZipFile, KeyError, ValueError, OSError, EOFError) as e:
        raise AuxParseError(f"The file could not be opened as an Excel workbook ({type(e).__name__}).") from e
    except Exception as e:                                                                  # noqa: BLE001
        raise AuxParseError(f"The file could not be opened as an Excel workbook ({type(e).__name__}).") from e
    sheets = []
    for ws in wb.worksheets:
        if not hasattr(ws, "iter_rows"):
            continue
        cells, formulas = {}, 0
        wsf = wf[ws.title] if ws.title in wf.sheetnames else None
        for row in ws.iter_rows(min_row=1, max_row=MAX_ROWS, max_col=MAX_COLS):
            for c in row:
                v = _clean(c.value)
                if v is not None:
                    cells[(c.row, c.column)] = v
                elif wsf is not None:
                    fv = wsf.cell(row=c.row, column=c.column).value
                    if isinstance(fv, str) and fv.startswith("="):
                        formulas += 1
        sheets.append(dict(title=ws.title, cells=cells, formulas_without_value=formulas))
    return sheets


def _grid_xls(data: bytes) -> list[dict]:
    try:
        import xlrd
    except ImportError as e:                                                                # pragma: no cover
        raise AuxParseError("Legacy .xls files need the 'xlrd' package (pip install xlrd).") from e
    try:
        book = xlrd.open_workbook(file_contents=data)
    except Exception as e:                                                                  # noqa: BLE001
        msg = str(e)
        if "ncrypt" in msg or "password" in msg.lower():
            raise AuxParseError("The workbook is password-protected.") from e
        raise AuxParseError(f"The file could not be opened as an Excel workbook ({type(e).__name__}).") from e
    sheets = []
    for sh in book.sheets():
        cells = {}
        for r in range(min(sh.nrows, MAX_ROWS)):
            for c in range(min(sh.ncols, MAX_COLS)):
                ct, v = sh.cell_type(r, c), sh.cell_value(r, c)
                if ct in (0, 5, 6):                                   # empty / error / blank
                    continue
                if ct == 3:
                    try:
                        v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                    except Exception:                                                         # noqa: BLE001
                        continue
                elif ct == 4:
                    continue
                cv = _clean(v)
                if cv is not None:
                    cells[(r + 1, c + 1)] = cv
        sheets.append(dict(title=sh.name, cells=cells, formulas_without_value=0))
    return sheets


def read_workbook(data: bytes) -> list[dict]:
    if not data:
        raise AuxParseError("The file is empty.")
    if len(data) > MAX_BYTES:
        raise AuxParseError("The file is too large to be a TEC-014 report.")
    if data[:2] == b"PK":
        return _grid_xlsx(data)
    if data[:4] == b"\xd0\xcf\x11\xe0":
        return _grid_xls(data)
    raise AuxParseError("This is not an Excel workbook (.xlsx / .xlsm / .xls).")


_FINGERPRINT = [r"AUX\.? ENGINE PERFORMANCE", r"\bDG MAKER", r"EXH\w*\.? GAS TEMP", r"COMBUSTION PRESS", r"RACK POSITION", r"SHOP TRIALS?",
                r"INJECTION TIMING", r"TESTED AT MAX", r"BOOST AIR PRESS", r"GENERATOR BEARING"]


def _score_sheet(cells: dict) -> int:
    texts = [nt(v) for v in cells.values() if isinstance(v, str)]
    return sum(1 for rx in _FINGERPRINT if any(re.search(rx, t) for t in texts))


# ───────────────────────────── main parser ─────────────────────────────
def parse_tec_014(data: bytes) -> dict:
    try:
        return _parse(data)
    except AuxParseError:
        raise
    except Exception as e:                                                                  # noqa: BLE001
        raise AuxParseError(f"Unexpected problem while reading the workbook ({type(e).__name__}: {e}).") from e


def _parse(data: bytes) -> dict:
    sheets = read_workbook(data)
    if not sheets:
        raise AuxParseError("The workbook has no worksheets.")
    best = max(sheets, key=lambda s: _score_sheet(s["cells"]))
    score = _score_sheet(best["cells"])
    if score < 4:
        raise AuxParseError("No TEC-014 (aux. engine performance report) layout was found in this workbook.")
    G: dict = best["cells"]
    issues: list[tuple[str, str]] = []
    claimed: set = set()
    refs: dict[str, str] = {}

    def claim(r, c):
        claimed.add((r, c))

    def get(r, c):
        return G.get((r, c))

    rows = sorted({r for r, _ in G})
    if best["formulas_without_value"]:
        issues.append(("warn", f"{best['formulas_without_value']} formula cell(s) have no stored result (the file was never recalculated in Excel) and were read as empty."))

    # ---- column groups, anchored on the printed "SHOP TRIALS" header -----------------
    sc = None
    for (r, c), v in sorted(G.items()):
        if isinstance(v, str) and re.fullmatch(r"SHOP TRIALS?", nt(v)):
            sc = c if sc is None else min(sc, c)
    if sc is None:
        sc = 6
        issues.append(("info", "The 'SHOP TRIALS' header text is missing on this sheet - the standard column layout (D · F · I) was assumed."))
    GCOLS = {"act": [sc - 2, sc - 1], "shopA": [sc, sc + 1, sc + 2], "shopB": [sc + 3, sc + 4, sc + 5]}

    def group_value(r, g, claim_it=True):
        """first non-empty cell in the group's columns; extra different values are left unclaimed."""
        hit = [(c, G[(r, c)]) for c in GCOLS[g] if (r, c) in G]
        if not hit:
            return None, None
        c0, v0 = hit[0]
        if claim_it:
            claim(r, c0)
            for c, v in hit[1:]:
                if v == v0:
                    claim(r, c)
        return v0, addr(r, c0)

    # ---- locate the "TESTED AT MAX. SAFE LOAD" row (start of the body) ------------------
    body0 = next((r for r in rows if re.search(r"TESTED AT MAX", nt(get(r, 1)))), None)
    if body0 is None:
        body0 = next((r for r in rows if any(re.search(r"TESTED AT MAX", nt(get(r, c))) for c in (1, 2))), None)
    if body0 is None:
        raise AuxParseError("The 'TESTED AT MAX. SAFE LOAD' row was not found - this does not look like a TEC-014 report.")

    # ---- header block ------------------------------------------------------------------
    ident: dict[str, Any] = {k: None for k in ("vessel", "maker", "type", "dg_no", "serial", "output_text", "rated_kw", "rated_kva", "rpm",
                                              "run_hours", "date", "time", "datetime_raw", "prepared_by", "fuel_used", "density", "viscosity",
                                              "fo_cons", "fo_cons_mt", "lo_cons", "lo_cons_l", "ambient_t", "sea_t")}
    raw_hdr: dict[str, Any] = {}

    def next_right(r, c0, cmax=12):
        for c in range(c0 + 1, cmax + 1):
            if (r, c) in G:
                return G[(r, c)], c
        return None, None

    for r in [x for x in rows if x < body0]:
        a = nt(get(r, 1))
        if a:
            matched = False
            if "SERIAL" in a or (DG_NO.match(a) and not any(re.search(rx, a) for _, rx in HDR_LEFT)):
                m = DG_NO.match(a)
                if m:
                    ident["dg_no"] = int(m.group(1))
                    claim(r, 1)
                    v, c = next_right(r, 1, 5)
                    if v is not None:
                        ident["serial"] = str(v).strip() if not isinstance(v, float) else f"{v:g}"
                        claim(r, c)
                        refs["serial"] = addr(r, c)
                    matched = True
            if not matched:
                for key, rx in HDR_LEFT:
                    if re.search(rx, a):
                        claim(r, 1)
                        v, c = next_right(r, 1, 5)
                        if v is not None:
                            raw_hdr[key] = v
                            claim(r, c)
                            refs[key] = addr(r, c)
                        break
        for c in sorted(cc for (rr, cc) in G if rr == r and cc >= 5):
            t = nt(G[(r, c)])
            for key, rx in HDR_RIGHT:
                if re.search(rx, t):
                    claim(r, c)
                    v, c2 = next_right(r, c, 14)
                    if v is not None:
                        raw_hdr[key] = v
                        claim(r, c2)
                        refs[key] = addr(r, c2)
                    break
    ident["vessel"] = re.sub(r"^\s*M\s*/\s*V\s+|^\s*MV\s+|^\s*M\.V\.\s+", "", str(raw_hdr["vessel"]), flags=re.I).strip() if raw_hdr.get("vessel") is not None else None
    for k in ("maker", "type", "fuel_used", "viscosity"):
        if raw_hdr.get(k) is not None:
            ident[k] = str(raw_hdr[k]).strip() if not isinstance(raw_hdr[k], float) else f"{raw_hdr[k]:g}"
    # output "700 KW/ 875 KVA"
    if raw_hdr.get("output") is not None:
        o = str(raw_hdr["output"]) if not isinstance(raw_hdr["output"], float) else f"{raw_hdr['output']:g}"
        ident["output_text"] = o
        mk = re.search(r"(\d[\d,.]*)\s*KVA", o, re.I)
        if mk:
            ident["rated_kva"] = float(mk.group(1).replace(",", ""))
        mw = re.search(r"(\d[\d,.]*)\s*K\.?W(?!A)", o, re.I)
        if mw:
            ident["rated_kw"] = float(mw.group(1).replace(",", ""))
        elif not mk:
            mn = re.match(r"^\s*(\d[\d,.]*)\s*$", o)
            if mn:
                ident["rated_kw"] = float(mn.group(1).replace(",", ""))
    for k in ("rpm", "run_hours", "density", "ambient_t", "sea_t"):
        if k in raw_hdr:
            n, kind = to_num(raw_hdr[k])
            ident[k] = n
            if n is None and kind == "text":
                issues.append(("info", f"Header field '{k}' holds text ('{raw_hdr[k]}') and could not be read as a number."))
    for k, kk in (("fo_cons", "fo_cons_mt"), ("lo_cons", "lo_cons_l")):
        if k in raw_hdr:
            ident[k] = str(raw_hdr[k]).strip() if not isinstance(raw_hdr[k], float) else f"{raw_hdr[k]:g}"
            n, _ = to_num(raw_hdr[k])
            ident[kk] = n
    ident["datetime_raw"] = raw_hdr.get("datetime")
    d, tm = parse_datetime(raw_hdr.get("datetime"))
    ident["date"], ident["time"] = d, tm
    if raw_hdr.get("datetime") is not None and d is None:
        issues.append(("warn", f"The report date '{raw_hdr['datetime']}' could not be understood."))
    if isinstance(ident["datetime_raw"], (dt.datetime, dt.date)):
        ident["datetime_raw"] = ident["datetime_raw"].isoformat()

    # ---- tested load row ----------------------------------------------------------------
    test = dict(kw=None, amps=None, load_frac=None, raw=None, ref=None)
    shop = {"shopA": dict(load=None, raw=None, ref=None), "shopB": dict(load=None, raw=None, ref=None)}
    claim(body0, 1)
    if (body0, 2) in G:
        claim(body0, 2)                                    # unit hint ("kw / A / %") - or a stray label the crew typed over it
    v, ref = group_value(body0, "act")
    if v is not None:
        test["raw"], test["ref"] = str(v) if not isinstance(v, float) else f"{v:g}", ref
        tt = test["raw"]
        nums = [float(x.replace(",", "")) for x in re.findall(r"\d[\d,]*(?:\.\d+)?", tt)]
        if "%" in tt and nums:
            test["load_frac"] = nums[0] / 100
        elif nums:
            rk = ident["rated_kw"]
            if len(nums) == 1 and rk and 1 < nums[0] <= 110 and rk > 110:
                test["load_frac"] = nums[0] / 100
                issues.append(("info", f"Tested load '{tt}' has no unit - read as {nums[0]:g} % of rated."))
            else:
                test["kw"] = nums[0]
                if len(nums) > 1:
                    test["amps"] = nums[1]
    elif v is None:
        issues.append(("warn", "The tested load (kW / A / %) was not entered - load-dependent checks cannot run."))
    if test["kw"] is not None and ident["rated_kw"]:
        test["load_frac"] = test["kw"] / ident["rated_kw"]
    for g in ("shopA", "shopB"):
        v, ref = group_value(body0, g)
        if v is None:
            continue
        shop[g]["raw"], shop[g]["ref"] = v, ref
        n, _ = to_num(v)
        if n is not None and n > 0:
            if n <= 1.5:
                shop[g]["load"] = n
            elif n <= 110:
                shop[g]["load"] = n / 100
            elif ident["rated_kw"]:
                shop[g]["load"] = n / ident["rated_kw"]
        else:
            issues.append(("info", f"Shop-trial load '{v}' was not understood."))

    # ---- body rows ----------------------------------------------------------------------
    cyl: dict[str, dict[str, dict[int, Optional[float]]]] = {g: {m: {} for m in CYL_METRICS} for g in GROUPS}
    cyl_kind: dict[str, dict[str, dict[int, str]]] = {g: {m: {} for m in CYL_METRICS} for g in GROUPS}
    cyl_ref: dict[str, str] = {}
    avgs_doc: dict[str, dict[str, tuple[float, int]]] = {g: {} for g in GROUPS}
    vals: dict[str, dict[str, Optional[float]]] = {g: {} for g in GROUPS}
    kinds: dict[str, dict[str, str]] = {g: {} for g in GROUPS}
    texts: dict[str, dict[str, str]] = {g: {} for g in GROUPS}
    sections_seen: set = set()
    section = None
    footer_started = False
    footer_rows: list[int] = []

    def store_single(r, key, groups=GROUPS):
        for g in groups:
            v, ref = group_value(r, g)
            if v is None:
                continue
            n, kind = to_num(v)
            vals[g][key], kinds[g][key] = n, kind
            if kind in ("text",) or isinstance(v, str):
                texts[g][key] = str(v)
            refs[f"{g}.{key}"] = ref
            if kind == "text" and key not in ("smoke", "fuel_mode"):
                issues.append(("info", f"{PARAMS.get(key, (key,))[0]} · {g}: '{v}' is text, not a number."))

    for r in [x for x in rows if x > body0]:
        a, b = nt(get(r, 1)), nt(get(r, 2))
        if re.match(r"^(PREPARED|CHECKED|APPROVED|NOTE)", a) or re.match(r"^NOTE", b):
            footer_started = True
        if footer_started:
            footer_rows.append(r)
            continue
        if a:
            section = None
            for key, rx in SECTION_RX:
                if re.search(rx, a):
                    section = key
                    break
            claim(r, 1)
            if section is None:
                section = f"?{a}"
                issues.append(("info", f"Unknown section '{get(r, 1)}' - its rows were not mapped."))
            else:
                sections_seen.add(section)
        if section is None or str(section).startswith("?"):
            continue
        # single-row sections (value sits directly in the value columns)
        if section in SINGLE_SECTIONS and not b:
            key = SINGLE_SECTIONS[section]
            store_single(r, key)
            if (r, 3) in G:
                claim(r, 3)
            continue
        if section in SINGLE_SECTIONS and b:
            key = SINGLE_SECTIONS[section]
            store_single(r, key)
            claim(r, 2)
            if (r, 3) in G:
                claim(r, 3)
            continue
        if not b:
            continue
        # cylinder / average rows
        m = CYL_ROW.match(b)
        if section in ("inj", "exh", "pmax", "rack") and (m or b == "AVERAGE" or b == "MEAN"):
            claim(r, 2)
            if (r, 3) in G:
                claim(r, 3)
            if m:
                cid = int(m.group(1))
                for g in GROUPS:
                    v, ref = group_value(r, g)
                    if v is None:
                        continue
                    n, kind = to_num(v)
                    if cid in cyl[g][section]:
                        issues.append(("warn", f"{CYL_METRICS[section][0]}: cylinder {cid} appears twice - the first value was kept."))
                        continue
                    cyl[g][section][cid] = n
                    cyl_kind[g][section][cid] = kind
                    cyl_ref[f"{g}.{section}.{cid}"] = ref
                    if kind == "text":
                        issues.append(("warn", f"{CYL_METRICS[section][0]} · cylinder {cid} · {g}: '{v}' is not a number."))
            else:
                for g in GROUPS:
                    v, ref = group_value(r, g)
                    n, kind = to_num(v)
                    if n is not None and section in AVG_METRICS:
                        avgs_doc[g][section] = (n, decimals(n))
            continue
        # exhaust-section extras
        if section == "exh":
            claim(r, 2)
            if (r, 3) in G:
                claim(r, 3)
            if re.search(r"T/?C\s*IN\w*\s*-\s*T/?C\s*OUT|^[Δ∆]\s*[ΤT]|DELTA", b):
                for g in GROUPS:
                    v, ref = group_value(r, g)
                    if v is None:
                        continue
                    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(v))]
                    if len(nums) >= 2:
                        vals[g]["tc_pair_in"], vals[g]["tc_pair_out"] = nums[0], nums[1]
                        kinds[g]["tc_pair_in"] = kinds[g]["tc_pair_out"] = "num"
                        texts[g]["tc_pair"] = str(v)
                        refs[f"{g}.tc_pair_in"] = refs[f"{g}.tc_pair_out"] = ref
                    elif to_num(v)[1] == "na":
                        kinds[g]["tc_pair_in"] = kinds[g]["tc_pair_out"] = "na"
                        vals[g]["tc_pair_in"] = vals[g]["tc_pair_out"] = None
                    else:
                        issues.append(("info", f"The T/C ΔT row holds '{v}' - expected 'inlet/outlet'."))
                continue
            if re.search(r"T/?C\s*IN", b):
                key = "tc_in_upper" if re.search(r"UPPER|4.*6", b) else "tc_in_lower" if re.search(r"LOWER|1.*3", b) else "tc_in"
                store_single(r, key)
                continue
            if re.search(r"T/?C\s*OUT", b):
                store_single(r, "tc_out")
                continue
        # label-mapped singles
        key = None
        for k, rx in LABEL_RX.get(section, []):
            if re.search(rx, b):
                key = k
                break
        if key is None:
            continue                                       # left unclaimed -> shows in "not mapped"
        claim(r, 2)
        if key == "fuel_mode":
            if (r, 3) in G:
                texts["act"]["fuel_mode"] = str(G[(r, 3)])
                refs["act.fuel_mode"] = addr(r, 3)
                claim(r, 3)
            v, ref = group_value(r, "act")
            if v is not None:
                texts["act"]["fuel_grade"] = str(v) if not isinstance(v, float) else f"{v:g}"
                refs["act.fuel_grade"] = ref
            continue
        if (r, 3) in G:
            claim(r, 3)
        store_single(r, key)

    # ---- footer -------------------------------------------------------------------------
    for r in footer_rows:
        a = get(r, 1)
        if isinstance(a, str) and re.match(r"(?i)^PREPARED BY", a.strip()):
            claim(r, 1)
            rest = re.sub(r"(?i)^PREPARED BY\s*:?\s*", "", a.strip())
            if not rest:
                v, c = next_right(r, 1, 8)
                rest = str(v) if v is not None else ""
                if v is not None:
                    claim(r, c)
            ident["prepared_by"] = rest or None

    # ---- unit normalisation ---------------------------------------------------------------
    def _pm_unit(g):
        xs = sorted(v for v in cyl[g]["pmax"].values() if v is not None)
        if not xs:
            return None, None
        med = xs[len(xs) // 2]
        return ("MPa" if med < 30 else "kgf/cm²"), med
    gunits = {g: _pm_unit(g) for g in GROUPS}
    present = {g: u for g, (u, _m) in gunits.items() if u}
    tgt = present.get("act") or next(iter(present.values()), "kgf/cm²")
    FAC = {"MPa": 10.0, "kgf/cm²": 0.980665}
    units: dict[str, Any] = {"pmax": tgt, "pmax_factor": FAC[tgt], "fo_p_in": "bar"}
    for g, u in present.items():
        if u == "MPa" and g == "act" or (g != "act" and u == "MPa" and tgt == "MPa" and "act" not in present):
            issues.append(("info", f"Pmax is printed in 'KGR/CM2' but the values (median {gunits[g][1]:g}) are of the order of MPa - read as MPa (×10 → bar)."))
        if u != tgt:
            scale = FAC[u] / FAC[tgt]
            for i_, v_ in list(cyl[g]["pmax"].items()):
                if v_ is not None:
                    cyl[g]["pmax"][i_] = v_ * scale
            if "pmax" in avgs_doc[g]:
                avgs_doc[g]["pmax"] = (avgs_doc[g]["pmax"][0] * scale, min(avgs_doc[g]["pmax"][1] + 1, 4))
            issues.append(("info", f"Shop-trial Pmax ({g}) is in {u} while the actual column is in {tgt} - converted for the comparison."))
    # FO inlet pressure: the actual column says Mpa; shop columns usually carry bar / kgf/cm2
    for g in GROUPS:
        v = vals[g].get("fo_p_in")
        if v is None:
            continue
        if v < 1.5:
            vals[g]["fo_p_in"] = round(v * 10.0, 6)
            issues.append(("info", f"FO pressure ({'actual' if g == 'act' else g}): {v:g} read as MPa → {v * 10:g} bar."))

    # ---- cylinder bookkeeping -------------------------------------------------------------
    ids = sorted({i for m in ("exh", "pmax", "rack", "inj") for i in cyl["act"][m]})
    n_data = max(ids) if ids else 0
    n_type = None
    if ident["type"]:
        mt = re.match(r"^\s*(\d{1,2})\s*[A-Z]", str(ident["type"]).upper())
        if mt:
            n_type = int(mt.group(1))
    n_cyl = n_data
    if n_type and n_data and n_type != n_data:
        issues.append(("warn", f"The engine type '{ident['type']}' suggests {n_type} cylinders but the sheet lists {n_data}."))
    if n_data and ids != list(range(1, n_data + 1)):
        issues.append(("warn", f"Cylinder numbering has gaps ({', '.join(map(str, ids))})."))
    for m in ("inj", "exh", "pmax", "rack"):
        miss = [i for i in range(1, n_cyl + 1) if cyl["act"][m].get(i) is None]
        if miss and len(miss) < n_cyl:
            issues.append(("warn", f"{CYL_METRICS[m][0]}: no usable value for cylinder {', '.join(map(str, miss))}."))

    # ---- cross-check with the sheet's own Average rows -------------------------------------
    xval: dict[str, dict] = {}
    for g in GROUPS:
        for m in AVG_METRICS:
            if m not in avgs_doc[g]:
                continue
            xs = [v for v in cyl[g][m].values() if v is not None]
            if len(xs) < 2:
                continue
            calc = sum(xs) / len(xs)
            doc, dec = avgs_doc[g][m]
            tier, label = _tier(calc, doc, dec)
            xval[f"{g}.{m}"] = dict(doc=doc, calc=calc, dec=dec, tier=tier, label=label, n=len(xs), group=g, metric=m)
    for k, x in xval.items():
        if x["label"] == "mismatch":
            issues.append(("warn", f"{CYL_METRICS[x['metric']][0]} ({'actual' if x['group'] == 'act' else x['group']}): the printed Average {x['doc']:g} "
                                   f"does not match the mean of the cylinder values ({x['calc']:.2f}) - a mis-read value or an arithmetic slip on the sheet."))

    # ---- shop-trial sanity -------------------------------------------------------------------
    la, lb = shop["shopA"]["load"], shop["shopB"]["load"]
    have_shop = any(v is not None for g in ("shopA", "shopB") for m in AVG_METRICS for v in cyl[g][m].values())
    if have_shop and (la is None or lb is None):
        issues.append(("info", "A shop-trial column has values but no load point - the shop-trial comparison is limited."))
    if la is not None and lb is not None and abs(la - lb) < 0.05:
        issues.append(("warn", "The two shop-trial load points are (almost) identical - interpolation is not possible."))

    # ---- plausibility -------------------------------------------------------------------------
    bad_range = 0
    n_checked = 0
    for g in GROUPS:
        for k, v in vals[g].items():
            if v is None or k not in PLAUSIBLE:
                continue
            n_checked += 1
            lo, hi = PLAUSIBLE[k]
            if not (lo <= v <= hi):
                bad_range += 1
                if g == "act":
                    issues.append(("warn", f"{PARAMS[k][0]} = {v:g} {PARAMS[k][1]} is outside the physically plausible range ({lo:g} – {hi:g})."))
        for m, arr in cyl[g].items():
            for i, v in arr.items():
                if v is None:
                    continue
                n_checked += 1
                if m == "pmax":
                    vv, (lo, hi) = v * units["pmax_factor"], CYL_PLAUSIBLE["pmax_bar"]
                else:
                    vv, (lo, hi) = v, CYL_PLAUSIBLE[m]
                if not (lo <= vv <= hi):
                    bad_range += 1
                    if g == "act":
                        issues.append(("warn", f"{CYL_METRICS[m][0]} · cylinder {i} = {v:g} is outside the plausible range."))

    # ---- load plausibility --------------------------------------------------------------------
    if test["load_frac"] is not None:
        if test["load_frac"] > 1.15:
            issues.append(("warn", f"The tested load ({test['kw']:g} kW) is {test['load_frac'] * 100:.0f} % of the rated output - check the entry."))
        elif test["load_frac"] < 0.05:
            issues.append(("warn", "The tested load is below 5 % of rated - check the entry."))
    if ident["rated_kw"] is None:
        issues.append(("warn", "The rated output could not be read - the load percentage cannot be computed."))

    # ---- blank detection & status --------------------------------------------------------------
    n_cyl_vals = sum(1 for m in ("exh", "pmax", "rack", "inj") for v in cyl["act"][m].values() if v is not None)
    n_single_vals = sum(1 for v in vals["act"].values() if v is not None)
    status = "ok"
    if n_cyl_vals == 0 and n_single_vals == 0:
        status = "blank"
    elif n_cyl == 0:
        status = "no_cylinders"
        issues.append(("error", "No cylinder rows ('No. 1', 'No. 2' ...) were found in the performance sections."))

    # ---- coverage audit --------------------------------------------------------------------------
    def is_furniture(t: str) -> bool:
        return bool(PUNCT_ONLY.match(t)) or any(re.match(p, t) for p in FURNITURE) or t in NA_TOKENS
    tot = cl_n = fur = 0
    unclaimed = []
    for (r, c), v in sorted(G.items()):
        tot += 1
        if (r, c) in claimed:
            cl_n += 1
            continue
        t = nt(v)
        if is_furniture(t):
            fur += 1
            continue
        left = next((nt(G[(r, cc)]) for cc in range(c - 1, 0, -1) if (r, cc) in G), "")
        above = next((nt(G[(rr, c)]) for rr in range(r - 1, max(r - 6, 0), -1) if (rr, c) in G), "")
        unclaimed.append(dict(cell=addr(r, c), row=r, text=str(v)[:160], left=left[:50], above=above[:50]))
    coverage = dict(total=tot, claimed=cl_n, template=fur, unmapped=len(unclaimed),
                    pct=round(100.0 * (cl_n + fur) / tot, 1) if tot else 100.0)

    # ---- confidence -----------------------------------------------------------------------------------
    expected_sections = {"inj", "cw_t", "lo_t", "exh", "lo_p", "fuel", "pmax", "rack", "boost_p", "gen_brg"}
    hdr_found = sum(1 for k in ("vessel", "maker", "type", "rated_kw", "date", "dg_no") if ident.get(k) not in (None, ""))
    structure = 0.55 * len(expected_sections & sections_seen) / len(expected_sections) + 0.45 * hdr_found / 6
    if n_cyl:
        exp_c = n_cyl * 4
        cyl_comp = sum(1 for m in CYL_METRICS for i in range(1, n_cyl + 1) if cyl["act"][m].get(i) is not None) / exp_c
    else:
        cyl_comp = 0.0
    single_keys = [k for k in PARAMS if k not in ("tc_in_upper", "tc_in_lower", "tc_in", "tc_pair_in", "tc_pair_out", "exh_p", "tc_out")]
    answered = sum(1 for k in single_keys if kinds["act"].get(k) in ("num", "na") or (k in vals["act"] and vals["act"][k] is not None))
    completeness = 0.7 * cyl_comp + 0.3 * answered / len(single_keys)
    plaus = 1.0 if n_checked == 0 else 1.0 - bad_range / n_checked
    wsum = nsum = 0.0
    for k, x in xval.items():
        w = 2.0 if x["group"] == "act" else 1.0
        wsum += w
        nsum += w * x["tier"]
    crossval = nsum / wsum if wsum else 0.6
    total = 100 * (0.20 * structure + 0.30 * completeness + 0.15 * plaus + 0.35 * crossval)
    warns = sum(1 for s, _ in issues if s == "warn")
    total = max(0.0, min(100.0, total - min(8.0, 2.0 * warns)))
    if status == "blank":
        conf = dict(total=0.0, structure=0.0, completeness=0.0, plausibility=0.0, crossval=0.0, level="blank", label="Blank form", verified_by_sheet=False)
    else:
        level, label = (("verified", "Verified") if total >= 97 else ("high", "High confidence") if total >= 90
                        else ("review", "Review recommended") if total >= 75 else ("low", "Low confidence"))
        conf = dict(total=round(total, 1), structure=round(structure, 3), completeness=round(completeness, 3), plausibility=round(plaus, 3),
                    crossval=round(crossval, 3), level=level, label=label, verified_by_sheet=bool(wsum))

    return dict(
        parser_version=PARSER_VERSION, sheet=best["title"], status=status, identity=ident, test=test, shop=shop, units=units,
        cyl=cyl, cyl_kind=cyl_kind, cyl_ref=cyl_ref, values=vals, kinds=kinds, texts=texts, refs=refs,
        avgs_doc=avgs_doc, xval=xval, n_cylinders=n_cyl, n_type=n_type, issues=issues, unclaimed=unclaimed, coverage=coverage,
        confidence=conf, sections=sorted(sections_seen), columns={g: [col_letter(c) for c in cs] for g, cs in GCOLS.items()},
    )


# ====================================================================================================
# 2. DIAGNOSIS
#    input checks  →  shop-trial reference  →  engine findings  →  probable causes  →  verdict  →  crew action list
# ====================================================================================================
#  Every finding is a dict:
#     sev        critical | mechanical | warning | advisory
#     cat        engine | reference | diagnosis | load   (the engine itself)      data  (the sheet - "check the entry")
#     title      short heading                 fault     what was seen, with the figures
#     causes     likely causes (list)          actions   what the crew should do (list)
#     when       now | soon | planned | retest | recheck        cyl   cylinder number or None
#     uses       readings the finding is based on              provisional   True if one of them failed a consistency check
# ----------------------------------------------------------------------------------------------------

SEV_RANK = {"critical": 5, "mechanical": 4, "warning": 3, "advisory": 2, "nominal": 1}
SEV_LABEL = {"critical": "Critical", "mechanical": "Mechanical", "warning": "Warning", "advisory": "Advisory", "nominal": "Nominal"}
SEV_CHIP = {"critical": "crit", "mechanical": "orange", "warning": "warn", "advisory": "info", "nominal": "ok"}

# when to act: key -> (sort order, label)
URGENCY = {"recheck": (0, "Re-check the entries first"), "now": (1, "Act now"), "soon": (2, "This week / before the next port"),
           "planned": (3, "At the next planned maintenance"), "retest": (4, "At the next monthly test")}
URGENCY_BY_SEV = {"critical": "now", "mechanical": "soon", "warning": "soon", "advisory": "retest"}

# ───────────────────────────── settings (editable in the sidebar) ─────────────────────────────
TH_DEFAULT: dict[str, Any] = dict(
    low_load=0.50,            # test below this fraction of rated output is flagged
    exh_dev=25.0, exh_dev_crit=40.0, exh_abs=480.0, exh_abs_crit=520.0,
    pmax_dev=4.0, pmax_dev_warn=7.0, rack_dev=6.0, rack_dev_warn=10.0, inj_spread=0.5, inj_spread_warn=1.0,
    jacket_out_max=85.0, lo_out_max=75.0, boost_air_max=55.0, lo_p_min=3.0, lo_p_tc_min=1.5, brg_max=80.0,
    ref_exh=30.0, ref_pmax=8.0, ref_boost=12.0, ref_rack=12.0,
    shop_mode="auto",         # auto = a typed shop-trial load close to a standard point is read as that point | printed = use as typed
    shop_points="50, 75",     # the % loads at which the maker's shop trials were run
)
TH_SPECS = [  # key, label, step, format
    ("low_load", "Low-load test below (fraction of rated)", 0.05, "%.2f"),
    ("exh_dev", "Exhaust deviation from mean (± °C)", 1.0, "%.0f"), ("exh_dev_crit", "Exhaust deviation, severe (± °C)", 1.0, "%.0f"),
    ("exh_abs", "Exhaust absolute, warning (°C)", 5.0, "%.0f"), ("exh_abs_crit", "Exhaust absolute, critical (°C)", 5.0, "%.0f"),
    ("pmax_dev", "Pmax deviation, advisory (± %)", 0.5, "%.1f"), ("pmax_dev_warn", "Pmax deviation, warning (± %)", 0.5, "%.1f"),
    ("rack_dev", "Fuel rack deviation, advisory (± %)", 0.5, "%.1f"), ("rack_dev_warn", "Fuel rack deviation, warning (± %)", 0.5, "%.1f"),
    ("inj_spread", "Injection timing spread, advisory (°)", 0.1, "%.1f"), ("inj_spread_warn", "Injection timing spread, warning (°)", 0.1, "%.1f"),
    ("jacket_out_max", "Jacket CW outlet max (°C)", 1.0, "%.0f"), ("lo_out_max", "L.O. cooler outlet max (°C)", 1.0, "%.0f"),
    ("boost_air_max", "Boost air at cylinder inlet max (°C)", 1.0, "%.0f"), ("lo_p_min", "L.O. pressure (engine) min (kgf/cm²)", 0.1, "%.1f"),
    ("lo_p_tc_min", "L.O. pressure (T/C) min (kgf/cm²)", 0.1, "%.1f"), ("brg_max", "Generator bearing max (°C)", 1.0, "%.0f"),
    ("ref_exh", "vs shop trial · exhaust (± °C)", 1.0, "%.0f"), ("ref_pmax", "vs shop trial · Pmax (± %)", 0.5, "%.1f"),
    ("ref_boost", "vs shop trial · boost (± %)", 0.5, "%.1f"), ("ref_rack", "vs shop trial · rack (± %)", 0.5, "%.1f"),
]
MAX_EXTRAPOLATION = 0.10      # beyond this distance (fraction of rated load) outside the shop-trial range no reference is given
SHOP_SNAP_TOLERANCE = 0.12    # a typed shop-trial load this close to a standard point is read as that point


def _ok(x) -> bool:
    return x is not None and x == x


def _mean(xs) -> Optional[float]:
    xs = [x for x in xs if _ok(x)]
    return sum(xs) / len(xs) if xs else None


def _spread(xs) -> Optional[float]:
    xs = [x for x in xs if _ok(x)]
    return (max(xs) - min(xs)) if len(xs) >= 2 else None


def _pct(x: float) -> str:
    return f"{x * 100:.0f} %"


def _g(x: float) -> str:
    """A reading for a sentence: at most one decimal."""
    return f"{x:.1f}".rstrip("0").rstrip(".")


def parse_points(text) -> list[float]:
    """'50, 75' -> [0.5, 0.75].  Unreadable input falls back to the default."""
    out = []
    for tok in re.split(r"[;,/\s]+", str(text or "")):
        try:
            x = float(tok)
        except ValueError:
            continue
        x = x / 100 if x > 1.5 else x
        if 0.05 <= x <= 1.3:
            out.append(round(x, 4))
    return sorted(set(out)) or [0.5, 0.75]


# ───────────────────────────── per-cylinder table ─────────────────────────────
def cyl_frame(p: dict, group: str = "act") -> pd.DataFrame:
    f = p["units"]["pmax_factor"]
    rows = []
    for i in range(1, p["n_cylinders"] + 1):
        row = {"id": i, **{m: p["cyl"][group][m].get(i) for m in ("inj", "exh", "pmax", "rack")}}
        row["pmax_bar"] = row["pmax"] * f if row["pmax"] is not None else None
        rows.append(row)
    df = pd.DataFrame(rows, columns=["id", "inj", "exh", "pmax", "pmax_bar", "rack"])
    for m in ("inj", "exh", "pmax", "pmax_bar", "rack"):
        df[m] = pd.to_numeric(df[m], errors="coerce")
    return df


def load_fraction(p: dict) -> Optional[float]:
    return p["test"].get("load_frac")


# ───────────────────────────── shop-trial load points ─────────────────────────────
def shop_loads(p: dict, th: dict) -> dict:
    """The load each shop-trial block was really run at.

    The crew types the two load points by hand, and they are often wrong - typically the engine's own test load is typed
    there (59 %) instead of the load of the maker's shop trial (50 %).  In 'auto' mode a typed value that is close to one of
    the standard points is read as that point; the app always shows what was typed and what was used."""
    L = load_fraction(p)
    std = parse_points(th["shop_points"])
    out = {}
    for g in ("shopA", "shopB"):
        printed = p["shop"][g]["load"]
        d = dict(printed=printed, used=printed, std=None, note="", status="missing" if printed is None else "as_printed")
        if printed is not None:
            near = min(std, key=lambda s: abs(s - printed))
            gap = abs(near - printed)
            if gap <= 0.005:
                d.update(status="standard", std=near)
            elif th["shop_mode"] != "auto":
                d["note"] = f"typed {_pct(printed)}, which is not a standard shop-trial point (nearest {_pct(near)}) - used as typed"
            elif gap <= SHOP_SNAP_TOLERANCE:
                d.update(status="corrected", used=near, std=near, note=f"typed {_pct(printed)}, read as the {_pct(near)} shop-trial point")
                if L is not None and abs(printed - L) <= 0.03:
                    d["note"] += " (the typed value equals this engine's own tested load - it was probably copied from there)"
            else:
                d.update(status="non_standard", note=f"{_pct(printed)} is not one of the standard points ({', '.join(f'{s * 100:.0f}' for s in std)} %) - used as typed")
        out[g] = d
    a, b = out["shopA"], out["shopB"]
    if a["used"] is not None and b["used"] is not None and abs(a["used"] - b["used"]) < 0.05 and (a["status"] == "corrected" or b["status"] == "corrected"):
        for d in (a, b):                                    # both blocks would collapse onto one point: keep what was typed
            d.update(used=d["printed"], status="non_standard", note="both blocks map to the same standard point - used as typed")
    return out


# ───────────────────────────── shop-trial reference ─────────────────────────────
class RefSpec(NamedTuple):
    key: str
    label: str
    unit: Optional[str]          # None -> the Pmax unit of the sheet
    src: str                     # "cyl:<metric>" (mean of the cylinders) or "single:<reading>"
    tol_kind: str                # abs | pct
    th_key: Optional[str]        # threshold that turns a deviation into a finding (None = shown only)
    core: bool                   # core parameters drive the findings
    trend: str                   # rise = must not fall as load rises | steady = hardly changes with load | free
    swing: float = 0.0           # steady: largest believable change between the two blocks (°C, or a fraction for pressures)


REF_SPECS = [
    RefSpec("exh_mean", "Exhaust gas temperature (mean)", "°C", "cyl:exh", "abs", "ref_exh", True, "rise"),
    RefSpec("pmax_mean", "Pmax (mean)", None, "cyl:pmax", "pct", "ref_pmax", True, "rise"),
    RefSpec("rack_mean", "Fuel rack (mean)", "mm", "cyl:rack", "pct", "ref_rack", True, "rise"),
    RefSpec("boost_p", "Boost air pressure", "kgf/cm²", "single:boost_p", "pct", "ref_boost", True, "rise"),
    RefSpec("tc_out", "T/C outlet temperature", "°C", "single:tc_out", "abs", None, False, "rise"),
    RefSpec("ba_cyl_in", "Boost air · cylinder inlet", "°C", "single:ba_cyl_in", "abs", None, False, "steady", 12),
    RefSpec("cw_jacket_in", "Jacket CW inlet", "°C", "single:cw_jacket_in", "abs", None, False, "steady", 15),
    RefSpec("cw_engine_out", "Engine CW outlet", "°C", "single:cw_engine_out", "abs", None, False, "steady", 15),
    RefSpec("cw_ac_in", "Air cooler CW inlet", "°C", "single:cw_ac_in", "abs", None, False, "steady", 12),
    RefSpec("lo_t_out", "L.O. cooler outlet", "°C", "single:lo_t_out", "abs", None, False, "steady", 15),
    RefSpec("lo_p_engine", "L.O. pressure · engine", "kgf/cm²", "single:lo_p_engine", "pct", None, False, "steady", 0.35),
    RefSpec("lo_p_tc", "L.O. pressure · T/C", "kgf/cm²", "single:lo_p_tc", "pct", None, False, "steady", 0.35),
    RefSpec("fo_p_in", "FO pressure, engine inlet", "bar", "single:fo_p_in", "pct", None, False, "steady", 0.35),
    RefSpec("fo_t_in", "FO temperature, engine inlet", "°C", "single:fo_t_in", "abs", None, False, "steady", 20),
    RefSpec("cw_p_jacket", "Cooling water pressure · jacket", "kgf/cm²", "single:cw_p_jacket", "pct", None, False, "steady", 0.35),
    RefSpec("gen_brg_de", "Generator bearing · engine side", "°C", "single:gen_brg_de", "abs", None, False, "free"),
    RefSpec("gen_brg_nde", "Generator bearing · opposite side", "°C", "single:gen_brg_nde", "abs", None, False, "free"),
]


def _src_value(p: dict, group: str, src: str) -> Optional[float]:
    kind, key = src.split(":")
    if kind == "cyl":
        return _mean(list(p["cyl"][group][key].values()))
    return p["values"][group].get(key)


def _block_problem(spec: RefSpec, v1: float, v2: float) -> str:
    """Why the two shop-trial values (low-load point -> high-load point) cannot both be right, or ''."""
    if spec.trend == "rise" and v2 < v1 - 0.02 * abs(v1):
        return "falls as the load rises"
    if spec.trend == "steady":
        change = abs(v2 - v1) if spec.tol_kind == "abs" else abs(v2 - v1) / max(abs(v1), abs(v2), 1e-9)
        if change > spec.swing:
            return f"changes by {abs(v2 - v1):.3g} between the two load points"
    return ""


def shop_reference(p: dict, th: dict, loads: dict) -> list[dict]:
    """One row per comparable parameter: actual, both shop points, the reference at the tested load, deviation, implied load."""
    L = load_fraction(p)
    la, lb = loads["shopA"]["used"], loads["shopB"]["used"]
    out = []
    for s in REF_SPECS:
        a = _src_value(p, "act", s.src)
        sa, sb = _src_value(p, "shopA", s.src), _src_value(p, "shopB", s.src)
        if a is None and sa is None and sb is None:
            continue
        row = dict(key=s.key, label=s.label, unit=s.unit if s.unit is not None else p["units"]["pmax"], act=a, shopA=sa, shopB=sb, ref=None, delta=None,
                   delta_pct=None, implied=None, method="none", core=s.core, tol_kind=s.tol_kind, tol=th.get(s.th_key) if s.th_key else None, note="")
        pts = [(l, v) for l, v in ((la, sa), (lb, sb)) if l is not None and v is not None]
        if L is None:
            row["note"] = "tested load unknown"
        elif len(pts) == 2 and abs(pts[0][0] - pts[1][0]) >= 0.05:
            (l1, v1), (l2, v2) = sorted(pts)
            problem = _block_problem(s, v1, v2)
            if problem:
                row["method"], row["note"] = "inconsistent", f"shop-trial value {problem} - reference not used"
            else:
                ref = v1 + (v2 - v1) * (L - l1) / (l2 - l1)
                far = max(l1 - L, L - l2, 0.0)
                if far <= 1e-9:
                    row["ref"], row["method"] = ref, "interpolated"
                elif far <= MAX_EXTRAPOLATION:
                    row["ref"], row["method"] = ref, "extrapolated"
                    row["note"] = f"load {_pct(L)} lies {far * 100:.0f} points outside the shop-trial range {_pct(l1)}-{_pct(l2)}"
                else:
                    row["method"] = "out_of_range"
                    row["note"] = f"load {_pct(L)} is {far * 100:.0f} points outside the shop-trial range {_pct(l1)}-{_pct(l2)} - no reference"
                if s.core and s.trend == "rise" and a is not None and v2 != v1:        # which shop-trial load would this reading correspond to?
                    row["implied"] = l1 + (a - v1) / (v2 - v1) * (l2 - l1)
        elif pts:
            l1, v1 = min(pts, key=lambda t: abs(t[0] - L))
            if abs(L - l1) <= 0.12:
                row["ref"], row["method"] = v1, "single"
                row["note"] = f"single shop-trial point at {_pct(l1)}"
            else:
                row["note"] = "shop-trial point too far from the tested load"
        if row["ref"] is not None and a is not None:
            row["delta"] = a - row["ref"]
            row["delta_pct"] = (a - row["ref"]) / abs(row["ref"]) * 100 if row["ref"] else None
        out.append(row)
    return out


# ───────────────────────────── analysis context ─────────────────────────────
LABELS = {**{k: v[0] for k, v in PARAMS.items()}, "exh_mean": "Mean cylinder exhaust temperature", "ambient_t": "Ambient temperature",
          "sea_t": "Sea water temperature"}


SYSTEM_BY_TOPIC = {"exhaust": "combustion", "pmax": "combustion", "smoke": "combustion", "rack": "fuel", "timing": "fuel", "fuel": "fuel", "air": "air",
                   "cooling": "cooling", "lube": "lube", "generator": "gen", "load": "gen", "reference": "shop"}


class Ctx:
    """Everything the checks need, plus the list they add findings to."""

    def __init__(self, p: dict, th: dict):
        self.p, self.th = p, th
        self.v = p["values"]["act"]
        self.i = p["identity"]
        self.L = load_fraction(p)
        self.df = cyl_frame(p)
        d, nan = self.df, float("nan")
        self.m = dict(exh=_mean(d["exh"]), pmax=_mean(d["pmax"]), rack=_mean(d["rack"]), inj=_mean(d["inj"]))
        d["exh_dev"] = d["exh"] - self.m["exh"] if self.m["exh"] is not None else nan
        d["pmax_dev_pct"] = (d["pmax"] - self.m["pmax"]) / self.m["pmax"] * 100 if self.m["pmax"] else nan
        d["rack_dev_pct"] = (d["rack"] - self.m["rack"]) / self.m["rack"] * 100 if self.m["rack"] else nan
        d["inj_dev"] = d["inj"] - self.m["inj"] if self.m["inj"] is not None else nan
        self.loads: dict = {}
        self.ref: list[dict] = []
        self.F: list[dict] = []
        self.checks_run = 0

    def reading(self, key: str) -> Optional[float]:
        if key == "exh_mean":
            return self.m["exh"]
        if key in ("ambient_t", "sea_t"):
            return self.i.get(key)
        return self.v.get(key)

    def cells(self, *keys: str) -> str:
        refs = self.p["refs"]
        return ", ".join(dict.fromkeys(c for c in (refs.get(f"act.{k}") or refs.get(k) for k in keys) if c))

    def usable_ref(self, key: str) -> Optional[dict]:
        r = next((x for x in self.ref if x["key"] == key), None)
        return r if r and r["delta"] is not None and r["method"] in ("interpolated", "extrapolated", "single") else None

    def add(self, sev, title, fault, actions, causes=(), cyl=None, topic="engine", when=None, uses=(), cells="", system=None, rows=()) -> None:
        cat = topic if topic in ("data", "diagnosis", "reference", "load") else "engine"
        system = "entries" if cat == "data" else system or SYSTEM_BY_TOPIC.get(topic, "combustion")
        self.F.append(dict(sev=sev, title=title, fault=fault, actions=list(actions), causes=list(causes), cyl=cyl, topic=topic, cat=cat, system=system, uses=tuple(uses),
                           cells=cells, rows=[tuple(r) for r in rows], when=when or ("recheck" if cat == "data" else URGENCY_BY_SEV[sev]), provisional=False))

    def flag(self, sev, title, rows, steps, uses=(), cells="") -> None:
        """An entry on the sheet that cannot be right. rows = [(entry, typed, expected / why)]."""
        fault = "\n".join(f"{e}: {t} - {w}" for e, t, w in rows)
        self.add(sev, f"Check entry · {title}", fault, [*steps, FIX_SHEET], topic="data",
                 uses=uses, cells=cells or self.cells(*uses), rows=rows)


SAME_TEST = "Repeat the test at the same load to confirm."
FIX_SHEET = "Correct the sheet and re-upload it."


def _smoke_class(text: Optional[str]) -> str:
    t = (text or "").upper()
    if not t.strip():
        return "unknown"
    for cls, rx in (("black", r"BLACK|DARK|HEAVY|THICK|SOOT"), ("blue", r"BLUE"), ("white", r"WHITE|GREY|GRAY"),
                    ("light", r"HAZY|LIGHT|CLEAR|CLEAN|INVISIBLE|NIL|NONE|SLIGHT"), ("brown", r"BROWN")):
        if re.search(rx, t):
            return cls
    return "unknown"


# ───────────────────────────── A. input checks ("is the sheet itself believable?") ─────────────────────────────
# (group, hot reading, cold reading, minimum gap, why).  The hot reading must exceed the cold reading by at least the gap.
# A gap below zero is a tolerance: the 'hot' reading may be at most that much below the 'cold' one.
PAIR_RULES = [
    ("jacket", "cw_engine_out", "cw_jacket_in", 1, "water must leave warmer than it enters"),
    ("airc_w", "cw_ac_out", "cw_ac_in", 1, "water must leave warmer than it enters"),
    ("loc_w", "cw_lo_out", "cw_lo_in", 1, "water must leave warmer than it enters"),
    ("loc_o", "lo_t_in", "lo_t_out", 1, "oil must enter hotter than it leaves"),
    ("loc_o", "lo_t_out", "cw_lo_in", 0, "oil cannot leave colder than its cooling water"),
    ("air", "ba_ac_in", "ba_cyl_in", 1, "air must enter the cooler hotter than it leaves"),
    ("air", "ba_cyl_in", "cw_ac_in", 2, "air cannot be cooled to the cooling-water temperature"),
    ("sea", "cw_ac_in", "sea_t", 0, "cooling water cannot be colder than the sea"),
    ("sea", "cw_lo_in", "sea_t", 0, "cooling water cannot be colder than the sea"),
    ("tc", "exh_mean", "tc_out", 30, "T/C outlet must be well below cylinder exhaust"),
    ("tc", "exh_mean", "tc_pair_out", 30, "T/C outlet must be well below cylinder exhaust"),
    ("tc", "tc_pair_in", "tc_pair_out", 60, "turbine drop is normally above 60 °C"),
    ("tc", "tc_in", "tc_out", 60, "turbine drop is normally above 60 °C"),
    *[("tc", "exh_mean", k, -30, "T/C inlet cannot be much hotter than cylinder exhaust") for k in ("tc_pair_in", "tc_in", "tc_in_upper", "tc_in_lower")],
    ("bearing", "gen_brg_de", "ambient_t", -3, "a running bearing cannot be colder than the air"),
    ("bearing", "gen_brg_nde", "ambient_t", -3, "a running bearing cannot be colder than the air"),
]
PAIR_GROUPS = {   # group -> (title, what the crew should do)
    "jacket": ("Jacket cooling-water temperatures", ["Re-read jacket water in / out at steady load."]),
    "airc_w": ("Air-cooler water temperatures", ["Re-read air-cooler water in / out."]),
    "loc_w": ("L.O.-cooler water temperatures", ["Re-read L.O.-cooler water in / out."]),
    "loc_o": ("L.O.-cooler oil temperatures", ["Re-read L.O. before / after the cooler and the cooling-water inlet."]),
    "air": ("Charge-air temperatures", ["Re-read charge air after the cooler and the cooler water inlet; check the sensor."]),
    "sea": ("Cooling water versus sea water", ["Re-read the cooling-water and sea-water temperatures."]),
    "tc": ("Turbocharger and exhaust temperatures", ["Re-read T/C in / out and cylinder exhaust (same load, same time).",
                                                    "If confirmed: small turbine drop = fouled turbine / nozzle ring or exhaust leak. Inspect and water-wash."]),
    "bearing": ("Generator bearing temperatures", ["Check the bearing sensors and re-read them."]),
}


def check_pairs(c: Ctx) -> None:
    """Physical direction checks between two readings (cooler in/out, air before/after cooler, T/C versus cylinders ...)."""
    bad: dict[str, list] = {}
    for grp, hot, cold, gap_min, why in PAIR_RULES:
        a, b = c.reading(hot), c.reading(cold)
        if a is None or b is None:
            continue
        c.checks_run += 1
        gap = a - b
        if gap >= gap_min:
            continue
        bad.setdefault(grp, []).append(dict(hot=hot, cold=cold, a=a, b=b, why=why, impossible=gap < min(gap_min, 0)))
    for grp, items in bad.items():
        title, steps = PAIR_GROUPS[grp]
        rows = [(f"{LABELS[x['hot']]} vs {LABELS[x['cold']]}", f"{_g(x['a'])} vs {_g(x['b'])} °C", x["why"]) for x in items]
        keys = sorted({k for x in items for k in (x["hot"], x["cold"]) if k != "exh_mean"})
        c.flag("warning" if any(x["impossible"] for x in items) else "advisory", title, rows, steps, keys)


def check_header(c: Ctx) -> None:
    """Load / current / rating / date / fuel entries that cannot be right."""
    i, t, v = c.i, c.p["test"], c.v
    if t.get("kw") and t.get("amps"):           # even at the extremes of the usual voltage range the power factor must lie between 0.5 and 1.0
        c.checks_run += 1
        pf_hi, pf_lo = t["kw"] / (1.732 * 380 * t["amps"] / 1000), t["kw"] / (1.732 * 480 * t["amps"] / 1000)
        if pf_lo > 1.0 or pf_hi < 0.5:
            c.flag("warning", "Generator kW versus amps", [("kW / A", f"{t['kw']:g} kW / {t['amps']:g} A", "needs power factor " + ("above 1.0" if pf_lo > 1.0 else "below 0.5") + " at 380-480 V")],
                   ["Re-read the kW and ampere meters."], ["load"])
    kw_r, kva_r = i.get("rated_kw"), i.get("rated_kva")
    if kw_r and kva_r:
        c.checks_run += 1
        if not 0.6 <= kw_r / kva_r <= 1.0:
            c.flag("advisory", "Rated output", [("Rated kW / kVA", f"{kw_r:g} / {kva_r:g}", f"power factor {kw_r / kva_r:.2f}, normally 0.8")], ["Compare with the name-plate."])
    if c.L is not None and c.L > 1.05:
        c.flag("advisory", "Tested load above the rating", [("Tested load", _pct(c.L), "above 105 % of rated")], ["Check the kW reading and the rated output."], ["load"], c.p["test"].get("ref") or "")
    d = i.get("date")
    if d:
        c.checks_run += 1
        age = (dt.date.today() - d).days
        if age < -1:
            c.flag("warning", "Report date", [("Report date", f"{d:%d %b %Y}", "is in the future")], ["Correct the date."])
        elif age > 400:
            c.flag("advisory", "Report date", [("Report date", f"{d:%d %b %Y}", "is over a year old")], ["Check the date."])
    rho = i.get("density")
    if rho is not None:
        c.checks_run += 1
        if not 0.80 <= rho <= 1.10:
            c.flag("advisory", "Fuel density", [("Fuel density", f"{rho:g}", "expected 0.80-1.10 t/m³ (kg/m³ typed?)")], ["Enter the density in t/m³, e.g. 0.98."])
    rpm = i.get("rpm")
    if rpm is not None:
        c.checks_run += 1
        if not 300 <= rpm <= 1900:
            c.flag("advisory", "Engine speed", [("Engine speed", f"{rpm:g} rpm", "not a usual generator speed")], ["Check the rpm entry."])
    fo_t, tank = v.get("fo_t_in"), v.get("hfo_tank_t")
    if fo_t is not None and tank is not None and (c.p["texts"]["act"].get("fuel_mode") or "").upper().startswith("HFO"):
        c.checks_run += 1
        if fo_t < tank - 5:
            c.flag("advisory", "Fuel temperatures", [("FO inlet vs service tank", f"{fo_t:g} vs {tank:g} °C", "heater only adds heat")], ["Re-read both temperatures."], ["fo_t_in", "hfo_tank_t"])
    ex = [x for x in c.df["exh"] if _ok(x)]
    if len(ex) >= 4:
        c.checks_run += 1
        if len(set(ex)) == 1:
            c.flag("advisory", "Identical exhaust temperatures", [("Exhaust, all cylinders", f"{ex[0]:g} °C", "identical - copied?")], ["Read every cylinder separately."], ["exh"])


def check_ranges(c: Ctx) -> None:
    """Readings outside any physically believable range (typing slips such as 720 instead of 72)."""
    out = []
    for k, x in c.v.items():
        if x is None or k not in PLAUSIBLE:
            continue
        c.checks_run += 1
        lo, hi = PLAUSIBLE[k]
        if not lo <= x <= hi:
            out.append((k, (PARAMS[k][0], f"{x:g} {PARAMS[k][1]}", f"believable {lo:g}-{hi:g}")))
    for m, arr in c.p["cyl"]["act"].items():
        for i, x in arr.items():
            if x is None:
                continue
            c.checks_run += 1
            lo, hi = CYL_PLAUSIBLE["pmax_bar" if m == "pmax" else m]
            xx = x * c.p["units"]["pmax_factor"] if m == "pmax" else x
            if not lo <= xx <= hi:
                out.append((m, (f"{CYL_METRICS[m][0]}, cyl {i}", f"{x:g}", f"believable {lo:g}-{hi:g}" + (" bar" if m == "pmax" else ""))))
    if out:
        c.flag("warning", "Readings outside the believable range", [r for _, r in out], ["Re-read these instruments; check decimal point and unit."],
               sorted({k for k, _ in out}), c.cells(*[k for k, _ in out if k in c.v]))


def check_sheet_arithmetic(c: Ctx) -> None:
    for x in c.p["xval"].values():
        if x["group"] == "act":
            c.checks_run += 1
            if x["label"] in ("weak", "mismatch"):
                c.flag("advisory", f"{CYL_LABEL[x['metric']].capitalize()} average on the sheet", [(f"{CYL_LABEL[x['metric']].capitalize()} average", f"printed {x['doc']:g}", f"cylinders give {x['calc']:.2f}")],
                       ["Compare with the original log: a reading or the average is mistyped."], [x["metric"]])


def check_shop_block(c: Ctx) -> None:
    """The shop-trial block is typed in by hand too: wrong load points and impossible values are reported here."""
    bad_pts = [(g, d) for g, d in c.loads.items() if d["status"] in ("corrected", "non_standard", "as_printed")]
    if bad_pts:
        c.checks_run += 1
        std = " / ".join(_pct(s) for s in parse_points(c.th["shop_points"]))
        rows = []
        for g, d in bad_pts:
            why = "equals today's test load - copied?" if "copied" in d["note"] else f"shop trials ran at {std}"
            used = f"used {_pct(d['used'])}" if d["status"] == "corrected" else "used as typed"
            rows.append((f"Shop trial block {'A' if g == 'shopA' else 'B'} load", f"typed {_pct(d['printed'])}", f"{why}; {used}"))
        c.flag("warning", "Shop-trial load points", rows, ["Copy the two load points exactly as in the maker's shop-trial report."], ["shop"],
               ", ".join(c.p["shop"][g]["ref"] for g, _ in bad_pts if c.p["shop"][g].get("ref")))
    c.checks_run += len(c.ref)
    bad = [r for r in c.ref if r["method"] == "inconsistent"]
    if bad:
        rows = [(r["label"], f"{_g(r['shopA'])} → {_g(r['shopB'])} {r['unit']}", r["note"].replace("shop-trial value ", "").split(" - ")[0]) for r in bad]
        c.flag("advisory", "Shop-trial figures look wrong", rows, ["Copy the shop-trial figures again from the maker's test report."], ["shop"])


def check_inputs(c: Ctx) -> None:
    check_pairs(c)
    check_header(c)
    check_ranges(c)
    check_sheet_arithmetic(c)
    check_shop_block(c)


# ───────────────────────────── B. engine findings ─────────────────────────────
CYL_LABEL = {"exh": "exhaust", "pmax": "Pmax", "rack": "fuel rack"}


def check_load(c: Ctx) -> None:
    L, th = c.L, c.th
    raw = c.p["test"]["raw"]
    if L is None:
        c.add("advisory", "Tested load unknown", "Load not readable from the sheet - load-based comparisons skipped.",
              ["Check the 'TESTED AT MAX. SAFE LOAD' entry and the rated output.", FIX_SHEET], topic="data", uses=["load"])
    elif L < 0.25:
        c.add("warning", f"Performance test at only {L * 100:.0f} % load", f"Recorded at {raw}. Too low to compare with the shop trial or with other months.",
              ["Repeat at the maximum safe load (at least 50 %, ideally 75 %)."], topic="load", when="retest")
    elif L < th["low_load"]:
        c.add("advisory", f"Part-load test · {L * 100:.0f} % of rated", f"Recorded at {raw}. Below {th['low_load'] * 100:.0f} % the readings trend poorly.",
              ["Use the maximum safe load for the monthly test."], topic="load")


def check_cylinders(c: Ctx) -> None:
    th, df, m = c.th, c.df, c.m
    for _, r in df.iterrows():
        i, e, d = int(r["id"]), r["exh"], r["exh_dev"]
        if _ok(e):
            if e >= th["exh_abs_crit"]:
                c.add("critical", f"Cyl {i} · exhaust {e:.0f} °C", f"Above the {th['exh_abs_crit']:.0f} °C critical limit.",
                      ["Reduce load on this generator now; transfer load if possible.", "Confirm on a second instrument, then check injector, exhaust valve, air cooler, turbocharger."],
                      ["Leaking or burnt exhaust valve", "Late combustion (injector, pump, timing)", "Too little charge air", "Overload"], i, "exhaust", uses=["exh"])
            elif e >= th["exh_abs"]:
                c.add("warning", f"Cyl {i} · high exhaust {e:.0f} °C", f"At or above the {th['exh_abs']:.0f} °C limit.",
                      ["Check charge-air temperature, air cooler and turbocharger.", "Check injection timing and injector of the hot cylinder."],
                      ["Dirty air cooler or turbocharger", "Hot charge air / cooling water", "Injector or timing", "Exhaust valve seating"], i, "exhaust", uses=["exh"])
        if _ok(d):
            hi = d > 0
            if abs(d) >= th["exh_dev_crit"]:
                c.add("mechanical", f"Cyl {i} · exhaust {d:+.0f} °C vs mean", f"{e:.0f} °C, engine mean {m['exh']:.0f} °C (limit ±{th['exh_dev_crit']:.0f}).",
                      (["Confirm the thermocouple.", "Check exhaust valve seating, injector and fuel pump; plan overhaul if confirmed."] if hi else
                       ["Confirm the thermocouple on a portable instrument.", "Test the injector; check the fuel pump."]),
                      (["Late combustion", "Leaking exhaust valve", "Worn injector"] if hi else ["Blocked or weak injector", "Low fuel delivery (pump wear, rack)", "Faulty thermocouple"]),
                      i, "exhaust", uses=["exh"])
            elif abs(d) >= th["exh_dev"]:
                c.add("warning", f"Cyl {i} · exhaust {d:+.0f} °C vs mean", f"{e:.0f} °C, engine mean {m['exh']:.0f} °C (limit ±{th['exh_dev']:.0f}).",
                      ["Confirm the thermocouple.", "Compare injector, pump and timing with the other cylinders."], ["Injector or pump difference", "Thermocouple fault"], i, "exhaust", uses=["exh"])
        d = r["pmax_dev_pct"]
        if _ok(d) and abs(d) >= th["pmax_dev"]:
            sev = "warning" if abs(d) >= th["pmax_dev_warn"] else "advisory"
            c.add(sev, f"Cyl {i} · Pmax {d:+.1f} % vs mean", f"{r['pmax']:g} {c.p['units']['pmax']}, engine mean {m['pmax']:.2f}.",
                  ["Re-measure Pmax on this cylinder.", "Check timing, injector and compression before adjusting."],
                  (["Advanced injection timing", "Pump delivering early"] if d > 0 else ["Retarded timing", "Injector or pump fault", "Compression leak (rings, liner, valves)"]), i, "pmax", uses=["pmax"])
        d = r["rack_dev_pct"]
        if _ok(d) and abs(d) >= th["rack_dev"]:
            sev = "warning" if abs(d) >= th["rack_dev_warn"] else "advisory"
            c.add(sev, f"Cyl {i} · fuel rack {d:+.1f} % vs mean", f"{r['rack']:g} mm, engine mean {m['rack']:.2f} mm.",
                  ["Check pump and rack linkage of this cylinder; calibrate the pump.", "Compare the injector with the other cylinders."],
                  (["Worn pump plunger / barrel", "Rack linkage out of adjustment"] if d > 0 else ["Pump over-delivering (calibration)", "Linkage friction or sticking"]), i, "rack", uses=["rack"],
                  when="soon" if sev == "warning" else "planned")
    sp = _spread(df["inj"])
    if sp is not None and sp >= th["inj_spread"]:
        sev = "warning" if sp >= th["inj_spread_warn"] else "advisory"
        lo_c = [int(i) for i, x in zip(df["id"], df["inj"]) if x == df["inj"].min()]
        hi_c = [int(i) for i, x in zip(df["id"], df["inj"]) if x == df["inj"].max()]
        c.add(sev, f"Injection timing differs by {sp:.1f}° between cylinders",
              f"Lowest {df['inj'].min():g}° (cyl {', '.join(map(str, lo_c))}), highest {df['inj'].max():g}° (cyl {', '.join(map(str, hi_c))}).",
              ["Check pump timing marks / VIT of the outlying cylinders against the maker's value."], ["Pump timing set differently per unit"], None, "timing", uses=["inj"], when="planned")


def check_plant(c: Ctx) -> None:
    """Cooling water, lube oil, charge air, fuel, generator bearings and smoke."""
    th, v = c.th, c.v
    jin, jout = v.get("cw_jacket_in"), v.get("cw_engine_out")
    if jout is not None and jout >= th["jacket_out_max"]:
        c.add("warning", f"Jacket cooling water outlet {jout:.0f} °C", f"Limit {th['jacket_out_max']:.0f} °C.",
              ["Check cooling-water flow: pump, strainers, valves, air in system.", "Check the thermostatic valve and cooler."], ["Low flow", "Thermostat / controller", "Fouled cooler"], None, "cooling", uses=["cw_engine_out"])
    if jin is not None and jout is not None and jout - jin > 12:
        c.add("advisory", f"Jacket cooling water rise {jout - jin:.0f} °C", f"{jin:.0f} → {jout:.0f} °C; usual rise is 5-10 °C.",
              ["Check jacket-water pump, strainers and venting."], ["Low circulating flow"], None, "cooling", uses=["cw_jacket_in", "cw_engine_out"])
    lo_o = v.get("lo_t_out")
    if lo_o is not None and lo_o >= th["lo_out_max"]:
        c.add("warning", f"L.O. cooler outlet {lo_o:.0f} °C", f"Limit {th['lo_out_max']:.0f} °C.",
              ["Clean / check the L.O. cooler, oil and water side.", "Check cooling-water flow and the thermostatic valve."], ["Fouled L.O. cooler", "Low cooling-water flow", "Thermostatic valve"], None, "lube", uses=["lo_t_out"])
    bc = v.get("ba_cyl_in")
    if bc is not None and bc >= th["boost_air_max"]:
        c.add("warning", f"Boost air at cylinder inlet {bc:.0f} °C", f"Limit {th['boost_air_max']:.0f} °C.",
              ["Clean the air cooler, air and water side.", "Check temperature and flow of the cooler water."], ["Fouled air cooler", "Warm or low cooling water"], None, "air", uses=["ba_cyl_in"])
    lp = v.get("lo_p_engine")
    if lp is not None:
        if lp < th["lo_p_min"] - 1.0:
            c.add("critical", f"L.O. pressure (engine) {lp:g} kgf/cm²", f"Well below the {th['lo_p_min']:g} minimum.",
                  ["Unload / stop and check: oil level, pump, filter differential, bearings."], ["Low oil level", "Pump or relief valve", "Blocked filter", "Worn bearings"], None, "lube", uses=["lo_p_engine"])
        elif lp < th["lo_p_min"]:
            c.add("warning", f"L.O. pressure (engine) {lp:g} kgf/cm²", f"Below the {th['lo_p_min']:g} minimum.",
                  ["Check filter differential, oil temperature and the pressure regulating valve."], ["Dirty filter", "Hot, thin oil", "Regulating valve setting"], None, "lube", uses=["lo_p_engine"])
    lpt = v.get("lo_p_tc")
    if lpt is not None and lpt < th["lo_p_tc_min"]:
        c.add("warning", f"L.O. pressure (T/C) {lpt:g} kgf/cm²", f"Below the {th['lo_p_tc_min']:g} minimum.",
              ["Check T/C oil supply line, filter and orifice."], ["Blocked supply line / orifice"], None, "lube", uses=["lo_p_tc"])
    fo_t = v.get("fo_t_in")
    if fo_t is not None and (c.p["texts"]["act"].get("fuel_mode") or "").upper().startswith("HFO") and fo_t < 80:
        c.add("advisory", f"FO inlet temperature {fo_t:.0f} °C on HFO", "Low for heavy fuel - viscosity at the injectors may be too high.",
              ["Check the heater and the viscosity controller set point."], ["Heater or controller fault"], None, "fuel", uses=["fo_t_in"])
    for k, lab in (("gen_brg_de", "engine side"), ("gen_brg_nde", "opposite side")):
        b = v.get(k)
        if b is not None and b >= th["brg_max"]:
            c.add("warning", f"Generator bearing ({lab}) {b:.0f} °C", f"Limit {th['brg_max']:.0f} °C.",
                  ["Inspect the bearing: lubrication, alignment, cooling."], ["Lubrication", "Misalignment", "Cooling"], None, "generator", uses=[k])
    sm, raw = _smoke_class(c.p["texts"]["act"].get("smoke")), c.p["texts"]["act"].get("smoke")
    if sm == "black":
        c.add("warning", f"Exhaust smoke: {raw}", "Dark smoke = poor combustion or overload.",
              ["Check air supply (air cooler, turbocharger), injectors and fuel quality."], ["Too little air", "Injector fault", "Poor fuel"], None, "smoke")
    elif sm == "blue":
        c.add("advisory", f"Exhaust smoke: {raw}", "Blue smoke = lube oil in the combustion chamber.",
              ["Check piston rings / liners; watch L.O. consumption."], ["Worn rings or liner", "Valve guide leakage"], None, "smoke")
    elif sm in ("white", "brown"):
        c.add("advisory", f"Exhaust smoke: {raw}", "Visible smoke noted on the sheet.",
              ["Review at the next test (white: unburned fuel / water / cold running; brown: slightly rich or part load)."], (), None, "smoke")


REF_ADVICE = {   # key -> (causes if above reference, causes if below, action if above, action if below)
    "exh_mean": (["Dirty air cooler or turbocharger", "Injection timing", "Exhaust valve seating", "High ambient / cooling-water temperature"], ["Thermocouple fault", "Load reading too high"],
                 "Check air cooler and turbocharger fouling, injection timing, exhaust valve seating.", "Check the thermocouples and the load reading."),
    "pmax_mean": (["Advanced timing", "High ignition-quality fuel"], ["Retarded timing", "Injector or compression leak"],
                  "Check pump timing / VIT.", "Check pump timing and compression (rings, liner, valves)."),
    "rack_mean": (["Worn pumps / injectors", "Low calorific-value fuel", "Shop trials usually run on MDO"], ["Load reading too high", "Rack calibration"],
                  "Check pumps and injectors for wear; review the fuel quality.", "Check the load reading, fuel quality and rack calibration."),
    "boost_p": (["Gauge / unit error", "High exhaust back-pressure", "Turbocharger over-speeding"], ["Dirty air cooler or filter", "Fouled turbocharger", "Gauge error"],
                "Check gauge calibration and units, exhaust back-pressure, turbocharger.", "Check air filter, air cooler and turbocharger for fouling."),
}
REF_SYSTEM = {"exh_mean": "combustion", "pmax_mean": "combustion", "rack_mean": "fuel", "boost_p": "air"}


def check_reference(c: Ctx) -> None:
    """Deviation from the shop-trial figures at the tested load."""
    trust_low = any(r["method"] == "inconsistent" and r["core"] for r in c.ref)
    for r in c.ref:
        if not r["core"] or r["ref"] is None or r["delta"] is None or r["tol"] is None:
            continue
        mag = abs(r["delta"]) if r["tol_kind"] == "abs" else abs(r["delta_pct"] or 0.0)
        if mag < r["tol"]:
            continue
        sev = "warning" if (mag >= 1.8 * r["tol"] and r["method"] in ("interpolated", "single") and not trust_low) else "advisory"
        up = r["delta"] > 0
        dtxt = f"{r['delta']:+.0f} °C" if r["tol_kind"] == "abs" else f"{r['delta_pct']:+.1f} %"
        extra = []
        if r["method"] == "extrapolated":
            extra.append("extrapolated")
        if trust_low:
            extra.append("shop data inconsistent - indicative only")
        if r["implied"] is not None and abs(r["implied"] - c.L) >= 0.08:
            extra.append(f"matches ~{_pct(r['implied'])} load on the shop curve")
        up_c, dn_c, up_a, dn_a = REF_ADVICE[r["key"]]
        c.add(sev, f"{r['label']} {dtxt} vs shop trial", f"{r['act']:.2f} {r['unit']} vs {r['ref']:.2f} expected at {_pct(c.L)} load" + (f" ({'; '.join(extra)})" if extra else ""),
              [up_a if up else dn_a, SAME_TEST], up_c if up else dn_c, None, "reference", uses=[r["key"], "shop", "load"], when="soon" if sev == "warning" else "retest",
              system=REF_SYSTEM.get(r["key"], "shop"))


def diagnose(c: Ctx) -> None:
    """Combine the individual findings into probable causes - several symptoms pointing at one fault."""
    th, v, df = c.th, c.v, c.df
    n_hot = int((df["exh"] >= th["exh_abs"]).sum()) if len(df) else 0

    def dev(key: str) -> Optional[float]:          # deviation from the shop-trial reference, in the unit its tolerance is given in
        r = c.usable_ref(key)
        return None if r is None else (r["delta"] if r["tol_kind"] == "abs" else r["delta_pct"])
    d_exh, d_boost, d_rack, d_pmax = dev("exh_mean"), dev("boost_p"), dev("rack_mean"), dev("pmax_mean")
    exh_hi = (d_exh is not None and d_exh >= th["ref_exh"]) or (len(df) > 0 and n_hot >= max(1, len(df) // 2))
    boost_lo = d_boost is not None and d_boost <= -th["ref_boost"]
    air_hot = v.get("ba_cyl_in") is not None and v["ba_cyl_in"] >= th["boost_air_max"]
    rack_hi = d_rack is not None and d_rack >= th["ref_rack"]
    pmax_lo = d_pmax is not None and d_pmax <= -th["ref_pmax"]

    def add(sev, title, fault, causes, actions, uses, system):
        c.add(sev, title, fault, actions, causes, None, "diagnosis", uses=uses, system=system)

    if exh_hi and (boost_lo or air_hot):
        add("warning", "Probable diagnosis · too little charge air",
            "Exhaust high and " + ("boost low." if boost_lo else "charge air hot.") + " The engine is short of cool air.",
            ["Fouled air cooler (air or water side)", "Fouled turbocharger or air filter", "Warm or low cooling water to the air cooler"],
            ["Clean the air cooler; check its pressure drop.", "Water-wash the turbocharger; check the air filter.", "Check temperature and flow of the cooler water."],
            ["exh", "ba_cyl_in", "boost_p", "shop"], "air")
    if exh_hi and rack_hi:
        add("warning", "Probable diagnosis · more fuel burnt than at the shop trial", "Exhaust high and rack above the shop-trial value at this load.",
            ["Worn pump plungers / injectors", "Poor or low-calorific fuel", "Fouling or friction losses"],
            ["Test injectors; check pumps for wear.", "Review fuel quality (calorific value, density, CCAI) and temperature."], ["exh", "rack", "shop"], "fuel")
    if pmax_lo and (exh_hi or rack_hi):
        add("warning", "Probable diagnosis · late or poor combustion", "Pmax below the shop-trial value while exhaust or rack is high - fuel burns late.",
            ["Retarded injection timing", "Worn injectors or pumps", "Poor ignition quality"],
            ["Check injection timing / VIT.", "Test injectors; review fuel quality."], ["pmax", "exh", "rack", "shop"], "combustion")
    for _, r in df.iterrows():
        if _ok(r["exh_dev"]) and r["exh_dev"] >= th["exh_dev"] and _ok(r["pmax_dev_pct"]) and r["pmax_dev_pct"] <= -th["pmax_dev"]:
            i = int(r["id"])
            add("mechanical", f"Probable diagnosis · fault on cylinder {i}",
                f"Cyl {i}: exhaust {r['exh_dev']:+.0f} °C, Pmax {r['pmax_dev_pct']:+.1f} % vs mean - fuel burns late or is lost.",
                ["Leaking injector / worn nozzle", "Leaking or burnt exhaust valve", "Compression loss (rings, liner)"],
                [f"Pull and test the injector of cyl {i}.", f"Check exhaust valve seating and compression of cyl {i} at the next opportunity."], ["exh", "pmax"], "combustion")


# ───────────────────────────── C. verdict and crew action list ─────────────────────────────
VERDICT_LEVELS = {   # level -> (label, banner class, icon)
    "act_now": ("Act now", "low", "⛔"), "inspect": ("Inspect and correct", "review", "🛠️"), "watch": ("Correct at the next opportunity", "review", "⚠️"),
    "monitor": ("Monitor", "high", "👁️"), "normal": ("Normal", "ok", "✅"),
}


def build_plan(F: list[dict]) -> list[dict]:
    """The crew action list: every distinct step once, most urgent first, with the findings that ask for it."""
    steps: dict[str, dict] = {}
    for order, f in enumerate(F):
        for s in f["actions"]:
            cur = steps.get(s)
            if cur is None:
                steps[s] = dict(step=s, when=f["when"], sev=f["sev"], because=[f["title"]], order=order)
            else:
                if f["title"] not in cur["because"]:
                    cur["because"].append(f["title"])
                if URGENCY[f["when"]][0] < URGENCY[cur["when"]][0]:
                    cur["when"] = f["when"]
                if SEV_RANK[f["sev"]] > SEV_RANK[cur["sev"]]:
                    cur["sev"] = f["sev"]
    return sorted(steps.values(), key=lambda s: (URGENCY[s["when"]][0], s["step"] == FIX_SHEET, -SEV_RANK[s["sev"]], s["order"]))


def short_title(t: str) -> str:
    return t.removeprefix("Probable diagnosis · ").removeprefix("Check entry · ")


def make_verdict(F: list[dict], n_checks: int) -> dict:
    eng = [f for f in F if f["cat"] not in ("data", "diagnosis")]
    data = [f for f in F if f["cat"] == "data"]
    diag = [f for f in F if f["cat"] == "diagnosis"]
    worst = max([SEV_RANK[f["sev"]] for f in eng + diag] or [1])
    level = "act_now" if worst >= 5 else "inspect" if worst == 4 else "watch" if worst == 3 else "monitor" if worst == 2 else "normal"
    top = [f["title"] for f in sorted(eng + diag, key=lambda f: -SEV_RANK[f["sev"]])[:3]]
    if level == "normal":
        head = "All readings are within the limits and consistent with the shop-trial figures."
    else:
        head = "; ".join(short_title(t) for t in top) + "."
    d_warn = sum(1 for f in data if f["sev"] == "warning")
    dq = "many" if d_warn >= 3 else "some" if (d_warn >= 1 or len(data) >= 2) else "reliable"
    dq_text = {"reliable": "Entries look consistent", "some": f"{len(data)} entr{'y' if len(data) == 1 else 'ies'} to re-check",
               "many": f"{len(data)} entries to re-check - treat the figures as provisional"}[dq]
    return dict(level=level, label=VERDICT_LEVELS[level][0], cls=VERDICT_LEVELS[level][1], icon=VERDICT_LEVELS[level][2], headline=head,
                data_quality=dq, data_text=dq_text, n_data=len(data), n_checks=n_checks, provisional=dq != "reliable")


def missing_entries(p: dict) -> list[str]:
    """Readings the sheet leaves empty or marks N/A - not an error, but the diagnosis cannot use them."""
    out = []
    for k, (lab, _) in PARAMS.items():
        if p["kinds"]["act"].get(k) == "na":
            out.append(lab + " (N/A)")
    i = p["identity"]
    for k, lab in (("viscosity", "Fuel viscosity"), ("lo_cons", "L.O. consumption"), ("fo_cons", "F.O. consumption"), ("ambient_t", "Ambient temperature"),
                   ("sea_t", "Sea water temperature"), ("run_hours", "Running hours")):
        if i.get(k) in (None, ""):
            out.append(lab + " (blank)")
    return out


# ───────────────────────────── the analysis ─────────────────────────────
def analyze(p: dict, th: Optional[dict] = None) -> dict:
    th = {**TH_DEFAULT, **(th or {})}
    c = Ctx(p, th)
    c.loads = shop_loads(p, th)
    c.ref = shop_reference(p, th, c.loads)
    check_load(c)
    check_cylinders(c)
    check_plant(c)
    check_reference(c)
    diagnose(c)
    check_inputs(c)
    F = c.F

    # an engine finding that rests on a reading that failed an input check is marked "verify first"
    bad_keys = {k for f in F if f["cat"] == "data" for k in f["uses"]}
    for f in F:
        f["provisional"] = f["cat"] != "data" and bool(bad_keys & set(f["uses"]))
    F.sort(key=lambda f: (f["cat"] == "data", f["cat"] != "diagnosis", -SEV_RANK[f["sev"]], f["cyl"] if f["cyl"] is not None else 99, f["title"]))

    real = [f for f in F if f["cat"] not in ("data", "diagnosis")]
    weights = {"critical": 30, "mechanical": 14, "warning": 7, "advisory": 2.5}
    score = 100.0 / (1.0 + sum(weights[f["sev"]] for f in real) / 40.0)
    sevs = {f["sev"] for f in real}
    score = min(score, 74.0) if "critical" in sevs else min(score, 84.0) if "mechanical" in sevs else score
    score = int(round(max(0, min(100, score))))
    grade = "Excellent" if score >= 85 else "Good" if score >= 70 else "Needs attention" if score >= 50 else "Critical"

    df = c.df
    worst = {int(i): "nominal" for i in df["id"]}
    for f in real:
        if f["cyl"] in worst and SEV_RANK[f["sev"]] > SEV_RANK[worst[f["cyl"]]]:
            worst[f["cyl"]] = f["sev"]
    df["status"] = df["id"].map(worst)
    counts = {s: sum(1 for f in real if f["sev"] == s) for s in SEV_RANK}
    plan = build_plan(F)
    return dict(df=df, mean=dict(c.m, pmax_bar=_mean(df["pmax_bar"])), spread=dict(exh=_spread(df["exh"]), pmax=_spread(df["pmax"]), rack=_spread(df["rack"]), inj=_spread(df["inj"])),
                findings=F, score=score, grade=grade, load=c.L, ref=c.ref, counts=counts, th=th, n=p["n_cylinders"], shop=c.loads,
                verdict=make_verdict(F, c.checks_run), plan=plan, missing=missing_entries(p))


# ====================================================================================================
# 3. FLEET  -  comparing reports with each other (same vessel / same generator)
# ====================================================================================================

def dg_label(p: dict) -> str:
    i = p["identity"]
    return f"DG {i['dg_no']}" if i.get("dg_no") else "DG"


def summary_row(name: str, p: dict, an: Optional[dict]) -> dict:
    i, c = p["identity"], p["confidence"]
    row = dict(file=name, vessel=i.get("vessel") or name, dg=dg_label(p), dg_no=i.get("dg_no"), date=i.get("date"), time=i.get("time"),
               load=p["test"].get("load_frac"), run_hours=i.get("run_hours"), confidence=c["total"], level=c["level"], n=p["n_cylinders"])
    if an:
        row.update(score=an["score"], crit=an["counts"]["critical"] + an["counts"]["mechanical"], warn=an["counts"]["warning"], adv=an["counts"]["advisory"],
                   verdict=an["verdict"]["label"], n_data=an["verdict"]["n_data"], exh=an["mean"]["exh"], exh_spread=an["spread"]["exh"],
                   pmax_bar=an["mean"]["pmax_bar"], rack=an["mean"]["rack"], boost=p["values"]["act"].get("boost_p"))
    return row


# readings that belong to one engine (not to a shared system): the same value on engines at different loads is suspicious
OWN_READINGS = ("cw_jacket_in", "cw_engine_out", "cw_ac_out", "cw_lo_in", "cw_lo_out", "lo_t_in", "lo_t_out", "lo_p_engine", "lo_p_tc", "ba_cyl_in",
                "gen_brg_de", "gen_brg_nde")
COMPARE_ACTION = "Compare air cooler, turbocharger, injectors, pumps with the sister unit; repeat both tests at the same load."


def peer_notes(items: list[tuple[str, dict, dict]], load_window: float = 0.12) -> list[dict]:
    """Compare generators of the same vessel (items: (name, parsed, analysis)).
    -> [{vessel, dgs, about (the DGs the note concerns), text, sev, cat, action}]"""
    notes: list[dict] = []
    by_vessel: dict[str, list] = {}
    for name, p, an in items:
        by_vessel.setdefault((p["identity"].get("vessel") or name), []).append((name, p, an))

    def add(vessel, dgs, about, text, sev="advisory", cat="engine", action=COMPARE_ACTION):
        notes.append(dict(vessel=vessel, dgs=tuple(dgs), about=tuple(about), text=text, sev=sev, cat=cat, action=action))

    for vessel, lst in by_vessel.items():
        for a in range(len(lst)):
            for b in range(a + 1, len(lst)):
                (_, p1, a1), (_, p2, a2) = lst[a], lst[b]
                l1, l2 = a1["load"], a2["load"]
                if l1 is None or l2 is None:
                    continue
                n1, n2 = dg_label(p1), dg_label(p2)
                e1, e2 = a1["mean"]["exh"], a2["mean"]["exh"]
                if abs(l1 - l2) <= load_window:
                    if e1 is not None and e2 is not None and abs(e1 - e2) >= 25:
                        hot, cold = (n1, n2) if e1 > e2 else (n2, n1)
                        add(vessel, (n1, n2), (hot,), f"{hot} exhaust is {abs(e1 - e2):.0f} °C hotter than {cold} at similar load ({l1 * 100:.0f} / {l2 * 100:.0f} %).", "warning")
                    b1, b2 = p1["values"]["act"].get("boost_p"), p2["values"]["act"].get("boost_p")
                    if b1 and b2 and abs(b1 - b2) / max(b1, b2) >= 0.15:
                        lo, hi = (n1, n2) if b1 < b2 else (n2, n1)
                        add(vessel, (n1, n2), (lo,), f"{lo} boost is {abs(b1 - b2) / max(b1, b2) * 100:.0f} % lower than {hi} at similar load.")
                    r1, r2 = a1["mean"]["rack"], a2["mean"]["rack"]
                    if r1 and r2 and abs(r1 - r2) / max(r1, r2) >= 0.10:
                        hi, lo = (n1, n2) if r1 > r2 else (n2, n1)
                        add(vessel, (n1, n2), (hi,), f"{hi} rack is {abs(r1 - r2) / max(r1, r2) * 100:.0f} % higher than {lo} at similar load.")
                else:
                    if e1 is not None and e2 is not None:
                        (lo_n, lo_l, lo_e), (hi_n, hi_l, hi_e) = sorted([(n1, l1, e1), (n2, l2, e2)], key=lambda t: t[1])
                        if lo_e >= hi_e + 25:
                            add(vessel, (n1, n2), (lo_n,), f"{lo_n} exhaust {lo_e:.0f} °C at {lo_l * 100:.0f} % load is hotter than {hi_n} {hi_e:.0f} °C at {hi_l * 100:.0f} % - too hot for its load.", "warning")
                    # identical readings on engines running at different loads: probably copied from one engine to the other
                    v1, v2 = p1["values"]["act"], p2["values"]["act"]
                    same = [k for k in OWN_READINGS if v1.get(k) is not None and v1.get(k) == v2.get(k)]
                    if len(same) >= 4:
                        add(vessel, (n1, n2), (n1, n2), f"{n1} and {n2}: {len(same)} identical readings ({', '.join(PARAMS[k][0].lower() for k in same[:3])}…) at different loads ({l1 * 100:.0f} / {l2 * 100:.0f} %) - copied?",
                            "advisory", "data", "Read each generator's own instruments; never copy between generators.")
    return notes


def history_notes(items: list[tuple[str, dict, dict]], load_window: float = 0.12) -> list[dict]:
    """Same generator, several dates: running hours must advance, and a drift at a similar load is flagged."""
    notes: list[dict] = []
    by_dg: dict[tuple, list] = {}
    for name, p, an in items:
        i = p["identity"]
        if i.get("date") and i.get("dg_no"):
            by_dg.setdefault(((i.get("vessel") or name), i["dg_no"]), []).append((i["date"], p, an))
    for (vessel, _), lst in by_dg.items():
        lst.sort(key=lambda t: t[0])
        for (d1, p1, a1), (d2, p2, a2) in zip(lst, lst[1:]):
            if d1 == d2:
                continue
            lab = dg_label(p2)
            span = f"{d1:%d %b %Y} → {d2:%d %b %Y}"

            def add(text, sev="advisory", cat="engine", action=COMPARE_ACTION):
                notes.append(dict(vessel=vessel, dgs=(lab,), about=(lab,), text=text, sev=sev, cat=cat, action=action))
            h1, h2 = p1["identity"].get("run_hours"), p2["identity"].get("run_hours")
            if h1 is not None and h2 is not None:
                days = (d2 - d1).days
                if h2 < h1:
                    add(f"{lab} running hours fell {h1:,.0f} → {h2:,.0f} h ({span}).", "warning", "data", "Check both hour entries against the hour meter and PMS.")
                elif h2 - h1 > days * 24 * 1.02 + 1:
                    add(f"{lab} running hours +{h2 - h1:,.0f} h in {days} days ({span}) - more than the calendar allows.", "warning", "data",
                        "Check both hour entries against the hour meter and PMS.")
            l1, l2 = a1["load"], a2["load"]
            if l1 is None or l2 is None or abs(l1 - l2) > load_window:
                continue
            e1, e2 = a1["mean"]["exh"], a2["mean"]["exh"]
            if e1 is not None and e2 is not None and e2 - e1 >= 25:
                add(f"{lab} exhaust +{e2 - e1:.0f} °C at similar load ({span}).", "warning",
                    action="Look for fouled air cooler / turbocharger, injector wear, exhaust valve leak.")
            r1, r2 = a1["mean"]["rack"], a2["mean"]["rack"]
            if r1 and r2 and (r2 - r1) / r1 >= 0.08:
                add(f"{lab} fuel rack +{(r2 - r1) / r1 * 100:.0f} % at similar load ({span}).", "advisory",
                    action="More fuel for the same load: check pump / injector wear, fuel quality.")
            b1, b2 = p1["values"]["act"].get("boost_p"), p2["values"]["act"].get("boost_p")
            if b1 and b2 and (b1 - b2) / b1 >= 0.12:
                add(f"{lab} boost −{(b1 - b2) / b1 * 100:.0f} % at similar load ({span}).", "advisory",
                    action="Check the air filter, the air cooler and the turbocharger for fouling.")
    return notes


# ───────────────────────────── export helpers ─────────────────────────────
def fmt_value(p: dict, group: str, key: str):
    v = p["values"][group].get(key)
    if v is not None:
        return v
    if p["kinds"][group].get(key) == "na":
        return "N/A"
    t = p["texts"][group].get(key)
    return t if t else None


def flat_rows(p: dict, an: Optional[dict]) -> list[dict]:
    """Every extracted reading as one flat table (CSV export)."""
    ref = {r["key"]: r for r in (an["ref"] if an else [])}
    shop = an["shop"] if an else {}
    rows: list[dict] = []
    i = p["identity"]
    for lab, val in (("Vessel", i["vessel"]), ("DG", i["dg_no"]), ("Maker", i["maker"]), ("Type", i["type"]), ("Serial no.", i["serial"]),
                     ("Rated output", i["output_text"]), ("RPM", i["rpm"]), ("Running hours", i["run_hours"]), ("Report date", i["date"]),
                     ("Report time", i["time"]), ("Tested load (raw)", p["test"]["raw"]), ("Tested kW", p["test"]["kw"]), ("Tested A", p["test"]["amps"]),
                     ("Load fraction", p["test"]["load_frac"]), ("Fuel used", i["fuel_used"]), ("Density", i["density"]), ("Viscosity", i["viscosity"]),
                     ("Fuel oil cons. (24 h)", i["fo_cons"]), ("Lub oil cons. (24 h)", i["lo_cons"]), ("Ambient temperature", i["ambient_t"]),
                     ("Sea water temperature", i["sea_t"]), ("Prepared by", i["prepared_by"]),
                     ("Shop-trial load A (typed)", p["shop"]["shopA"]["load"]), ("Shop-trial load B (typed)", p["shop"]["shopB"]["load"]),
                     ("Shop-trial load A (used)", shop.get("shopA", {}).get("used")), ("Shop-trial load B (used)", shop.get("shopB", {}).get("used"))):
        rows.append(dict(section="Header", parameter=lab, unit="", actual=val, shop_a="", shop_b="", reference="", cell=""))
    for title, keys in PARAM_LAYOUT:
        for k in keys:
            vals = [fmt_value(p, g, k) for g in GROUPS]
            if all(x is None for x in vals):
                continue
            r = ref.get(k)
            rows.append(dict(section=title, parameter=PARAMS[k][0], unit=PARAMS[k][1], actual=vals[0], shop_a=vals[1], shop_b=vals[2],
                             reference=(r["ref"] if r else None), cell=p["refs"].get(f"act.{k}", "")))
    for extra in ("fuel_mode", "fuel_grade", "smoke"):
        t = p["texts"]["act"].get(extra)
        if t:
            rows.append(dict(section="Text entries", parameter=extra.replace("_", " "), unit="", actual=t, shop_a="", shop_b="", reference="", cell=p["refs"].get(f"act.{extra}", "")))
    for m, (lab, unit) in CYL_METRICS.items():
        for cid in range(1, p["n_cylinders"] + 1):
            vals = [p["cyl"][g][m].get(cid) for g in GROUPS]
            if all(x is None for x in vals):
                continue
            rows.append(dict(section="Cylinders", parameter=f"{lab} · cyl {cid}", unit=unit or p["units"]["pmax"], actual=vals[0], shop_a=vals[1], shop_b=vals[2],
                             reference="", cell=p["cyl_ref"].get(f"act.{m}.{cid}", "")))
    for m in AVG_METRICS:
        for g, lab in (("act", "actual"), ("shopA", "shop A"), ("shopB", "shop B")):
            if m in p["avgs_doc"][g]:
                rows.append(dict(section="Printed averages", parameter=f"{CYL_METRICS[m][0]} average ({lab})", unit=CYL_METRICS[m][1] or p["units"]["pmax"],
                                 actual=p["avgs_doc"][g][m][0], shop_a="", shop_b="", reference="", cell=""))
    return rows


# ====================================================================================================
# 4. UI
# ====================================================================================================


DASH = '<span class="dim">—</span>'


# ═════════════════════════════════════════════════════════════════════════════
#  state, parsing, sidebar
# ═════════════════════════════════════════════════════════════════════════════
def init_state() -> None:
    ss = st.session_state
    ss.setdefault("dg_gen", 0)


def reset_thresholds() -> None:
    for k in [f"dg_th_{s[0]}" for s in TH_SPECS] + ["dg_shop_mode", "dg_shop_pts"]:
        st.session_state.pop(k, None)             # dropping the key makes each widget fall back to its default value


def clear_all() -> None:
    st.session_state["dg_gen"] += 1


@st.cache_data(show_spinner=False, max_entries=64)
def parse_cached(data: bytes):
    try:
        return parse_tec_014(data), None
    except AuxParseError as e:
        return None, str(e)
    except Exception as e:                                                       # noqa: BLE001
        return None, f"Unexpected problem while reading the workbook ({type(e).__name__}). The file may be damaged."


SHOP_MODES = {"auto": "Read typed loads as the standard points (recommended)", "printed": "Use the loads exactly as typed"}


def render_sidebar_thresholds() -> dict:
    with st.sidebar:
        md('<div style="height:1px;background:var(--line);margin:1.1rem 0 .9rem"></div>'
           '<p class="lbl" style="margin-bottom:.4rem">Auxiliary engines · TEC-014</p>')
        th: dict = {}
        with st.expander("Shop-trial load points"):
            th["shop_mode"] = st.radio("Loads typed in the shop-trial columns", list(SHOP_MODES), format_func=SHOP_MODES.get, key="dg_shop_mode")
            th["shop_points"] = st.text_input("Shop trials were run at (% load)", value=TH_DEFAULT["shop_points"], key="dg_shop_pts")
            st.caption("The crew types the two shop-trial loads by hand and often enters the engine's own test load (e.g. 59 %) instead. "
                       "A typed value close to one of these points is read as that point, and the sheet is flagged.")
        with st.expander("Aux. diagnostic thresholds"):
            for k, lab, step, fmt in TH_SPECS:
                th[k] = st.number_input(lab, value=float(TH_DEFAULT[k]), step=step, format=fmt, key=f"dg_th_{k}")
            st.button("Reset to defaults", on_click=reset_thresholds, key="dg_reset_th")
            st.caption("Limits are generic for medium-speed generator sets; adjust to the maker's values.")
    return th


def render_sidebar_actions(n_reports: int) -> None:
    with st.sidebar:
        if n_reports:
            st.button("Clear all aux. reports", on_click=clear_all, key="dg_clear")


def load_reports(files, th: dict) -> tuple[list, list]:
    reps, dups, seen = [], [], set()
    for f in files or []:
        data = f.getvalue()
        h = hashlib.sha1(data).hexdigest()[:12]
        if h in seen:
            dups.append(f.name)
            continue
        seen.add(h)
        parsed, err = parse_cached(data)
        an, an_err = None, None
        if parsed and parsed["status"] == "ok":
            try:
                an = analyze(parsed, th)
            except Exception as e:                                               # noqa: BLE001
                an_err = f"The readings were extracted but could not be analysed ({type(e).__name__})."
        reps.append(dict(id=h, name=f.name, parsed=parsed, error=err or an_err, an=an))

    def key(r):
        p = r["parsed"]
        if not p:
            return ("~", 99, _date.max.isoformat())
        i = p["identity"]
        return ((i.get("vessel") or "~"), i.get("dg_no") or 99, (i.get("date") or _date.max).isoformat())
    reps.sort(key=key)
    return reps, dups


def rep_label(r: dict) -> str:
    p = r["parsed"]
    if not p:
        return r["name"]
    i = p["identity"]
    d = i.get("date")
    return " · ".join(x for x in [i.get("vessel") or r["name"], dg_label(p), d.strftime("%d %b %y") if d else ""] if x)


# ═════════════════════════════════════════════════════════════════════════════
#  header, empty state
# ═════════════════════════════════════════════════════════════════════════════
def render_header(n: int) -> None:
    right = chip(f"v{APP_VERSION}", "gold") + " " + (chip(f"{n} report{'s' if n != 1 else ''} loaded", "info") if n else chip("awaiting data"))
    md(f'<div class="hdr"><div class="hdr-l"><div class="logo">DG</div><div><div class="brand">AUXILIARY ENGINES</div>'
       f'<div class="tag">Generator performance intelligence · TEC-014</div></div></div><div>{right}</div></div>')


def render_steps(step: int) -> None:
    out = []
    for i, nm in enumerate(["Upload", "Extract", "Analyze"], 1):
        cls = "done" if i < step else "on" if i == step else ""
        out.append(f'<div class="step {cls}"><span class="n">{"✓" if i < step else i}</span><span class="t">{nm}</span></div>')
    md('<div class="steps">' + "".join(out) + "</div>")


def render_empty() -> None:
    md("""<div class="empty"><div class="big">Drop your aux. engine reports</div>
    <div class="sm">Drag one or many monthly <b>TEC-014</b> Excel reports (.xlsx) into the zone above - one workbook per generator. The sheets are read in memory, matched by label
    (not by cell position), reconciled with their own printed averages and compared with the shop-trial figures printed on the same form.</div></div>
    <div class="feat">
    <div class="card"><div class="fi">🧭</div><div class="fh">Reads by label</div><div class="fd">Crews insert and delete rows. Every field is found by its section and label, so a shifted layout cannot put a value in the wrong place.</div></div>
    <div class="card"><div class="fi">🛡️</div><div class="fh">Nothing lost silently</div><div class="fd">Every filled cell is mapped to a field, recognised as template text, or listed under "not mapped".</div></div>
    <div class="card"><div class="fi">🔎</div><div class="fh">Checks the entries</div><div class="fd">Cooler in / out reversed, air colder than its cooling water, T/C hotter than the cylinders, bearings colder than the room, wrong shop-trial loads: every sheet is tested against physics before it is trusted.</div></div>
    <div class="card"><div class="fi">🛠️</div><div class="fh">Verdict and crew actions</div><div class="fd">One verdict, the probable causes and a prioritised action list for the engine room - compared with the shop-trial figures at the tested load.</div></div>
    <div class="card"><div class="fi">🔌</div><div class="fh">DG vs DG</div><div class="fd">Load all generators of a vessel together to spot the unit that runs hotter, needs more fuel or loses boost - and readings copied from one engine to another.</div></div>
    </div>""")


# ═════════════════════════════════════════════════════════════════════════════
#  charts
# ═════════════════════════════════════════════════════════════════════════════
def _ylim(vals, mean):
    xs = [v for v in vals if v is not None and v == v]
    if not xs:
        return None
    lo, hi = min(xs), max(xs)
    sp = hi - lo
    pad_lo = max(sp * 1.6, abs(mean or hi) * 0.05)
    return [max(0.0, lo - pad_lo), hi + max(sp * 0.7, abs(mean or hi) * 0.03)]


def chart_cyl(df, col, dev_col, title, unit, warn, crit, th_unit, fmt, ref=None, ref_label="", mean=None, height=330):
    x = [str(int(i)) for i in df["id"]]
    y = df[col].tolist()
    colors = []
    for d in df[dev_col].tolist():
        if d is None or d != d:
            colors.append(C_GOLD)
        elif crit is not None and abs(d) >= crit:
            colors.append(C_ORANGE)
        elif abs(d) >= warn:
            colors.append(C_WARN)
        else:
            colors.append(C_GOLD)
    fig = go.Figure(go.Bar(x=x, y=y, marker=dict(color=colors, opacity=.92), text=[fmt.format(v) if v == v and v is not None else "" for v in y],
                           textposition="outside", textfont=dict(size=11, color="#B8B4A8"), name=title, showlegend=False,
                           hovertemplate="Cyl %{x}<br>%{y:" + (".1f" if "f" in fmt else "") + "} " + unit + "<extra></extra>"))
    if mean is not None:
        fig.add_hline(y=mean, line=dict(color=C_GOLD, width=1, dash="dash"), annotation_text=f"mean {fmt.format(mean)}", annotation_font=dict(color=C_GOLD, size=10),
                      annotation_position="top left")
    if ref is not None:
        fig.add_hline(y=ref, line=dict(color=C_INFO, width=1.4, dash="dot"), annotation_text=ref_label, annotation_font=dict(color=C_INFO, size=10),
                      annotation_position="bottom right")
    _style(fig, height, legend=False)
    fig.update_layout(title=dict(text=f"{title} · {unit}" if unit else title, font=dict(family="Barlow Condensed, sans-serif", size=15, color="#8E8B80"), x=0.0, xanchor="left"),
                      hovermode="closest")
    rng = _ylim(y + ([ref] if ref is not None else []), mean)
    if rng:
        fig.update_yaxes(range=rng)
    fig.update_xaxes(title_text="Cylinder")
    return fig


def chart_ref(rows, th):
    items = [r for r in rows if r["ref"] is not None and r["delta_pct"] is not None and r["method"] in ("interpolated", "extrapolated", "single")]
    if not items:
        return None
    items = sorted(items, key=lambda r: (not r["core"], r["label"]))
    labels, vals, cols = [], [], []
    for r in items:
        labels.append(r["label"] + (" *" if r["method"] == "extrapolated" else ""))
        vals.append(r["delta_pct"])
        within = None
        if r["tol"] is not None:
            within = (abs(r["delta"]) < r["tol"]) if r["tol_kind"] == "abs" else (abs(r["delta_pct"]) < r["tol"])
        cols.append(C_OK if within else C_WARN if within is False else C_MUTED)
    fig = go.Figure(go.Bar(y=labels[::-1], x=vals[::-1], orientation="h", marker=dict(color=cols[::-1], opacity=.92),
                           text=[f"{v:+.1f} %" for v in vals[::-1]], textposition="outside", textfont=dict(size=10, color="#B8B4A8"),
                           hovertemplate="%{y}<br>%{x:+.1f} % vs reference<extra></extra>"))
    _style(fig, max(260, 36 * len(items) + 90), legend=False)
    fig.update_layout(hovermode="closest", margin=dict(l=12, r=40, t=36, b=10),
                      title=dict(text="Deviation from the shop-trial reference at the tested load · %", font=dict(family="Barlow Condensed, sans-serif", size=15, color="#8E8B80"),
                                 x=0.0, xanchor="left"))
    fig.add_vline(x=0, line=dict(color=C_MUTED, width=1))
    fig.update_xaxes(title_text="% above (+) / below (−) reference")
    return fig


def chart_dg_compare(rows):
    x = [r["label"] for r in rows]
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=x, y=[r["exh"] for r in rows], name="Exhaust mean °C", marker=dict(color=C_GOLD, opacity=.9),
                         text=[f"{r['exh']:.0f}" if r["exh"] is not None else "" for r in rows], textposition="outside", textfont=dict(size=11, color="#B8B4A8")), secondary_y=False)
    fig.add_trace(go.Scatter(x=x, y=[(r["load"] * 100 if r["load"] is not None else None) for r in rows], name="Tested load %", mode="lines+markers",
                             line=dict(color=C_INFO, width=2.4, dash="dot"), marker=dict(size=10, color=C_INFO, line=dict(color="#09090D", width=2))), secondary_y=True)
    _style(fig, 340)
    ex = [r["exh"] for r in rows if r["exh"] is not None]
    if ex:
        fig.update_yaxes(range=[max(0, min(ex) - 80), max(ex) + 40], secondary_y=False, title_text="Exhaust mean °C")
    fig.update_yaxes(range=[0, 110], secondary_y=True, title_text="Load %", showgrid=False, tickvals=[0, 25, 50, 75, 100])
    return fig


def chart_dg_trend(pts):
    """pts: list of dict(date, exh, pmax, rack, boost, load) ascending."""
    fig = make_subplots(rows=2, cols=2, shared_xaxes=False, vertical_spacing=0.18, horizontal_spacing=0.09,
                        subplot_titles=("Exhaust mean · °C", "Pmax mean · bar", "Fuel rack mean · mm", "Boost pressure · kgf/cm²"))
    x = [p["date"] for p in pts]
    for k, (r, c) in {"exh": (1, 1), "pmax": (1, 2), "rack": (2, 1), "boost": (2, 2)}.items():
        fig.add_trace(go.Scatter(x=x, y=[p[k] for p in pts], mode="lines+markers", line=dict(color=C_GOLD2, width=2.4),
                                 marker=dict(size=9, color=C_GOLD2, line=dict(color="#09090D", width=2)), showlegend=False,
                                 customdata=[[(p["load"] or 0) * 100] for p in pts], hovertemplate="%{x|%d %b %Y}<br>%{y:.2f}<br>load %{customdata[0]:.0f} %<extra></extra>"), r, c)
    _style(fig, 470, legend=False)
    fig.update_layout(hovermode="closest")
    fig.update_annotations(font=dict(family="Barlow Condensed, sans-serif", size=13, color="#8E8B80"), xanchor="left")
    return fig


# ═════════════════════════════════════════════════════════════════════════════
#  report pieces
# ═════════════════════════════════════════════════════════════════════════════
def _num(v, nd=1):
    return fnum(v, nd) if v is not None else "—"


def render_card(rep: dict) -> None:
    p = rep["parsed"]
    i, conf, cov = p["identity"], p["confidence"], p["coverage"]
    level = conf["level"]
    tone = {"verified": "#5BB98C", "high": "#7BAFD9", "review": "#E5B54A", "low": "#E8645E", "blank": "#C4A35A"}[level]
    ring = ring_svg(conf["total"], tone, f"{conf['total']:.0f}" if level != "blank" else "—", "CONFIDENCE")

    def bar(label, frac, extra=""):
        return (f'<div class="bar-row"><div class="bl"><span>{label}</span><b>{frac * 100:.0f}%{extra}</b></div><div class="track"><div class="fill" style="width:{frac * 100:.0f}%"></div></div></div>')
    bars = (bar("Structure located", conf["structure"]) + bar("Completeness", conf["completeness"]) + bar("Physical plausibility", conf["plausibility"])
            + bar("Sheet cross-check", conf["crossval"], "" if conf.get("verified_by_sheet") else " (no averages)"))
    t = p["test"]
    load_txt = None
    if t.get("raw"):
        load_txt = t["raw"] + (f" · {t['load_frac'] * 100:.0f} % of rated" if t.get("load_frac") is not None else "")
    fuel = " · ".join(x for x in [p["texts"]["act"].get("fuel_mode"), p["texts"]["act"].get("fuel_grade") or i.get("fuel_used"),
                                  f"ρ {i['density']:g}" if i.get("density") else "", f"{i['fo_cons']} / 24 h" if i.get("fo_cons") else ""] if x)
    when = " ".join(x for x in [i["date"].strftime("%d %b %Y") if i.get("date") else (str(i.get("datetime_raw") or "") or None), i.get("time")] if x)
    facts = (fact("Test date / time", when) + fact("Generator", dg_label(p) + (f" · S/N {i['serial']}" if i.get("serial") else "")) + fact("Rated output", i.get("output_text"))
             + fact("Tested load", load_txt) + fact("Engine speed", f"{i['rpm']:.0f} rpm" if i.get("rpm") else None)
             + fact("Running hours", f"{i['run_hours']:,.0f}" if i.get("run_hours") else None) + fact("Fuel", fuel)
             + fact("Ambient / sea water", " / ".join(f"{x:g} °C" for x in (i.get("ambient_t"), i.get("sea_t")) if x is not None) or None)
             + fact("Prepared by", i.get("prepared_by"))
             + fact("Sheet coverage", f'{cov["pct"]:.0f} % of cells mapped' + (f' · {cov["unmapped"]} not mapped' if cov["unmapped"] else "")))
    sub = " · ".join(x for x in [i.get("maker"), i.get("type"), f'{p["n_cylinders"]} cylinders' if p["n_cylinders"] else ""] if x)
    md(f'<div class="card gold"><div class="vs"><div><p class="lbl">Vessel</p><div class="vname">{esc(i.get("vessel") or rep["name"])}</div>'
       f'<div class="vsub">{esc(sub)}</div><div class="facts">{facts}</div></div>'
       f'<div class="ring-wrap"><div>{ring}</div><div class="bars">{bars}</div></div></div></div>')


def render_banner(p: dict) -> None:
    conf, issues = p["confidence"], p["issues"]
    warns = [m for s, m in issues if s == "warn"]
    n_ok = sum(1 for v in p["xval"].values() if v["label"] == "match")
    n_x = len(p["xval"])
    lvl = conf["level"]
    if lvl == "blank":
        md('<div class="banner blank"><div class="ic">📄</div><div><div class="bt">Blank form - no measurements entered</div><div class="bd">'
           'This workbook is a TEC-014 form without readings.</div></div></div>')
    elif lvl == "verified":
        md(f'<div class="banner ok"><div class="ic">✅</div><div><div class="bt">Extraction verified</div><div class="bd">Every cylinder reading reconciles with the averages '
           f'printed on the sheet ({n_ok} of {n_x} cross-checks matched exactly) and every filled cell was mapped.</div></div></div>')
    elif lvl == "high":
        extra = f" · {esc(warns[0])}" if warns else ""
        md(f'<div class="banner high"><div class="ic">ℹ️</div><div><div class="bt">High confidence</div><div class="bd">Values were extracted cleanly; {n_ok} of {n_x} cross-checks matched exactly{extra}. See the extraction audit.</div></div></div>')
    elif lvl == "review":
        items = "".join(f"<br>• {esc(w)}" for w in warns[:3]) or "<br>• Some values could not be reconciled with the sheet."
        md(f'<div class="banner review"><div class="ic">⚠️</div><div><div class="bt">Review recommended</div><div class="bd">Please compare the readings with the original sheet.{items}</div></div></div>')
    else:
        items = "".join(f"<br>• {esc(w)}" for w in warns[:3])
        md(f'<div class="banner low"><div class="ic">⛔</div><div><div class="bt">Low confidence - manual verification required</div><div class="bd">The extraction could not be reconciled with the sheet.{items}</div></div></div>')
    for s, m in issues:
        if s == "error":
            md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Structure error</div><div class="bd">{esc(m)}</div></div></div>')


def _steps(items) -> str:
    return "<br>".join(f"☐ {esc(a)}" for a in items)


def finding_card(f: dict) -> str:
    """One finding: what was seen, what to do, likely causes."""
    if f["cat"] == "data":
        tag = chip("Check entry", "warn" if f["sev"] == "warning" else "info")
    elif f["cat"] == "diagnosis":
        tag = chip("Diagnosis", SEV_CHIP[f["sev"]])
    else:
        tag = chip(SEV_LABEL[f["sev"]], SEV_CHIP[f["sev"]])
    if f.get("provisional"):
        tag += " " + chip("verify entry first", "warn")
    title = f["title"].removeprefix("Check entry · ")
    fault = "<br>".join(esc(x) for x in f["fault"].split("\n"))
    do = "".join(f'<div>☐ {esc(a)}</div>' for a in f["actions"])
    causes = f'<div class="note" style="margin-top:6px"><b>Likely:</b> {" · ".join(esc(c) for c in f["causes"])}</div>' if f["causes"] else ""
    cells = f'<div class="note">Cells {esc(f["cells"])}</div>' if f.get("cells") else ""
    return (f'<div class="fcard {f["sev"]}"><div class="ft"><span class="fh">{esc(title)}</span>{tag}{chip(URGENCY[f["when"]][1].split(" / ")[0], "")}</div>'
            f'<div class="ff">{fault}</div><div class="fa">{do}</div>{causes}{cells}</div>')


def note_card(n: dict) -> str:
    tag = chip("Check entry", "warn" if n["sev"] == "warning" else "info") if n["cat"] == "data" else chip(SEV_LABEL[n["sev"]], SEV_CHIP[n["sev"]])
    return (f'<div class="fcard {n["sev"]}"><div class="ft"><span class="fh">Other reports</span>{tag}</div><div class="ff">{esc(n["text"])}</div>'
            f'<div class="fa">☐ {esc(n["action"])}</div></div>')


def render_verdict(an: dict) -> None:
    v, c = an["verdict"], an["counts"]
    dq_tone = {"reliable": "ok", "some": "warn", "many": "crit"}[v["data_quality"]]
    chips = " ".join(chip(f"{c[s]} {SEV_LABEL[s]}", SEV_CHIP[s]) for s in ("critical", "mechanical", "warning", "advisory") if c.get(s)) or chip("No findings", "ok")
    chips += " " + chip(v["data_text"], dq_tone)
    prov = " · provisional" if v["provisional"] else ""
    md(f'<div class="banner {v["cls"]}"><div class="ic">{v["icon"]}</div><div><div class="bt">{esc(v["label"])}{prov}</div><div class="bd">{esc(v["headline"])}'
       + f'<div style="margin-top:.55rem">{chips}</div></div></div></div>')


PLAN_TONE = {"recheck": "advisory", "now": "critical", "soon": "warning", "planned": "advisory", "retest": "nominal"}


def render_plan(an: dict) -> None:
    """The crew action list: one line per step, grouped by when it is due. Hover a step to see what asks for it."""
    if not an["plan"]:
        md('<div class="fcard nominal"><div class="ft"><span class="fh">Nothing to correct</span></div><div class="ff">Repeat the test next month at the same load.</div></div>')
        return
    for key, (_, label) in sorted(URGENCY.items(), key=lambda kv: kv[1][0]):
        items = [x for x in an["plan"] if x["when"] == key]
        if not items:
            continue
        rows = "".join(f'<div class="ff" title="{esc("; ".join(x["because"][:3]))}">☐ {esc(x["step"])}</div>' for x in items)
        md(f'<div class="fcard {PLAN_TONE[key]}"><div class="ft"><span class="fh">{esc(label)}</span>{chip(str(len(items)), "gold")}</div>{rows}</div>')


def shop_panel(an: dict) -> str:
    """Typed -> used load for each shop-trial block, and whether the shop data can be trusted."""
    tone = {"standard": "ok", "corrected": "warn", "non_standard": "warn", "as_printed": "info"}
    chips = []
    for g, name in (("shopA", "Block A"), ("shopB", "Block B")):
        d = an["shop"][g]
        if d["printed"] is None:
            chips.append(chip(f"{name}: no load typed", "info"))
        elif d["status"] == "standard":
            chips.append(chip(f"{name}: {d['printed'] * 100:.0f} % ✓", "ok"))
        else:
            chips.append(chip(f"{name}: typed {d['printed'] * 100:.0f} % → used {d['used'] * 100:.0f} %", tone[d["status"]]))
    n_bad = sum(1 for r in an["ref"] if r["method"] == "inconsistent")
    n_use = sum(1 for r in an["ref"] if r["ref"] is not None)
    ok = chip(f"{n_use} references usable", "ok") if n_use else chip("no usable reference", "crit")
    bad = " " + chip(f"{n_bad} figures contradict each other", "warn") if n_bad else ""
    return f'<div style="margin:.2rem 0 .8rem">{" ".join(chips)} {ok}{bad}</div>'


def _c(v, nd=1):
    return f"{v:.{nd}f}" if v is not None and v == v else DASH


def _d(v, nd=1, suffix=""):
    if v is None or v != v:
        return DASH
    return f"{v:+.{nd}f}{suffix}"


def cylinder_table(p: dict, an: dict) -> str:
    unit = p["units"]["pmax"]
    rows = []
    for _, r in an["df"].iterrows():
        s = r["status"]
        rows.append([f'<span class="cyl-chip" style="min-width:30px;height:28px;font-size:.85rem">{int(r["id"])}</span>', _c(r["inj"], 1), _c(r["exh"], 0), _d(r["exh_dev"], 0),
                     _c(r["pmax"], 2), _d(r["pmax_dev_pct"], 1, " %"), _c(r["rack"], 1), _d(r["rack_dev_pct"], 1, " %"), chip(SEV_LABEL[s], SEV_CHIP[s])])
    m = an["mean"]
    rows.append(['<b style="color:#C4A35A;font-family:var(--disp);letter-spacing:.14em">MEAN</b>', _c(m["inj"], 2), _c(m["exh"], 1), "", _c(m["pmax"], 3), "", _c(m["rack"], 2), "", ""])
    return table_html(["Cyl", "Inj. ° BTDC", "Exhaust °C", "Δ °C", f"Pmax {unit}", "Δ %", "Rack mm", "Δ %", "Status"], rows)


def ref_table(an: dict) -> str:
    rows = []
    for r in an["ref"]:
        meth = {"interpolated": chip("interpolated", "ok"), "extrapolated": chip("extrapolated", "warn"), "single": chip("single point", "info"),
                "inconsistent": chip("inconsistent", "crit"), "out_of_range": chip("out of range", "warn"), "none": DASH}[r["method"]]
        dl = DASH if r["delta"] is None else (f'{r["delta"]:+.1f}' + (f' ({r["delta_pct"]:+.1f} %)' if r["delta_pct"] is not None else ""))
        imp = DASH if r.get("implied") is None else f'{r["implied"] * 100:.0f} %'
        rows.append([f'{esc(r["label"])} <span class="dim">· {esc(r["unit"])}{" · core" if r["core"] else ""}</span>', _c(r["act"], 2), _c(r["shopA"], 2), _c(r["shopB"], 2),
                     _c(r["ref"], 2), dl, imp, meth])
    return table_html(["Parameter", "Actual", "Shop A", "Shop B", "Reference", "Δ", "Implied load", "Method"], rows)


def _fv(p: dict, g: str, key: str) -> str:
    v = p["values"][g].get(key)
    if v is not None:
        return f"{v:g}"
    if p["kinds"][g].get(key) == "na":
        return '<span class="dim">N/A</span>'
    t = p["texts"][g].get(key)
    return esc(t) if t else DASH


def readings_tables(p: dict, an: Optional[dict] = None) -> None:
    la, lb = (an["shop"]["shopA"]["used"], an["shop"]["shopB"]["used"]) if an else (p["shop"]["shopA"]["load"], p["shop"]["shopB"]["load"])
    ha = f"Shop A{f' · {la * 100:.0f} %' if la else ''}"
    hb = f"Shop B{f' · {lb * 100:.0f} %' if lb else ''}"
    for title, keys in PARAM_LAYOUT:
        rows = []
        for k in keys:
            if all(_fv(p, g, k) == DASH for g in GROUPS):
                continue
            rows.append([esc(PARAMS[k][0]), _fv(p, "act", k), _fv(p, "shopA", k), _fv(p, "shopB", k), esc(PARAMS[k][1]), f'<span class="dim">{esc(p["refs"].get("act." + k, ""))}</span>'])
        if rows:
            md(f'<p class="lbl" style="margin-top:1rem">{esc(title)}</p>')
            md(table_html(["Parameter", "Actual", ha, hb, "Unit", "Cell"], rows))
    extras = [(lab, p["texts"]["act"].get(k)) for k, lab in (("fuel_mode", "Fuel mode"), ("fuel_grade", "Fuel grade"), ("smoke", "Exhaust smoke colour")) if p["texts"]["act"].get(k)]
    if extras:
        md('<p class="lbl" style="margin-top:1rem">Text entries</p><div class="card"><div class="facts">' + "".join(fact(a, b) for a, b in extras) + "</div></div>")


def render_audit(p: dict) -> None:
    names = {"exh": "Exhaust temperature", "pmax": "Pmax", "rack": "Fuel rack"}
    gname = {"act": "Actual", "shopA": "Shop A", "shopB": "Shop B"}
    tone = {"match": "ok", "near": "info", "weak": "warn", "mismatch": "crit"}
    rows = []
    for k, v in p["xval"].items():
        dec = max(v["dec"], 1)
        rows.append([names[v["metric"]], gname[v["group"]], f'{v["doc"]:.{v["dec"]}f}', f'{v["calc"]:.{dec + 1}f}', f'{v["calc"] - v["doc"]:+.{dec + 1}f}', chip(v["label"], tone[v["label"]])])
    md('<p class="lbl">Reconciliation with the averages printed on the sheet</p>')
    md(table_html(["Parameter", "Block", "Sheet average", "Calculated mean", "Δ", "Verdict"], rows) if rows
       else '<div class="card"><p class="note">The sheet contains no Average values to reconcile against.</p></div>')
    c = p["confidence"]
    md(f'<p class="note" style="margin-top:.5rem">Confidence {c["total"]:.1f}% = 20% structure · 30% completeness · 15% plausibility · 35% sheet cross-check, minus 2 points per warning. '
       'A mismatch can mean a mis-read value <i>or</i> an arithmetic slip by the crew on the sheet itself.</p>')
    md('<p class="lbl" style="margin-top:1.1rem">Extraction notes</p>')
    if p["issues"]:
        sevtone = {"info": "info", "warn": "warn", "error": "crit"}
        md("".join(f'<div style="margin:5px 0">{chip(s.upper(), sevtone[s])} <span style="font-size:.88rem">{esc(m)}</span></div>' for s, m in p["issues"]))
    else:
        md('<div class="note">No extraction notes - nothing unusual was found.</div>')
    cols = p["columns"]
    md('<p class="lbl" style="margin-top:1.1rem">Where the values were read</p>'
       f'<div class="card"><div class="facts">{fact("Worksheet", p["sheet"])}{fact("Actual values", "columns " + "–".join([cols["act"][0], cols["act"][-1]]))}'
       f'{fact("Shop trial A", "columns " + "–".join([cols["shopA"][0], cols["shopA"][-1]]))}{fact("Shop trial B", "columns " + "–".join([cols["shopB"][0], cols["shopB"][-1]]))}'
       f'{fact("Pmax unit read as", p["units"]["pmax"])}{fact("Sections found", str(len(p["sections"])))}</div></div>')
    cov = p["coverage"]
    md(f'<p class="lbl" style="margin-top:1.1rem">Sheet coverage</p><div class="note">{cov["claimed"]} of {cov["total"]} filled cells mapped to a field · '
       f'{cov["template"]} template / unit labels · <b>{cov["unmapped"]} not mapped</b>.</div>')
    if p["unclaimed"]:
        md(table_html(["Cell", "Text on the sheet", "Next to", "Above"], [[esc(u["cell"]), esc(u["text"][:120]), esc(u["left"][:40]), esc(u["above"][:40])] for u in p["unclaimed"]])
           + '<p class="note">These cells hold text the reader could not assign to a field. Nothing is hidden - check them against the workbook.</p>')


# ═════════════════════════════════════════════════════════════════════════════
#  exports
# ═════════════════════════════════════════════════════════════════════════════
def export_csv(p: dict, an) -> bytes:
    df = pd.DataFrame(flat_rows(p, an))
    return df.to_csv(index=False).encode("utf-8-sig")


def export_json(rep: dict) -> bytes:
    p, an = rep["parsed"], rep["an"]
    cyl = {g: {m: {str(k): v for k, v in arr.items()} for m, arr in blocks.items()} for g, blocks in p["cyl"].items()}
    doc = dict(generator=f"M.E. Command Center v{APP_VERSION} · TEC-014 reader v{PARSER_VERSION}", file=rep["name"], generated=time.strftime("%Y-%m-%d %H:%M:%S"),
               identity=p["identity"], tested_load=p["test"], shop_trial_loads=p["shop"], units=p["units"], cylinders=cyl, readings=p["values"], text_entries=p["texts"],
               printed_averages={g: {m: v[0] for m, v in d.items()} for g, d in p["avgs_doc"].items()},
               cross_checks={k: {kk: vv for kk, vv in v.items() if kk != "tier"} for k, v in p["xval"].items()}, confidence=p["confidence"],
               coverage=p["coverage"], unmapped_cells=p["unclaimed"], source_cells=p["refs"], notes=[dict(severity=s, message=m) for s, m in p["issues"]])
    if an:
        doc.update(health_score=an["score"], grade=an["grade"], verdict=an["verdict"], crew_actions=an["plan"], findings=an["findings"], shop_trial_loads_used=an["shop"],
                   not_measured=an["missing"], shop_trial_reference=an["ref"],
                   cylinder_table=an["df"].where(an["df"].notna(), None).to_dict(orient="records"))
    return json.dumps(doc, indent=2, default=str, ensure_ascii=False).encode("utf-8")


def export_html(rep: dict, notes: Optional[list] = None) -> bytes:
    """One printable page for the ship: verdict, action list, entries to re-check, findings, cylinders."""
    p, an = rep["parsed"], rep["an"]
    i, v = p["identity"], an["verdict"]
    unit = p["units"]["pmax"]

    def f(x, nd=1):
        return "—" if x is None or x != x else f"{x:.{nd}f}"
    cyl = "".join(f"<tr><td>{int(r['id'])}</td><td>{f(r['inj'])}</td><td>{f(r['exh'], 0)}</td><td>{f(r['pmax'], 2)}</td><td>{f(r['rack'])}</td><td>{SEV_LABEL[r['status']]}</td></tr>"
                  for _, r in an["df"].iterrows())
    plan = ""
    for key, (_, label) in sorted(URGENCY.items(), key=lambda kv: kv[1][0]):
        items = [s for s in an["plan"] if s["when"] == key]
        if items:
            plan += f"<h3>{esc(label)}</h3>" + "".join(f"<div class='st'>&#9744; {esc(s['step'])}<small> - {esc('; '.join(s['because'][:2]))}</small></div>" for s in items)
    data = "".join(f"<div class='f advisory'><b>{esc(x['title'])}</b><br>{'<br>'.join(esc(l) for l in x['fault'].split(chr(10)))}" + (f"<br><small>Sheet cells: {esc(x['cells'])}</small>" if x.get("cells") else "") + "</div>"
                   for x in an["findings"] if x["cat"] == "data")
    data += "".join(f"<div class='f advisory'>{esc(n['text'])}</div>" for n in (notes or []) if n["cat"] == "data")
    fnd = "".join(f"<div class='f {x['sev']}'><b>{esc(x['title'])}</b>{' (verify the entry first)' if x.get('provisional') else ''}<br>{esc(x['fault'])}"
                  + (f"<br><small>Likely causes: {esc(' / '.join(x['causes']))}</small>" if x["causes"] else "") + "</div>" for x in an["findings"] if x["cat"] != "data")
    ref = "".join(f"<tr><td>{esc(r['label'])}</td><td>{f(r['act'], 2)}</td><td>{f(r['shopA'], 2)}</td><td>{f(r['shopB'], 2)}</td><td>{f(r['ref'], 2)}</td><td>{esc(r['method'])}</td></tr>" for r in an["ref"])
    load = f"{an['load'] * 100:.0f} %" if an["load"] is not None else "unknown"
    shop = " and ".join(f"{d['used'] * 100:.0f} %" for d in an["shop"].values() if d["used"] is not None) or "-"
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>Aux. engine check - {esc(i.get('vessel'))} {esc(dg_label(p))}</title><style>
body{{font-family:Arial,Helvetica,sans-serif;color:#1b1b1b;margin:28px auto;max-width:900px;font-size:14px}}h1{{margin:0;letter-spacing:.03em}}h3{{margin:18px 0 6px}}.sub{{color:#666;margin-bottom:14px}}
table{{border-collapse:collapse;width:100%;margin:10px 0}}th,td{{border:1px solid #ccc;padding:5px 8px;text-align:right}}th{{background:#f3efe2}}td:first-child,th:first-child{{text-align:left}}
.f{{border-left:5px solid #999;background:#fafafa;padding:7px 11px;margin:7px 0}}.f.critical{{border-color:#d33}}.f.warning{{border-color:#d9a500}}.f.mechanical{{border-color:#e07b2a}}.f.advisory{{border-color:#4a8ac4}}
.v{{border:2px solid #444;padding:10px 14px;margin:12px 0}}.st{{margin:4px 0}}small{{color:#666}}.k{{display:inline-block;margin-right:20px}}@media print{{body{{margin:10px}}}}</style></head><body>
<h1>{esc(i.get('vessel') or rep['name'])} - {esc(dg_label(p))}: performance check and action sheet</h1>
<div class="sub">{esc(i.get('maker'))} {esc(i.get('type'))} · S/N {esc(i.get('serial'))} · {esc(i.get('output_text'))} · test {esc(str(i.get('date') or ''))} {esc(i.get('time') or '')} · source {esc(rep['name'])} · M.E. Command Center v{APP_VERSION}</div>
<div class="v"><b>Verdict: {esc(v['label'])}</b>{' (provisional)' if v['provisional'] else ''}<br>{esc(v['headline'])}<br><small>{esc(v['data_text'])}</small></div>
<div><span class="k"><b>Health</b> {an['score']}/100 ({an['grade']})</span><span class="k"><b>Load</b> {load}</span><span class="k"><b>Exhaust mean</b> {f(an['mean']['exh'], 0)} °C</span>
<span class="k"><b>Pmax mean</b> {f(an['mean']['pmax'], 2)} {esc(unit)}</span><span class="k"><b>Rack mean</b> {f(an['mean']['rack'])} mm</span></div>
<h2>What to do</h2>{plan or '<p>No action needed. Repeat the test next month at the same load.</p>'}
<h2>Entries to re-check on the sheet</h2>{data or '<p>No inconsistent entries found.</p>'}
<h2>Findings</h2>{fnd or '<p>No findings.</p>'}
<h3>Cylinders</h3><table><tr><th>Cyl</th><th>Inj. ° BTDC</th><th>Exhaust °C</th><th>Pmax {esc(unit)}</th><th>Rack mm</th><th>Status</th></tr>{cyl}</table>
<h3>Versus shop trial (compared at {shop} load points)</h3><table><tr><th>Parameter</th><th>Actual</th><th>Shop A</th><th>Shop B</th><th>Reference</th><th>Method</th></tr>{ref}</table>
<p style="color:#888;font-size:12px">Causes are indications drawn from the readings, not proof: the Chief Engineer decides. Limits are generic and editable; shop-trial figures are those printed on the sheet.</p></body></html>"""
    return page.encode("utf-8")


# ═════════════════════════════════════════════════════════════════════════════
#  report / overview
# ═════════════════════════════════════════════════════════════════════════════
#  cockpit: the part of the report that is read in a few seconds
# ═════════════════════════════════════════════════════════════════════════════
COCKPIT_CSS = """<style>
.board{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;margin:.2rem 0 1rem}
.tile{background:linear-gradient(180deg,var(--card),var(--bg2));border:1px solid var(--line);border-radius:13px;padding:.7rem .9rem .75rem;position:relative;overflow:hidden;min-height:86px}
.tile::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--faint)}
.tile.ok::before{background:var(--ok)}.tile.advisory::before{background:var(--info)}.tile.warning::before{background:var(--warn)}
.tile.mechanical::before{background:var(--orange)}.tile.critical::before{background:var(--crit)}.tile.none::before{background:var(--faint)}
.tile .tn{font-family:var(--disp);font-size:.74rem;letter-spacing:.16em;text-transform:uppercase;color:var(--muted);font-weight:600;display:flex;justify-content:space-between;gap:6px}
.tile .tv{font-family:var(--disp);font-size:1.12rem;font-weight:700;letter-spacing:.06em;margin-top:3px}
.tile.ok .tv{color:var(--ok)}.tile.advisory .tv{color:var(--info)}.tile.warning .tv{color:var(--warn)}.tile.mechanical .tv{color:var(--orange)}.tile.critical .tv{color:var(--crit)}.tile.none .tv{color:var(--muted)}
.tile .ts{font-size:.78rem;color:var(--muted);margin-top:2px;line-height:1.3}
.tile .fl{color:var(--warn);font-size:.72rem;margin-top:3px}
.heat{width:100%;border-collapse:separate;border-spacing:4px;table-layout:fixed}
.heat th{font-family:var(--disp);font-size:.72rem;letter-spacing:.14em;color:var(--muted);font-weight:600;text-transform:uppercase;text-align:center;padding:2px}
.heat th.rl,.heat td.rl{text-align:left;width:19%;color:var(--muted);font-size:.78rem}
.heat td{text-align:center;border-radius:9px;padding:.42rem .2rem;border:1px solid var(--line);background:var(--card)}
.heat td b{display:block;font-family:var(--mono);font-size:.98rem;font-weight:500;color:var(--text)}
.heat td i{display:block;font-style:normal;font-family:var(--mono);font-size:.72rem;color:var(--muted)}
.heat td.h-ok{background:var(--ok-s);border-color:rgba(91,185,140,.28)}
.heat td.h-advisory{background:var(--info-s);border-color:rgba(123,175,217,.3)}
.heat td.h-warning{background:var(--warn-s);border-color:rgba(229,181,74,.4)}
.heat td.h-mechanical{background:var(--orange-s);border-color:rgba(234,139,74,.45)}
.heat td.h-critical{background:var(--crit-s);border-color:rgba(232,100,94,.5)}
.heat td.mean{background:transparent;border-style:dashed}
.topact{display:flex;gap:10px;align-items:flex-start;padding:.5rem .1rem;border-bottom:1px solid var(--line);font-size:.9rem}
.topact:last-child{border-bottom:0}.topact .n{font-family:var(--mono);color:var(--gold2);min-width:1.2rem}
.topact .w{margin-left:auto;white-space:nowrap}
</style>"""

SYSTEMS = [("combustion", "Combustion"), ("fuel", "Fuel & injection"), ("air", "Charge air & T/C"), ("cooling", "Cooling water"),
           ("lube", "Lube oil"), ("gen", "Generator & load"), ("shop", "Shop-trial match"), ("entries", "Sheet entries")]
def short(text: str, n: int = 46) -> str:
    text = re.sub(r"^(Probable diagnosis · |Check entry · )", "", text)
    return text if len(text) <= n else text[:n - 1].rstrip() + "…"


def status_board(an: dict, notes: list) -> str:
    by: dict[str, list] = {k: [] for k, _ in SYSTEMS}
    for f in an["findings"]:
        by[f["system"]].append(f)
    n_data = len(by["entries"]) + sum(1 for n in notes if n["cat"] == "data")
    tiles = []
    for key, name in SYSTEMS:
        fs = sorted(by[key], key=lambda f: -SEV_RANK[f["sev"]])
        if key == "entries":
            tone = "warning" if any(f["sev"] == "warning" for f in by[key]) else "advisory" if n_data else "ok"
            val = f"{n_data} to re-check" if n_data else "Consistent"
            sub = short(fs[0]["title"], 52) if fs else (short(notes[0]["text"], 52) if n_data else f'{an["verdict"]["n_checks"]} checks passed')
            tiles.append(f'<div class="tile {tone}"><div class="tn"><span>{name}</span></div><div class="tv">{val}</div><div class="ts">{esc(sub)}</div></div>')
            continue
        if key == "shop":
            parts = [f'{d["printed"] * 100:.0f} → {d["used"] * 100:.0f} %' for d in an["shop"].values() if d["status"] == "corrected"]
            parts += [f'{d["printed"] * 100:.0f} % not standard' for d in an["shop"].values() if d["status"] == "non_standard"]
            n_bad = sum(1 for r in an["ref"] if r["method"] == "inconsistent")
            if n_bad:
                parts.append(f"{n_bad} figures inconsistent")
            if not any(r["delta"] is not None for r in an["ref"]) and not fs:
                tiles.append(f'<div class="tile none"><div class="tn"><span>{name}</span></div><div class="tv">No reference</div><div class="ts">{esc("; ".join(parts) or "shop trial not usable")}</div></div>')
                continue
            if not fs and parts:
                tiles.append(f'<div class="tile advisory"><div class="tn"><span>{name}</span></div><div class="tv">Check</div><div class="ts">shop load typed {esc("; ".join(parts))}</div></div>')
                continue
        if not fs:
            tiles.append(f'<div class="tile ok"><div class="tn"><span>{name}</span></div><div class="tv">OK</div><div class="ts">within limits</div></div>')
            continue
        w = fs[0]
        extra = f'<span>+{len(fs) - 1}</span>' if len(fs) > 1 else ""
        flag = '<div class="fl">verify entry first</div>' if w.get("provisional") else ""
        tiles.append(f'<div class="tile {w["sev"]}"><div class="tn"><span>{name}</span>{extra}</div><div class="tv">{SEV_LABEL[w["sev"]]}</div>'
                     f'<div class="ts">{esc(short(w["title"]))}</div>{flag}</div>')
    return '<div class="board">' + "".join(tiles) + "</div>"


def _cell_tone(kind: str, dev, th: dict, val=None) -> str:
    if dev is None or dev != dev:
        return ""
    a = abs(dev)
    if kind == "exh":
        if val is not None and val >= th["exh_abs_crit"]:
            return "critical"
        lim, crit = th["exh_dev"], th["exh_dev_crit"]
        if val is not None and val >= th["exh_abs"]:
            return "warning" if a < crit else "mechanical"
    elif kind == "pmax":
        lim, crit = th["pmax_dev"], th["pmax_dev_warn"]
    elif kind == "rack":
        lim, crit = th["rack_dev"], th["rack_dev_warn"]
    else:
        lim, crit = th["inj_spread"] / 2, th["inj_spread_warn"] / 2
    return "mechanical" if a >= crit else "warning" if a >= lim else "ok"


def cylinder_heatmap(p: dict, an: dict, th: dict) -> str:
    df, m, unit = an["df"], an["mean"], p["units"]["pmax"]
    ref = {r["key"]: r["ref"] for r in an["ref"] if r["ref"] is not None}
    spec = [("Exhaust °C", "exh", "exh_dev", "{:.0f}", "{:+.0f}", "exh_mean", ""), (f"Pmax {unit}", "pmax", "pmax_dev_pct", "{:.2f}", "{:+.1f}%", "pmax_mean", ""),
            ("Rack mm", "rack", "rack_dev_pct", "{:.1f}", "{:+.1f}%", "rack_mean", ""), ("Inj. ° BTDC", "inj", "inj_dev", "{:.1f}", "{:+.1f}°", None, "")]
    head = '<th class="rl"></th>' + "".join(f'<th>Cyl {int(i)}</th>' for i in df["id"]) + "<th>Mean</th><th>Shop ref</th>"
    rows = []
    for lab, col, dcol, fmt, dfmt, rk, _ in spec:
        if not df[col].notna().any():
            continue
        cells = []
        for _, r in df.iterrows():
            v, d = r[col], r[dcol]
            if v != v:
                cells.append("<td>—</td>")
                continue
            cells.append(f'<td class="h-{_cell_tone(col, d, th, v)}"><b>{fmt.format(v)}</b><i>{dfmt.format(d) if d == d else ""}</i></td>')
        rv = ref.get(rk) if rk else None
        rows.append(f'<tr><td class="rl">{esc(lab)}</td>{"".join(cells)}<td class="mean"><b>{fmt.format(m[col]) if m[col] is not None else "—"}</b></td>'
                    f'<td class="mean"><b>{fmt.format(rv) if rv is not None else "—"}</b></td></tr>')
    return f'<table class="heat"><thead><tr>{head}</tr></thead><tbody>{"".join(rows)}</tbody></table>'


def top_actions(an: dict, n: int = 3) -> str:
    """The few things to do first: sheet re-checks are one line (the list is in its tab), then the most urgent engine actions."""
    plan = an["plan"]
    if not plan:
        return '<p class="note">Nothing to correct. Repeat the performance test next month at the same load.</p>'
    check = [s for s in plan if s["when"] == "recheck"]
    rest = [s for s in plan if s["when"] != "recheck"]
    rows = []
    if check:
        rows.append(f'<div class="topact"><span class="n">!</span><span>{chip("Re-check first", "info")}<br>{len(check)} sheet entr{"y" if len(check) == 1 else "ies"} '
                    f'cannot be right - see the “Entries to re-check” tab before trusting the figures.</span></div>')
    for k, s in enumerate(rest[:n], 1):
        rows.append(f'<div class="topact"><span class="n">{k}</span><span>{chip(URGENCY[s["when"]][1].split(" / ")[0], SEV_CHIP[PLAN_TONE[s["when"]]])}<br>{esc(s["step"])}</span></div>')
    if len(rest) > n:
        rows.append(f'<p class="note" style="margin-top:.4rem">+ {len(rest) - n} more in the Actions tab</p>')
    return "".join(rows)


def findings_list(items: list, notes: list, empty: str) -> None:
    """Findings grouped by plant system, one collapsed line each (critical / mechanical start open)."""
    if not items and not notes:
        md(f'<div class="fcard nominal"><div class="ft"><span class="fh">{esc(empty)}</span></div></div>')
        return
    icon = {"critical": "🔴", "mechanical": "🟠", "warning": "🟡", "advisory": "🔵"}
    for key, name in SYSTEMS:
        grp = [f for f in items if f["system"] == key]
        if not grp:
            continue
        md(f'<p class="lbl" style="margin:.9rem 0 .3rem">{esc(name)} · {len(grp)}</p>')
        for f in grp:
            prov = " · verify entry" if f.get("provisional") else ""
            with st.expander(f'{icon[f["sev"]]}  {short(f["title"], 80)}  ·  {SEV_LABEL[f["sev"]]}{prov}', expanded=f["sev"] in ("critical", "mechanical")):
                md(finding_card(f))
    if notes:
        md('<p class="lbl" style="margin:.9rem 0 .3rem">Other reports</p>')
    for n in notes:
        with st.expander(f'{icon[n["sev"]]}  {short(n["text"], 90)}'):
            md(note_card(n))


def entries_table(an: dict, notes: list, rid: str = "") -> None:
    """Entries that cannot be right: what was typed, what the rule expects, which cells. Tick them off as the sheet is corrected."""
    data = [f for f in an["findings"] if f["cat"] == "data"]
    nd = [n for n in notes if n["cat"] == "data"]
    if not data and not nd:
        md(f'<div class="fcard nominal"><div class="ft"><span class="fh">Entries consistent</span></div><div class="ff">{an["verdict"]["n_checks"]} checks, none failed.</div></div>')
    for k, f in enumerate(data):
        rows = [[esc(e), f'<span style="font-family:var(--mono)">{esc(t)}</span>', f'<div style="text-align:left;font-family:var(--body)">{esc(w)}</div>'] for e, t, w in f["rows"]] or [[esc(f["fault"]), "", ""]]
        cells = f' <span class="note">· cells {esc(f["cells"])}</span>' if f.get("cells") else ""
        md(f'<div class="fcard {f["sev"]}" style="padding:.7rem 1.1rem"><div class="ft"><span class="fh" style="font-size:.95rem">{esc(f["title"].removeprefix("Check entry · "))}</span>{cells}</div>'
           + table_html(["Entry", "Typed", "Expected / rule"], rows) + f'<div class="fa" style="margin-top:.4rem">☐ {esc(f["actions"][0])}</div></div>')
        st.checkbox("Fixed on the sheet", key=f"dg_{rid}_fix_{k}")
    for n in nd:
        md(note_card(n))
    if data:
        md(f'<p class="note">{FIX_SHEET}</p>')
    if an["missing"]:
        md('<p class="note"><b>Not measured:</b> ' + esc(" · ".join(an["missing"])) + "</p>")


# ═════════════════════════════════════════════════════════════════════════════
def render_report(rep: dict, th: dict, notes_all: list) -> None:
    rid = rep["id"]
    p, an = rep["parsed"], rep["an"]
    if not p:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Could not read {esc(rep["name"])}</div><div class="bd">{esc(rep["error"] or "Unknown error")}</div></div></div>')
        return
    render_steps(3 if an else 2)
    lvl = p["confidence"]["level"]
    if lvl in ("review", "low", "blank") or any(sv == "error" for sv, _ in p["issues"]):
        render_banner(p)                                    # a problem with the extraction itself is never hidden
    if p["status"] == "blank":
        return
    if p["status"] != "ok" or an is None:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Analysis unavailable</div><div class="bd">{esc(rep["error"] or "The readings could not be analysed.")}</div></div></div>')
        with st.expander("Extraction audit · how every number was obtained"):
            render_audit(p)
        return
    lab = dg_label(p)
    vessel = p["identity"].get("vessel")
    notes = [n for n in notes_all if n["vessel"] == (vessel or rep["name"]) and lab in n["about"]]     # comparison notes, shown on the DG they concern
    L, t, sc, unit = an["load"], p["test"], an["score"], p["units"]["pmax"]
    v = an["verdict"]
    tone = "ok" if sc >= 85 else "" if sc >= 70 else "warn" if sc >= 50 else "crit"
    cards = [
        kpi("Engine health", f"{sc}", "/100", an["grade"], tone),
        kpi("Tested load", f"{L * 100:.0f}" if L is not None else "—", "%" if L is not None else "", f"{t['kw']:g} kW" + (f" · {t['amps']:g} A" if t.get("amps") else "") if t.get("kw") else "",
            "" if L is None else "ok" if L >= th["low_load"] else "warn"),
        kpi("Mean exhaust", _num(an["mean"]["exh"], 0), "°C" if an["mean"]["exh"] is not None else "", f"spread {_num(an['spread']['exh'], 0)} °C" if an["spread"]["exh"] is not None else ""),
        kpi("Mean Pmax", _num(an["mean"]["pmax_bar"], 1), "", f"bar · {_num(an['mean']['pmax'], 2)} {unit} · spread {_num(an['spread']['pmax'], 2)}"),
        kpi("Mean fuel rack", _num(an["mean"]["rack"], 1), "mm", f"spread {_num(an['spread']['rack'], 1)} mm" if an["spread"]["rack"] is not None else ""),
        kpi("Boost pressure", _num(p["values"]["act"].get("boost_p"), 2), "", "kgf/cm²"),
        kpi("Entries to re-check", str(v["n_data"] + sum(1 for n in notes if n["cat"] == "data")), "", f"{v['n_checks']} checks run", {"reliable": "ok", "some": "warn", "many": "crit"}[v["data_quality"]]),
    ]
    md(COCKPIT_CSS)
    i = p["identity"]
    when = i["date"].strftime("%d %b %Y") if i.get("date") else ""
    bits = [lab, when, f"{L * 100:.0f} % load" if L is not None else "", f"{i['rpm']:.0f} rpm" if i.get("rpm") else "", f"{i['run_hours']:,.0f} h" if i.get("run_hours") else ""]
    md(f'<div style="display:flex;align-items:baseline;gap:.9rem;flex-wrap:wrap;margin:.4rem 0 .2rem"><span style="font-family:var(--disp);font-size:1.5rem;font-weight:700;letter-spacing:.08em;'
       f'text-transform:uppercase">{esc(i.get("vessel") or rep["name"])}</span><span class="note">{esc(" · ".join(b for b in bits if b))}</span></div>')
    render_verdict(an)
    md(status_board(an, notes))
    left, right = st.columns([5, 4])
    with left:
        md('<p class="lbl" style="margin:.1rem 0 .3rem">Cylinders · colour = deviation from the engine mean</p>' + cylinder_heatmap(p, an, th))
    with right:
        md('<p class="lbl" style="margin:.1rem 0 .3rem">Do first</p>' + top_actions(an))
    md('<div class="kpis" style="grid-template-columns:repeat(auto-fit,minmax(128px,1fr));margin:.8rem 0 .4rem">' + "".join(cards) + "</div>")
    with st.expander(f"Sheet details · extraction confidence {p['confidence']['total']:.0f} %"):
        render_card(rep)
        if lvl not in ("review", "low", "blank"):
            render_banner(p)
    eng = [f for f in an["findings"] if f["cat"] != "data"]
    eng_notes = [n for n in notes if n["cat"] != "data"]
    n_chk = v["n_data"] + sum(1 for n in notes if n["cat"] == "data")
    t_act, t_find, t_chk, t_shop, t_ch, t_rd = st.tabs([f"Actions ({len(an['plan'])})", f"Findings ({len(eng) + len(eng_notes)})", f"Entries to re-check ({n_chk})",
                                                       "Shop trial", "Charts", "Readings & export"])
    with t_act:
        render_plan(an)
    with t_find:
        findings_list(eng, eng_notes, "All readings within limits")
    with t_chk:
        entries_table(an, notes, rid)
    with t_shop:
        md(shop_panel(an))
        fr = chart_ref(an["ref"], th)
        if fr is not None:
            plot(fr, f"dg_{rid}_refdev")
            md('<p class="note">* extrapolated beyond the shop-trial loads · grey = no tolerance defined.</p>')
        with st.expander("Full comparison table · implied load = shop-trial load at which the reading would be normal"):
            md(ref_table(an))
    with t_ch:
        df = an["df"]
        ref = {r["key"]: r for r in an["ref"]}

        def refv(k):
            r = ref.get(k)
            return (r["ref"], f"shop ref {r['ref']:.4g}") if r and r["ref"] is not None else (None, "")
        a, b = st.columns(2)
        with a:
            rv, rl = refv("exh_mean")
            plot(chart_cyl(df, "exh", "exh_dev", "Exhaust gas temperature", "°C", th["exh_dev"], th["exh_dev_crit"], "°C", "{:.0f}", rv, rl, an["mean"]["exh"]), f"dg_{rid}_exh")
        with b:
            rv, rl = refv("pmax_mean")
            plot(chart_cyl(df, "pmax", "pmax_dev_pct", "Pmax", unit, th["pmax_dev"], th["pmax_dev_warn"], "%", "{:.2f}", rv, rl, an["mean"]["pmax"]), f"dg_{rid}_pmax")
        a, b = st.columns(2)
        with a:
            rv, rl = refv("rack_mean")
            plot(chart_cyl(df, "rack", "rack_dev_pct", "Fuel rack position", "mm", th["rack_dev"], th["rack_dev_warn"], "%", "{:.1f}", rv, rl, an["mean"]["rack"]), f"dg_{rid}_rack")
        with b:
            plot(chart_cyl(df, "inj", "inj_dev", "Injection timing", "° BTDC", th["inj_spread"] / 2, th["inj_spread_warn"] / 2, "°", "{:.1f}", None, "", an["mean"]["inj"]), f"dg_{rid}_inj")
        md(cylinder_table(p, an))
    with t_rd:
        readings_tables(p, an)
        c1, c2, c3 = st.columns(3)
        base = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{vessel or rep['name']}_{lab}") + "_" + (p["identity"]["date"].strftime("%Y%m%d") if p["identity"].get("date") else "report")
        c1.download_button("All readings · CSV", export_csv(p, an), f"{base}.csv", "text/csv", key=f"dg_{rid}_csv", **stretch(st.download_button))
        c2.download_button("Full report · JSON", export_json(rep), f"{base}.json", "application/json", key=f"dg_{rid}_json", **stretch(st.download_button))
        c3.download_button("Crew action sheet · HTML", export_html(rep, notes), f"{base}.html", "text/html", key=f"dg_{rid}_html", **stretch(st.download_button))
    with st.expander("Extraction audit · how every number was obtained"):
        render_audit(p)


def render_overview(reps: list, th: dict, peers_all: list) -> None:
    rows, ok_rows, by_dg, per_vessel = [], [], {}, {}
    for r in reps:
        p = r["parsed"]
        if not p:
            rows.append([esc(r["name"]), "—", "—", "—", "—", chip("unreadable", "crit"), "—", "—", "—", "—", "—", "—", "—", "—"])
            continue
        s = summary_row(r["name"], p, r["an"])
        an = r["an"]
        lvl_tone = {"verified": "ok", "high": "info", "review": "warn", "low": "crit", "blank": "gold"}[p["confidence"]["level"]]
        rows.append([f'<b>{esc(s["vessel"])}</b>', esc(s["dg"]), s["date"].strftime("%d %b %Y") if s["date"] else "—", f'{s["load"] * 100:.0f} %' if s["load"] is not None else "—",
                     f'{s["run_hours"]:,.0f}' if s["run_hours"] else "—", chip(f'{s["confidence"]:.0f}%' if p["confidence"]["level"] != "blank" else "blank", lvl_tone),
                     esc(s["verdict"]) if an else "—", str(s["score"]) if an else "—", (chip(str(s["crit"]), "crit") if s["crit"] else '<span class="dim">0</span>') if an else "—",
                     (chip(str(s["warn"]), "warn") if s["warn"] else '<span class="dim">0</span>') if an else "—",
                     (chip(str(s["n_data"]), "info") if s["n_data"] else '<span class="dim">0</span>') if an else "—", _num(s.get("exh"), 0) if an else "—",
                     _num(s.get("pmax_bar"), 1) if an else "—", _num(s.get("boost"), 2) if an else "—"])
        if an:
            ok_rows.append(dict(label=f'{s["vessel"]} · {s["dg"]}', score=an["score"]))
            per_vessel.setdefault(s["vessel"], []).append(dict(label=s["dg"], exh=s["exh"], load=s["load"]))
            if s["date"] and s["dg_no"]:
                by_dg.setdefault((s["vessel"], s["dg"]), []).append(dict(date=s["date"], exh=s["exh"], pmax=s["pmax_bar"], rack=s["rack"], boost=s["boost"], load=s["load"]))
    section("F1", "Auxiliary engines overview", f"{len(reps)} report(s) loaded")
    md(table_html(["Vessel", "DG", "Test date", "Load", "Run. hours", "Confidence", "Verdict", "Health", "Crit. / mech.", "Warnings", "Re-check", "Exhaust °C", "Pmax bar", "Boost"], rows))
    if peers_all:
        section("F2", "Comparisons between reports", "generator vs generator, and the same generator over time")
        for n in peers_all:
            md(f'<div class="fcard {n["sev"]}"><div class="ft"><span class="fh">{esc(n["vessel"])}</span>'
               f'{chip("Check entry", "warn" if n["sev"] == "warning" else "info") if n["cat"] == "data" else chip(SEV_LABEL[n["sev"]], SEV_CHIP[n["sev"]])}</div>'
               f'<div class="ff">{esc(n["text"])}</div><div class="fa"><b>Do</b>{esc(n["action"])}</div></div>')
    for vessel, lst in per_vessel.items():
        if len(lst) >= 2:
            md(f'<p class="lbl" style="margin-top:1rem">{esc(vessel)} · exhaust temperature versus load</p>')
            plot(chart_dg_compare(lst), f"dg_cmp_{vessel}")
    if ok_rows:
        section("F3", "Health ranking")
        from me_engine import chart_fleet_health
        plot(chart_fleet_health(sorted(ok_rows, key=lambda r: -r["score"])), "dg_fleet_health")
    trends = {k: sorted(v, key=lambda t: t["date"]) for k, v in by_dg.items() if len({t["date"] for t in v}) >= 2}
    if trends:
        section("F4", "Trends", "same generator · several dates")
        md('<p class="note">Compare like with like: a trend is only meaningful when the tests were made at a similar load (hover shows the load).</p>')
        for (vessel, dg), pts in trends.items():
            md(f'<p class="lbl">{esc(vessel)} · {esc(dg)}</p>')
            plot(chart_dg_trend(pts), f"dg_trend_{vessel}_{dg}")
    else:
        md('<p class="note" style="margin-top:1rem">Upload several months for the same generator to unlock month-to-month trend charts.</p>')


# ═════════════════════════════════════════════════════════════════════════════
#  entry point
# ═════════════════════════════════════════════════════════════════════════════
def render_aux_tab() -> None:
    init_state()
    th = render_sidebar_thresholds()
    header = st.container()
    files = st.file_uploader("Drop TEC-014 aux. engine reports (.xlsx) here", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True,
                             key=f"dg_up_{st.session_state['dg_gen']}", label_visibility="collapsed",
                             help="One Excel workbook per generator. You can drop several generators and vessels at once.")
    reps, dups = load_reports(files, th)
    with header:
        render_header(len(reps))
    render_sidebar_actions(len(reps))
    if dups:
        md(f'<p class="note">Identical file(s) ignored: {esc(", ".join(dups))}.</p>')
    if not reps:
        render_steps(1)
        render_empty()
        return
    items = [(r["name"], r["parsed"], r["an"]) for r in reps if r["parsed"] and r["an"]]
    peers_all = peer_notes(items) + history_notes(items)
    if len(reps) == 1:
        render_report(reps[0], th, peers_all)
        return
    ids = [r["id"] for r in reps]
    labels = {r["id"]: rep_label(r) for r in reps}
    if st.session_state.get("dg_active") not in ids:
        st.session_state["dg_active"] = ids[0]
    t1, t2 = st.tabs(["Report", "Overview"])
    with t1:
        st.selectbox("Report", ids, format_func=lambda i: labels[i], key="dg_active")
        render_report(next(r for r in reps if r["id"] == st.session_state["dg_active"]), th, peers_all)
    with t2:
        render_overview(reps, th, peers_all)
