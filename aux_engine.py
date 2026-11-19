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

    def add(self, sev, title, fault, actions, causes=(), cyl=None, topic="engine", when=None, uses=(), cells="") -> None:
        cat = topic if topic in ("data", "diagnosis", "reference", "load") else "engine"
        self.F.append(dict(sev=sev, title=title, fault=fault, actions=list(actions), causes=list(causes), cyl=cyl, topic=topic, cat=cat, uses=tuple(uses),
                           cells=cells, when=when or ("recheck" if cat == "data" else URGENCY_BY_SEV[sev]), provisional=False))


SAME_TEST = "Repeat the test at the same load to confirm the reading."
FIX_SHEET = "Correct the sheet and upload the report again."


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
    ("jacket", "cw_engine_out", "cw_jacket_in", 1, "cooling water picks up heat in the engine, so it must leave warmer than it enters"),
    ("airc_w", "cw_ac_out", "cw_ac_in", 1, "cooling water picks up heat in the air cooler, so it must leave warmer than it enters"),
    ("loc_w", "cw_lo_out", "cw_lo_in", 1, "cooling water picks up heat in the L.O. cooler, so it must leave warmer than it enters"),
    ("loc_o", "lo_t_in", "lo_t_out", 1, "lube oil is cooled in the cooler, so it must enter hotter than it leaves"),
    ("loc_o", "lo_t_out", "cw_lo_in", 0, "lube oil cannot leave the cooler colder than the water that cools it"),
    ("air", "ba_ac_in", "ba_cyl_in", 1, "charge air is cooled in the air cooler, so it must enter hotter than it leaves"),
    ("air", "ba_cyl_in", "cw_ac_in", 2, "charge air cannot be cooled to (or below) the temperature of the cooling water entering the air cooler - "
                                        "a gap of several degrees is normal"),
    ("sea", "cw_ac_in", "sea_t", 0, "the low-temperature cooling water is cooled by the sea, so it cannot be colder than the sea water"),
    ("sea", "cw_lo_in", "sea_t", 0, "the low-temperature cooling water is cooled by the sea, so it cannot be colder than the sea water"),
    ("tc", "exh_mean", "tc_out", 30, "gas cools down in the turbine, so the T/C outlet must be well below the cylinder exhaust temperatures"),
    ("tc", "exh_mean", "tc_pair_out", 30, "gas cools down in the turbine, so the T/C outlet must be well below the cylinder exhaust temperatures"),
    ("tc", "tc_pair_in", "tc_pair_out", 60, "the turbine takes a large temperature drop out of the gas (normally more than 60 °C at this load)"),
    ("tc", "tc_in", "tc_out", 60, "the turbine takes a large temperature drop out of the gas (normally more than 60 °C at this load)"),
    *[("tc", "exh_mean", k, -30, "the gas arriving at the turbine cannot be much hotter than the cylinder exhaust temperatures it is made of")
      for k in ("tc_pair_in", "tc_in", "tc_in_upper", "tc_in_lower")],
    ("bearing", "gen_brg_de", "ambient_t", -3, "a running generator bearing cannot be colder than the air around it"),
    ("bearing", "gen_brg_nde", "ambient_t", -3, "a running generator bearing cannot be colder than the air around it"),
]
PAIR_GROUPS = {   # group -> (title, what the crew should do)
    "jacket": ("Jacket cooling-water temperatures", ["Re-read the jacket-water inlet and outlet thermometers at steady load."]),
    "airc_w": ("Air-cooler water temperatures", ["Re-read the air-cooler water inlet and outlet thermometers."]),
    "loc_w": ("L.O.-cooler water temperatures", ["Re-read the L.O.-cooler water inlet and outlet thermometers."]),
    "loc_o": ("L.O.-cooler oil temperatures", ["Re-read the L.O. temperatures before and after the cooler (and the cooling-water inlet)."]),
    "air": ("Charge-air temperatures", ["Re-read the charge-air temperature after the air cooler and the cooling-water temperature entering it; "
                                        "check the thermometer / sensor."]),
    "sea": ("Cooling water versus sea water", ["Re-read the low-temperature cooling-water and the sea-water temperatures."]),
    "tc": ("Turbocharger and exhaust temperatures", ["Re-read the T/C inlet and outlet thermometers and the cylinder exhaust thermocouples "
                                                    "(same load, same time).", "If they are confirmed, a very small temperature drop across the turbine points to a fouled "
                                                    "turbine / nozzle ring or an exhaust leak: inspect and water-wash / clean the turbine side."]),
    "bearing": ("Generator bearing temperatures", ["Check the bearing temperature sensors / thermometers and re-read them."]),
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
        impossible = gap < min(gap_min, 0)
        bad.setdefault(grp, []).append(dict(hot=hot, cold=cold, a=a, b=b, gap=gap, why=why, impossible=impossible))
    for grp, items in bad.items():
        title, steps = PAIR_GROUPS[grp]
        lines, seen = [], set()
        for x in items:
            if x["why"] in seen:
                lines.append(f"{LABELS[x['hot']]} {_g(x['a'])} °C and {LABELS[x['cold']]} {_g(x['b'])} °C (difference {_g(x['gap'])} °C).")
            else:
                seen.add(x["why"])
                lines.append(f"{LABELS[x['hot']]} {_g(x['a'])} °C and {LABELS[x['cold']]} {_g(x['b'])} °C (difference {_g(x['gap'])} °C): {x['why']}.")
        keys = {k for x in items for k in (x["hot"], x["cold"]) if k != "exh_mean"}
        sev = "warning" if any(x["impossible"] for x in items) else "advisory"
        c.add(sev, f"Check entry · {title.lower()}", "\n".join(lines), [*steps, FIX_SHEET], topic="data", uses=sorted(keys), cells=c.cells(*sorted(keys)))


def check_header(c: Ctx) -> None:
    """Load / current / rating / date / fuel entries that cannot be right."""
    i, t, v = c.i, c.p["test"], c.v
    cells_t = c.p["test"].get("ref") or ""
    # kW against amps: even at the extremes of the usual voltage range the power factor must lie between 0.5 and 1.0
    if t.get("kw") and t.get("amps"):
        c.checks_run += 1
        pf_hi, pf_lo = t["kw"] / (1.732 * 380 * t["amps"] / 1000), t["kw"] / (1.732 * 480 * t["amps"] / 1000)
        if pf_lo > 1.0:
            c.add("warning", "Check entry · generator current too low for the power", f"{t['kw']:g} kW at only {t['amps']:g} A would need a power factor above 1.0 "
                  f"(at 380-480 V).", ["Re-read the generator kW and ampere meters.", FIX_SHEET], topic="data", uses=["load"])
        elif pf_hi < 0.5:
            c.add("warning", "Check entry · generator current too high for the power", f"{t['kw']:g} kW at {t['amps']:g} A would mean a power factor below 0.5 "
                  f"(at 380-480 V).", ["Re-read the generator kW and ampere meters.", FIX_SHEET], topic="data", uses=["load"])
    kw_r, kva_r = i.get("rated_kw"), i.get("rated_kva")
    if kw_r and kva_r:
        c.checks_run += 1
        if not 0.6 <= kw_r / kva_r <= 1.0:
            c.add("advisory", "Check entry · rated output", f"{kw_r:g} kW / {kva_r:g} kVA is a power factor of {kw_r / kva_r:.2f}; generators are normally rated at 0.8.",
                  ["Compare with the generator name-plate.", FIX_SHEET], topic="data")
    if c.L is not None and c.L > 1.05:
        c.add("advisory", "Check entry · tested load above the rating", f"The test was recorded at {_pct(c.L)} of the {i.get('rated_kw') or 0:g} kW rating ({cells_t}).",
              ["Check the kW reading and the rated output.", FIX_SHEET], topic="data", uses=["load"])
    d = i.get("date")
    if d:
        c.checks_run += 1
        age = (dt.date.today() - d).days
        if age < -1:
            c.add("warning", "Check entry · report date in the future", f"The report is dated {d:%d %b %Y}.", ["Correct the date on the sheet.", FIX_SHEET], topic="data")
        elif age > 400:
            c.add("advisory", "Check entry · report is more than a year old", f"The report is dated {d:%d %b %Y}.", ["Check the date, or use a current report."], topic="data")
    rho = i.get("density")
    if rho is not None:
        c.checks_run += 1
        if not 0.80 <= rho <= 1.10:
            c.add("advisory", "Check entry · fuel density", f"Density {rho:g} is outside the usual 0.80-1.10 t/m³ (a value in kg/m³ was probably typed).",
                  ["Enter the density in t/m³ (for example 0.98).", FIX_SHEET], topic="data")
    rpm = i.get("rpm")
    if rpm is not None:
        c.checks_run += 1
        if not 300 <= rpm <= 1900:
            c.add("advisory", "Check entry · engine speed", f"{rpm:g} rpm is not a usual generator-set speed.", ["Check the rpm entry.", FIX_SHEET], topic="data")
    mode = (c.p["texts"]["act"].get("fuel_mode") or "").upper()
    fo_t, tank = v.get("fo_t_in"), v.get("hfo_tank_t")
    if fo_t is not None and tank is not None and mode.startswith("HFO"):
        c.checks_run += 1
        if fo_t < tank - 5:
            c.add("advisory", "Check entry · fuel temperatures", f"Fuel at the engine inlet ({fo_t:g} °C) is colder than the service tank heating ({tank:g} °C); "
                  "the heater should only add heat.", ["Re-read both temperatures.", FIX_SHEET], topic="data", uses=["fo_t_in", "hfo_tank_t"])
    ex = [x for x in c.df["exh"] if _ok(x)]
    if len(ex) >= 4:
        c.checks_run += 1
        if len(set(ex)) == 1:
            c.add("advisory", "Check entry · identical exhaust temperatures", f"All {len(ex)} cylinders show exactly {ex[0]:g} °C - real engines never do; the same value was probably copied.",
                  ["Read every cylinder separately.", FIX_SHEET], topic="data", uses=["exh"])


def check_ranges(c: Ctx) -> None:
    """Readings outside any physically believable range (typing slips such as 720 instead of 72)."""
    out = []
    for k, x in c.v.items():
        if x is None or k not in PLAUSIBLE:
            continue
        c.checks_run += 1
        lo, hi = PLAUSIBLE[k]
        if not lo <= x <= hi:
            out.append((k, f"{PARAMS[k][0]} {x:g} {PARAMS[k][1]} (believable range {lo:g}-{hi:g})"))
    for m, arr in c.p["cyl"]["act"].items():
        for i, x in arr.items():
            if x is None:
                continue
            c.checks_run += 1
            lo, hi = CYL_PLAUSIBLE["pmax_bar" if m == "pmax" else m]
            xx = x * c.p["units"]["pmax_factor"] if m == "pmax" else x
            if not lo <= xx <= hi:
                out.append((m, f"{CYL_METRICS[m][0]}, cylinder {i}: {x:g}"))
    if out:
        c.add("warning", "Check entry · readings outside the believable range", "; ".join(t for _, t in out) + ".",
              ["Re-read these instruments, and check the decimal point and the unit.", FIX_SHEET], topic="data", uses=sorted({k for k, _ in out}),
              cells=c.cells(*[k for k, _ in out if k in c.v]))


def check_sheet_arithmetic(c: Ctx) -> None:
    for x in c.p["xval"].values():
        if x["group"] == "act":
            c.checks_run += 1
            if x["label"] in ("weak", "mismatch"):
                c.add("advisory", f"Check entry · {CYL_LABEL[x['metric']]} average on the sheet", f"Printed average {x['doc']:g}, calculated from the cylinders {x['calc']:.2f}.",
                      ["A cylinder reading was mis-typed or the average was mis-calculated: compare with the original log.", FIX_SHEET], topic="data", uses=[x["metric"]])


def check_shop_block(c: Ctx) -> None:
    """The shop-trial block is typed in by hand too: wrong load points and impossible values are reported here."""
    bad_pts = [(g, d) for g, d in c.loads.items() if d["status"] in ("corrected", "non_standard", "as_printed")]
    if bad_pts:
        c.checks_run += 1
        std = " and ".join(_pct(s) for s in parse_points(c.th["shop_points"]))
        lines = [f"Block {'A' if g == 'shopA' else 'B'}: {d['note']}." for g, d in bad_pts]
        used = " and ".join(_pct(d["used"]) for d in c.loads.values() if d["used"] is not None)
        tail = (f" The comparison below uses {used}." if c.th["shop_mode"] == "auto" and any(d["status"] == "corrected" for _, d in bad_pts)
                else " The comparison below uses the loads as typed.")
        c.add("warning", "Check entry · shop-trial load points", f"The shop trials were run at {std} load. " + " ".join(lines) + tail,
              ["Copy the two load points exactly as printed in the maker's shop-trial report - not the load of today's test.", FIX_SHEET],
              topic="data", uses=["shop"], cells=", ".join(c.p["shop"][g]["ref"] for g, _ in bad_pts if c.p["shop"][g].get("ref")))
    c.checks_run += len(c.ref)
    bad = [r for r in c.ref if r["method"] == "inconsistent"]
    if bad:
        names = "; ".join(f"{r['label']} ({_g(r['shopA'])} → {_g(r['shopB'])} {r['unit']})" for r in bad)
        c.add("advisory", "Check entry · shop-trial figures look wrong",
              f"These shop-trial readings do not behave like real engine data between the two load points: {names}. They were not used as a reference.",
              ["Copy the shop-trial figures again from the maker's test report.", FIX_SHEET], topic="data", uses=["shop"])


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
        c.add("advisory", "Tested load unknown", "The load could not be derived from the sheet, so load-dependent comparisons were skipped.",
              ["Check the 'TESTED AT MAX. SAFE LOAD' entry and the rated output.", FIX_SHEET], topic="data", uses=["load"])
    elif L < 0.25:
        c.add("warning", f"Performance test at only {L * 100:.0f} % load",
              f"The sheet was recorded at {raw} ({L * 100:.0f} % of the {c.i['rated_kw'] or 0:.0f} kW rating). Temperatures, Pmax and rack at this load are not "
              "representative and comparisons with the shop trial become extrapolations.",
              ["Repeat the test at the maximum safe load (at least 50 %, ideally 75 %) as the form requires."], topic="load", when="retest")
    elif L < th["low_load"]:
        c.add("advisory", f"Part-load test · {L * 100:.0f} % of rated",
              f"Recorded at {raw} ({L * 100:.0f} % of rated). Below {th['low_load'] * 100:.0f} % the readings are of limited value for trending.",
              ["Prefer the maximum safe load for the monthly test so that months can be compared."], topic="load")


def check_cylinders(c: Ctx) -> None:
    th, df, m = c.th, c.df, c.m
    for _, r in df.iterrows():
        i, e, d = int(r["id"]), r["exh"], r["exh_dev"]
        if _ok(e):
            if e >= th["exh_abs_crit"]:
                c.add("critical", f"Cyl {i} · exhaust {e:.0f} °C", f"Above the critical limit of {th['exh_abs_crit']:.0f} °C.",
                      ["Reduce the load on this generator now and transfer load if possible.", "Confirm the reading on a second instrument, then find the cause: injector, exhaust "
                       "valve seating, air cooler and turbocharger condition."],
                      ["Leaking or burnt exhaust valve", "Late combustion (injector, fuel pump, timing)", "Too little charge air: dirty air cooler or turbocharger", "Overload"],
                      i, "exhaust", uses=["exh"])
            elif e >= th["exh_abs"]:
                c.add("warning", f"Cyl {i} · high exhaust {e:.0f} °C", f"At or above the {th['exh_abs']:.0f} °C limit.",
                      ["Check the charge-air temperature, the air cooler (air and water side) and the turbocharger condition.", "Check injection timing and the injector of the hot cylinder(s)."],
                      ["Dirty air cooler or turbocharger", "High charge-air or cooling-water temperature", "Injector or timing", "Exhaust valve seating"], i, "exhaust", uses=["exh"])
        if _ok(d):
            hi = d > 0
            if abs(d) >= th["exh_dev_crit"]:
                c.add("mechanical", f"Cyl {i} · exhaust {d:+.0f} °C vs mean", f"{e:.0f} °C against an engine mean of {m['exh']:.0f} °C (limit ±{th['exh_dev_crit']:.0f} °C).",
                      (["Check the exhaust valve seating, injector and fuel pump of this cylinder; plan an overhaul if confirmed.", "Confirm the thermocouple first."] if hi else
                       ["Confirm the thermocouple against a portable instrument.", "Test the injector and check the fuel pump of this cylinder."]),
                      (["Late combustion", "Leaking exhaust valve", "Worn injector"] if hi else ["Blocked or weak injector", "Low fuel delivery (pump wear, rack)", "Faulty thermocouple"]),
                      i, "exhaust", uses=["exh"])
            elif abs(d) >= th["exh_dev"]:
                c.add("warning", f"Cyl {i} · exhaust {d:+.0f} °C vs mean", f"{e:.0f} °C against an engine mean of {m['exh']:.0f} °C (limit ±{th['exh_dev']:.0f} °C).",
                      ["Compare injector and fuel-pump condition and timing with the other cylinders.", "Confirm the thermocouple."],
                      ["Injector or fuel-pump difference", "Thermocouple fault"], i, "exhaust", uses=["exh"])
        d = r["pmax_dev_pct"]
        if _ok(d) and abs(d) >= th["pmax_dev"]:
            sev = "warning" if abs(d) >= th["pmax_dev_warn"] else "advisory"
            c.add(sev, f"Cyl {i} · Pmax {d:+.1f} % vs mean", f"{r['pmax']:g} {c.p['units']['pmax']} against a mean of {m['pmax']:.2f}.",
                  ["Re-measure Pmax on this cylinder to confirm.", "Check injection timing, injector and compression of this cylinder before adjusting anything."],
                  (["Advanced injection timing", "Fuel pump delivering early"] if d > 0 else ["Retarded injection timing", "Injector or fuel-pump fault", "Compression leakage (rings, liner, valves)"]),
                  i, "pmax", uses=["pmax"])
        d = r["rack_dev_pct"]
        if _ok(d) and abs(d) >= th["rack_dev"]:
            sev = "warning" if abs(d) >= th["rack_dev_warn"] else "advisory"
            c.add(sev, f"Cyl {i} · fuel rack {d:+.1f} % vs mean", f"{r['rack']:g} mm against a mean of {m['rack']:.2f} mm.",
                  ["Check the fuel pump and the rack linkage of this cylinder; calibrate the pump if needed.", "Compare the injector condition with the other cylinders."],
                  (["Worn pump plunger / barrel (more rack needed for the same delivery)", "Rack linkage out of adjustment"] if d > 0
                   else ["Pump delivering more per mm of rack (calibration)", "Linkage friction or sticking"]), i, "rack", uses=["rack"], when="soon" if sev == "warning" else "planned")
    sp = _spread(df["inj"])
    if sp is not None and sp >= th["inj_spread"]:
        sev = "warning" if sp >= th["inj_spread_warn"] else "advisory"
        lo_c = [int(i) for i, x in zip(df["id"], df["inj"]) if x == df["inj"].min()]
        hi_c = [int(i) for i, x in zip(df["id"], df["inj"]) if x == df["inj"].max()]
        c.add(sev, f"Injection timing differs by {sp:.1f}° between cylinders",
              f"Lowest {df['inj'].min():g}° BTDC (cyl {', '.join(map(str, lo_c))}), highest {df['inj'].max():g}° (cyl {', '.join(map(str, hi_c))}).",
              ["Re-check the fuel pump timing marks / VIT setting of the outlying cylinders against the maker's value."],
              ["Fuel pump timing set differently on the units"], None, "timing", uses=["inj"], when="planned")


def check_plant(c: Ctx) -> None:
    """Cooling water, lube oil, charge air, fuel, generator bearings and smoke."""
    th, v = c.th, c.v
    jin, jout = v.get("cw_jacket_in"), v.get("cw_engine_out")
    if jout is not None and jout >= th["jacket_out_max"]:
        c.add("warning", f"Jacket cooling water outlet {jout:.0f} °C", f"Limit {th['jacket_out_max']:.0f} °C.",
              ["Check the cooling-water flow (pump, strainers, valves, air in the system).", "Check the thermostatic valve and the cooler."],
              ["Low flow", "Thermostat / temperature controller", "Fouled cooler"], None, "cooling", uses=["cw_engine_out"])
    if jin is not None and jout is not None and jout - jin > 12:
        c.add("advisory", f"Jacket cooling water rise {jout - jin:.0f} °C", f"Inlet {jin:.0f} °C → outlet {jout:.0f} °C; the usual rise is 5-10 °C.",
              ["Check the jacket-water pump, strainers and venting."], ["Low circulating flow"], None, "cooling", uses=["cw_jacket_in", "cw_engine_out"])
    lo_o = v.get("lo_t_out")
    if lo_o is not None and lo_o >= th["lo_out_max"]:
        c.add("warning", f"L.O. cooler outlet {lo_o:.0f} °C", f"Limit {th['lo_out_max']:.0f} °C.",
              ["Clean / check the L.O. cooler (oil and water side).", "Check the cooling-water flow to the cooler and the thermostatic valve."],
              ["Fouled L.O. cooler", "Low cooling-water flow", "Thermostatic valve"], None, "lube", uses=["lo_t_out"])
    bc = v.get("ba_cyl_in")
    if bc is not None and bc >= th["boost_air_max"]:
        c.add("warning", f"Boost air at cylinder inlet {bc:.0f} °C", f"Limit {th['boost_air_max']:.0f} °C.",
              ["Clean the air cooler (air and water side).", "Check the temperature and flow of the cooling water entering the air cooler."],
              ["Fouled air cooler", "Warm or insufficient cooling water"], None, "air", uses=["ba_cyl_in"])
    lp = v.get("lo_p_engine")
    if lp is not None:
        if lp < th["lo_p_min"] - 1.0:
            c.add("critical", f"L.O. pressure (engine) {lp:g} kgf/cm²", f"Well below the {th['lo_p_min']:g} kgf/cm² minimum.",
                  ["Unload / stop the engine and investigate: oil level, pump, filter differential, bearing wear."],
                  ["Low oil level", "Pump or relief valve fault", "Blocked filter", "Worn bearings"], None, "lube", uses=["lo_p_engine"])
        elif lp < th["lo_p_min"]:
            c.add("warning", f"L.O. pressure (engine) {lp:g} kgf/cm²", f"Below the {th['lo_p_min']:g} kgf/cm² minimum.",
                  ["Check the filter differential, the oil temperature / viscosity and the pressure regulating valve."],
                  ["Dirty filter", "Hot, thin oil", "Regulating valve setting"], None, "lube", uses=["lo_p_engine"])
    lpt = v.get("lo_p_tc")
    if lpt is not None and lpt < th["lo_p_tc_min"]:
        c.add("warning", f"L.O. pressure (T/C) {lpt:g} kgf/cm²", f"Below the {th['lo_p_tc_min']:g} kgf/cm² minimum.",
              ["Check the T/C oil supply line, filter and orifice."], ["Blocked supply line / orifice"], None, "lube", uses=["lo_p_tc"])
    fo_t = v.get("fo_t_in")
    if fo_t is not None and (c.p["texts"]["act"].get("fuel_mode") or "").upper().startswith("HFO") and fo_t < 80:
        c.add("advisory", f"FO inlet temperature {fo_t:.0f} °C on HFO", "Low for heavy fuel - the viscosity at the injectors may be too high.",
              ["Check the heater and the viscosity controller set point."], ["Heater or viscosity-controller fault"], None, "fuel", uses=["fo_t_in"])
    for k, lab in (("gen_brg_de", "engine side"), ("gen_brg_nde", "opposite side")):
        b = v.get(k)
        if b is not None and b >= th["brg_max"]:
            c.add("warning", f"Generator bearing ({lab}) {b:.0f} °C", f"Limit {th['brg_max']:.0f} °C.",
                  ["Check bearing lubrication, alignment and cooling; inspect the bearing."], ["Lubrication", "Misalignment", "Cooling"], None, "generator", uses=[k])
    sm, raw = _smoke_class(c.p["texts"]["act"].get("smoke")), c.p["texts"]["act"].get("smoke")
    if sm == "black":
        c.add("warning", f"Exhaust smoke: {raw}", "Dark smoke indicates poor combustion or overload.",
              ["Check the air supply (air cooler, turbocharger), the injectors and the fuel quality."], ["Too little air", "Injector fault", "Poor fuel"], None, "smoke")
    elif sm == "blue":
        c.add("advisory", f"Exhaust smoke: {raw}", "Blue smoke points to lube oil in the combustion chamber.",
              ["Check the piston rings / liners and watch the lube-oil consumption."], ["Worn rings or liner", "Valve guide leakage"], None, "smoke")
    elif sm in ("white", "brown"):
        c.add("advisory", f"Exhaust smoke: {raw}", "A visible smoke colour is noted on the sheet.",
              ["Review at the next test: white = unburned fuel / water / cold running; brown = slightly rich or part-load combustion."], (), None, "smoke")


REF_ADVICE = {   # key -> (causes if above reference, causes if below, action if above, action if below)
    "exh_mean": (["Dirty air cooler or turbocharger", "Injection timing", "Exhaust valve seating", "High ambient / cooling-water temperature"], ["Thermocouple fault", "Load reading too high"],
                 "Check the air cooler and turbocharger for fouling, the injection timing and the exhaust valve seating.", "Check the thermocouples and the load reading."),
    "pmax_mean": (["Advanced timing", "High ignition quality fuel"], ["Retarded timing", "Injector or compression leakage"],
                  "Check the fuel pump timing / VIT setting.", "Check the fuel pump timing and the compression (rings, liner, valves)."),
    "rack_mean": (["Worn fuel pumps / injectors", "Low calorific-value fuel", "Shop trials are usually run on MDO"], ["Load reading too high", "Rack calibration"],
                  "Check the fuel pumps and injectors for wear and review the fuel quality.", "Check the load reading, the fuel quality and the rack calibration."),
    "boost_p": (["Gauge / unit error", "High exhaust back-pressure", "Turbocharger running fast (restricted flow)"], ["Dirty air cooler or air filter", "Fouled turbocharger", "Gauge error"],
                "Check the pressure gauge / transmitter calibration and units, the exhaust back-pressure and the turbocharger condition.",
                "Check the air filter, the air cooler and the turbocharger for fouling."),
}


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
        extra = (f" Indicative only: {r['note']}." if r["method"] == "extrapolated" else "") + \
                (" Part of the shop-trial block on this sheet is inconsistent, so treat this as indicative." if trust_low else "")
        if r["implied"] is not None and abs(r["implied"] - c.L) >= 0.08:
            extra += f" On the shop-trial curve this reading corresponds to about {_pct(r['implied'])} load."
        up_c, dn_c, up_a, dn_a = REF_ADVICE[r["key"]]
        c.add(sev, f"{r['label']} {dtxt} vs shop trial", f"{r['act']:.2f} {r['unit']} against {r['ref']:.2f} {r['unit']} expected at {_pct(c.L)} load ({r['method']}).{extra}",
              [up_a if up else dn_a, SAME_TEST], up_c if up else dn_c, None, "reference", uses=[r["key"], "shop", "load"], when="soon" if sev == "warning" else "retest")


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

    def add(sev, title, fault, causes, actions, uses):
        c.add(sev, title, fault, actions, causes, None, "diagnosis", uses=uses)

    if exh_hi and (boost_lo or air_hot):
        add("warning", "Probable diagnosis · too little charge air (air cooler / turbocharger side)",
            "Exhaust temperatures are high " + ("while boost pressure is low" if boost_lo else "and the air leaving the cooler is hot") + " - the engine is short of cool air.",
            ["Fouled air cooler (air or water side)", "Fouled turbocharger (compressor / turbine) or dirty air filter", "Warm or insufficient cooling water to the air cooler"],
            ["Clean the air cooler and check the pressure drop across it.", "Water-wash the turbocharger as per the maker's instructions and check the air filter.",
             "Check the temperature and flow of the cooling water entering the air cooler."], ["exh", "ba_cyl_in", "boost_p", "shop"])
    if exh_hi and rack_hi:
        add("warning", "Probable diagnosis · the engine burns more fuel than at the shop trial",
            "High exhaust temperature together with more fuel rack than the shop trial needs at this load.",
            ["Worn fuel pump plungers / injectors", "Poor-quality or low-calorific fuel", "Fouling or friction losses"],
            ["Test the injectors and check the fuel pumps for wear.", "Review the fuel quality (calorific value, density, CCAI) and the fuel temperature."], ["exh", "rack", "shop"])
    if pmax_lo and (exh_hi or rack_hi):
        add("warning", "Probable diagnosis · late or poor combustion",
            "Peak pressure is below the shop-trial value while exhaust temperature or rack is high: the fuel burns late.",
            ["Retarded injection timing", "Worn injectors or fuel pumps", "Poor ignition quality of the fuel"],
            ["Check the injection timing / VIT setting.", "Test the injectors; review the fuel quality."], ["pmax", "exh", "rack", "shop"])
    for _, r in df.iterrows():
        if _ok(r["exh_dev"]) and r["exh_dev"] >= th["exh_dev"] and _ok(r["pmax_dev_pct"]) and r["pmax_dev_pct"] <= -th["pmax_dev"]:
            i = int(r["id"])
            add("mechanical", f"Probable diagnosis · fault on cylinder {i} (injector, exhaust valve or compression)",
                f"Cylinder {i} runs hot on exhaust ({r['exh_dev']:+.0f} °C vs mean) with low peak pressure ({r['pmax_dev_pct']:+.1f} %): the fuel burns late or is lost.",
                ["Leaking injector or worn nozzle", "Leaking or burnt exhaust valve", "Compression loss (rings, liner)"],
                [f"Pull and test the injector of cylinder {i}.", f"Check the exhaust valve seating and the compression of cylinder {i} at the next opportunity."], ["exh", "pmax"])


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
        head = "Main points: " + "; ".join(top) + "."
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
COMPARE_ACTION = "Compare air cooler, turbocharger, injectors and fuel pumps with the sister unit; repeat both tests at the same load."


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
                        add(vessel, (n1, n2), (hot,), f"{hot} runs {abs(e1 - e2):.0f} °C hotter on exhaust than {cold} at a similar load ({l1 * 100:.0f} % / {l2 * 100:.0f} %).", "warning")
                    b1, b2 = p1["values"]["act"].get("boost_p"), p2["values"]["act"].get("boost_p")
                    if b1 and b2 and abs(b1 - b2) / max(b1, b2) >= 0.15:
                        lo, hi = (n1, n2) if b1 < b2 else (n2, n1)
                        add(vessel, (n1, n2), (lo,), f"{lo} boost pressure is {abs(b1 - b2) / max(b1, b2) * 100:.0f} % lower than {hi} at a similar load - compare air filters, coolers and T/C.")
                    r1, r2 = a1["mean"]["rack"], a2["mean"]["rack"]
                    if r1 and r2 and abs(r1 - r2) / max(r1, r2) >= 0.10:
                        hi, lo = (n1, n2) if r1 > r2 else (n2, n1)
                        add(vessel, (n1, n2), (hi,), f"{hi} needs {abs(r1 - r2) / max(r1, r2) * 100:.0f} % more fuel rack than {lo} at a similar load.")
                else:
                    if e1 is not None and e2 is not None:
                        (lo_n, lo_l, lo_e), (hi_n, hi_l, hi_e) = sorted([(n1, l1, e1), (n2, l2, e2)], key=lambda t: t[1])
                        if lo_e >= hi_e + 25:
                            add(vessel, (n1, n2), (lo_n,), f"{lo_n} runs hotter on exhaust ({lo_e:.0f} °C) than {hi_n} ({hi_e:.0f} °C) although it was tested at a much lower load "
                                                           f"({lo_l * 100:.0f} % vs {hi_l * 100:.0f} %) - abnormally hot for its load: check the air cooler, T/C and injection.", "warning")
                    # identical readings on engines running at different loads: probably copied from one engine to the other
                    v1, v2 = p1["values"]["act"], p2["values"]["act"]
                    same = [k for k in OWN_READINGS if v1.get(k) is not None and v1.get(k) == v2.get(k)]
                    if len(same) >= 4:
                        add(vessel, (n1, n2), (n1, n2), f"{n1} and {n2} report identical values for {len(same)} readings ({', '.join(PARAMS[k][0].lower() for k in same[:4])}"
                                                        f"{' ...' if len(same) > 4 else ''}) although their loads differ ({l1 * 100:.0f} % vs {l2 * 100:.0f} %) - they were probably copied from one engine.",
                            "advisory", "data", "Read each generator's own instruments; never copy readings from one generator to another.")
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
                    add(f"{lab} running hours went down from {h1:,.0f} to {h2:,.0f} h ({span}).", "warning", "data", "Check both running-hour entries against the hour meter and the PMS record.")
                elif h2 - h1 > days * 24 * 1.02 + 1:
                    add(f"{lab} running hours rose by {h2 - h1:,.0f} h in {days} days ({span}) - more than the calendar allows.", "warning", "data",
                        "Check both running-hour entries against the hour meter and the PMS record.")
            l1, l2 = a1["load"], a2["load"]
            if l1 is None or l2 is None or abs(l1 - l2) > load_window:
                continue
            e1, e2 = a1["mean"]["exh"], a2["mean"]["exh"]
            if e1 is not None and e2 is not None and e2 - e1 >= 25:
                add(f"{lab} mean exhaust temperature rose {e2 - e1:.0f} °C at a similar load ({span}).", "warning",
                    action="Look for a fouled air cooler / turbocharger, injector wear or exhaust valve leakage since the previous test.")
            r1, r2 = a1["mean"]["rack"], a2["mean"]["rack"]
            if r1 and r2 and (r2 - r1) / r1 >= 0.08:
                add(f"{lab} mean fuel rack rose {(r2 - r1) / r1 * 100:.0f} % at a similar load ({span}).", "advisory",
                    action="The engine needs more fuel for the same load: check pump / injector wear and fuel quality.")
            b1, b2 = p1["values"]["act"].get("boost_p"), p2["values"]["act"].get("boost_p")
            if b1 and b2 and (b1 - b2) / b1 >= 0.12:
                add(f"{lab} boost pressure fell {(b1 - b2) / b1 * 100:.0f} % at a similar load ({span}).", "advisory",
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
    """One finding: what was seen, likely causes, what to do."""
    if f["cat"] == "data":
        tag = chip("Check entry", "warn" if f["sev"] == "warning" else "info")
    elif f["cat"] == "diagnosis":
        tag = chip("Probable diagnosis", SEV_CHIP[f["sev"]])
    else:
        tag = chip(SEV_LABEL[f["sev"]], SEV_CHIP[f["sev"]])
    if f.get("provisional"):
        tag += " " + chip("verify entry first", "warn")
    causes = (f'<div class="ff" style="color:var(--muted);font-size:.84rem"><b style="color:var(--gold2);letter-spacing:.12em;font-size:.72rem;text-transform:uppercase">Likely causes</b> '
              f'{" · ".join(esc(c) for c in f["causes"])}</div>') if f["causes"] else ""
    cells = f'<div class="note" style="margin-bottom:4px">Sheet cells: {esc(f["cells"])}</div>' if f.get("cells") else ""
    title = f["title"].removeprefix("Check entry · ") if f["cat"] == "data" else f["title"]
    fault = "<br>".join(esc(x) for x in f["fault"].split("\n"))
    return (f'<div class="fcard {f["sev"]}"><div class="ft"><span class="fh">{esc(title)}</span>{tag}</div><div class="ff">{fault}</div>{causes}{cells}'
            f'<div class="fa"><b>Do</b><div style="margin-top:3px">{_steps(f["actions"])}</div><div class="note" style="margin-top:4px">{esc(URGENCY[f["when"]][1])}</div></div></div>')


def note_card(n: dict) -> str:
    tag = chip("Check entry", "warn" if n["sev"] == "warning" else "info") if n["cat"] == "data" else chip(SEV_LABEL[n["sev"]], SEV_CHIP[n["sev"]])
    return (f'<div class="fcard {n["sev"]}"><div class="ft"><span class="fh">Compared with other reports</span>{tag}</div><div class="ff">{esc(n["text"])}</div>'
            f'<div class="fa"><b>Do</b><div style="margin-top:3px">{_steps([n["action"]])}</div></div></div>')


def render_verdict(an: dict) -> None:
    v, c = an["verdict"], an["counts"]
    dq_tone = {"reliable": "ok", "some": "warn", "many": "crit"}[v["data_quality"]]
    chips = " ".join(chip(f"{c[s]} {SEV_LABEL[s]}", SEV_CHIP[s]) for s in ("critical", "mechanical", "warning", "advisory") if c.get(s)) or chip("No findings", "ok")
    chips += " " + chip(v["data_text"], dq_tone)
    prov = " · provisional" if v["provisional"] else ""
    md(f'<div class="banner {v["cls"]}"><div class="ic">{v["icon"]}</div><div><div class="bt">{esc(v["label"])}{prov}</div><div class="bd">{esc(v["headline"])}'
       + ('<br>Some entries on the sheet failed consistency checks - re-check them before relying on the figures.' if v["provisional"] else "")
       + f'<div style="margin-top:.55rem">{chips}</div></div></div></div>')


PLAN_TONE = {"recheck": "advisory", "now": "critical", "soon": "warning", "planned": "advisory", "retest": "nominal"}


def render_plan(an: dict) -> None:
    if not an["plan"]:
        md('<div class="fcard nominal"><div class="ft"><span class="fh">No action needed</span></div><div class="ff">Nothing to correct. Repeat the performance test next month at the same load.</div></div>')
        return
    for key, (_, label) in sorted(URGENCY.items(), key=lambda kv: kv[1][0]):
        items = [s for s in an["plan"] if s["when"] == key]
        if not items:
            continue
        rows = "".join(f'<div class="ff">☐ {esc(s["step"])}<span class="note"> - {esc("; ".join(s["because"][:2]))}</span></div>' for s in items)
        md(f'<div class="fcard {PLAN_TONE[key]}"><div class="ft"><span class="fh">{esc(label)}</span>{chip(str(len(items)), "gold")}</div>{rows}</div>')


def render_checks(an: dict, p: dict, notes: list) -> None:
    data = [f for f in an["findings"] if f["cat"] == "data"]
    n_notes = [n for n in notes if n["cat"] == "data"]
    v = an["verdict"]
    md(f'<p class="note">{v["n_checks"]} consistency checks were run on the readings (cooler in / out direction, charge air versus cooling water, turbocharger versus cylinders, '
       f'bearings versus ambient, kW versus amps, shop-trial loads and figures, printed averages). <b>{len(data) + len(n_notes)} flagged.</b></p>')
    if not data and not n_notes:
        md('<div class="fcard nominal"><div class="ft"><span class="fh">Entries look consistent</span></div><div class="ff">No reading failed a consistency check.</div></div>')
    for f in data:
        md(finding_card(f))
    for n in n_notes:
        md(note_card(n))
    if an["missing"]:
        md('<p class="note" style="margin-top:.4rem"><b>Not measured / left blank:</b> ' + esc(" · ".join(an["missing"])) + ". The diagnosis cannot use these.</p>")


def render_findings(an: dict, notes: list) -> None:
    eng = [f for f in an["findings"] if f["cat"] != "data"]
    eng_notes = [n for n in notes if n["cat"] != "data"]
    if not eng and not eng_notes:
        md('<div class="fcard nominal"><div class="ft"><span class="fh">All readings within limits</span></div><div class="ff">No deviations were found against the configured thresholds.</div></div>')
    for n in eng_notes:
        md(note_card(n))
    for f in eng:
        md(finding_card(f))


def shop_panel(an: dict) -> str:
    tone = {"standard": ("standard point", "ok"), "corrected": ("corrected", "warn"), "non_standard": ("not standard", "warn"), "as_printed": ("as typed", "info")}
    rows, notes = [], []
    for g, name in (("shopA", "Block A"), ("shopB", "Block B")):
        d = an["shop"][g]
        if d["printed"] is None:
            rows.append([name, DASH, DASH, chip("no load typed", "info")])
            continue
        lab, tn = tone[d["status"]]
        rows.append([name, f'{d["printed"] * 100:.0f} %', f'{d["used"] * 100:.0f} %', chip(lab, tn)])
        if d["note"]:
            notes.append(f'{name}: {esc(d["note"])}.')
    return (table_html(["Shop-trial block", "Typed on the sheet", "Used for comparison", "Status"], rows)
            + (f'<p class="note" style="margin-top:.4rem">{" ".join(notes)}</p>' if notes else ""))


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
def render_report(rep: dict, th: dict, notes_all: list) -> None:
    rid = rep["id"]
    p, an = rep["parsed"], rep["an"]
    if not p:
        md(f'<div class="banner err"><div class="ic">⛔</div><div><div class="bt">Could not read {esc(rep["name"])}</div><div class="bd">{esc(rep["error"] or "Unknown error")}</div></div></div>')
        return
    render_steps(3 if an else 2)
    render_card(rep)
    render_banner(p)
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
        kpi("Mean Pmax", _num(an["mean"]["pmax_bar"], 1), "bar", f"{_num(an['mean']['pmax'], 2)} {unit} · spread {_num(an['spread']['pmax'], 2)}"),
        kpi("Mean fuel rack", _num(an["mean"]["rack"], 1), "mm", f"spread {_num(an['spread']['rack'], 1)} mm" if an["spread"]["rack"] is not None else ""),
        kpi("Boost pressure", _num(p["values"]["act"].get("boost_p"), 2), "", "kgf/cm²"),
        kpi("Entries to re-check", str(v["n_data"] + sum(1 for n in notes if n["cat"] == "data")), "", f"{v['n_checks']} checks run", {"reliable": "ok", "some": "warn", "many": "crit"}[v["data_quality"]]),
    ]
    section("01", "Verdict", f"{an['n']} cylinders analysed")
    render_verdict(an)
    md('<div class="kpis" style="grid-template-columns:repeat(auto-fit,minmax(128px,1fr))">' + "".join(cards) + "</div>")
    section("02", "Crew action list", "most urgent first - every step once")
    render_plan(an)
    section("03", "Entries to re-check", "the sheet itself is tested before its figures are trusted")
    render_checks(an, p, notes)
    section("04", "Diagnostic findings", "probable diagnosis first, then by severity")
    render_findings(an, notes)
    section("05", "Performance versus shop trial", "reference interpolated between the two load points")
    md(shop_panel(an))
    fr = chart_ref(an["ref"], th)
    if fr is not None:
        plot(fr, f"dg_{rid}_refdev")
        md('<p class="note">* extrapolated slightly beyond the shop-trial load range · grey = no tolerance defined for that parameter.</p>')
    md(ref_table(an))
    md('<p class="note" style="margin-top:.5rem">“Core” parameters drive the findings. “Implied load” is the shop-trial load at which the reading would be normal: a figure far from the tested load '
       'means the engine needs more (or less) than the shop-trial engine - or the kW reading is off. Readings that behave unphysically between the two blocks are marked inconsistent and not used.</p>')
    section("06", "Cylinder charts", "bars turn amber / orange with the deviation from the engine mean")
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
    section("07", "Cylinder summary")
    md(cylinder_table(p, an))
    section("08", "All extracted readings", "every value with its source cell")
    readings_tables(p, an)
    section("09", "Export", "the HTML page is the printable crew action sheet")
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
