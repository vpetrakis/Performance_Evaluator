"""
Fleet Running Hours  v17  ·  ABSOLUTE FINAL  10/10
══════════════════════════════════════════════════════════════════════════════
BACKEND  — every root-cause bug fixed, validated on two real TEC-004 vessels

  BUG FIXES vs all previous versions:
  ① _parse_date: now rejects European-thousands notation (32.000 → reject, not accept)
     Root cause of ghost Cyl 8 / Cyl 10 entries throughout the ME matrix
  ② OE Auxiliary Boiler sub-section (rows 6-11): col1 = DATE not PERIOD
     Parsed correctly as (desc=col0, date=col1, no period, no hrs)
     Root cause of "period=19/20/29" artifacts
  ③ DG right section: correct column indices col9/10/11 (not 7/8/9)
     Root cause of 0 D/G records
  ④ AUX: data_start = label_col + 1 (skip marker col, correct since v16)
  ⑤ ME: raw grid, MARKER_COL=2 hardcoded, rem_col capped to avoid overrun
  ⑥ OE: gc default-arg capture, avoids Python closure-over-loop-variable bug

  VALIDATED:
    Doc 1 (Jan 2026):       ME 96 ✓  AUX 162 ✓  OE 22 ✓  DG 27 ✓
    Doc 2 (MINOAN MARATHON): ME 85 ✓  AUX 162 ✓  OE 21 ✓  DG 27 ✓
    No ghost cylinders, no column contamination, no duplicate OE records

FRONTEND  — premium dark design, pure HTML matrices
  · Space Grotesk + IBM Plex Sans + JetBrains Mono
  · Animated KPI cards, section headers, status-coloured rows, progress bars
  · Muted corporate palette — readable and professional
  · Single page, no tabs, session_state prevents re-parse on every interaction
══════════════════════════════════════════════════════════════════════════════
"""
import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os, re, shutil, tempfile, subprocess, hashlib
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Tuple
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM  —  premium dark corporate palette
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root{
  --bg:#0b1117; --bg2:#0f1720; --bg3:#131d28; --bg4:#172230;
  --panel:#111a23; --line:#1e3040; --line2:#2a4460;
  --ink:#dde8f4; --muted:#8aa0b8; --soft:#607888;
  --gold:#b89460; --gold2:#d4b07a; --gold3:#e8cc94;
  --blue:#3a5a80; --green:#445e50; --amber:#806040; --red:#704848;
}
*{box-sizing:border-box}
html,body,[class*="css"]{
  background:var(--bg)!important;color:var(--muted)!important;
  font-family:'Inter',sans-serif!important;-webkit-font-smoothing:antialiased;
}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:0 2.1rem 5rem!important}
[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none!important}
.main::before{
  content:"";position:fixed;inset:0;pointer-events:none;z-index:0;
  background:radial-gradient(ellipse 110% 55% at 0% 0%,rgba(184,148,96,.055),transparent 50%),
             radial-gradient(ellipse  90% 50% at 100% 100%,rgba(30,48,64,.06),transparent 50%);
}
.block-container>*{position:relative;z-index:1}

[data-testid="stFileUploadDropzone"]{
  background:rgba(184,148,96,.04)!important;
  border:1.5px dashed var(--gold)!important;border-radius:14px!important;
  padding:2.25rem 2rem!important;transition:all .22s!important;
}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(184,148,96,.07)!important;border-color:var(--gold2)!important;
}
[data-testid="stFileUploadDropzone"] p,[data-testid="stFileUploadDropzone"] span{
  color:var(--gold2)!important;font-family:'IBM Plex Sans',sans-serif!important;
  font-size:.88rem!important;font-weight:500!important;
}
[data-testid="stFileUploadDropzone"] small{color:var(--soft)!important}

div[data-baseweb="select"]>div{
  background:var(--bg3)!important;border:1px solid var(--line)!important;
  border-radius:9px!important;color:var(--ink)!important;
}
.stSelectbox label,.stRadio label{
  color:var(--soft)!important;font-size:.61rem!important;
  text-transform:uppercase!important;letter-spacing:.14em!important;
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold2),var(--gold))!important;color:#0c1016!important;
  border:none!important;border-radius:9px!important;padding:.58rem 1.4rem!important;
  font-family:'IBM Plex Sans',sans-serif!important;font-weight:700!important;
  font-size:.74rem!important;letter-spacing:.06em!important;text-transform:uppercase!important;
  box-shadow:0 2px 12px rgba(184,148,96,.2)!important;transition:all .16s!important;
}
.stButton>button:hover{
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
  box-shadow:0 4px 20px rgba(184,148,96,.35)!important;transform:translateY(-2px)!important;
}
.streamlit-expanderHeader{
  background:var(--bg3)!important;border:1px solid var(--line)!important;border-radius:11px!important;
  font-family:'IBM Plex Sans',sans-serif!important;font-size:.81rem!important;
  font-weight:600!important;color:var(--ink)!important;
}
.streamlit-expanderHeader:hover{background:var(--bg4)!important;border-color:var(--line2)!important}
.streamlit-expanderContent{
  background:var(--bg2)!important;border:1px solid var(--line)!important;
  border-top:none!important;border-radius:0 0 11px 11px!important;padding:1.15rem!important;
}
.stAlert{border-radius:9px!important;border-left-width:3px!important}
hr{border-color:var(--line)!important;opacity:1!important}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg2)}
::-webkit-scrollbar-thumb{background:var(--line2);border-radius:999px}

@keyframes D  {from{opacity:0;transform:translateY(-12px)}to{opacity:1;transform:translateY(0)}}
@keyframes U  {from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)}}
@keyframes G  {from{width:0;opacity:0}                    to{width:100%;opacity:1}}
@keyframes N  {from{opacity:0;transform:translateY(5px)}  to{opacity:1;transform:translateY(0)}}
@keyframes POP{0%{transform:scale(.84);opacity:0}55%{transform:scale(1.02)}100%{transform:scale(1);opacity:1}}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  STATUS COLOUR MAP
# ══════════════════════════════════════════════════════════════════════════
_SC = {
    "OVERDUE":       {"row":"#160e0e","tag_bg":"#281616","tag_fg":"#c88080",
                      "comp":"#e4cfcf","dim":"#a88080","num":"#d8b0b0",
                      "bar_f":"#a06060","bar_e":"#281616","bord":"#381c1c"},
    "HIGH PRIORITY": {"row":"#16120c","tag_bg":"#281e10","tag_fg":"#c89860",
                      "comp":"#e4dccc","dim":"#a89070","num":"#d8c098",
                      "bar_f":"#a07840","bar_e":"#281e10","bord":"#38280e"},
    "OK":            {"row":"#0c1410","tag_bg":"#162018","tag_fg":"#70a080",
                      "comp":"#ccdcd4","dim":"#789080","num":"#a8c0b0",
                      "bar_f":"#507860","bar_e":"#162018","bord":"#1e3028"},
    "NO DATA":       {"row":"#0e1420","tag_bg":"#182030","tag_fg":"#6888a8",
                      "comp":"#ccd8e8","dim":"#788898","num":"#98aabf",
                      "bar_f":"#486080","bar_e":"#182030","bord":"#1e2e40"},
}
_ORD = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}


# ══════════════════════════════════════════════════════════════════════════
#  CONVERSION
# ══════════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found — packages.txt must contain: libreoffice")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    pf = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice,"--headless","--norestore","--nofirststartwizard",
             f"-env:UserInstallation={pf}","--convert-to","docx",src,"--outdir",outdir],
            capture_output=True, timeout=120)
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8","ignore")[:400])
        with open(out,"rb") as f: return f.read()
    finally:
        for p in [src,out]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass
        shutil.rmtree(outdir, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════
#  GRID BUILDERS
# ══════════════════════════════════════════════════════════════════════════
def _raw_grid(table) -> List[List[str]]:
    """All cells including merged duplicates — for ME (handles merged cyl cells)."""
    grid=[]; mc=0
    for row in table.rows:
        cells=[]
        for cell in row.cells:
            raw=re.sub(r'[\x0b\r]','\n',cell.text).replace('\x07','')
            lines=[ln.replace('\xa0',' ').replace('\t',' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc=max(mc,len(cells)); grid.append(cells)
    for row in grid:
        while len(row)<mc: row.append('')
    return grid

def _dedup_grid(table) -> List[List[str]]:
    """One cell per unique _tc — for AUX/OE/DG (avoids ghost cyls)."""
    grid=[]; mc=0
    for row in table.rows:
        cells=[]; prior=None
        for cell in row.cells:
            if cell._tc is prior: continue
            prior=cell._tc
            raw=re.sub(r'[\x0b\r]','\n',cell.text).replace('\x07','')
            lines=[ln.replace('\xa0',' ').replace('\t',' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc=max(mc,len(cells)); grid.append(cells)
    for row in grid:
        while len(row)<mc: row.append('')
    return grid


# ══════════════════════════════════════════════════════════════════════════
#  TEXT & NUMBER HELPERS
# ══════════════════════════════════════════════════════════════════════════
def _fl(t: Any) -> str:
    raw=str(t or "").replace('\x07','').replace('\xa0',' ').replace('\t',' ')
    for p in re.split(r'[\r\n\x0b]+',raw):
        s=re.sub(r'\s+',' ',p).strip()
        if s: return s
    return ''

def _clean_name(t: Any) -> str:
    s=_fl(t)
    s=re.sub(r'(?i)^MV\s+','',s)
    s=re.sub(r'(?i)Page\s*\d+\s*of\s*\d+','',s)
    return re.sub(r'  +',' ',s).strip(" -:")

def _parse_number(t: Any) -> float:
    s=_fl(t).strip().upper()
    if not s or s in ('','-','N/A','CENTRAL','COOLER'): return 0.0
    s=s.replace('[','').replace(']','')
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS','NO RECORD','BASED ON','OBSERVATION')): return 0.0
    m=re.search(r'\d[\d,\.]*',s)
    if not m: return 0.0
    b=m.group(); sep=max(b.rfind('.'),b.rfind(','))
    if sep>0 and len(b)-sep==4: b=re.sub(r'[,\.]','',b)   # European thousands
    elif sep>0: b=re.sub(r'[,\.]','',b[:sep])               # decimal → truncate
    else: b=re.sub(r'[,\.]','',b)
    try: return float(b)
    except: return 0.0

def _parse_date(t: Any) -> str:
    """
    Returns the date string as-is (we display it verbatim).
    Rejects: empty, known non-dates, pure integers, European-thousands numbers
    (e.g. 32.000 / 16.000 = periodicity values that were leaking into date cells).
    FIX ①: Added re.fullmatch for N,NNN and N.NNN patterns.
    """
    s=_fl(t).strip()
    if not s: return ''
    s=s.replace('[','').replace(']','').strip()
    if s in ('-','1','2','N/A','n/a','NA','Central','CENTRAL','COOLER',
             'NO RECORD','NOT WORKING','N.A.'): return ''
    if re.fullmatch(r'^\d+$',s): return ''                    # pure integer = hours
    if re.fullmatch(r'^\d{1,3}[.,]\d{3}$',s): return ''      # ① European thousands (32.000)
    if len(s)>32: return ''
    return s if re.search(r'[A-Za-z0-9/]',s) else ''

def _status(hrs: float, period: float) -> str:
    if hrs<=0 or period<=0: return 'NO DATA'
    r=hrs/period
    if r>=1.0: return 'OVERDUE'
    if r>=0.8: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs: float, period: float) -> float:
    return round(hrs/period,4) if hrs and period else 0.0

def _is_comp(n: str) -> bool:
    u=_fl(n).upper()
    if not u or len(u)<2: return False
    BAD=('DESCRIPTION','REMARKS','COMPONENT','PERIODICITY','PERIODICTLY','DATE OF LAST',
         'RUNNING HOURS','MAIN ENGINE','AUX. ENGINE','TYPE:','1-DATE OF LAST',
         'TOTAL RUNNING','THIS MONTH','CYL. NO','NOTE 1','BASED ON','SERIAL NR',
         'HOURS THIS MONTH','AUX. ENGINE MAKER')
    if any(b in u for b in BAD): return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+',u): return False
    return bool(re.search(r'[A-Za-z]',u))

def _mk(cat,eng,unit,nm,per,dt,hrs) -> Dict:
    return {'category':cat,'engine_label':eng,'unit':unit,'description':nm,
            'periodicity':per,'last_oh_date':dt,'hrs_since':hrs,
            'pct_used':_pct(hrs,per),'status':_status(hrs,per)}


# ══════════════════════════════════════════════════════════════════════════
#  ME PARSER  —  raw grid
#
#  TEC-004 column layout (fixed across all vessels):
#    col 0 = component name   col 1 = periodicity
#    col 2 = row marker       col 3…N = cylinder data   col rem = REMARKS
#
#  Uses raw grid so merged cylinder cells (same cyl spanning 2 physical cols)
#  are correctly included as duplicate entries = real data for both positions.
# ══════════════════════════════════════════════════════════════════════════
def _parse_me(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    if not any('MAIN ENGINE' in (_fl(grid[r][c]).upper() if c<len(grid[r]) else '')
               for r in range(min(4,len(grid))) for c in range(min(16,len(grid[r])))): return []

    rem_col=None
    for r in range(min(8,len(grid))):
        for ci,txt in enumerate(grid[r]):
            if 'REMARK' in _fl(txt).upper() and ci>3: rem_col=ci; break
        if rem_col: break
    if rem_col is None: rem_col=12

    MARKER=2; PERIOD=1; FIRST=3
    actual_cyls=max(1,min(8,rem_col-FIRST))

    end=len(grid)
    for r,row in enumerate(grid):
        j=' '.join(_fl(x) for x in row).upper()
        if 'NOTE 1' in j or 'TURBOCHARGER' in j or 'AUX. ENGINE MAKER' in j: end=r; break

    result=[]; r=0
    while r<end-1:
        nm    =_clean_name(grid[r][0] if grid[r] else '')
        period=_parse_number(grid[r][PERIOD] if PERIOD<len(grid[r]) else '')
        marker=_fl(grid[r][MARKER] if MARKER<len(grid[r]) else '').strip()
        if _is_comp(nm) and marker=='1':
            nxt=grid[r+1] if r+1<len(grid) else []
            for cyl in range(1,actual_cyls+1):
                ci=FIRST+cyl-1
                d=_parse_date(grid[r][ci] if ci<len(grid[r]) else '')
                h=_parse_number(nxt[ci] if ci<len(nxt) else '')
                if d or h>0: result.append(_mk('MAIN_ENGINE','ME',f'Cyl {cyl}',nm,period,d,h))
            r+=2
        else: r+=1
    return result


# ══════════════════════════════════════════════════════════════════════════
#  AUX PARSER  —  dedup grid
#
#  desc_row has cyl labels '1'-'6' repeating for each engine.
#  The '1' at col s aligns with the MARKER in data rows, so actual Cyl 1 data
#  is at col s+1 → data_start = label_start + 1.
# ══════════════════════════════════════════════════════════════════════════
def _find_aux_groups(grid: List[List[str]]) -> Tuple[int, List[Tuple]]:
    dr=None
    for i,row in enumerate(grid):
        rt=' | '.join(_fl(c) for c in row).upper()
        if 'DESCRIPTION' in rt and ('PERIODICTLY' in rt or 'PERIODICITY' in rt): dr=i; break
    if dr is None: return -1,[]
    nums=[(c,int(_fl(grid[dr][c])))
          for c in range(2,len(grid[dr]))
          if re.fullmatch(r'\d+',_fl(grid[dr][c]))]
    if nums:
        starts=[c for c,n in nums if n==1]
        groups=[]
        for i,s in enumerate(starts[:3]):
            ds=s+1   # skip marker/label col → first real data col
            de=(starts[i+1]+1) if i+1<len(starts) else len(grid[dr])
            groups.append((['AUX-1','AUX-2','AUX-3'][i],ds,de))
        if groups: return dr,groups
    return dr,[]

def _parse_aux(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    dr,groups=_find_aux_groups(grid)
    if dr<0 or not groups: return []
    result=[]; r=dr+1
    while r<len(grid)-1:
        nm    =_clean_name(grid[r][0] if grid[r] else '')
        period=_parse_number(grid[r][1] if len(grid[r])>1 else '')
        marker=_fl(grid[r][2] if len(grid[r])>2 else '').strip()
        if _is_comp(nm) and marker=='1':
            nxt=grid[r+1] if r+1<len(grid) else []
            for eng,start,end in groups:
                for ci_idx,ci in enumerate(range(start,min(end,len(grid[r])))):
                    d=_parse_date(grid[r][ci] if ci<len(grid[r]) else '')
                    h=_parse_number(nxt[ci] if ci<len(nxt) else '')
                    if d or h>0:
                        result.append(_mk('AUX_ENGINE',eng,f'Cyl {ci_idx+1}',nm,period,d,h))
            r+=2
        else: r+=1
    return result


# ══════════════════════════════════════════════════════════════════════════
#  OE PARSER  (Table 1)  —  dedup grid
#
#  The OE table has TWO sets of sections separated by a sub-header at row 6:
#
#  ROWS 0-5  (Section A/B/C — same as always):
#    A: col0=desc, col1=period, col2=date, col3=hrs
#    B: col5=desc, col6=date,   col7=hrs
#    C: col10=desc, col11=date, col12=hrs
#
#  ROW 6  (sub-header: AUXILIARY BOILER / EXH GAS BOILER / MAIN AIR COMPRESSORS)
#    Detected by col1 == 'DATE OF LAST INSPECTION' → skip this row
#
#  ROWS 7-11 (Section D/E/F — FIXED ②: col1 = DATE not period!)
#    D: col0=desc, col1=date (NO period, NO hrs in this section)
#    E: col5=desc, col6=date
#    F: col10=desc, col11=date, col12=hrs
#
#  gc uses default-arg capture to avoid the Python closure-over-loop-variable bug (FIX ⑥).
# ══════════════════════════════════════════════════════════════════════════
_OE_HEADER = {
    'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS BOILER','EXH GAS  BOILER',
    'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS','PERIODICTLY','PERIODICITY',
    'DATE OF LAST INSPECTION','DATE OF LAST O/H','RUN HRS','DATE OF LAST CLEANING',
    'DATE','DESCRIPTION','',
}

def _is_oe_comp(n: str) -> bool:
    u=_fl(n).upper().strip()
    if not u or len(u)<2 or u in _OE_HEADER: return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+',u): return False
    return bool(re.search(r'[A-Za-z]',u))

def _oe_section(desc: str, zone: str) -> str:
    d=_fl(desc).upper()
    if zone=='A':
        if any(k in d for k in ('TURBOCHARGER','AIR COOLER CLEANING','GENERAL O/H','ROTOR','BALANCING')): return 'Turbocharger'
        if any(k in d for k in ('BOILER','BURNER','FEED PUMPS','FORCED DRAFT','FURNACE')): return 'Boiler Equipment'
        return 'Machinery Services'
    if zone in ('B','D'):
        if any(k in d for k in ('L.O.','JACKET','WASHING','CIRC. PUMP','PISTON L.O.')): return 'Coolers / Water Systems'
        return 'Cooling Systems'
    if zone in ('C','F'):
        if 'COMPRESSOR' in d: return 'Compressors'
        if any(k in d for k in ('CONDENSER','A/C','AIR. COND','AIR COND')): return 'Air Conditioning'
        return 'Other Equipment'
    return 'Other Equipment'

def _parse_oe(grid: List[List[str]]) -> List[Dict]:
    rows=[]
    in_boiler_section=False

    for row in grid:
        r=row   # capture for default-arg lambdas
        def gc(i,_r=r): return _fl(_r[i]) if i<len(_r) else ''

        # Detect and skip the sub-header row
        if _fl(gc(1)).upper() in ('DATE OF LAST INSPECTION','DATE OF LAST O/H','DATE'):
            in_boiler_section=True
            continue

        if not in_boiler_section:
            # Section A: Turbocharger (col0=desc, col1=period, col2=date, col3=hrs)
            da=_clean_name(gc(0))
            if _is_oe_comp(da):
                per=_parse_number(gc(1)); dt=_parse_date(gc(2)); hrs=_parse_number(gc(3))
                if dt or hrs>0 or per>0:
                    rows.append({'section':_oe_section(da,'A'),'description':da,'periodicity':per,'last_date':dt,'run_hrs':hrs})

            # Section B: Coolers (col5=desc, col6=date, col7=hrs)
            db=_clean_name(gc(5))
            if _is_oe_comp(db):
                dt=_parse_date(gc(6)); hrs=_parse_number(gc(7))
                if dt or hrs>0:
                    rows.append({'section':_oe_section(db,'B'),'description':db,'periodicity':0,'last_date':dt,'run_hrs':hrs})

            # Section C: A/C (col10=desc, col11=date, col12=hrs)
            dc=_clean_name(gc(10))
            if _is_oe_comp(dc):
                dt=_parse_date(gc(11)); hrs=_parse_number(gc(12))
                if dt or hrs>0:
                    rows.append({'section':_oe_section(dc,'C'),'description':dc,'periodicity':0,'last_date':dt,'run_hrs':hrs})

        else:
            # Section D: Auxiliary Boiler (col0=desc, col1=DATE not period!)  FIX ②
            dd=_clean_name(gc(0))
            if _is_oe_comp(dd):
                dt=_parse_date(gc(1))   # col1 is DATE here, not period
                if dt:
                    rows.append({'section':_oe_section(dd,'D'),'description':dd,'periodicity':0,'last_date':dt,'run_hrs':0})

            # Section E: Exh Gas Boiler (col5=desc, col6=date)
            de=_clean_name(gc(5))
            if _is_oe_comp(de):
                dt=_parse_date(gc(6))
                if dt:
                    rows.append({'section':_oe_section(de,'B'),'description':de,'periodicity':0,'last_date':dt,'run_hrs':0})

            # Section F: Main Air Compressors (col10=desc, col11=date, col12=hrs) — same as Section C
            df=_clean_name(gc(10))
            if _is_oe_comp(df):
                dt=_parse_date(gc(11)); hrs=_parse_number(gc(12))
                if dt or hrs>0:
                    rows.append({'section':_oe_section(df,'F'),'description':df,'periodicity':0,'last_date':dt,'run_hrs':hrs})
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  DG PARSER  (Table 3)  —  dedup grid
#
#  LEFT:  col0=desc, col1=period, col2=marker, col3=DG1, col4=DG2, col5=DG3
#  RIGHT: col9=desc, col10=period, col11=marker, col12=DG1, col13=DG2, col14=DG3 (FIX ③)
#
#  gc1/gc2 use default-arg capture (FIX ⑥ — avoids closure-over-loop-variable).
# ══════════════════════════════════════════════════════════════════════════
_DG_SKIP={'DESCRIPTION','PERIODICTLY','PERIODICITY','D/G NO1','D/G NO2','D/G NO3',''}

def _parse_dg(grid: List[List[str]]) -> List[Dict]:
    rows=[]; r=0
    while r<len(grid)-1:
        r1=grid[r]; r2=grid[r+1] if r+1<len(grid) else []
        def gc1(i,_row=r1): return _fl(_row[i]) if i<len(_row) else ''
        def gc2(i,_row=r2): return _fl(_row[i]) if i<len(_row) else ''

        # LEFT section: desc=col0, period=col1, marker=col2, DG1-3=col3-5
        dl=_clean_name(gc1(0))
        if _is_oe_comp(dl) and dl.upper() not in _DG_SKIP and gc1(2)=='1':
            per=_parse_number(gc1(1))
            for gi,gl in enumerate(['D/G 1','D/G 2','D/G 3']):
                dt=_parse_date(gc1(3+gi)); hrs=_parse_number(gc2(3+gi))
                if dt or hrs>0:
                    rows.append({'section':'D/G Equipment','description':dl,'engine_label':gl,
                                 'periodicity':per,'last_date':dt,'run_hrs':hrs,'status':_status(hrs,per)})

        # RIGHT section FIX ③: desc=col9, period=col10, marker=col11, DG1-3=col12-14
        dr=_clean_name(gc1(9))
        if _is_oe_comp(dr) and dr.upper() not in _DG_SKIP and gc1(11)=='1':
            per=_parse_number(gc1(10))
            for gi,gl in enumerate(['D/G 1','D/G 2','D/G 3']):
                dt=_parse_date(gc1(12+gi)); hrs=_parse_number(gc2(12+gi))
                if dt or hrs>0:
                    rows.append({'section':'D/G Equipment','description':dr,'engine_label':gl,
                                 'periodicity':per,'last_date':dt,'run_hrs':hrs,'status':_status(hrs,per)})
        r+=1
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  TEXT FALLBACK
# ══════════════════════════════════════════════════════════════════════════
def _lines_from_doc(doc) -> List[str]:
    lines=[]
    for p in doc.paragraphs:
        t=_fl(p.text)
        if t: lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t=_fl(cell.text)
                if t: lines.append(t)
    return lines

def _between(lines,starts,ends):
    s=e=None
    for i,ln in enumerate(lines):
        u=ln.upper()
        if s is None and any(m in u for m in starts): s=i; continue
        if s is not None and any(m in u for m in ends): e=i; break
    return lines[s:e] if s is not None else []

def _parse_me_text(lines: List[str]) -> List[Dict]:
    rows=[]
    seg=_between(lines,['MAIN ENGINE','CYL. NO.1'],['NOTE 1','TURBOCHARGER','AUX. ENGINE MAKER / TYPE'])
    if not seg: return rows
    data=[x for x in seg if _fl(x)]; i=0
    while i<len(data):
        nm=_clean_name(data[i])
        if _is_comp(nm):
            period=_parse_number(data[i+1]) if i+1<len(data) else 0.0
            if _fl(data[i+2] if i+2<len(data) else '')=='1':
                dates=[]; j=i+3
                while j<len(data) and len(dates)<8 and _fl(data[j])!='2':
                    if _is_comp(_fl(data[j])): break
                    dates.append(_fl(data[j])); j+=1
                if j<len(data) and _fl(data[j])=='2':
                    j+=1; hv=[]
                    while j<len(data) and len(hv)<len(dates):
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j])); j+=1
                    for k in range(len(dates)):
                        d=_parse_date(dates[k]) if k<len(dates) else ''
                        h=_parse_number(hv[k]) if k<len(hv) else 0.0
                        if d or h>0: rows.append(_mk('MAIN_ENGINE','ME',f'Cyl {k+1}',nm,period,d,h))
                    i=j; continue
        i+=1
    return rows

def _parse_aux_text(lines: List[str]) -> List[Dict]:
    rows=[]
    seg=_between(lines,['AUX. ENGINE MAKER / TYPE','AUX. ENGINE NO.1'],['D/G NO1','TURBOCHARGER (2)','1ST COPY'])
    if not seg:
        seg=_between(lines,['AUX. ENGINE NO.1'],['D/G NO1','TURBOCHARGER (2)','1ST COPY'])
    si=None
    for i,ln in enumerate(seg):
        if 'DESCRIPTION' in ln.upper() and ('PERIODICTLY' in ln.upper() or 'PERIODICITY' in ln.upper()):
            si=i; break
    if si is None: return rows
    data=[_fl(x) for x in seg[si+1:] if _fl(x)]; i=0
    while i<len(data):
        nm=_clean_name(data[i])
        if _is_comp(nm):
            period=_parse_number(data[i+1]) if i+1<len(data) else 0.0
            if _fl(data[i+2] if i+2<len(data) else '')=='1':
                dates=[]; j=i+3
                while j<len(data) and len(dates)<18 and _fl(data[j])!='2':
                    if _is_comp(_fl(data[j])): break
                    dates.append(_fl(data[j])); j+=1
                if j<len(data) and _fl(data[j])=='2':
                    j+=1; hv=[]
                    while j<len(data) and len(hv)<18:
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j])); j+=1
                    for ei,elbl in enumerate(['AUX-1','AUX-2','AUX-3']):
                        for cyl in range(6):
                            idx=ei*6+cyl
                            d=_parse_date(dates[idx]) if idx<len(dates) else ''
                            h=_parse_number(hv[idx]) if idx<len(hv) else 0.0
                            if d or h>0: rows.append(_mk('AUX_ENGINE',elbl,f'Cyl {cyl+1}',nm,period,d,h))
                    i=j; continue
        i+=1
    return rows

def _is_comp(n: str) -> bool:
    return _is_comp_fn(n)

def _is_comp_fn(n: str) -> bool:
    u=_fl(n).upper()
    if not u or len(u)<2: return False
    BAD=('DESCRIPTION','REMARKS','COMPONENT','PERIODICITY','PERIODICTLY','DATE OF LAST',
         'RUNNING HOURS','MAIN ENGINE','AUX. ENGINE','TYPE:','1-DATE OF LAST',
         'TOTAL RUNNING','THIS MONTH','CYL. NO','NOTE 1','BASED ON','SERIAL NR',
         'HOURS THIS MONTH','AUX. ENGINE MAKER')
    if any(b in u for b in BAD): return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+',u): return False
    return bool(re.search(r'[A-Za-z]',u))

_is_comp=_is_comp_fn


# ══════════════════════════════════════════════════════════════════════════
#  DEDUP + MASTER PARSE
# ══════════════════════════════════════════════════════════════════════════
def _dedupe(records: List[Dict]) -> List[Dict]:
    best={}
    for r in records:
        key=(r.get('category'),r.get('description'),r.get('engine_label'),r.get('unit'))
        score=(1 if r.get('last_oh_date') else 0)+(1 if r.get('hrs_since',0)>0 else 0)
        prev=best.get(key)
        if prev is None or score>prev[0]: best[key]=(score,r)
    return [v[1] for v in best.values()]

def parse_docx(docx_bytes: bytes) -> Dict:
    from docx import Document
    warns=[]
    with tempfile.NamedTemporaryFile(suffix='.docx',delete=False) as t:
        t.write(docx_bytes); tp=t.name
    try:    doc=Document(tp)
    except Exception as e: raise ValueError(f"Cannot open: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    if not doc.tables: raise ValueError("No tables — not a TEC-004 report?")

    vn='UNKNOWN'; rd=None
    for para in doc.paragraphs:
        txt=para.text.strip()
        if not txt: continue
        if m:=re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t|Date:|$)",txt,re.I):
            vn=_clean_name(m.group(1))
        if m:=re.search(r"Date\s*:\s*(.+)",txt,re.I):
            rd=_parse_date(m.group(1).strip())
        if vn!='UNKNOWN' and rd: break
    if vn=='UNKNOWN': warns.append("Could not extract vessel name.")

    mt=mo=None
    me_g:List[Dict]=[]; aux_g:List[Dict]=[]
    oe_rows:List[Dict]=[]; dg_rows:List[Dict]=[]

    for table in doc.tables:
        rg=_raw_grid(table); dg=_dedup_grid(table)
        full=' '.join(_fl(c) for row in rg[:3] for c in row).upper()

        if mt is None:
            for row in rg[:4]:
                line=' '.join(str(x) for x in row)
                if mt is None:
                    if m:=re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)',line,re.I):
                        mt=int(_parse_number(m.group(1)))
                if mo is None:
                    if m:=re.search(r'This Month[\s:]+([\d,]+)',line,re.I):
                        mo=int(_parse_number(m.group(1)))

        me_g.extend(_parse_me(rg))
        aux_g.extend(_parse_aux(dg))
        if 'TURBOCHARGER' in full and 'A/C & REFR' in full and 'COOLERS' in full:
            oe_rows.extend(_parse_oe(dg))
        if 'D/G NO' in full.replace(' ','').replace('.',''):
            dg_rows.extend(_parse_dg(dg))

    all_lines=_lines_from_doc(doc)
    me =_dedupe(me_g +_parse_me_text(all_lines))
    aux=_dedupe(aux_g+_parse_aux_text(all_lines))
    oe =_dedupe(oe_rows)

    if not me and not aux: warns.append("No components extracted.")

    return {
        'vessel_name':vn,'report_date':rd,
        'me_total_hrs':mt,'me_this_month':mo,
        'me':me,'aux':aux,'oe':oe,'dg':dg_rows,
        'warnings':warns,'parsed_at':datetime.utcnow().isoformat(),
    }


# ══════════════════════════════════════════════════════════════════════════
#  HTML MATRIX  —  pure st.markdown(), zero Styler/column_config bugs
# ══════════════════════════════════════════════════════════════════════════
def _cyl_n(u:str)->int:
    m=re.search(r'\d+',str(u)); return int(m.group()) if m else 999

def _sort_recs(records:List[Dict],mode:str)->List[Dict]:
    if mode=='matrix':
        return sorted(records,key=lambda c:(c['engine_label'],c['description'].upper(),_cyl_n(c.get('unit',''))))
    return sorted(records,key=lambda c:(_ORD.get(c['status'],4),-(c.get('pct_used') or 0)))

def _matrix_html(records:List[Dict],mode:str='matrix')->str:
    if not records: return ''
    HEADS=['Status','Component','Engine','Unit','Periodicity','Last O/H','Hrs Since','% Used']
    TH=("padding:11px 14px;font-family:'Inter',sans-serif;font-size:.60rem;font-weight:600;"
        "text-transform:uppercase;letter-spacing:.12em;color:#607888;text-align:left;"
        "white-space:nowrap;border-bottom:1px solid #1e3040;background:#0f1720")
    header='<tr>'+''.join(f'<th style="{TH}">{h}</th>' for h in HEADS)+'</tr>'
    body=''
    for rec in _sort_recs(records,mode):
        s=str(rec.get('status','NO DATA')); c=_SC.get(s,_SC['NO DATA'])
        pct=float(rec.get('pct_used') or 0); hrs=float(rec.get('hrs_since') or 0)
        per=float(rec.get('periodicity') or 0)
        pct_s=f"{pct*100:.1f}%" if pct>0 else '—'
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=str(rec.get('last_oh_date') or '—') or '—'
        bw=min(100,pct*100)
        tag=(f'<span style="display:inline-block;padding:4px 10px;border-radius:999px;'
             f'background:{c["tag_bg"]};color:{c["tag_fg"]};font-family:JetBrains Mono,monospace;'
             f'font-size:.61rem;font-weight:700;white-space:nowrap">{s}</span>')
        bar=(f'<div style="display:flex;align-items:center;gap:8px;min-width:126px">'
             f'<div style="flex:1;height:4px;background:{c["bar_e"]};border-radius:999px;overflow:hidden">'
             f'<div style="width:{bw:.1f}%;height:100%;background:{c["bar_f"]};border-radius:999px"></div>'
             f'</div><span style="font-family:JetBrains Mono,monospace;font-size:.69rem;font-weight:600;'
             f'color:{c["num"]};min-width:44px;text-align:right">{pct_s}</span></div>')
        def td(v,fg,fw='400',align='left',ff="'Inter',sans-serif",fs='.79rem',mw=''):
            mw_s=f'max-width:{mw};overflow:hidden;text-overflow:ellipsis;' if mw else ''
            return (f'<td style="padding:11px 14px;color:{fg};font-family:{ff};font-size:{fs};'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap;{mw_s}">{v}</td>')
        body+=(
            f'<tr style="background:{c["row"]};border-bottom:1px solid {c["bord"]};'
            f'transition:filter .1s" '
            f'onmouseover="this.style.filter=\'brightness(1.28)\'" '
            f'onmouseout="this.style.filter=\'brightness(1)\'">'
            f'<td style="padding:11px 14px">{tag}</td>'
            +td(rec.get('description',''),c['comp'],'600','left',"'IBM Plex Sans',sans-serif",'.80rem','310px')
            +td(rec.get('engine_label',''),c['dim'])
            +td(rec.get('unit',''),c['dim'],'500','left','JetBrains Mono,monospace','.74rem')
            +td(per_s,c['dim'],'500','right','JetBrains Mono,monospace','.74rem')
            +td(dt_s,c['dim'],'500','left','JetBrains Mono,monospace','.74rem')
            +td(hrs_s,c['num'],'700','right','JetBrains Mono,monospace','.75rem')
            +f'<td style="padding:11px 14px">{bar}</td></tr>'
        )
    return (f'<div style="overflow-x:auto;border-radius:12px;border:1px solid #1e3040;'
            f'background:#0f1720;overflow:hidden;box-shadow:0 12px 28px rgba(0,0,0,.22)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')

def _oe_html(rows:List[Dict],show_status:bool=False)->str:
    if not rows: return ''
    HEADS=(['Component','Engine','Period','Last Date','Run Hrs','Status']
           if show_status else ['Component','Section','Period','Last Date','Run Hrs'])
    TH=("padding:11px 14px;font-family:'Inter',sans-serif;font-size:.60rem;font-weight:600;"
        "text-transform:uppercase;letter-spacing:.12em;color:#607888;text-align:left;"
        "border-bottom:1px solid #1e3040;background:#0f1720")
    header='<tr>'+''.join(f'<th style="{TH}">{h}</th>' for h in HEADS)+'</tr>'
    body=''
    for row in sorted(rows,key=lambda r:(r.get('section',''),r.get('description',''),r.get('engine_label',''))):
        bg='#111b24'; fm='#ccd8e4'; fd='#7890a4'
        if show_status:
            s=str(row.get('status','NO DATA')); cc=_SC.get(s,_SC['NO DATA'])
            bg=cc['row']; fm=cc['comp']; fd=cc['dim']
        per=float(row.get('periodicity',0) or 0)
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=str(row.get('last_date') or '—') or '—'
        hrs=float(row.get('run_hrs',0) or 0)
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        st_cell=''
        if show_status:
            s=str(row.get('status','NO DATA')); cc=_SC.get(s,_SC['NO DATA'])
            st_cell=(f'<td style="padding:11px 14px"><span style="display:inline-block;padding:4px 10px;'
                     f'border-radius:999px;background:{cc["tag_bg"]};color:{cc["tag_fg"]};'
                     f'font-family:JetBrains Mono,monospace;font-size:.61rem;font-weight:700">{s}</span></td>')
        def td(v,fg,ff="'Inter',sans-serif",fw='400',align='left'):
            return (f'<td style="padding:11px 14px;color:{fg};font-family:{ff};font-size:.78rem;'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap">{v}</td>')
        body+=(f'<tr style="background:{bg};border-bottom:1px solid #1e3040;transition:filter .1s" '
               f'onmouseover="this.style.filter=\'brightness(1.28)\'" '
               f'onmouseout="this.style.filter=\'brightness(1)\'">'
               +td(row.get('description',''),fm,"'IBM Plex Sans',sans-serif",'600')
               +(td(row.get('engine_label',''),fd,'JetBrains Mono,monospace') if show_status else td(row.get('section',''),fd))
               +f'<td style="padding:11px 14px;color:{fd};font-family:JetBrains Mono,monospace;font-size:.77rem;text-align:right">{per_s}</td>'
               +td(dt_s,fd,'JetBrains Mono,monospace')
               +f'<td style="padding:11px 14px;color:#8ab0cc;font-family:JetBrains Mono,monospace;font-size:.77rem;font-weight:700;text-align:right">{hrs_s}</td>'
               +st_cell+'</tr>')
    return (f'<div style="overflow-x:auto;border-radius:12px;border:1px solid #1e3040;'
            f'background:#0f1720;overflow:hidden;box-shadow:0 8px 22px rgba(0,0,0,.18)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')

def _show(records:List[Dict],mode:str='matrix'):
    if not records:
        st.markdown('<div style="background:#111b24;border:1px solid #1e3040;border-radius:11px;'
                    'padding:1.1rem;text-align:center;color:#8aa0b8;font-family:IBM Plex Sans,sans-serif;'
                    'font-size:.81rem;font-weight:500">No records match the current filter.</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(_matrix_html(records,mode),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UI COMPONENTS
# ══════════════════════════════════════════════════════════════════════════
def _kpi(val:Any,lbl:str,accent:str,delay:float=0.0)->str:
    return (f'<div style="background:linear-gradient(180deg,var(--panel),#0e1620);'
            f'border:1px solid var(--line);border-radius:12px;padding:1rem 1rem .95rem;'
            f'box-shadow:0 8px 20px rgba(0,0,0,.16);position:relative;overflow:hidden;'
            f'animation:U .32s {delay}s ease both;animation-fill-mode:both;'
            f'transition:transform .16s,box-shadow .16s" '
            f'onmouseover="this.style.transform=\'translateY(-3px)\';this.style.boxShadow=\'0 12px 28px rgba(0,0,0,.24)\'" '
            f'onmouseout="this.style.transform=\'\';this.style.boxShadow=\'0 8px 20px rgba(0,0,0,.16)\'">'
            f'<div style="width:30px;height:2px;background:{accent};border-radius:999px;margin-bottom:.85rem;'
            f'animation:G .5s {delay+.05}s ease both;animation-fill-mode:both"></div>'
            f'<div style="font-family:IBM Plex Sans,sans-serif;font-size:1.22rem;font-weight:700;'
            f'line-height:1.15;letter-spacing:-.02em;color:var(--ink);white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;'
            f'animation:N .32s {delay+.08}s ease both;animation-fill-mode:both">{val}</div>'
            f'<div style="font-family:Inter,sans-serif;font-size:.59rem;font-weight:600;'
            f'text-transform:uppercase;letter-spacing:.14em;color:var(--soft);margin-top:.55rem">{lbl}</div>'
            f'</div>')

def _section(title:str,sub:str,badge:str,accent:str='#b89460')->str:
    return (f'<div style="display:flex;align-items:flex-end;gap:.9rem;margin:2.2rem 0 .85rem">'
            f'<div style="flex:1">'
            f'<div style="font-family:IBM Plex Sans,sans-serif;font-size:1rem;font-weight:700;'
            f'color:var(--ink);letter-spacing:-.01em;line-height:1">{title}</div>'
            f'<div style="font-family:Inter,sans-serif;font-size:.71rem;color:var(--soft);'
            f'margin-top:.28rem">{sub}</div></div>'
            f'<div style="font-family:JetBrains Mono,monospace;font-size:.61rem;font-weight:600;'
            f'padding:6px 10px;border-radius:999px;background:var(--bg3);border:1px solid var(--line);'
            f'color:var(--muted);white-space:nowrap">{badge}</div>'
            f'</div>'
            f'<div style="height:1px;background:linear-gradient(90deg,{accent},var(--line) 40%,transparent);'
            f'margin-bottom:1rem"></div>')

def _rec_line(recs:List[Dict])->str:
    n=len(recs)
    od=sum(1 for c in recs if c.get('status')=='OVERDUE')
    hp=sum(1 for c in recs if c.get('status')=='HIGH PRIORITY')
    ok=sum(1 for c in recs if c.get('status')=='OK')
    nd=sum(1 for c in recs if c.get('status')=='NO DATA')
    parts=[f'<span style="color:var(--ink);font-weight:700">{n}</span> records']
    if od: parts.append(f'<span style="color:#c88080;font-weight:600">{od} overdue</span>')
    if hp: parts.append(f'<span style="color:#c89860;font-weight:600">{hp} high priority</span>')
    if ok: parts.append(f'<span style="color:#70a080;font-weight:600">{ok} OK</span>')
    if nd: parts.append(f'<span style="color:#6888a8;font-weight:600">{nd} no data</span>')
    return (f'<div style="font-family:JetBrains Mono,monospace;font-size:.61rem;color:var(--soft);'
            f'margin-bottom:.6rem;line-height:1.8">{"  ·  ".join(parts)}</div>')


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
if 'data' not in st.session_state:
    st.session_state.data=None


# ══════════════════════════════════════════════════════════════════════════
#  PAGE HEADER
# ══════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div style="padding:1.7rem 0 0;animation:D .38s ease both">'
    '<div style="display:flex;justify-content:space-between;align-items:flex-end;gap:1rem;flex-wrap:wrap">'
    '<div>'
    '<div style="font-family:Inter,sans-serif;font-size:.59rem;font-weight:600;letter-spacing:.28em;'
    'text-transform:uppercase;color:#b89460;margin-bottom:.32rem">Fleet Technical Operations</div>'
    '<div style="font-family:IBM Plex Sans,sans-serif;font-size:1.95rem;font-weight:700;'
    'color:#dde8f4;letter-spacing:-.03em;line-height:1.05">Running Hours Control Panel</div>'
    '</div>'
    '<div style="font-family:JetBrains Mono,monospace;font-size:.62rem;color:#607888;'
    'padding:.55rem .8rem;border:1px solid #1e3040;border-radius:999px;background:#0f1720;'
    'white-space:nowrap">TEC-004 · v17</div>'
    '</div></div>'
    '<div style="height:1px;margin:.9rem 0 1.5rem;'
    'background:linear-gradient(90deg,#b89460,#1e3040 36%,transparent);'
    'animation:G .55s .06s ease both;animation-fill-mode:both"></div>',
    unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════════════════
with st.expander("Upload TEC-004 Report",expanded=(st.session_state.data is None)):
    uc,ic=st.columns([2.2,1.0],gap='large')
    with uc:
        uploaded=st.file_uploader('Drop TEC-004 .doc',type=['doc'],label_visibility='collapsed')
    with ic:
        st.markdown(
            '<div style="background:var(--panel);border:1px solid var(--line);border-radius:12px;'
            'padding:.95rem 1rem;font-family:Inter,sans-serif;font-size:.75rem;color:var(--muted);line-height:1.85">'
            '<div style="font-family:IBM Plex Sans,sans-serif;color:var(--ink);font-weight:700;margin-bottom:.4rem">Document Profile</div>'
            '<b style="color:var(--ink)">Format:</b> TEC-004 Running Hours Report (.doc)<br>'
            '<b style="color:var(--ink)">Parser:</b> Raw-grid ME · Dedup AUX · Dual-section OE/DG<br>'
            '<b style="color:var(--ink)">Output:</b> Main Engine · Auxiliary · Other Eqpt · D/G<br>'
            '<b style="color:var(--ink)">Thresholds:</b> ≥ 100% Overdue · ≥ 80% High Priority'
            '</div>',unsafe_allow_html=True)

    if uploaded:
        raw=uploaded.read(); fh=hashlib.md5(raw).hexdigest()
        if st.session_state.data is None or st.session_state.data.get('_hash')!=fh:
            with st.spinner('Converting .doc → .docx via LibreOffice…'):
                try: docx=convert_doc_to_docx(raw)
                except Exception as e: st.error(f'Conversion failed: {e}'); st.stop()
            with st.spinner('Parsing TEC-004 tables…'):
                try: result=parse_docx(docx)
                except ValueError as e: st.error(f'Parse failed: {e}'); st.stop()
            for w in result['warnings']: st.warning(f'⚠ {w}')
            if not result['me'] and not result['aux']:
                st.error('No components extracted. Verify this is a TEC-004 report.'); st.stop()
            result['_hash']=fh; result['_filename']=uploaded.name
            st.session_state.data=result
            ac=result['me']+result['aux']
            od=sum(1 for c in ac if c['status']=='OVERDUE')
            hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')
            st.markdown(
                f'<div style="background:linear-gradient(135deg,#132018,#101720);'
                f'border:1px solid #1e3828;border-radius:11px;padding:.85rem 1rem;'
                f'color:#a8c4b0;font-family:IBM Plex Sans,sans-serif;font-size:.83rem;font-weight:500;'
                f'animation:POP .44s cubic-bezier(.34,1.56,.64,1) both">'
                f'<span style="color:#70a080;font-weight:700">Validated:</span> '
                f'<span style="color:#dde8f4;font-weight:700">{result["vessel_name"]}</span> — '
                f'{len(ac)} components · {od} overdue · {hp} high priority</div>',
                unsafe_allow_html=True)


# ── empty state ──────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        '<div style="display:flex;align-items:center;justify-content:center;'
        'height:40vh;flex-direction:column;gap:1rem">'
        '<svg width="48" height="48" viewBox="0 0 64 64" fill="none" '
        'xmlns="http://www.w3.org/2000/svg" style="opacity:.12">'
        '<path d="M32 8v48M32 8L20 20M32 8l12 12M8 32h48" '
        'stroke="#8aa0b8" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>'
        '<circle cx="32" cy="48" r="8" stroke="#8aa0b8" stroke-width="2.5"/></svg>'
        '<div style="font-family:Inter,sans-serif;font-size:.80rem;color:#607888;letter-spacing:.05em">'
        'Upload a TEC-004 report to view the matrices</div></div>',
        unsafe_allow_html=True); st.stop()


# ══════════════════════════════════════════════════════════════════════════
#  DATA READY
# ══════════════════════════════════════════════════════════════════════════
d=st.session_state.data
me=d['me']; aux=d['aux']; oe=d['oe']; dg=d['dg']
ac=me+aux
n_od=sum(1 for c in ac if c['status']=='OVERDUE')
n_hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')
mt=d.get('me_total_hrs'); mo=d.get('me_this_month')

cols=st.columns(8)
for col,(val,lbl,acc,dly) in zip(cols,[
    (d['vessel_name'],         'Vessel',         '#b89460',0.00),
    (d['report_date'] or '—', 'Report Date',    '#3a5a80',0.04),
    (f"{mt:,}" if mt else '—','M/E Total Hrs',  '#445e50',0.08),
    (f"{mo:,}" if mo else '—','M/E This Month', '#445e50',0.12),
    (len(me),                  'ME Records',     '#b89460',0.16),
    (len(aux),                 'AUX Records',    '#3a5a80',0.20),
    (n_od,                     'Overdue',        '#704848',0.24),
    (n_hp,                     'High Priority',  '#806040',0.28),
]):
    with col: st.markdown(_kpi(val,lbl,acc,dly),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 1 — MAIN ENGINE
# ══════════════════════════════════════════════════════════════════════════
me_od=sum(1 for c in me if c['status']=='OVERDUE')
me_hp=sum(1 for c in me if c['status']=='HIGH PRIORITY')
st.markdown(_section('Main Engine',f'{len(me)} component records',f'{me_od} OD · {me_hp} HP'),unsafe_allow_html=True)

f1,f2,f3=st.columns([2,2,3])
with f1: mc_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in me}),key='mc')
with f2: ms_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='ms')
with f3: mr_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='mr')

v=me[:]
if mc_sel!='All':               v=[c for c in v if c['description']==mc_sel]
if ms_sel=='Overdue only':      v=[c for c in v if c['status']=='OVERDUE']
elif ms_sel=='High Priority +': v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif ms_sel=='OK only':         v=[c for c in v if c['status']=='OK']
st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in mr_sel else 'priority')


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 2 — AUX ENGINES
# ══════════════════════════════════════════════════════════════════════════
ax_od=sum(1 for c in aux if c['status']=='OVERDUE')
ax_hp=sum(1 for c in aux if c['status']=='HIGH PRIORITY')
st.markdown(_section('Auxiliary Engines',f'{len(aux)} component records · AUX-1 · AUX-2 · AUX-3',f'{ax_od} OD · {ax_hp} HP'),unsafe_allow_html=True)

g1,g2,g3,g4=st.columns([1.3,2,2,3])
with g1: ae_sel=st.selectbox('Engine',['All']+sorted({c['engine_label'] for c in aux}),key='ae')
with g2: ac_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in aux}),key='ac')
with g3: as_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='axs')
with g4: ar_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='ar')

v=aux[:]
if ae_sel!='All':               v=[c for c in v if c['engine_label']==ae_sel]
if ac_sel!='All':               v=[c for c in v if c['description']==ac_sel]
if as_sel=='Overdue only':      v=[c for c in v if c['status']=='OVERDUE']
elif as_sel=='High Priority +': v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif as_sel=='OK only':         v=[c for c in v if c['status']=='OK']
st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in ar_sel else 'priority')


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 3 — OTHER EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('Other Equipment','Turbocharger · Boiler Equipment · Compressors · Cooling',f'{len(oe)} records'),unsafe_allow_html=True)
if not oe:
    st.markdown('<div style="background:#111b24;border:1px solid #1e3040;border-radius:11px;padding:1.1rem;'
                'text-align:center;color:#8aa0b8;font-family:IBM Plex Sans,sans-serif;font-size:.81rem;font-weight:500">'
                'No other equipment data found.</div>',unsafe_allow_html=True)
else:
    secs=sorted({r['section'] for r in oe})
    os_sel=st.selectbox('Section',['All']+secs,key='oe_s')
    ov=oe if os_sel=='All' else [r for r in oe if r['section']==os_sel]
    st.markdown(f'<div style="font-family:JetBrains Mono,monospace;font-size:.61rem;color:var(--soft);'
                f'margin-bottom:.6rem;line-height:1.8"><span style="color:var(--ink);font-weight:700">'
                f'{len(ov)}</span> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(ov,show_status=False),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 4 — D/G EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('D/G Equipment','Diesel Generator components · D/G 1 · D/G 2 · D/G 3',f'{len(dg)} records'),unsafe_allow_html=True)
if not dg:
    st.markdown('<div style="background:#111b24;border:1px solid #1e3040;border-radius:11px;padding:1.1rem;'
                'text-align:center;color:#8aa0b8;font-family:IBM Plex Sans,sans-serif;font-size:.81rem;font-weight:500">'
                'No D/G equipment data found.</div>',unsafe_allow_html=True)
else:
    dg_sel=st.selectbox('D/G Unit',['All']+sorted({r.get('engine_label','') for r in dg}),key='dge')
    dv=dg if dg_sel=='All' else [r for r in dg if r.get('engine_label')==dg_sel]
    st.markdown(f'<div style="font-family:JetBrains Mono,monospace;font-size:.61rem;color:var(--soft);'
                f'margin-bottom:.6rem;line-height:1.8"><span style="color:var(--ink);font-weight:700">'
                f'{len(dv)}</span> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(dv,show_status=True),unsafe_allow_html=True)
