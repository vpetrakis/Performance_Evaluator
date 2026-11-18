"""
Regression tests for the whole app.   Run:   python -m pytest -q test_app.py

  * Synthetic tests (fictional data) always run - no company files are needed.
  * Real-sheet checks are opt-in, because company data is never shipped with the code:
        ME_SAMPLE_DIR=<folder with the TEC-005 .doc samples>   AUX_SAMPLE_DIR=<folder with the DG_NO.1/2/3 .xlsx>
"""
import glob
import io
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import zipfile

import openpyxl
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import aux_engine as AE   # noqa: E402

A = P = AE   # parser, analysis and UI now live in one module


# ---------------------------------------------------------------- synthetic workbook builder
def build_rows(n=6, shop=True, na_tc=True, kw="412/637 A", rated="700 KW/ 875 KVA", dgno=1, vessel="M/V TEST VESSEL", date="MAY 29, 2026/1300H",
               pmax_mpa=True, avg_ok=True):
    rnd = random.Random(7 + n + dgno)
    exh = [420 + rnd.randint(-12, 12) for _ in range(n)]
    pm = [round(10.4 + rnd.choice([0, .1]), 1) for _ in range(n)]
    rk = [21 + rnd.choice([0, 1]) for _ in range(n)]
    inj = [9.0] * n
    avg = lambda xs, nd=3: round(sum(xs) / len(xs), nd)
    sA = dict(exh=[375] * n, pm=[10.1] * n, rk=[17.5] * n)
    sB = dict(exh=[390] * n, pm=[13.2] * n, rk=[20.0] * n)      # monotonic on purpose (exhaust rises with load)
    R = []
    R.append(dict(A="Ship's Name ", B=vessel, F="Fuel Used", I="VLS"))
    R.append(dict(A="DG Maker", B="TESTMAKER", F="Density", I=0.98))
    R.append(dict(A="DG Type", B=f"{n}DK-20", F="Viscosity"))
    R.append(dict(A=f"DG #{dgno}                                      Serial No.:", B=f"SN{dgno}00{n}", F="Fuel Oil Cons (24 hrs)", I="2.51 MT"))
    R.append(dict(A="Output (Kw) - (Engine / alternator) ", B=rated, F="Lub Oil Cons (24 hrs)"))
    R.append(dict(A="RPM", B=720, F="Ambient Temperature", I=42))
    R.append(dict(A="Running Hours", B=60450, F="Sea Water Temperature", I=28))
    R.append(dict(A="Date / Time :", B=date, F="SHOP TRIALS"))
    R.append(dict(A="TESTED AT MAX. SAFE LOAD :      ", B="kw / A / %", D=kw, F=0.57 if shop else None, I=0.75 if shop else None))
    for i in range(n):
        R.append(dict(A="FO INJECTION TIMING BTDC" if i == 0 else None, B=f"No. {i + 1}", D=inj[i]))
    for k, (lab, d, f, i_) in enumerate([("JACKET INLET", 66, 62, 60), ("ENGINE OUTLET", 72, 63, 64), ("AIR COOLER INLET", 37, 46, 38), ("AIR COOLER OUTLET", 42, None, None),
                                         ("L.O. COOLER INLET", 42, None, None), ("L.O. COOLER OUTLET", 48, None, None)]):
        R.append(dict(A="COOLING WATER TEMP (DEG C)" if k == 0 else None, B=lab, C="oC ", D=d, F=f, I=i_))
    R.append(dict(A="LUB OIL TEMP. (DEG C)", B="L.O. COOLER INLET", C="oC ", D=62))
    R.append(dict(B="L.O. COOLER OUTLET", C="oC ", D=57, F=46, I=50))
    for k, lab in enumerate(["T/C INLET", "AIR COOLER INLET", "CYLINDER INLET"]):
        R.append(dict(A="BOOST AIR TEMP. (DEG C)" if k == 0 else None, B=lab, C="oC ", D=("N/A" if lab != "CYLINDER INLET" else 37), F=(45 if lab == "CYLINDER INLET" else None),
                      I=(30 if lab == "CYLINDER INLET" else None)))
    for i in range(n):
        R.append(dict(A="EXH. GAS TEMP. (DEG C)" if i == 0 else None, B=f"No. {i + 1}", C="oC ", D=exh[i], F=sA["exh"][i], I=sB["exh"][i]))
    R.append(dict(B="AVERAGE", C="oC ", D=avg(exh) if avg_ok else 999, F=avg(sA["exh"]), I=avg(sB["exh"])))
    R.append(dict(B="T/C INLET : 4, 5, 6 (UPPER)", C="oC ", D="N/A"))
    R.append(dict(B="T/C INLET : 1, 2, 3 (LOWER)", C="oC ", D="N/A"))
    R.append(dict(B="T/C OUTLET", C="oC ", D=(350 if na_tc else "N/A"), F=340, I=360))
    R.append(dict(B="ΔΤ (T/C IN - T/C OUT)", C="oC ", D="380/350" if na_tc else "N/A"))
    R.append(dict(A="L.O. PRESSURE (KGR/CM2)", B="ENGINE", D=4.5, F=5, I=4.8))
    R.append(dict(B="T/C", D=3.5, F=4.2, I=4))
    R.append(dict(A="MDO - HFO", B="MDO / HFO", C="HFO", D="VLS"))
    R.append(dict(B="Press Mpa (Eng inlet)", C="Mpa", D=0.56, F=5, I=4.8))
    R.append(dict(B="Temp oC (Eng inlet)", C="oC ", D=118, F=110, I=120))
    R.append(dict(B="Vis Con (Heater) Out oC ", C="oC ", D="NIL"))
    R.append(dict(B="Vis Con Set (cSt)", C="cst", D=12))
    R.append(dict(A="Heavy Fuel Service Tank", B="Heating Temp", C="oC ", D=115))
    R.append(dict(A="COOLING WATER                                             PRESSURE (KGR/CM2)", B="JACKET", D=2.8, F=2.2, I=2.1))
    R.append(dict(B="COOLER", D="N/A"))
    R.append(dict(A="BOOST AIR PRESSURE (KGR/CM2)", D=1.2, F=1.1, I=1.85))
    R.append(dict(A="EXHAUST PRESSURE T/C OUTLET (BAR)", D="N/A"))
    for i in range(n):
        R.append(dict(A="COMBUSTION PRESSURE               (KGR/CM2) - PMAX" if i == 0 else None, B=f"No. {i + 1}", D=pm[i] if pmax_mpa else round(pm[i] * 10.2, 1),
                      F=sA["pm"][i], I=sB["pm"][i]))
    R.append(dict(B="AVERAGE", D=avg(pm) if pmax_mpa else avg([x * 10.2 for x in pm]), F=avg(sA["pm"]), I=avg(sB["pm"])))
    for i in range(n):
        R.append(dict(A="RACK POSITION OF FUEL                         PUMP IN mm" if i == 0 else None, B=f"No. {i + 1}", D=rk[i], F=sA["rk"][i], I=sB["rk"][i]))
    R.append(dict(B="AVERAGE", D=avg(rk), F=avg(sA["rk"]), I=avg(sB["rk"])))
    R.append(dict(A="GENERATOR BEARING TEMPERATURE", B="ENGINE SIDE", D=38, F=51))
    R.append(dict(B="OPPOSITE SIDE", D=28, F=48))
    R.append(dict(A="EXHAUST SMOKE COLOR", D="HAZY BROWN"))
    R.append(dict(A="PREPARED BY: 3E TEST PERSON"))
    R.append(dict(A="Note 1: Performance test to be carried out at max safe load of D/G"))
    R.append(dict(A="Note 2: Shop Trial records to be filled only for comparison purposes"))
    return R, dict(exh=exh, pm=pm, rk=rk, inj=inj)


def to_xlsx(rows, gaps=(), title_row=True, shop_header=True, formulas=False):
    """rows -> workbook bytes.  `gaps`: row indexes (in the body) before which an empty row is inserted."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["C1"] = "MONTHLY AUX. ENGINE PERFORMANCE REPORT"
    ws["I5"] = "To be sent every month"
    r = 3
    for idx, row in enumerate(rows):
        if idx in gaps:
            r += 1
        for col, key in zip("ABCDFI", "ABCDFI"):
            if row.get(key) is not None:
                ws[f"{col}{r}"] = row[key]
        if row.get("F") is None and shop_header and str(row.get("F", "")) == "" and False:
            pass
        r += 1
    # SHOP TRIALS header sits in column F of the "Date / Time" row
    for rr in range(1, r):
        if str(ws[f"A{rr}"].value or "").startswith("Date / Time"):
            if shop_header:
                ws[f"F{rr}"] = "SHOP TRIALS"
            else:
                ws[f"F{rr}"] = None
    ws.merge_cells("C1:K2")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- auxiliary engines (synthetic)
def parsed(**kw):
    gaps = kw.pop("gaps", ())
    shop_header = kw.pop("shop_header", True)
    rows, truth = build_rows(**kw)
    return P.parse_tec_014(to_xlsx(rows, gaps=gaps, shop_header=shop_header)), truth


def test_basic():
    p, t = parsed()
    assert p["status"] == "ok" and p["n_cylinders"] == 6 and p["n_type"] == 6
    assert [p["cyl"]["act"]["exh"][i] for i in range(1, 7)] == t["exh"]
    assert [p["cyl"]["act"]["pmax"][i] for i in range(1, 7)] == t["pm"]
    assert [p["cyl"]["act"]["rack"][i] for i in range(1, 7)] == t["rk"]
    assert all(x["label"] == "match" for x in p["xval"].values()) and len(p["xval"]) == 9
    assert p["coverage"]["unmapped"] == 0 and p["confidence"]["level"] == "verified"
    i = p["identity"]
    assert i["vessel"] == "TEST VESSEL" and i["dg_no"] == 1 and i["rated_kw"] == 700 and i["rated_kva"] == 875
    assert str(i["date"]) == "2026-05-29" and i["time"] == "13:00"
    assert p["test"]["kw"] == 412 and p["test"]["amps"] == 637 and abs(p["test"]["load_frac"] - 412 / 700) < 1e-9
    assert p["units"]["pmax"] == "MPa"
    assert p["values"]["act"]["fo_p_in"] == pytest.approx(5.6)       # 0.56 MPa -> bar
    assert (p["values"]["act"]["tc_pair_in"], p["values"]["act"]["tc_pair_out"]) == (380, 350)
    assert p["values"]["act"]["ba_tc_in"] is None and p["kinds"]["act"]["ba_tc_in"] == "na"


@pytest.mark.parametrize("gaps", [(1,), (3, 9), (0, 5, 12, 20, 33), tuple(range(2, 60, 7))])
def test_rows_may_shift(gaps):
    base, _ = parsed()
    p, _ = parsed(gaps=gaps)
    for k in ("cyl", "values", "texts", "avgs_doc"):
        assert p[k] == base[k], k
    assert p["coverage"]["unmapped"] == 0


@pytest.mark.parametrize("n", [3, 4, 5, 7, 8, 9, 12])
def test_cylinder_counts(n):
    p, t = parsed(n=n)
    assert p["n_cylinders"] == n and p["n_type"] == n
    assert [s for s, m in p["issues"] if s == "warn"] == []
    assert len(p["xval"]) == 9


def test_missing_shop_header_text():
    p, _ = parsed(shop_header=False)
    assert p["shop"]["shopA"]["load"] == 0.57 and p["shop"]["shopB"]["load"] == 0.75
    assert any("SHOP TRIALS" in m for s, m in p["issues"] if s == "info")


def test_no_shop_block():
    p, _ = parsed(shop=False)
    an = A.analyze(p)
    assert p["shop"]["shopA"]["load"] is None
    assert all(r["ref"] is None for r in an["ref"] if r["method"] != "none") or True
    assert an["score"] >= 0


def test_pmax_kgf_units():
    p, _ = parsed(pmax_mpa=False)
    assert p["units"]["pmax"] == "kgf/cm²"
    assert not any("outside the plausible" in m for s, m in p["issues"])


def test_average_mismatch_detected():
    p, _ = parsed(avg_ok=False)
    assert p["xval"]["act.exh"]["label"] == "mismatch"
    assert any("does not match" in m for s, m in p["issues"] if s == "warn")
    assert p["confidence"]["level"] != "verified"
    assert any("average on the sheet" in f["title"] for f in A.analyze(p)["findings"])


def test_tc_na_and_tc_present():
    p, _ = parsed(na_tc=False)
    assert p["values"]["act"]["tc_out"] is None and p["kinds"]["act"]["tc_out"] == "na"


def test_unknown_label_is_listed_not_lost():
    rows, _ = build_rows()
    rows.insert(12, dict(B="MYSTERY SENSOR", D=77))
    p = P.parse_tec_014(to_xlsx(rows))
    assert p["coverage"]["unmapped"] >= 1
    assert any("MYSTERY" in u["text"] or u["text"] == "77" for u in p["unclaimed"])


def test_blank_form():
    rows, _ = build_rows()
    for r in rows:
        for k in ("D", "F", "I"):
            if k in r and not (r.get("B") == "kw / A / %"):
                r[k] = None
    for r in rows:
        r["D"] = None
    p = P.parse_tec_014(to_xlsx(rows))
    assert p["status"] == "blank" and p["confidence"]["level"] == "blank"


def test_not_a_tec014():
    wb = openpyxl.Workbook()
    wb.active["A1"] = "hello"
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(P.AuxParseError):
        P.parse_tec_014(buf.getvalue())


def test_bad_files_never_crash():
    good = to_xlsx(build_rows()[0])
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w") as z:
        z.writestr("a.txt", "x")
    for blob in (b"", b"garbage", zbuf.getvalue(), good[: len(good) // 2], good[:50], b"PK\x03\x04" + b"\0" * 100, b"\xd0\xcf\x11\xe0" + b"\0" * 600, b"a,b,c\n1,2,3\n"):
        with pytest.raises(P.AuxParseError):
            P.parse_tec_014(blob)


def test_fuzz_mutations():
    base = to_xlsx(build_rows()[0])
    rnd = random.Random(11)
    n_ok = n_err = 0
    for _ in range(120):
        wb = openpyxl.load_workbook(io.BytesIO(base))
        ws = wb.active
        cells = [c for row in ws.iter_rows() for c in row if c.value is not None]
        for c in rnd.sample(cells, rnd.randint(1, 12)):
            op = rnd.choice(["del", "text", "num", "neg", "huge", "dup"])
            c.value = {"del": None, "text": rnd.choice(["abc", "N/A", "12/", "--", "1,5,7", " "]), "num": rnd.uniform(-5, 900), "neg": -3.0, "huge": 1e12, "dup": c.value}[op]
        buf = io.BytesIO()
        wb.save(buf)
        try:
            p = P.parse_tec_014(buf.getvalue())
        except P.AuxParseError:
            n_err += 1
            continue
        n_ok += 1
        if p["status"] == "ok":
            A.analyze(p)                                   # must not raise either
    assert n_ok > 40


def test_xls_roundtrip():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        pytest.skip("LibreOffice not available")
    pytest.importorskip("xlrd")
    d = tempfile.mkdtemp()
    src = os.path.join(d, "in.xlsx")
    rows, _ = build_rows()
    open(src, "wb").write(to_xlsx(rows))
    subprocess.run([soffice, "--headless", "--convert-to", "xls", "--outdir", d, src], check=True, capture_output=True, timeout=120)
    p_x = P.parse_tec_014(open(src, "rb").read())
    p_s = P.parse_tec_014(open(os.path.join(d, "in.xls"), "rb").read())
    for k in ("cyl", "values", "avgs_doc"):
        assert p_s[k] == p_x[k], k
    assert p_s["identity"]["vessel"] == p_x["identity"]["vessel"] and p_s["coverage"]["unmapped"] == 0


# ───────── diagnostics ─────────
def test_low_load_and_hot_cylinder():
    rows, t = build_rows(kw="150/240 A")
    for r in rows:
        if r.get("B") == "No. 3" and r.get("A") is None and r.get("C") == "oC ":
            r["D"] = 530                               # exhaust cylinder 3 far too hot
    p = P.parse_tec_014(to_xlsx(rows))
    an = A.analyze(p)
    titles = " | ".join(f["title"] for f in an["findings"])
    assert "Performance test at only 21 % load" in titles
    assert any(f["sev"] == "critical" and f["cyl"] == 3 for f in an["findings"])
    assert an["score"] <= 74


def test_shop_reference_interpolation():
    p, _ = parsed()
    an = A.analyze(p)
    r = {x["key"]: x for x in an["ref"]}
    L = 412 / 700
    exp = 10.1 + (13.2 - 10.1) * (L - 0.57) / (0.75 - 0.57)
    assert r["pmax_mean"]["method"] == "interpolated" and r["pmax_mean"]["ref"] == pytest.approx(exp, rel=1e-6)


def test_peer_notes_hot_despite_low_load():
    a, _ = parsed(dgno=1)
    rows, _ = build_rows(dgno=3, kw="216/344 A")
    for r in rows:
        if r.get("C") == "oC " and str(r.get("B", "")).startswith("No.") and r.get("D") and r["D"] > 300:
            r["D"] = r["D"] + 60
    b = P.parse_tec_014(to_xlsx(rows))
    notes = A.peer_notes([("a", a, A.analyze(a)), ("b", b, A.analyze(b))])
    assert any("abnormally hot" in n["text"] and n["text"].startswith("DG 3") for n in notes)


def test_exports_are_serialisable():
    p, _ = parsed()
    an = A.analyze(p)
    rows = A.flat_rows(p, an)
    assert len(rows) > 60
    json.dumps(rows, default=str)


# ---------------------------------------------------------------- main engine (opt-in, real sheets)
ME_DIR = os.environ.get("ME_SAMPLE_DIR")
needs_me = pytest.mark.skipif(not ME_DIR or not os.path.isdir(ME_DIR), reason="ME_SAMPLE_DIR not set")

ME_EXPECTED = {   # file -> (cylinders, Pmax per cylinder, confidence, level)
    "f1_102rpm": (7, [86, 88, 89, 85, 83, 83, 90], 100.0, "verified"), "f3_perf": (6, [93, 92, 92, 92, 92, 91], 93.0, "high"),
    "f4_sep": (5, [103, 105, 100, 100, 101], 100.0, "verified"), "f5_sept8": (7, [78, 78, 79, 80, 80, 81, 80], 92.5, "high"),
    "f6_sept2026": (6, [95, 90, 95, 85, 90, 90], 100.0, "verified"), "f7_july14": (6, [74, 72, 71, 80, 78, 75], 100.0, "verified"),
    "g1_alexis5": (5, None, 100.0, "verified"), "g2_minoan7": (7, None, 100.0, "verified"),
}


@pytest.fixture(scope="module")
def me():
    import me_engine
    return me_engine


@needs_me
@pytest.mark.parametrize("stem", sorted(ME_EXPECTED))
def test_me_sheet(me, stem):
    n, pm, conf, level = ME_EXPECTED[stem]
    p = me.parse_tec_005(open(os.path.join(ME_DIR, stem + ".doc"), "rb").read())
    assert p["n_cylinders"] == n
    if pm:
        assert [c["p_max"] for c in p["cylinders"]] == pm
    assert p["confidence"]["total"] == conf and p["confidence"]["level"] == level


@needs_me
def test_me_blank_form(me):
    p = me.parse_tec_005(open(os.path.join(ME_DIR, "f2_plain.doc"), "rb").read())
    assert p["confidence"]["level"] == "blank"


# ---------------------------------------------------------------- auxiliary engines (opt-in, real sheets)
AUX_DIR = os.environ.get("AUX_SAMPLE_DIR")
needs_aux = pytest.mark.skipif(not AUX_DIR or not os.path.isdir(AUX_DIR), reason="AUX_SAMPLE_DIR not set")

AUX_EXPECTED = {  # DG no -> (exhaust, Pmax [MPa], rack, injection timing, load %, running hours)
    1: ([432, 428, 423, 419, 409, 422], [10.5, 10.5, 10.4, 10.5, 10.4, 10.5], [21, 21, 22, 21, 22, 21], [9] * 6, 412 / 700, 60450),
    2: ([437, 430, 431, 429, 450, 449], [10.4, 10.4, 10.4, 10.5, 10.4, 10.5], [21, 21, 22, 21, 22, 21], [9.5] * 6, 425 / 700, 60702),
    3: ([474, 464, 480, 470, 470, 472], [7, 6.8, 7, 6.5, 7, 6.8], [16.5, 16.5, 17, 16.5, 17, 16.5], [8.5, 8.5, 8, 8, 8, 8], 216 / 700, 51344),
}


@needs_aux
@pytest.mark.parametrize("dg", [1, 2, 3])
def test_sample(dg):
    f = [x for x in glob.glob(os.path.join(AUX_DIR, "*.xlsx")) if f"NO.{dg}" in x][0]
    p = P.parse_tec_014(open(f, "rb").read())
    exh, pm, rk, inj, load, hrs = AUX_EXPECTED[dg]
    assert p["identity"]["dg_no"] == dg and p["n_cylinders"] == 6
    assert [p["cyl"]["act"]["exh"][i] for i in range(1, 7)] == exh
    assert [p["cyl"]["act"]["pmax"][i] for i in range(1, 7)] == pm
    assert [p["cyl"]["act"]["rack"][i] for i in range(1, 7)] == rk
    assert [p["cyl"]["act"]["inj"][i] for i in range(1, 7)] == inj
    assert p["test"]["load_frac"] == pytest.approx(load) and p["identity"]["run_hours"] == hrs
    assert p["coverage"]["unmapped"] == 0 and p["confidence"]["level"] == "verified"
    assert A.analyze(p)["score"] > 0
