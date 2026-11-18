# M.E. Command Center

Marine diesel performance intelligence - one app, two sections (tabs):

| Tab | Input | Module |
|---|---|---|
| **Main engine · TEC-005** | legacy Word 97-2003 `.doc` Performance Data Sheets | `me_engine.py` |
| **Auxiliary engines · TEC-014** | monthly Aux. Engine Performance Report `.xlsx` (also `.xlsm` / `.xls`), one workbook per generator | `aux_engine.py` |

Everything runs in memory: nothing is stored, nothing is sent anywhere. Because the data is company information,
run it **locally or on your own server** (see *Run* below) - Streamlit Community Cloud is cloud hosting.

## Structure

```
app.py             entry point: page config + the two tabs
me_engine.py       main-engine app (TEC-005)  - unchanged v4 engine
aux_engine.py      auxiliary-engine app (TEC-014): reader, diagnostics and tab, in three marked sections
test_app.py        pytest suite (synthetic workbooks; real-sheet checks are opt-in)
requirements.txt   dependencies
.streamlit/config.toml   dark theme, 50 MB upload limit, no telemetry
```

The auxiliary tab keeps all its widget / session keys under the `dg_` prefix, so it cannot interfere with the main-engine tab.

## Auxiliary engines - how it works

* **Reads by label, not by cell address.** Crews insert and delete rows, so the same field sits on a different row in different files.
  Fields are located by section title + row label; value columns are anchored on the printed *SHOP TRIALS* header
  (actual = D:E, shop-trial A = F:H, shop-trial B = I:K).
* **Nothing is lost silently.** Every filled cell is mapped to a field, recognised as printed template text, or listed under *not mapped*.
* **Cross-checked.** Exhaust, Pmax and fuel-rack values are reconciled with the *Average* rows printed on the sheet (actual and both shop-trial blocks).
* **Units.** Pmax printed under a `KGR/CM2` heading but entered in MPa is detected and converted; FO inlet pressure typed in MPa is converted to bar.
* **Diagnostics.** Per-cylinder deviation from the engine mean (exhaust, Pmax, rack, injection timing), absolute limits (editable in the sidebar),
  comparison with the shop-trial figures **interpolated to the tested load** (extrapolation is limited and flagged; an inconsistent shop block is reported
  and not used), data-consistency checks, and generator-vs-generator comparison for the same vessel.
* **Exports.** CSV (every reading with its source cell), JSON (everything incl. coverage and unmapped cells), printable HTML.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py          # opens http://localhost:8501
```

For several users, run the same command on a company server / VM and open port 8501 inside your network.
No container is needed. `.gitignore` keeps `*.doc`, `*.xls*` out of git on purpose.

## Tests

```bash
python -m pytest -q test_app.py                                   # synthetic data, no company files needed
ME_SAMPLE_DIR=/path/to/doc/sheets AUX_SAMPLE_DIR=/path/to/dg/xlsx python -m pytest -q test_app.py   # + real-sheet checks
```
