"""
M.E. Command Center  -  application entry point.

  Tab 1  Main engine        (TEC-005 .doc performance sheets)   ->  me_engine.py   (unchanged v4 engine)
  Tab 2  Auxiliary engines  (TEC-014 .xlsx generator reports)   ->  aux_engine.py

st.set_page_config must be the very first Streamlit call and may only run once, so it lives here.
Both tabs are rendered on every rerun, which keeps each tab's uploaded files and state alive while the
other tab is in use; every widget / session key of the aux tab is prefixed "dg_" to stay clear of the main engine.
"""
import streamlit as st

st.set_page_config(page_title="M.E. Command Center", page_icon="⚙️", layout="wide", initial_sidebar_state="expanded")

import aux_engine                      # noqa: E402
import me_engine                   # noqa: E402

me_engine.inject_css()
tab_me, tab_aux = st.tabs(["Main engine · TEC-005", "Auxiliary engines · TEC-014"])
with tab_me:
    me_engine.main()
with tab_aux:
    aux_engine.render_aux_tab()
