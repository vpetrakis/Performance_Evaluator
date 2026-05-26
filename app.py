import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pydantic import BaseModel, Field, ValidationError
import olefile
from io import BytesIO
import sqlite3
from datetime import datetime
from fpdf import FPDF
import base64

# --- 10/10 ENTERPRISE CSS & THEMING ---
st.set_page_config(page_title="M.E. Command Center", page_icon="⚓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    .stApp { background-color: #050b14; color: #e2e8f0; font-family: 'Inter', sans-serif; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    h1, h2, h3 { color: #38bdf8 !important; font-weight: 800 !important; letter-spacing: -1px; text-transform: uppercase; }
    div[data-testid="metric-container"] {
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(56, 189, 248, 0.3); padding: 1.5rem; border-radius: 16px; 
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.5); backdrop-filter: blur(4px);
    }
    /* Hide row indices in dataframe */
    .row_heading.level0 {display:none} .blank {display:none}
    </style>
""", unsafe_allow_html=True)

# --- 1. HISTORICAL DATABASE (FLEET MEMORY) ---
def init_db():
    conn = sqlite3.connect('fleet_memory.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS engine_logs
                 (date TEXT, vessel TEXT, rpm REAL, cylinder INTEGER, 
                  p_max REAL, p_comp REAL, fuel_index REAL)''')
    conn.commit()
    return conn

def save_to_db(df, vessel, rpm):
    conn = sqlite3.connect('fleet_memory.db')
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    for _, row in df.iterrows():
        conn.execute("INSERT INTO engine_logs VALUES (?, ?, ?, ?, ?, ?, ?)",
                     (date_str, vessel, rpm, int(row['id']), row['p_max_iso'], row['p_comp_iso'], row['fuel_index']))
    conn.commit()
    conn.close()

def load_history(vessel, cylinder):
    conn = sqlite3.connect('fleet_memory.db')
    df = pd.read_sql_query(f"SELECT date, p_max, p_comp FROM engine_logs WHERE vessel='{vessel}' AND cylinder={cylinder} ORDER BY date ASC", conn)
    conn.close()
    return df

init_db()

# --- 2. THERMODYNAMIC ENGINE (ISO CORRECTIONS & SHIELD) ---
class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0)
    p_comp: float = Field(..., gt=0)
    fuel_index: float = Field(..., gt=0, lt=150)

def apply_iso_corrections(df, t_scav, p_baro):
    """Normalizes pressures based on ambient conditions (Simplified Marine ISO Heuristic)"""
    T_REF = 25.0  # ISO Standard 25C
    P_REF = 1000.0 # ISO Standard 1000 mbar
    
    # 0.3 bar drop per 1 degree over 25C, and 0.01 bar drop per 1 mbar drop
    t_correction = (t_scav - T_REF) * 0.3
    p_correction = (P_REF - p_baro) * 0.01
    total_corr = t_correction + p_correction
    
    df['p_max_iso'] = df['p_max'] + total_corr
    df['p_comp_iso'] = df['p_comp'] + (total_corr * 0.7) # Compression affected slightly less
    df['ratio'] = df['p_max_iso'] / df['p_comp_iso']
    return df

# --- 3. THE UNBREAKABLE BINARY PARSER ---
def extract_raw_binary(file_bytes):
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    if ole.exists('WordDocument'):
        return ole.openstream('WordDocument').read().decode('ascii', errors='ignore')
    return ""

def parse_tec_005_cells(raw_text):
    cells = [c.strip() for c in raw_text.split('\x07')]
    cylinders = [{"id": i, "p_max": 100.0, "p_comp": 80.0, "fuel_index": 50.0} for i in range(1, 7)]
    all_candidates = [0.0 if c == '' else float(c) for c in cells if c == '' or c.replace('.', '', 1).isdigit()]
    
    p_comp_candidates = [n for n in all_candidates if (45 <= n <= 95) or n == 0.0]
    p_max_candidates = [n for n in all_candidates if (96 <= n <= 150) or n == 0.0]
    fuel_candidates = [n for n in all_candidates if (30 <= n <= 130) or n == 0.0]

    for i in range(6):
        if i < len(p_max_candidates) and p_max_candidates[i] != 0.0: cylinders[i]['p_max'] = p_max_candidates[i]
        if i < len(p_comp_candidates) and p_comp_candidates[i] != 0.0: cylinders[i]['p_comp'] = p_comp_candidates[i]
        if i < len(fuel_candidates) and fuel_candidates[i] != 0.0: cylinders[i]['fuel_index'] = fuel_candidates[i]
            
    return cylinders

# --- 4. SMS PDF REPORT EXPORTER ---
def generate_sms_report(df, vessel, rpm):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(190, 10, "TEC-005 THERMODYNAMIC ANALYSIS REPORT", ln=True, align='C')
    pdf.set_font("Arial", '', 12)
    pdf.cell(190, 10, f"Vessel: {vessel} | Engine Load: {rpm} RPM | Date: {datetime.now().strftime('%Y-%m-%d')}", ln=True, align='C')
    pdf.ln(10)
    
    # Data Table
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(30, 10, "CYL", border=1); pdf.cell(40, 10, "Pmax (ISO) bar", border=1); pdf.cell(40, 10, "Pcomp (ISO) bar", border=1); pdf.cell(40, 10, "Fuel Index", border=1); pdf.cell(40, 10, "Ratio", border=1)
    pdf.ln()
    pdf.set_font("Arial", '', 10)
    for _, row in df.iterrows():
        pdf.cell(30, 10, str(int(row['id'])), border=1); pdf.cell(40, 10, f"{row['p_max_iso']:.1f}", border=1); pdf.cell(40, 10, f"{row['p_comp_iso']:.1f}", border=1); pdf.cell(40, 10, f"{row['fuel_index']:.1f}", border=1); pdf.cell(40, 10, f"{row['ratio']:.2f}", border=1)
        pdf.ln()
        
    return pdf.output(dest="S").encode("latin-1")

# --- UI COMMAND CENTER ---
with st.sidebar:
    st.markdown("### ⚙️ Engine Telemetry")
    vessel = st.selectbox("Vessel ID", ["M/V ALEXIS", "M/V BRISTOL"])
    rpm_input = st.number_input("Load (RPM)", min_value=30, max_value=110, value=87)
    st.markdown("---")
    st.markdown("### 🌍 ISO Ambient Conditions")
    t_scav = st.number_input("Scavenge Air Temp (°C)", value=42.0)
    p_baro = st.number_input("Barometric Press (mbar)", value=1012.0)

st.title("M.E. COMMAND CENTER")
st.markdown(f"### ⚓ Vessel: {vessel} | ISO Precision Diagnostics Enabled")

uploaded_file = st.file_uploader("INITIATE SECURE DATA UPLINK (.doc)", type=["doc"])

if uploaded_file:
    # Memory Flush & Extraction
    needs_update = False
    if "current_file" not in st.session_state or st.session_state.current_file != uploaded_file.name: needs_update = True
    elif "guessed_data" in st.session_state and len(st.session_state.guessed_data) > 0 and "fuel_index" not in st.session_state.guessed_data[0]: needs_update = True 

    if needs_update:
        raw_text = extract_raw_binary(uploaded_file.read())
        st.session_state.guessed_data = parse_tec_005_cells(raw_text)
        st.session_state.current_file = uploaded_file.name
        # Convert dict to DataFrame for the Interactive Grid
        st.session_state.df_edit = pd.DataFrame(st.session_state.guessed_data)

    st.markdown("---")
    st.markdown("### 🛡️ INTERACTIVE DATA MATRIX")
    st.info("The system has locked onto the binary cell borders. Edit the matrix directly below.")
    
    # UPGRADE: Streamlit Data Editor (Excel-like UI)
    edited_df = st.data_editor(
        st.session_state.df_edit, 
        use_container_width=True, hide_index=True,
        column_config={
            "id": st.column_config.NumberColumn("Cylinder", disabled=True),
            "p_max": st.column_config.NumberColumn("Pmax (bar)", min_value=0, max_value=200, format="%.1f"),
            "p_comp": st.column_config.NumberColumn("Pcomp (bar)", min_value=0, max_value=150, format="%.1f"),
            "fuel_index": st.column_config.NumberColumn("Fuel Pump Index", min_value=0, max_value=150, format="%.1f"),
        }
    )

    if st.button("🚀 EXECUTE CORE ANALYSIS", type="primary", use_container_width=True):
        errors = []
        valid_cylinders = []
        
        # Validation Loop
        for _, row in edited_df.iterrows():
            try:
                data = row.to_dict()
                if data['p_comp'] >= data['p_max']:
                    errors.append(f"Cylinder {data['id']}: Pcomp ({data['p_comp']}) >= Pmax ({data['p_max']}).")
                    continue
                valid_cylinders.append(Cylinder(**data))
            except ValidationError:
                errors.append(f"Cylinder {data['id']}: Thermodynamic bounds violated.")
        
        if errors:
            st.error("🚨 **INTEGRITY FAILURE:** Correct the matrix:")
            for e in errors: st.markdown(f"- {e}")
        else:
            # Process clean data
            base_df = pd.DataFrame([cyl.model_dump() for cyl in valid_cylinders])
            final_df = apply_iso_corrections(base_df, t_scav, p_baro)
            save_to_db(final_df, vessel, rpm_input) # Save to fleet memory
            
            st.markdown("---")
            tab1, tab2 = st.tabs(["📊 CURRENT DIAGNOSTICS", "📈 FLEET TRENDS (PREDICTIVE)"])
            
            with tab1:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Engine Load", f"{rpm_input} RPM")
                m2.metric("Mean Pmax (ISO)", f"{final_df['p_max_iso'].mean():.1f} bar", f"{final_df['p_max_iso'].mean() - final_df['p_max'].mean():.1f} ISO Adj")
                m3.metric("Mean Fuel Index", f"{final_df['fuel_index'].mean():.1f}")
                avg_ratio = final_df['ratio'].mean()
                m4.metric("Avg Ratio", f"{avg_ratio:.2f}", delta_color="normal" if 1.3<=avg_ratio<=1.5 else "inverse")
                
                col_chart, col_alerts = st.columns([1.6, 1])
                with col_chart:
                    # Dual Axis Plotly
                    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.08, row_heights=[0.7, 0.3])
                    fig.add_trace(go.Scatter(x=final_df['id'], y=final_df['p_max_iso'], name='Pmax ISO', mode='lines+markers', line=dict(color='#38bdf8', width=3), marker=dict(size=10)), row=1, col=1)
                    fig.add_trace(go.Scatter(x=final_df['id'], y=final_df['p_comp_iso'], name='Pcomp ISO', mode='lines+markers', line=dict(color='#8b5cf6', width=3), marker=dict(size=10)), row=1, col=1)
                    fig.add_trace(go.Bar(x=final_df['id'], y=final_df['fuel_index'], name='Fuel Index', marker=dict(color='#10b981')), row=2, col=1)
                    fig.update_layout(title="ISO-CORRECTED THERMODYNAMICS", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'), showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # PDF Export Button
                    pdf_bytes = generate_sms_report(final_df, vessel, rpm_input)
                    st.download_button(label="📥 DOWNLOAD OFFICIAL SMS REPORT (PDF)", data=pdf_bytes, file_name=f"TEC-005_{vessel}_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)

                with col_alerts:
                    st.markdown("### 🧠 Smart Diagnostics")
                    # UPGRADE: Progressive Disclosure (Expanders)
                    for _, row in final_df.iterrows():
                        cyl = int(row['id'])
                        ratio = row['ratio']
                        if ratio < 1.3:
                            with st.expander(f"🔴 CYL {cyl}: CRITICAL FAULT - Poor Combustion"):
                                st.write("**Mechanic Action:** Check fuel pump timing, worn plunger, or atomizers.")
                        elif ratio > 1.55:
                            with st.expander(f"🟡 CYL {cyl}: WARNING - Harsh Combustion"):
                                st.write("**Mechanic Action:** Check for early injection timing or fuel quality.")
                        elif row['fuel_index'] - final_df['fuel_index'].mean() > 3 and row['p_max_iso'] < final_df['p_max_iso'].mean():
                            with st.expander(f"🟠 CYL {cyl}: MECHANICAL - Worn Pump"):
                                st.write("**Mechanic Action:** Demanding higher fuel index for lower pressure. Overhaul pump.")
                        else:
                            st.success(f"🟢 CYL {cyl}: Nominal")

            with tab2:
                st.markdown("### 📉 6-Month Degradation Tracker")
                selected_cyl = st.selectbox("Select Cylinder to Trend", range(1, 7))
                hist_df = load_history(vessel, selected_cyl)
                if len(hist_df) > 1:
                    trend_fig = go.Figure()
                    trend_fig.add_trace(go.Scatter(x=hist_df['date'], y=hist_df['p_max'], mode='lines+markers', name="Pmax Degradation"))
                    trend_fig.update_layout(title=f"Cylinder {selected_cyl} Historical $P_{{max}}$", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'))
                    st.plotly_chart(trend_fig, use_container_width=True)
                else:
                    st.info("Insufficient historical data. Upload more monthly logs to establish a predictive trendline.")
