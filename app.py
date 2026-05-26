import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pydantic import BaseModel, Field, ValidationError
import olefile
import re
from io import BytesIO

# --- PREMIUM CSS INJECTION ---
st.set_page_config(page_title="M.E. Diagnostic Hub", page_icon="⚙️", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .stApp { background-color: #0b1120; color: #f8fafc; font-family: 'Inter', sans-serif; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    div[data-testid="metric-container"] {
        background-color: rgba(30, 41, 59, 0.5); border: 1px solid rgba(56, 189, 248, 0.2);
        padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    h1, h2, h3 { color: #38bdf8 !important; font-weight: 600 !important; letter-spacing: -0.5px; }
    </style>
""", unsafe_allow_html=True)

# --- DATA INTEGRITY SHIELD ---
class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0)
    p_comp: float = Field(..., gt=0)
    exhaust_temp: float = Field(..., gt=100, lt=600)
    
    @property
    def combustion_ratio(self) -> float:
        return self.p_max / self.p_comp

# --- HEURISTIC PARSER ---
def heuristic_decoder(raw_text):
    all_numbers = [float(x) for x in re.findall(r'\b\d{2,3}(?:\.\d+)?\b', raw_text)]
    cylinders = []
    p_comp_candidates = [n for n in all_numbers if 45 <= n <= 68]
    p_max_candidates = [n for n in all_numbers if 69 <= n <= 140]
    exh_candidates = [n for n in all_numbers if 280 <= n <= 450]
    
    for i in range(1, 7):
        cylinders.append({
            "id": i,
            "p_max": p_max_candidates[i-1] if len(p_max_candidates) >= i else 0.0,
            "p_comp": p_comp_candidates[i-1] if len(p_comp_candidates) >= i else 0.0,
            "exhaust_temp": exh_candidates[i-1] if len(exh_candidates) >= i else 0.0,
        })
    return cylinders

def extract_from_binary(file_bytes):
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    if ole.exists('WordDocument'):
        word_stream = ole.openstream('WordDocument').read()
        raw_text = word_stream.decode('ascii', errors='ignore')
        return heuristic_decoder(raw_text)
    return [{"id": i, "p_max": 0.0, "p_comp": 0.0, "exhaust_temp": 0.0} for i in range(1, 7)]

# --- ADVANCED THERMODYNAMIC ENGINE ---
def run_elite_diagnostics(df: pd.DataFrame):
    diagnostics = []
    avg_exh = df['exhaust_temp'].mean()
    for _, row in df.iterrows():
        cyl = int(row['id'])
        ratio = row['ratio']
        delta_exh = row['exhaust_temp'] - avg_exh
        
        if ratio < 1.3:
            diagnostics.append({"cyl": cyl, "status": "🔴 CRITICAL", "fault": "Poor Combustion (Ratio < 1.3)", "action": "Check fuel pump timing, worn plunger, or atomizers."})
        elif ratio > 1.55:
            diagnostics.append({"cyl": cyl, "status": "🟡 WARNING", "fault": "Harsh Combustion (Ratio > 1.55)", "action": "Check for early injection timing or fuel quality."})
            
        if ratio >= 1.3 and delta_exh > 15 and row['p_comp'] < df['p_comp'].mean() - 2:
            diagnostics.append({"cyl": cyl, "status": "🔴 CRITICAL", "fault": "Loss of Compression / Blow-by", "action": "Overhaul required. Inspect exhaust valve and piston rings."})
    return diagnostics

def create_pv_chart(df: pd.DataFrame):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_max'], name='Pmax (bar)', mode='lines+markers', line=dict(color='#0ea5e9', width=3), marker=dict(size=10, symbol='diamond')))
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_comp'], name='Pcomp (bar)', mode='lines+markers', line=dict(color='#8b5cf6', width=3), marker=dict(size=10)))
    ideal_pmax_lower = df['p_comp'] * 1.3
    ideal_pmax_upper = df['p_comp'] * 1.5
    fig.add_trace(go.Scatter(
        x=pd.concat([df['id'], df['id'][::-1]]), y=pd.concat([ideal_pmax_upper, ideal_pmax_lower[::-1]]),
        fill='toself', fillcolor='rgba(14, 165, 233, 0.1)', line=dict(color='rgba(255,255,255,0)'), name='Optimal Combustion Zone'
    ))
    fig.update_layout(
        title="Thermodynamic Pressure Profile", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8'), xaxis=dict(title="Cylinder", showgrid=False, tickmode='linear'),
        yaxis=dict(title="Pressure (bar)", showgrid=True, gridcolor='rgba(255,255,255,0.05)'), hovermode="x unified", margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

# --- THE COMMAND CENTER UI ---
st.title("M.E. PERFORMANCE COMMAND CENTER")
st.markdown("### Vessel: M/V ALEXIS | Engine: MAN-B&W 5S60MC-C MK8")

uploaded_file = st.file_uploader("Initiate Data Uplink (.doc / .docx)", type=["doc", "docx"])

if uploaded_file:
    raw_data = extract_from_binary(uploaded_file.read())
    
    st.markdown("---")
    st.markdown("### ⚠️ DATA INTEGRITY VERIFICATION")
    st.info("Please verify the extracted parameters. Correct any misaligned values, then click Execute.")
    
    # --- BULLETPROOF NATIVE GRID (Bypasses Pandas UI Bugs) ---
    col1, col2, col3, col4 = st.columns(4)
    col1.markdown("**Cylinder**")
    col2.markdown("**Pmax (bar)**")
    col3.markdown("**Pcomp (bar)**")
    col4.markdown("**Exhaust Temp (°C)**")
    
    verified_data = []
    
    for i, row in enumerate(raw_data):
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"<div style='padding-top: 10px;'><b>CYL {row['id']}</b></div>", unsafe_allow_html=True)
        p_max_val = c2.number_input("pmax", value=float(row['p_max']), key=f"pmax_{i}", label_visibility="collapsed")
        p_comp_val = c3.number_input("pcomp", value=float(row['p_comp']), key=f"pcomp_{i}", label_visibility="collapsed")
        exh_val = c4.number_input("exh", value=float(row['exhaust_temp']), key=f"exh_{i}", label_visibility="collapsed")
        
        verified_data.append({
            "id": row['id'],
            "p_max": p_max_val,
            "p_comp": p_comp_val,
            "exhaust_temp": exh_val
        })

    # --- EXECUTION GATE ---
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 EXECUTE THERMODYNAMIC ANALYSIS", type="primary", use_container_width=True):
        try:
            valid_cylinders = [Cylinder(**row) for row in verified_data]
            final_df = pd.DataFrame([{**cyl.model_dump(), "ratio": cyl.combustion_ratio} for cyl in valid_cylinders])
            
            st.markdown("---")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Engine Load Index", "87 RPM", "Ballast Condition")
            m2.metric("Mean Pmax", f"{final_df['p_max'].mean():.1f} bar")
            m3.metric("Mean Exhaust", f"{final_df['exhaust_temp'].mean():.0f} °C")
            
            avg_ratio = final_df['ratio'].mean()
            ratio_color = "normal" if 1.3 <= avg_ratio <= 1.5 else "inverse"
            m4.metric("Avg Combustion Ratio", f"{avg_ratio:.2f}", "Ideal: 1.3 - 1.5", delta_color=ratio_color)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_chart, col_alerts = st.columns([1.5, 1])
            
            with col_chart:
                st.plotly_chart(create_pv_chart(final_df), use_container_width=True)
                
            with col_alerts:
                st.markdown("### 🛠️ Root Cause Diagnostics")
                diagnostics = run_elite_diagnostics(final_df)
                
                if not diagnostics:
                    st.success("🟢 **SYSTEM NOMINAL:** All thermodynamic ratios and thermal gradients are optimal.")
                else:
                    for diag in diagnostics:
                        st.markdown(f"""
                        <div style="background-color: rgba(220, 38, 38, 0.1); border-left: 4px solid #ef4444; padding: 1rem; border-radius: 4px; margin-bottom: 1rem;">
                            <h4 style="color: #ef4444; margin: 0 0 0.5rem 0;">{diag['status']} | Cylinder {diag['cyl']}</h4>
                            <strong>Fault:</strong> {diag['fault']}<br>
                            <span style="color: #94a3b8; font-size: 0.9em;"><strong>Prescription:</strong> {diag['action']}</span>
                        </div>
                        """, unsafe_allow_html=True)
                        
        except ValidationError:
            st.error("CRITICAL: Data Integrity Failure. A value entered in the grid violates thermodynamic limits. Please correct the table.")
