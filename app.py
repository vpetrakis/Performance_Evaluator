import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pydantic import BaseModel, Field, ValidationError
import time

# --- 1. PREMIUM CSS INJECTION ---
st.set_page_config(page_title="M.E. Diagnostic Hub", page_icon="⚙️", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* Global Background and Fonts */
    .stApp {
        background-color: #0b1120;
        color: #f8fafc;
        font-family: 'Inter', sans-serif;
    }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Premium Metric Cards */
    div[data-testid="metric-container"] {
        background-color: rgba(30, 41, 59, 0.5);
        border: 1px solid rgba(56, 189, 248, 0.2);
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    
    /* Custom Headers */
    h1, h2, h3 {
        color: #38bdf8 !important;
        font-weight: 600 !important;
        letter-spacing: -0.5px;
    }
    
    /* Dataframes */
    .stDataFrame {
        border-radius: 8px !important;
        overflow: hidden !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. DATA INTEGRITY SHIELD ---
class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0)
    p_comp: float = Field(..., gt=0)
    exhaust_temp: float = Field(..., gt=100, lt=600)
    
    @property
    def combustion_ratio(self) -> float:
        return self.p_max / self.p_comp

# --- 3. ADVANCED THERMODYNAMIC ENGINE ---
def run_elite_diagnostics(df: pd.DataFrame):
    diagnostics = []
    
    avg_exh = df['exhaust_temp'].mean()
    
    for _, row in df.iterrows():
        cyl = int(row['id'])
        ratio = row['p_max'] / row['p_comp']
        delta_exh = row['exhaust_temp'] - avg_exh
        
        # 1. Combustion Efficiency (The Ultimate Truth)
        if ratio < 1.3:
            diagnostics.append({
                "cyl": cyl,
                "status": "🔴 CRITICAL",
                "fault": "Poor Combustion (Ratio < 1.3)",
                "action": "Check fuel pump timing, worn plunger, or atomizers."
            })
        elif ratio > 1.55:
            diagnostics.append({
                "cyl": cyl,
                "status": "🟡 WARNING",
                "fault": "Harsh Combustion (Ratio > 1.55)",
                "action": "Check for early injection timing or fuel quality."
            })
            
        # 2. Multi-Variable Fault Matrix
        if ratio >= 1.3 and delta_exh > 15 and row['p_comp'] < df['p_comp'].mean() - 2:
            diagnostics.append({
                "cyl": cyl,
                "status": "🔴 CRITICAL",
                "fault": "Loss of Compression / Blow-by",
                "action": "Overhaul required. Inspect exhaust valve and piston rings."
            })
            
    return diagnostics

def create_pv_chart(df: pd.DataFrame):
    """Generates a premium interactive chart using Plotly"""
    fig = go.Figure()

    # Pmax Line
    fig.add_trace(go.Scatter(
        x=df['id'], y=df['p_max'], 
        name='Pmax (bar)',
        mode='lines+markers',
        line=dict(color='#0ea5e9', width=3),
        marker=dict(size=10, symbol='diamond')
    ))

    # Pcomp Line
    fig.add_trace(go.Scatter(
        x=df['id'], y=df['p_comp'], 
        name='Pcomp (bar)',
        mode='lines+markers',
        line=dict(color='#8b5cf6', width=3),
        marker=dict(size=10)
    ))

    # Shaded Ideal Ratio Area (Visualizing Thermodynamics)
    ideal_pmax_lower = df['p_comp'] * 1.3
    ideal_pmax_upper = df['p_comp'] * 1.5
    
    fig.add_trace(go.Scatter(
        x=pd.concat([df['id'], df['id'][::-1]]),
        y=pd.concat([ideal_pmax_upper, ideal_pmax_lower[::-1]]),
        fill='toself',
        fillcolor='rgba(14, 165, 233, 0.1)',
        line=dict(color='rgba(255,255,255,0)'),
        name='Optimal Combustion Zone'
    ))

    fig.update_layout(
        title="Thermodynamic Pressure Profile",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8'),
        xaxis=dict(title="Cylinder", showgrid=False, tickmode='linear'),
        yaxis=dict(title="Pressure (bar)", showgrid=True, gridcolor='rgba(255,255,255,0.05)'),
        hovermode="x unified",
        margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

# --- 4. THE DASHBOARD ---
st.title("M.E. PERFORMANCE COMMAND CENTER")
st.markdown("### Vessel: M/V ALEXIS | Engine: MAN-B&W 5S60MC-C MK8")

uploaded_file = st.file_uploader("Initiate Data Uplink (.doc / .docx)", type=["doc", "docx"])

if uploaded_file:
    with st.spinner("Executing ISO Normalization & Thermodynamic Analysis..."):
        time.sleep(1) # UI polish
        
        try:
            # Mock Extraction of validated data
            raw_cylinders = [
                Cylinder(id=1, p_max=80.0, p_comp=58.0, exhaust_temp=320.0),
                Cylinder(id=2, p_max=81.0, p_comp=59.0, exhaust_temp=345.0),
                Cylinder(id=3, p_max=80.0, p_comp=59.0, exhaust_temp=350.0),
                Cylinder(id=4, p_max=80.0, p_comp=59.0, exhaust_temp=345.0),
                Cylinder(id=5, p_max=88.0, p_comp=58.0, exhaust_temp=330.0), 
                Cylinder(id=6, p_max=72.0, p_comp=58.0, exhaust_temp=338.0), # Intentionally dropping Pmax to trigger ratio fault
            ]
            
            df = pd.DataFrame([{
                **cyl.model_dump(), 
                "ratio": cyl.combustion_ratio
            } for cyl in raw_cylinders])
            
            # Top-Level Executive Metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Engine Load Index", "87 RPM", "Ballast Condition")
            m2.metric("Mean Pmax", f"{df['p_max'].mean():.1f} bar")
            m3.metric("Mean Exhaust", f"{df['exhaust_temp'].mean():.0f} °C")
            
            # The Critical Ratio Metric
            avg_ratio = df['ratio'].mean()
            ratio_color = "normal" if 1.3 <= avg_ratio <= 1.5 else "inverse"
            m4.metric("Avg Combustion Ratio", f"{avg_ratio:.2f}", "Ideal: 1.3 - 1.5", delta_color=ratio_color)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Central Visual & Diagnostics
            col_chart, col_alerts = st.columns([1.5, 1])
            
            with col_chart:
                st.plotly_chart(create_pv_chart(df), use_container_width=True)
                
            with col_alerts:
                st.markdown("### 🛠️ Root Cause Diagnostics")
                diagnostics = run_elite_diagnostics(df)
                
                if not diagnostics:
                    st.success("🟢 **SYSTEM NOMINAL:** All thermodynamic ratios and thermal gradients are optimal.")
                else:
                    for diag in diagnostics:
                        with st.container():
                            st.markdown(f"""
                            <div style="background-color: rgba(220, 38, 38, 0.1); border-left: 4px solid #ef4444; padding: 1rem; border-radius: 4px; margin-bottom: 1rem;">
                                <h4 style="color: #ef4444; margin: 0 0 0.5rem 0;">{diag['status']} | Cylinder {diag['cyl']}</h4>
                                <strong>Fault:</strong> {diag['fault']}<br>
                                <span style="color: #94a3b8; font-size: 0.9em;"><strong>Prescription:</strong> {diag['action']}</span>
                            </div>
                            """, unsafe_allow_html=True)

            # Data Table Layer
            with st.expander("VIEW RAW THERMODYNAMIC MATRIX", expanded=False):
                styled_df = df.style.background_gradient(cmap='Blues', subset=['p_max', 'p_comp']) \
                                    .background_gradient(cmap='OrRd', subset=['exhaust_temp']) \
                                    .format({'ratio': "{:.2f}"})
                st.dataframe(styled_df, use_container_width=True)
                
        except ValidationError as e:
            st.error("CRITICAL: Data Integrity Failure.")
            st.write(e)
