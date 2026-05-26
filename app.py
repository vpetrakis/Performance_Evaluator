import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pydantic import BaseModel, Field, ValidationError
import olefile
from io import BytesIO
import json
import os
from openai import OpenAI

# --- UI CONFIGURATION ---
st.set_page_config(page_title="M.E. Diagnostic Hub", page_icon="⚙️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0b1120; color: #f8fafc; font-family: 'Inter', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: rgba(30, 41, 59, 0.5); border: 1px solid rgba(56, 189, 248, 0.2);
        padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    h1, h2, h3 { color: #38bdf8 !important; font-weight: 600 !important; }
    </style>
""", unsafe_allow_html=True)

# --- UPGRADE 1: DYNAMIC BASELINE ENGINE ---
class EngineBaselines:
    @staticmethod
    def get_shop_trial_targets(rpm: float):
        """
        Dynamically calculates expected thermodynamic parameters based on engine RPM 
        (Approximated for MAN-B&W 5S60MC-C MK8).
        """
        if rpm >= 100:
            return {"pmax_target": 145.0, "pcomp_target": 105.0}
        elif rpm >= 85:
            return {"pmax_target": 120.0, "pcomp_target": 80.0}
        elif rpm >= 70:
            return {"pmax_target": 95.0, "pcomp_target": 65.0}
        else:
            return {"pmax_target": 75.0, "pcomp_target": 50.0}

class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0)
    p_comp: float = Field(..., gt=0)
    exhaust_temp: float = Field(..., gt=100, lt=600)
    
    @property
    def combustion_ratio(self) -> float:
        return self.p_max / self.p_comp

# --- UPGRADE 2: AI DATA PIPELINE ---
def extract_raw_binary(file_bytes):
    """Pulls the shattered binary text to feed to the AI."""
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    if ole.exists('WordDocument'):
        return ole.openstream('WordDocument').read().decode('ascii', errors='ignore')
    raise ValueError("Invalid Document Structure")

def ai_heuristic_extraction(raw_text: str) -> pd.DataFrame:
    """
    Sends the shattered text to an LLM to logically reconstruct the grid, 
    bypassing the 'invisible shift' bug entirely.
    """
    api_key = os.environ.get("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error("API Key missing. Add OPENAI_API_KEY to your environment secrets.")
        st.stop()
        
    client = OpenAI(api_key=api_key)
    
    prompt = f"""
    You are a marine engineer. Read this corrupted text stream from a MAN-B&W 5S60MC-C engine log.
    Reconstruct the thermodynamic parameters for Cylinders 1 through 6. 
    Account for missing or blank rows (e.g., Pi bar is often left blank).
    Normal ranges: Pcomp (45-70), Pmax (70-150), Exhaust Temp (280-450).
    
    Return ONLY a raw JSON array of objects with keys: id, p_max, p_comp, exhaust_temp.
    
    Raw Text:
    {raw_text[:2000]} 
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"} # Forces perfect JSON output
    )
    
    # Parse the AI's JSON output directly into Pydantic models for strict validation
    try:
        response_content = response.choices[0].message.content
        # Ensure it's a valid dictionary with a list inside, or just a list
        parsed_json = json.loads(response_content)
        data_list = parsed_json if isinstance(parsed_json, list) else list(parsed_json.values())[0]
        
        valid_cylinders = [Cylinder(**row) for row in data_list]
        return pd.DataFrame([{**cyl.model_dump(), "ratio": cyl.combustion_ratio} for cyl in valid_cylinders])
    except Exception as e:
        raise ValueError(f"AI Reconstruction Failed: {e}")

# --- ELITE DIAGNOSTICS ENGINE ---
def run_dynamic_diagnostics(df: pd.DataFrame, rpm: float):
    diagnostics = []
    avg_exh = df['exhaust_temp'].mean()
    targets = EngineBaselines.get_shop_trial_targets(rpm)
    target_ratio = targets['pmax_target'] / targets['pcomp_target']
    
    # Dynamic thresholds based on load
    lower_ratio_limit = target_ratio * 0.90
    upper_ratio_limit = target_ratio * 1.15
    
    for _, row in df.iterrows():
        cyl = int(row['id'])
        ratio = row['ratio']
        delta_exh = row['exhaust_temp'] - avg_exh
        
        if ratio < lower_ratio_limit:
            diagnostics.append({"cyl": cyl, "status": "🔴 CRITICAL", "fault": f"Poor Combustion (Ratio {ratio:.2f} below dynamic target {lower_ratio_limit:.2f})", "action": "Check fuel pump timing, worn plunger, or atomizers."})
        elif ratio > upper_ratio_limit:
            diagnostics.append({"cyl": cyl, "status": "🟡 WARNING", "fault": f"Harsh Combustion (Ratio {ratio:.2f} above dynamic target {upper_ratio_limit:.2f})", "action": "Check for early injection timing or fuel quality."})
            
        if delta_exh > 15 and row['p_comp'] < (targets['pcomp_target'] - 5):
            diagnostics.append({"cyl": cyl, "status": "🔴 CRITICAL", "fault": "Loss of Compression / Blow-by", "action": "Overhaul required. Inspect exhaust valve and piston rings."})
            
    return diagnostics

def create_pv_chart(df: pd.DataFrame, rpm: float):
    targets = EngineBaselines.get_shop_trial_targets(rpm)
    target_ratio = targets['pmax_target'] / targets['pcomp_target']
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_max'], name='Pmax (bar)', mode='lines+markers', line=dict(color='#0ea5e9', width=3), marker=dict(size=10, symbol='diamond')))
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_comp'], name='Pcomp (bar)', mode='lines+markers', line=dict(color='#8b5cf6', width=3), marker=dict(size=10)))

    ideal_pmax_lower = df['p_comp'] * (target_ratio * 0.90)
    ideal_pmax_upper = df['p_comp'] * (target_ratio * 1.15)
    
    fig.add_trace(go.Scatter(
        x=pd.concat([df['id'], df['id'][::-1]]), y=pd.concat([ideal_pmax_upper, ideal_pmax_lower[::-1]]),
        fill='toself', fillcolor='rgba(14, 165, 233, 0.1)', line=dict(color='rgba(255,255,255,0)'), name=f'Dynamic Combustion Zone ({rpm} RPM)'
    ))

    fig.update_layout(
        title="Thermodynamic Pressure Profile", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8'), xaxis=dict(title="Cylinder", showgrid=False, tickmode='linear'),
        yaxis=dict(title="Pressure (bar)", showgrid=True, gridcolor='rgba(255,255,255,0.05)'), hovermode="x unified", margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

# --- UI COMMAND CENTER ---
st.title("M.E. PERFORMANCE COMMAND CENTER")

# Sidebar for Dynamic Variables
with st.sidebar:
    st.markdown("### Engine Parameters")
    vessel = st.selectbox("Vessel", ["M/V ALEXIS", "M/V BRISTOL"])
    rpm_input = st.number_input("Current Load (RPM)", min_value=30.0, max_value=110.0, value=87.0, step=1.0)
    st.markdown("---")
    st.markdown("### Expected Shop Trials")
    targets = EngineBaselines.get_shop_trial_targets(rpm_input)
    st.metric("Target Pmax", f"{targets['pmax_target']} bar")
    st.metric("Target Pcomp", f"{targets['pcomp_target']} bar")

uploaded_file = st.file_uploader("Initiate Data Uplink (.doc / .docx)", type=["doc", "docx"])

if uploaded_file:
    with st.spinner("Executing AI Data Reconstruction..."):
        try:
            # Pipeline: Read Binary -> Extract Gibberish -> AI Reconstruction -> Validated DataFrame
            raw_text = extract_raw_binary(uploaded_file.read())
            final_df = ai_heuristic_extraction(raw_text)
            
            # --- RENDER DASHBOARD ---
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Engine Load Index", f"{rpm_input} RPM")
            m2.metric("Mean Pmax", f"{final_df['p_max'].mean():.1f} bar", f"{final_df['p_max'].mean() - targets['pmax_target']:.1f} vs Target")
            m3.metric("Mean Exhaust", f"{final_df['exhaust_temp'].mean():.0f} °C")
            m4.metric("Avg Combustion Ratio", f"{final_df['ratio'].mean():.2f}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_chart, col_alerts = st.columns([1.5, 1])
            
            with col_chart:
                st.plotly_chart(create_pv_chart(final_df, rpm_input), use_container_width=True)
                
            with col_alerts:
                st.markdown("### 🛠️ Dynamic Diagnostics")
                diagnostics = run_dynamic_diagnostics(final_df, rpm_input)
                
                if not diagnostics:
                    st.success("🟢 **SYSTEM NOMINAL:** Thermodynamics match Shop Trial baselines.")
                else:
                    for diag in diagnostics:
                        st.markdown(f"""
                        <div style="background-color: rgba(220, 38, 38, 0.1); border-left: 4px solid #ef4444; padding: 1rem; border-radius: 4px; margin-bottom: 1rem;">
                            <h4 style="color: #ef4444; margin: 0 0 0.5rem 0;">{diag['status']} | Cylinder {diag['cyl']}</h4>
                            <strong>Fault:</strong> {diag['fault']}<br>
                            <span style="color: #94a3b8; font-size: 0.9em;"><strong>Prescription:</strong> {diag['action']}</span>
                        </div>
                        """, unsafe_allow_html=True)
                        
        except Exception as e:
            st.error(f"CRITICAL: Pipeline Failure. Detail: {str(e)}")
