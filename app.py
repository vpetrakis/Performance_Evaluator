import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pydantic import BaseModel, Field, ValidationError
import olefile
from io import BytesIO

# --- UI CONFIGURATION ---
st.set_page_config(page_title="M.E. Diagnostic Hub", page_icon="⚙️", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .stApp { background-color: #0b1120; color: #f8fafc; font-family: 'Inter', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: rgba(30, 41, 59, 0.5); border: 1px solid rgba(56, 189, 248, 0.2);
        padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    h1, h2, h3 { color: #38bdf8 !important; font-weight: 600 !important; }
    .cyl-block { background-color: rgba(30, 41, 59, 0.3); padding: 1rem; border-radius: 8px; margin-bottom: 1rem; border: 1px solid rgba(255,255,255,0.05); }
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

# --- NEW: THE TEC-005 CELL PARSER ---
def extract_raw_binary(file_bytes):
    """Reads the raw OLE stream from the legacy .doc file."""
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    if ole.exists('WordDocument'):
        return ole.openstream('WordDocument').read().decode('ascii', errors='ignore')
    return ""

def parse_tec_005_cells(raw_text):
    """
    Uses the \x07 (Bell) character to split the Word document into exact table cells.
    Preserves blank cells to prevent the data from shifting left.
    """
    # Split the raw text exactly at the hidden table cell walls
    cells = [c.strip() for c in raw_text.split('\x07')]
    
    # Initialize with safe, neutral baselines to prevent instant Pydantic crashes
    cylinders = [{"id": i, "p_max": 100.0, "p_comp": 80.0, "exhaust_temp": 350.0} for i in range(1, 7)]
    
    all_candidates = []
    
    # Extract numbers, explicitly turning blank cells into 0.0 to lock the grid in place
    for c in cells:
        if c == '':
            all_candidates.append(0.0) 
        else:
            try:
                all_candidates.append(float(c))
            except ValueError:
                pass # Ignore text labels like "Pi bar" or "Exhaust"
                
    # Filter numbers into their thermal buckets, keeping the 0.0 placeholders
    p_comp_candidates = [n for n in all_candidates if (45 <= n <= 95) or n == 0.0]
    p_max_candidates = [n for n in all_candidates if (96 <= n <= 150) or n == 0.0]
    exh_candidates = [n for n in all_candidates if (280 <= n <= 450) or n == 0.0]

    # Map the isolated numbers to the cylinders. 
    # If a cell was blank (0.0), it leaves the safe baseline in place so the app doesn't crash.
    for i in range(6):
        if i < len(p_max_candidates) and p_max_candidates[i] != 0.0:
            cylinders[i]['p_max'] = p_max_candidates[i]
        if i < len(p_comp_candidates) and p_comp_candidates[i] != 0.0:
            cylinders[i]['p_comp'] = p_comp_candidates[i]
        if i < len(exh_candidates) and exh_candidates[i] != 0.0:
            cylinders[i]['exhaust_temp'] = exh_candidates[i]
            
    return cylinders

# --- ELITE DIAGNOSTICS ENGINE ---
def run_diagnostics(df: pd.DataFrame):
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

# --- UI COMMAND CENTER ---
st.title("M.E. PERFORMANCE COMMAND CENTER")
st.markdown("### Vessel: M/V ALEXIS | Engine: MAN-B&W 5S60MC-C MK8")

uploaded_file = st.file_uploader("Initiate Data Uplink (.doc)", type=["doc"])

if uploaded_file:
    # 1. SESSION STATE MEMORY (Prevents the UI from wiping your inputs)
    if "current_file" not in st.session_state or st.session_state.current_file != uploaded_file.name:
        raw_text = extract_raw_binary(uploaded_file.read())
        st.session_state.guessed_data = parse_tec_005_cells(raw_text)
        st.session_state.current_file = uploaded_file.name

    st.markdown("---")
    st.markdown("### ⚠️ DATA INTEGRITY VERIFICATION")
    st.info("Legacy `.doc` formatting detected. The system has extracted the following parameters using strict cell boundaries. Please verify before executing.")
    
    verified_data = []
    
    # 2. THE UNBREAKABLE VERTICAL UI
    with st.form("verification_form"):
        for row in st.session_state.guessed_data:
            st.markdown(f"<div class='cyl-block'><b>CYLINDER {row['id']}</b></div>", unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns(3)
            p_max_val = c1.number_input("Pmax (bar)", value=float(row['p_max']), key=f"pmax_{row['id']}", min_value=0.0)
            p_comp_val = c2.number_input("Pcomp (bar)", value=float(row['p_comp']), key=f"pcomp_{row['id']}", min_value=0.0)
            exh_val = c3.number_input("Exhaust Temp (°C)", value=float(row['exhaust_temp']), key=f"exh_{row['id']}", min_value=0.0)
            
            verified_data.append({
                "id": row['id'],
                "p_max": p_max_val,
                "p_comp": p_comp_val,
                "exhaust_temp": exh_val
            })
            st.markdown("<br>", unsafe_allow_html=True)

        # 3. EXECUTION GATE
        submitted = st.form_submit_button("🚀 EXECUTE THERMODYNAMIC ANALYSIS", type="primary", use_container_width=True)

    if submitted:
        # 4. GRANULAR ERROR TRACKING
        errors = []
        valid_cylinders = []
        
        for data in verified_data:
            try:
                if data['p_comp'] >= data['p_max']:
                    errors.append(f"Cylinder {data['id']}: Pcomp ({data['p_comp']} bar) cannot be equal to or higher than Pmax ({data['p_max']} bar).")
                    continue
                
                cyl = Cylinder(**data)
                valid_cylinders.append(cyl)
                
            except ValidationError as e:
                for err in e.errors():
                    field = err['loc'][0]
                    errors.append(f"Cylinder {data['id']} [{field}]: Value is outside physical engine limits.")
        
        if errors:
            st.error("🚨 **DATA INTEGRITY FAILURE:** Please correct the following errors in the grid above:")
            for error in errors:
                st.markdown(f"- {error}")
        else:
            final_df = pd.DataFrame([{**cyl.model_dump(), "ratio": cyl.combustion_ratio} for cyl in valid_cylinders])
            
            st.markdown("---")
            
            # --- RENDER DASHBOARD ---
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
                diagnostics = run_diagnostics(final_df)
                
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
