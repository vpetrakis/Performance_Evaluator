import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pydantic import BaseModel, Field, ValidationError
import olefile
from io import BytesIO

# --- 10/10 ENTERPRISE CSS INJECTION ---
st.set_page_config(page_title="M.E. Command Center", page_icon="⚙️", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    
    .stApp { background-color: #050b14; color: #e2e8f0; font-family: 'Inter', sans-serif; }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    
    /* Neon Typography */
    h1, h2, h3 { color: #38bdf8 !important; font-weight: 800 !important; letter-spacing: -1px; text-transform: uppercase; }
    
    /* Glassmorphism KPI Metrics */
    div[data-testid="metric-container"] {
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(56, 189, 248, 0.3);
        padding: 1.5rem; border-radius: 16px; 
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.5);
        backdrop-filter: blur(4px);
    }
    
    /* Verification Grid Blocks */
    .cyl-block { 
        background: rgba(15, 23, 42, 0.6); 
        padding: 1.2rem; 
        border-radius: 12px; 
        margin-bottom: 1rem; 
        border-left: 4px solid #38bdf8;
        box-shadow: inset 0 0 20px rgba(0,0,0,0.5);
    }
    
    /* Neon Input Boxes */
    .stNumberInput > div > div > input { color: #38bdf8 !important; font-weight: 600; background-color: rgba(0,0,0,0.2) !important; border: 1px solid rgba(56, 189, 248, 0.2) !important; }
    </style>
""", unsafe_allow_html=True)

# --- THE PYDANTIC SHIELD (Swapped to Fuel Index) ---
class Cylinder(BaseModel):
    id: int = Field(..., ge=1, le=12)
    p_max: float = Field(..., gt=0)
    p_comp: float = Field(..., gt=0)
    fuel_index: float = Field(..., gt=0, lt=150) # Physical limits of the fuel rack
    
    @property
    def combustion_ratio(self) -> float:
        return self.p_max / self.p_comp

# --- THE UNBREAKABLE BINARY PARSER ---
def extract_raw_binary(file_bytes):
    ole = olefile.OleFileIO(BytesIO(file_bytes))
    if ole.exists('WordDocument'):
        return ole.openstream('WordDocument').read().decode('ascii', errors='ignore')
    return ""

def parse_tec_005_cells(raw_text):
    # The magical \x07 split that prevents numbers from shifting
    cells = [c.strip() for c in raw_text.split('\x07')]
    
    # Safe baselines to prevent instant crashes
    cylinders = [{"id": i, "p_max": 100.0, "p_comp": 80.0, "fuel_index": 50.0} for i in range(1, 7)]
    all_candidates = []
    
    for c in cells:
        if c == '':
            all_candidates.append(0.0) 
        else:
            try:
                all_candidates.append(float(c))
            except ValueError:
                pass
                
    # Logic Buckets (Adjusted for Fuel Index)
    p_comp_candidates = [n for n in all_candidates if (45 <= n <= 95) or n == 0.0]
    p_max_candidates = [n for n in all_candidates if (96 <= n <= 150) or n == 0.0]
    fuel_candidates = [n for n in all_candidates if (30 <= n <= 130) or n == 0.0] # Fuel Index Range

    for i in range(6):
        if i < len(p_max_candidates) and p_max_candidates[i] != 0.0:
            cylinders[i]['p_max'] = p_max_candidates[i]
        if i < len(p_comp_candidates) and p_comp_candidates[i] != 0.0:
            cylinders[i]['p_comp'] = p_comp_candidates[i]
        if i < len(fuel_candidates) and fuel_candidates[i] != 0.0:
            cylinders[i]['fuel_index'] = fuel_candidates[i]
            
    return cylinders

# --- ELITE DIAGNOSTICS ENGINE ---
def run_diagnostics(df: pd.DataFrame):
    diagnostics = []
    avg_fuel = df['fuel_index'].mean()
    
    for _, row in df.iterrows():
        cyl = int(row['id'])
        ratio = row['ratio']
        delta_fuel = row['fuel_index'] - avg_fuel
        
        if ratio < 1.3:
            diagnostics.append({"cyl": cyl, "status": "🔴 CRITICAL", "fault": "Poor Combustion (Ratio < 1.3)", "action": "Check fuel pump timing, worn plunger, or atomizers."})
        elif ratio > 1.55:
            diagnostics.append({"cyl": cyl, "status": "🟡 WARNING", "fault": "Harsh Combustion (Ratio > 1.55)", "action": "Check for early injection timing or fuel quality."})
            
        # New Diagnostic Logic: High fuel index but low pressure output = Worn Pump
        if delta_fuel > 3 and row['p_max'] < df['p_max'].mean():
            diagnostics.append({"cyl": cyl, "status": "🟠 MECHANICAL", "fault": "Worn Fuel Pump Plunger", "action": "Cylinder is demanding higher index for lower pressure output. Overhaul pump."})
            
    return diagnostics

# --- BROADCAST-QUALITY PLOTLY UI ---
def create_dual_axis_chart(df: pd.DataFrame):
    # Create subplots: 70% height for Pressures, 30% for Fuel Index
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.08, row_heights=[0.7, 0.3])

    # TOP CHART: Pressures
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_max'], name='Pmax (bar)', mode='lines+markers', line=dict(color='#38bdf8', width=4), marker=dict(size=12, symbol='diamond', line=dict(width=2, color='white'))), row=1, col=1)
    fig.add_trace(go.Scatter(x=df['id'], y=df['p_comp'], name='Pcomp (bar)', mode='lines+markers', line=dict(color='#8b5cf6', width=4), marker=dict(size=12, line=dict(width=2, color='white'))), row=1, col=1)

    # Optimal Zone Polygon
    ideal_pmax_lower = df['p_comp'] * 1.3
    ideal_pmax_upper = df['p_comp'] * 1.5
    fig.add_trace(go.Scatter(x=pd.concat([df['id'], df['id'][::-1]]), y=pd.concat([ideal_pmax_upper, ideal_pmax_lower[::-1]]), fill='toself', fillcolor='rgba(56, 189, 248, 0.1)', line=dict(color='rgba(255,255,255,0)'), name='Optimal Zone'), row=1, col=1)

    # BOTTOM CHART: Fuel Index
    fig.add_trace(go.Bar(x=df['id'], y=df['fuel_index'], name='Fuel Index', marker=dict(color='#10b981', opacity=0.8, line=dict(width=1, color='#059669'))), row=2, col=1)

    # High-End Theming
    fig.update_layout(
        title=dict(text="THERMODYNAMIC & FUEL PROFILER", font=dict(size=20, color='#e2e8f0', family="Inter")),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8', family="Inter"), hovermode="x unified",
        margin=dict(l=20, r=20, t=60, b=20), showlegend=False
    )
    
    fig.update_yaxes(title_text="Pressure (bar)", gridcolor='rgba(255,255,255,0.05)', zeroline=False, row=1, col=1)
    fig.update_yaxes(title_text="Index", gridcolor='rgba(255,255,255,0.05)', zeroline=False, row=2, col=1)
    fig.update_xaxes(title_text="Cylinder Number", showgrid=False, tickmode='linear', row=2, col=1)

    return fig

# --- COMMAND CENTER BOOT SEQUENCE ---
st.title("M.E. COMMAND CENTER")
st.markdown("### ⚓ Vessel: M/V ALEXIS | Engine: MAN-B&W 5S60MC-C MK8")

uploaded_file = st.file_uploader("INITIATE SECURE DATA UPLINK (.doc)", type=["doc"])

if uploaded_file:
    # --- SMART MEMORY FLUSH ---
    # Checks if it's a new file OR if the old 'exhaust_temp' is stuck in memory
    needs_update = False
    if "current_file" not in st.session_state or st.session_state.current_file != uploaded_file.name:
        needs_update = True
    elif "guessed_data" in st.session_state and len(st.session_state.guessed_data) > 0 and "fuel_index" not in st.session_state.guessed_data[0]:
        needs_update = True 

    if needs_update:
        raw_text = extract_raw_binary(uploaded_file.read())
        st.session_state.guessed_data = parse_tec_005_cells(raw_text)
        st.session_state.current_file = uploaded_file.name

    st.markdown("---")
    st.markdown("### 🛡️ DATA INTEGRITY VERIFICATION")
    st.info("System has isolated parameters via internal grid markers. Verify blank cells before execution.")
    
    verified_data = []
    
    # 10/10 Verification UI
    with st.form("verification_form"):
        for row in st.session_state.guessed_data:
            st.markdown(f"<div class='cyl-block'><b>CYLINDER {row['id']}</b></div>", unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns(3)
            p_max_val = c1.number_input("Pmax (bar)", value=float(row['p_max']), key=f"pmax_{row['id']}", min_value=0.0)
            p_comp_val = c2.number_input("Pcomp (bar)", value=float(row['p_comp']), key=f"pcomp_{row['id']}", min_value=0.0)
            fuel_val = c3.number_input("Fuel Pump Index", value=float(row['fuel_index']), key=f"fuel_{row['id']}", min_value=0.0)
            
            verified_data.append({
                "id": row['id'],
                "p_max": p_max_val,
                "p_comp": p_comp_val,
                "fuel_index": fuel_val
            })

        submitted = st.form_submit_button("🚀 EXECUTE CORE ANALYSIS", type="primary", use_container_width=True)

    if submitted:
        errors = []
        valid_cylinders = []
        
        for data in verified_data:
            try:
                if data['p_comp'] >= data['p_max']:
                    errors.append(f"Cylinder {data['id']}: Pcomp ({data['p_comp']} bar) cannot be equal to or higher than Pmax.")
                    continue
                cyl = Cylinder(**data)
                valid_cylinders.append(cyl)
            except ValidationError as e:
                for err in e.errors():
                    field = err['loc'][0]
                    errors.append(f"Cylinder {data['id']} [{field}]: Value violates physical engine limits.")
        
        if errors:
            st.error("🚨 **DATA INTEGRITY FAILURE:** Correct the following parameters:")
            for error in errors:
                st.markdown(f"- {error}")
        else:
            final_df = pd.DataFrame([{**cyl.model_dump(), "ratio": cyl.combustion_ratio} for cyl in valid_cylinders])
            
            st.markdown("---")
            
            # 10/10 KPI Dashboard
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Engine RPM", "87 RPM", "Steady State", delta_color="off")
            m2.metric("Mean Pmax", f"{final_df['p_max'].mean():.1f} bar")
            m3.metric("Mean Fuel Index", f"{final_df['fuel_index'].mean():.1f}")
            
            avg_ratio = final_df['ratio'].mean()
            ratio_color = "normal" if 1.3 <= avg_ratio <= 1.5 else "inverse"
            m4.metric("Combustion Ratio", f"{avg_ratio:.2f}", "Target: 1.3 - 1.5", delta_color=ratio_color)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_chart, col_alerts = st.columns([1.6, 1])
            
            with col_chart:
                st.plotly_chart(create_dual_axis_chart(final_df), use_container_width=True)
                
            with col_alerts:
                st.markdown("### 🧠 AI Diagnostics")
                diagnostics = run_diagnostics(final_df)
                
                if not diagnostics:
                    st.success("🟢 **SYSTEM NOMINAL:** Thermodynamics and fuel delivery are optimal.")
                else:
                    for diag in diagnostics:
                        color = "#ef4444" if "CRITICAL" in diag['status'] else "#f59e0b"
                        bg_color = "rgba(220, 38, 38, 0.1)" if "CRITICAL" in diag['status'] else "rgba(245, 158, 11, 0.1)"
                        st.markdown(f"""
                        <div style="background-color: {bg_color}; border-left: 4px solid {color}; padding: 1rem; border-radius: 8px; margin-bottom: 1rem;">
                            <h4 style="color: {color}; margin: 0 0 0.5rem 0;">{diag['status']} | CYL {diag['cyl']}</h4>
                            <strong>Fault:</strong> {diag['fault']}<br>
                            <span style="color: #94a3b8; font-size: 0.85em;"><strong>Action:</strong> {diag['action']}</span>
                        </div>
                        """, unsafe_allow_html=True)
