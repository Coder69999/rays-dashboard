import streamlit as st
import pandas as pd
import altair as alt

# Load and clean data
@st.cache_data
def load_data():
    data_path = "D-V1.xlsx"
    sheet_name = "Sheet1"
    df = pd.read_excel(data_path, sheet_name=sheet_name)
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Convert percentage columns
    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption', 'Percent Green Consumption']
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '').astype(float) / 100
    
    return df

df = load_data()

# Client selection
client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

# Main display
st.title(f"📊 {selected['Client Name']}")

# Custom CSS for vertical separation
st.markdown("""
<style>
    .section {
        padding: 20px;
        margin-bottom: 30px;
        border-left: 5px solid #4CAF50;
        background-color: #f9f9f9;
        border-radius: 5px;
    }
    .metric-label {
        font-weight: bold;
        margin-bottom: 5px;
    }
    .metric-value {
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# 1. Basic Information Section
with st.container():
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("Basic Information")
    
    st.markdown('<div class="metric-label">Voltage Level</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Voltage Level"]} kV</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-label">Sanctioned Load</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Sanctioned Load (kVA)"]:,.0f} kVA</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-label">Contract Demand</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Contract Demand (kVA)"]:,.0f} kVA</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# 2. Load Information Section
with st.container():
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("Load Information")
    
    st.markdown('<div class="metric-label">Average Load Factor</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Average Load Factor"]*100:.2f}%</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-label">Annual Consumption</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Annual Consumption"]:,.0f} kWh</div>', unsafe_allow_html=True)
    
    # Peak hour calculation
    pm_consumption = selected['6-10 PM Consumption'] if '6-10 PM Consumption' in selected else 0
    am_consumption = selected['6-8 AM Consumption'] if '6-8 AM Consumption' in selected else 0
    peak_pct = (pm_consumption + am_consumption) * 100
    
    st.markdown('<div class="metric-label">Peak Hour Consumption</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{peak_pct:.2f}% (6-10 PM + 6-8 AM)</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# 3. Solar Information Section
with st.container():
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("Solar Information")
    
    st.markdown('<div class="metric-label">Installed Solar</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Installed Solar Capacity (DC)"]} kW</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-label">Annual Setoff</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Annual Setoff"]:,.0f} kWh</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-label">Green %</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-value">{selected["Percent Green Consumption"]*100:.2f}%</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# ROI Analysis Section (same as before)
# ... [rest of your ROI analysis code]
