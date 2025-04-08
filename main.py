import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# Load and clean data
@st.cache_data
def load_data():
    data_path = "D-V1.xlsx"
    sheet_name = "Sheet1"
    df = pd.read_excel(data_path, sheet_name=sheet_name)
    
    # Clean column names - handle spaces and special characters
    df.columns = df.columns.str.strip()
    
    # Convert percentage columns to numeric values
    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption', 'Percent Green Consumption']
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '').astype(float)
    
    return df

df = load_data()

# Sidebar for client selection
client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

# Helper function to safely get percentage values
def get_percentage(value):
    if isinstance(value, str):
        return float(value.replace('%', ''))
    return float(value)

# Main display
st.title(f"📊 Client Overview: {client}")

# First section - Basic info (grouped together)
st.subheader("Basic Information")
col1, col2, col3 = st.columns(3)
with col1:
    st.write(f"**Voltage Level**  \n{selected['Voltage Level']} kV")
with col2:
    st.write(f"**Sanctioned Load**  \n{selected['Sanctioned Load (kVA)']:,.0f} kVA")
with col3:
    st.write(f"**Contract Demand**  \n{selected['Contract Demand (kVA)']:,.0f} kVA")

# Horizontal separator line after Basic Info
st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# Second section - Load and consumption
st.subheader("Load Information")
col1, col2 = st.columns(2)
with col1:
    st.write(f"**Average Load Factor**  \n{get_percentage(selected['Average Load Factor']):.2f}%")
    st.write(f"**Annual Consumption**  \n{selected['Annual Consumption']:,.0f} kWh")
    
    # Calculate and display peak hour consumption
    try:
        pm_consumption = get_percentage(selected['6-10 PM Consumption'])
        am_consumption = get_percentage(selected['6-8 AM Consumption'])
        peak_pct = pm_consumption + am_consumption
        st.write(f"**Peak Hour Consumption**  \n{peak_pct:.2f}% (6-10 PM + 6-8 AM)")
    except KeyError as e:
        st.error(f"Missing data column: {e}")

# Horizontal separator line after Load Info
st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# Third section - Solar metrics
st.subheader("Solar Information")
col1, col2, col3 = st.columns(3)
with col1:
    st.write(f"**Installed Solar**  \n{selected['Installed Solar Capacity (DC)']} kW")
with col2:
    st.write(f"**Annual Setoff**  \n{selected['Annual Setoff']:,.0f} kWh")
with col3:
    st.write(f"**Green %**  \n{get_percentage(selected['Percent Green Consumption']):.2f}%")

# Horizontal separator line after Solar Info
st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# --- ROI Analysis Section ---
st.title("📈 ROI Analysis")

# Sidebar controls for ROI analysis
bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

# ROI Calculation Logic
capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5  # kWh/year
bess_impact_rate = 0.91 + 0.74  # ₹/unit baseline charge
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

# Calculations
try:
    available_cd = selected['Contract Demand (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
    available_sl = selected['Sanctioned Load (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
    
    base_tariff = selected['Base Tariff']
    
    solar_to_cd_roi = ((available_cd * solar_gen_per_mw * base_tariff) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
    solar_to_sl_roi = ((available_sl * solar_gen_per_mw * base_tariff) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

    # BESS
    bess_mw = selected['Installed Solar Capacity (DC)']/1000 * bess_pct / 100
    bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    # Wind
    wind_mw = selected['Contract Demand (kVA)']/1000
    wind_saving = wind_mw * wind_gen_per_mw * base_tariff
    wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)

except KeyError as e:
    st.error(f"Missing required data column: {e}")
    st.stop()

# ROI Chart
roi_data = pd.DataFrame({
    'Option': ['Solar to CD', 'Solar to SL', f'BESS ({bess_pct}%)', 'Wind'],
    'ROI (%)': [solar_to_cd_roi * 100, solar_to_sl_roi * 100, bess_roi * 100, wind_roi * 100]
})

chart = alt.Chart(roi_data).mark_bar().encode(
    x=alt.X('Option', sort=None),
    y='ROI (%)',
    color='Option'
).properties(height=400)

st.altair_chart(chart, use_container_width=True)

# Recommendation
best_option = roi_data.loc[roi_data['ROI (%)'].idxmax()]
st.success(f"Recommended Option: {best_option['Option']} (ROI: {best_option['ROI (%)']:.2f}%)")
