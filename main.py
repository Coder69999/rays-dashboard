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
    
    # Clean column names
    df.columns = df.columns.str.strip()  # Remove whitespace
    df.columns = df.columns.str.replace(' ', '_')  # Replace spaces with underscores
    df.columns = df.columns.str.replace('[^a-zA-Z0-9_]', '', regex=True)  # Remove special chars
    
    # Convert percentage strings to numeric values
    percent_cols = [col for col in df.columns if '%' in col or 'Percent' in col]
    for col in percent_cols:
        df[col] = df[col].astype(str).str.replace('%', '').astype(float) / 100
    
    return df

df = load_data()

# Debug: Show available columns
# st.write("Available columns:", df.columns.tolist())

# Sidebar for client selection
client = st.sidebar.selectbox("Select Client", df['Client_Name'].unique())
selected = df[df['Client_Name'] == client].iloc[0]

# Main display
st.title(f"📊 Client Overview: {client}")

# First section - Basic info
col1, col2, col3 = st.columns(3)
with col1:
    st.subheader("Voltage Level")
    st.write(f"{selected['Voltage_Level']} kV")
    
with col2:
    st.subheader("Sanctioned Load")
    st.write(f"{selected['Sanctioned_Load_kVA']:,.0f} kVA")
    
with col3:
    st.subheader("Contract Demand")
    st.write(f"{selected['Contract_Demand_kVA']:,.0f} kVA")

# Second section - Load and consumption
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Average Load Factor")
    st.write(f"{selected['Average_Load_Factor']*100:.2f}%")
    
    st.subheader("Annual Consumption")
    st.write(f"{selected['Annual_Consumption']:,.0f} kWh")
    
    # Calculate and display peak hour consumption
    pm_consumption = selected['6-10_PM_Consumption']
    am_consumption = selected['6-8_AM_Consumption']
    peak_pct = (pm_consumption + am_consumption) * 100
    st.subheader("Peak Hour Consumption")
    st.write(f"{peak_pct:.2f}% (6-10 PM + 6-8 AM)")

with col2:
    # Solar metrics in a horizontal layout
    cols = st.columns(3)
    with cols[0]:
        st.subheader("Installed Solar")
        st.write(f"{selected['Installed_Solar_Capacity_DC']} kW")
    with cols[1]:
        st.subheader("Annual Setoff")
        st.write(f"{selected['Annual_Setoff']:,.0f} kWh")
    with cols[2]:
        st.subheader("Green %")
        st.write(f"{selected['Percent_Green_Consumption']*100:.2f}%")

# Add a separator line
st.markdown("---")

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
available_cd = selected['Contract_Demand_kVA']/1000 - selected['Installed_Solar_Capacity_DC']/1000
available_sl = selected['Sanctioned_Load_kVA']/1000 - selected['Installed_Solar_Capacity_DC']/1000
solar_to_cd_roi = ((available_cd * solar_gen_per_mw * selected['Base_Tariff']) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
solar_to_sl_roi = ((available_sl * solar_gen_per_mw * selected['Base_Tariff']) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

# BESS
bess_mw = selected['Installed_Solar_Capacity_DC']/1000 * bess_pct / 100
bess_waiver_saving = selected['Annual_Consumption'] * (bess_impact_rate * waiver_pct / 100)
bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

# Wind
wind_mw = selected['Contract_Demand_kVA']/1000
wind_saving = wind_mw * wind_gen_per_mw * selected['Base_Tariff']
wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)

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
