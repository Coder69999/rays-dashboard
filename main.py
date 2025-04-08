import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# Load processed Excel file
data_path = "ChatgPT-1.xlsx"  # Replace with uploaded file path
sheet_name = "Sheet1"
df = pd.read_excel(data_path, sheet_name=sheet_name)

# Clean column names and handle missing values
df.columns = df.columns.str.strip()
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
df.replace('-', 0, inplace=True)

# Sidebar for client selection
client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

# Main display
st.title(f"📊 Client Overview: {client}")

# First section - Basic info
col1, col2, col3 = st.columns(3)
with col1:
    st.subheader("Voltage Level")
    st.write(f"{selected['Voltage Level']} kV")
    
with col2:
    st.subheader("Sanctioned Load")
    st.write(f"{selected['Sanctioned Load (kVA)']:,.0f} kVA")
    
with col3:
    st.subheader("Contract Demand")
    st.write(f"{selected['Contract Demand (kVA)']:,.0f} kVA")

# Second section - Load and consumption
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Average Load Factor")
    st.write(f"{float(selected['Average Load Factor'].strip('%')) if isinstance(selected['Average Load Factor'], str) else selected['Average Load Factor']:.2f}%")
    
    st.subheader("Annual Consumption")
    st.write(f"{selected['Annual Consumption']:,.0f} kWh")
    
    # Calculate and display peak hour consumption
    pm_consumption = float(selected['6-10 PM Consumption'].strip('%')) if isinstance(selected['6-10 PM Consumption'], str) else selected['6-10 PM Consumption']
    am_consumption = float(selected['6-8 AM Consumption'].strip('%')) if isinstance(selected['6-8 AM Consumption'], str) else selected['6-8 AM Consumption']
    peak_pct = pm_consumption + am_consumption
    st.subheader("Peak Hour Consumption")
    st.write(f"{peak_pct:.2f}% (6-10 PM + 6-8 AM)")

with col2:
    # Solar metrics in a horizontal layout
    cols = st.columns(3)
    with cols[0]:
        st.subheader("Installed Solar")
        st.write(f"{selected['Installed Solar Capacity (DC)']} kW")
    with cols[1]:
        st.subheader("Annual Setoff")
        st.write(f"{selected['Annual Setoff']:,.0f} kWh")
    with cols[2]:
        st.subheader("Green %")
        st.write(f"{float(selected['Percent Green Consumption'].strip('%')) if isinstance(selected['Percent Green Consumption'], str) else selected['Percent Green Consumption']:.2f}%")

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
available_cd = selected['Contract Demand (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
available_sl = selected['Sanctioned Load (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
solar_to_cd_roi = ((available_cd * solar_gen_per_mw * selected['Base Tariff']) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
solar_to_sl_roi = ((available_sl * solar_gen_per_mw * selected['Base Tariff']) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

# BESS
bess_mw = selected['Installed Solar Capacity (DC)']/1000 * bess_pct / 100
bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

# Wind
wind_mw = selected['Contract Demand (kVA)']/1000
wind_saving = wind_mw * wind_gen_per_mw * selected['Base Tariff']
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
