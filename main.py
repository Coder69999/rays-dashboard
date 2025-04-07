import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# Load processed Excel file
data_path = "ChatgPT-1.xlsx"  # Replace with uploaded file path
sheet_name = "Sheet1"
df = pd.read_excel(data_path, sheet_name=sheet_name)

# Rename columns and clean up
df = df.rename(columns={
    'Client Name': 'Client',
    'Contract Demand (kVA)': 'CD',
    'Sanctioned Load (kVA)': 'SL',
    'Installed Solar Capacity (kW)': 'Installed_kW',
    'Monthly Consumption (kWh)': 'Monthly_kWh',
    'Tariff (₹/unit)': 'Tariff',
    'Voltage Level': 'Voltage',
    'Evening Load %': 'Evening_Load_Pct',
    'ToD Details (if any)': 'ToD_kWh'
})

# Derived columns
df['Installed_MW'] = df['Installed_kW'] / 1000
df['Annual_kWh'] = df['Monthly_kWh'] * 12
df['Evening_Load_Pct'] = df['Evening_Load_Pct'].fillna(0)
df['ToD_kWh'] = df['ToD_kWh'].fillna(0)

# Sidebar filters
client = st.sidebar.selectbox("Select Client", df['Client'].unique())
bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

selected = df[df['Client'] == client].iloc[0]

# --- LOGIC ---
capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5  # kWh/year
bess_impact_rate = 0.91 + 0.74  # ₹/unit baseline charge
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

# Calculations
available_cd = selected['CD'] - selected['Installed_kW'] / 1000
available_sl = selected['SL'] - selected['Installed_kW'] / 1000
solar_to_cd_roi = ((available_cd * solar_gen_per_mw * selected['Tariff']) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
solar_to_sl_roi = ((available_sl * solar_gen_per_mw * selected['Tariff']) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

# BESS
bess_mw = selected['Installed_MW'] * bess_pct / 100
bess_waiver_saving = selected['Annual_kWh'] * (bess_impact_rate * waiver_pct / 100)
bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

# Wind
wind_mw = selected['CD']
wind_saving = wind_mw * wind_gen_per_mw * selected['Tariff']
wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)

# Display Metrics
st.title(f"📊 Extension Opportunities for {client}")
st.metric("Installed Capacity (MW)", f"{selected['Installed_MW']:.2f}")
st.metric("Monthly Usage (lakh kWh)", f"{selected['Monthly_kWh'] / 100000:.2f}")
st.metric("Annual Usage (Cr kWh)", f"{selected['Annual_kWh'] / 1e7:.2f}")

# ROI Chart
roi_data = pd.DataFrame({
    'Option': ['Solar to CD', 'Solar to SL', f'BESS ({bess_pct}%)', 'Wind'],
    'ROI (%)': [solar_to_cd_roi * 100, solar_to_sl_roi * 100, bess_roi * 100, wind_roi * 100]
})

st.subheader("📈 ROI Comparison")
chart = alt.Chart(roi_data).mark_bar().encode(
    x=alt.X('Option', sort=None),
    y='ROI (%)',
    color='Option'
).properties(height=400)

st.altair_chart(chart, use_container_width=True)

# Recommendation
best_option = roi_data.loc[roi_data['ROI (%)'].idxmax()]
st.success(f"Recommended Option: {best_option['Option']} (ROI: {best_option['ROI (%)']:.2f}%)")
