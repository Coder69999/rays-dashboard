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
    try:
        return float(str(value).replace('%', ''))
    except:
        return 0.0

# Main display
st.title(f"\U0001F4CA Client Overview: {client}")

# Prepare data for table display
load_info = {
    "Parameter": [
        "Voltage Level",
        "Sanctioned Load",
        "Contract Demand",
        "Average Load Factor",
        "Annual Consumption",
        "Peak Hour Consumption"
    ],
    "Value": [
        f"{selected['Voltage Level']} kV",
        f"{selected['Sanctioned Load (kVA)']:,.0f} kVA",
        f"{selected['Contract Demand (kVA)']:,.0f} kVA",
        f"{get_percentage(selected['Average Load Factor']*100):.2f}%",
        f"{selected['Annual Consumption']:,.0f} kWh",
        f"{get_percentage(selected['6-10 PM Consumption'])*100 + get_percentage(selected['6-8 AM Consumption'])*100:.2f}% (6-10 PM + 6-8 AM)"
    ]
}

solar_info = {
    "Parameter": [
        "Installed Solar Capacity (AC)",
        "Installed Solar Capacity (DC)",
        "Annual Setoff",
        "Green Energy Contribution"
    ],
    "Value": [
        f"{selected.get('Installed Solar Capacity (AC)', 0):,.2f} kW",
        f"{selected.get('Installed Solar Capacity (DC)', 0):,.2f} kW",
        f"{selected['Annual Setoff']:,.0f} kWh",
        f"{get_percentage(selected['Percent Green Consumption'])*100:.2f}%"
    ]
}

df_load = pd.DataFrame(load_info)
df_solar = pd.DataFrame(solar_info)

# Columns for tables
col1, col_sep, col2 = st.columns([6, 0.1, 6])

with col1:
    st.subheader("\u26A1 Basic Load Information")
    st.markdown(
        df_load.to_html(index=False, escape=False, formatters={
            "Parameter": lambda x: f"<th style='text-align:center;background:#f0f0f0'>{x}</th>",
            "Value": lambda x: f"<td style='color:#e67300;font-weight:bold'>{x}</td>"
        }),
        unsafe_allow_html=True
    )

with col_sep:
    st.markdown("<div style='height:100%; border-left: 3px solid #bbb;'></div>", unsafe_allow_html=True)

with col2:
    st.subheader("\U0001F31E Existing Solar Setup")
    st.markdown(
        df_solar.to_html(index=False, escape=False, formatters={
            "Parameter": lambda x: f"<th style='text-align:center;background:#f0f0f0'>{x}</th>",
            "Value": lambda x: f"<td style='color:#e67300;font-weight:bold'>{x}</td>"
        }),
        unsafe_allow_html=True
    )

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# --- Opportunity Assessment Section ---
st.title("\U0001F4A1 Available Extension Opportunities")

opportunities = []

# --- 1. Solar to Contract Demand Opportunity ---
installed_ac = selected.get('Installed Solar Capacity (AC)', 0)
contract_demand = selected.get('Contract Demand (kVA)', 0)

available_cd_ac = contract_demand - installed_ac
threshold_cd = 0.2 * contract_demand  # 20%

if available_cd_ac > threshold_cd:
    available_cd_dc = 1.4 * available_cd_ac
    opportunities.append({
        "Option": "Solar to Contract Demand",
        "Available AC Capacity (kW)": f"{available_cd_ac:.2f}",
        "Estimated DC Capacity (kW)": f"{available_cd_dc:.2f}"
    })

if opportunities:
    df_opps = pd.DataFrame(opportunities)
    st.markdown("<h5 style='margin-top: 0.5em;'>Potential Options Based on Capacity</h5>", unsafe_allow_html=True)
    st.dataframe(df_opps, use_container_width=True)
else:
    st.warning("No significant extension opportunities available based on current contract demand.")

# --- ROI Analysis Section ---
st.title("\U0001F4C8 ROI Analysis")

bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5  # kWh/year
bess_impact_rate = 0.91 + 0.74  # ₹/unit baseline charge
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

try:
    available_cd = contract_demand / 1000 - installed_ac / 1000
    available_sl = selected['Sanctioned Load (kVA)']/1000 - installed_ac / 1000

    base_tariff = selected['Base Tariff']

    solar_to_cd_roi = ((available_cd * solar_gen_per_mw * base_tariff) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
    solar_to_sl_roi = ((available_sl * solar_gen_per_mw * base_tariff) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

    bess_mw = installed_ac / 1000 * bess_pct / 100
    bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    wind_mw = contract_demand / 1000
    wind_saving = wind_mw * wind_gen_per_mw * base_tariff
    wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)

except KeyError as e:
    st.error(f"Missing required data column: {e}")
    st.stop()

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

best_option = roi_data.loc[roi_data['ROI (%)'].idxmax()]
st.success(f"Recommended Option: {best_option['Option']} (ROI: {best_option['ROI (%)']:.2f}%)")
