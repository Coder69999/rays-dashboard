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
    df.columns = df.columns.str.strip()
    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption', 'Percent Green Consumption']
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '').astype(float)
    return df

df = load_data()

client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

def get_percentage(value):
    try:
        return float(str(value).replace('%', ''))
    except:
        return 0.0

st.title(f"\U0001F4CA Client Overview: {client}")

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

solar_info = {"Parameter": [], "Value": []}
solar_info["Parameter"].extend(["Installed Solar Capacity", "Annual Setoff", "Green Energy Contribution"])
solar_info["Value"].extend([
    f"{selected['Installed Solar Capacity (DC)']:,.0f} kW",
    f"{selected['Annual Setoff']:,.0f} kWh",
    f"{get_percentage(selected['Percent Green Consumption'])*100:.2f}%"
])
if 'Solar Utilization' in selected and pd.notna(selected['Solar Utilization']):
    solar_info["Parameter"].append("Solar Utilization")
    solar_info["Value"].append(f"{get_percentage(selected['Solar Utilization']):.2f}%")
if 'Solar ROI' in selected and pd.notna(selected['Solar ROI']):
    solar_info["Parameter"].append("Solar ROI")
    solar_info["Value"].append(f"{get_percentage(selected['Solar ROI']):.2f}%")

df_load = pd.DataFrame(load_info)
df_solar = pd.DataFrame(solar_info)

# Style to center headers and darken background
table_style = """
<style>
thead th {
    text-align: center !important;
    background-color: #1e1e1e !important;
    color: white !important;
    font-weight: bold !important;
}
td {
    text-align: left;
}
</style>
"""

col1, col_sep, col2 = st.columns([6, 0.1, 6])

with col1:
    st.subheader("\u26A1 Basic Load Information")
    st.markdown(table_style + df_load.to_html(index=False, escape=False, formatters={
        "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
    }), unsafe_allow_html=True)

with col_sep:
    st.markdown("<div style='height:100%; border-left: 3px solid #bbb;'></div>", unsafe_allow_html=True)

with col2:
    st.subheader("\U0001F31E Existing Solar Setup")
    st.markdown(table_style + df_solar.to_html(index=False, escape=False, formatters={
        "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
    }), unsafe_allow_html=True)

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# ROI Analysis Section
st.title("\U0001F4C8 ROI Analysis")
bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5
bess_impact_rate = 0.91 + 0.74
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

try:
    available_cd = selected['Contract Demand (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
    available_sl = selected['Sanctioned Load (kVA)']/1000 - selected['Installed Solar Capacity (DC)']/1000
    base_tariff = selected['Base Tariff']

    solar_to_cd_roi = ((available_cd * solar_gen_per_mw * base_tariff) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
    solar_to_sl_roi = ((available_sl * solar_gen_per_mw * base_tariff) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

    bess_mw = selected['Installed Solar Capacity (DC)']/1000 * bess_pct / 100
    bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    wind_mw = selected['Contract Demand (kVA)']/1000
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
