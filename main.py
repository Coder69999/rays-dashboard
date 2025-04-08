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

    # Normalize column names
    df.columns = df.columns.str.strip().str.replace('\u00a0', ' ').str.replace(r'\s+', ' ', regex=True)

    # Convert percentage columns to numeric values
    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption', 'Percent Green Consumption', 'Solar Utilization', 'Solar ROI']
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

solar_ac = selected.get('Installed Solar Capacity (AC)', np.nan)
solar_dc = selected.get('Installed Solar Capacity (DC)', np.nan)

solar_info = {
    "Parameter": [
        "Installed Solar Capacity (AC)",
        "Installed Solar Capacity (DC)",
        "Annual Setoff",
        "Green Energy Contribution",
        "Solar Utilization" if 'Solar Utilization' in selected else "",
        "Solar ROI" if 'Solar ROI' in selected else ""
    ],
    "Value": [
        f"{solar_ac:,.2f} kW" if not pd.isna(solar_ac) else "N/A",
        f"{solar_dc:,.2f} kW" if not pd.isna(solar_dc) else "N/A",
        f"{selected['Annual Setoff']:,.0f} kWh",
        f"{get_percentage(selected['Percent Green Consumption'])*100:.2f}%",
        f"{get_percentage(selected['Solar Utilization']):.2f}%" if 'Solar Utilization' in selected else "",
        f"{get_percentage(selected['Solar ROI']):.2f}%" if 'Solar ROI' in selected else ""
    ]
}

# Create dataframes
df_load = pd.DataFrame(load_info)
df_solar = pd.DataFrame(solar_info).dropna()

# Columns for tables
col1, col_sep, col2 = st.columns([6, 0.1, 6])

with col1:
    st.subheader("\u26A1 Basic Load Information")
    st.markdown(
        df_load.to_html(index=False, escape=False, formatters={
            "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
        }),
        unsafe_allow_html=True
    )

with col_sep:
    st.markdown("<div style='height:100%; border-left: 3px solid #bbb;'></div>", unsafe_allow_html=True)

with col2:
    st.subheader("\U0001F31E Existing Solar Setup")
    styled_table = df_solar.to_html(index=False, escape=False, formatters={
        "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
    })
    styled_table = styled_table.replace('<th>', '<th style="text-align:center;background-color:#f0f0f0">')
    st.markdown(styled_table, unsafe_allow_html=True)

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# --- Options with Capacity ---
try:
    solar_to_cd_mw = max(0, selected['Contract Demand (kVA)']/1000 - solar_ac/1000)
    solar_to_sl_mw = max(0, selected['Sanctioned Load (kVA)']/1000 - solar_ac/1000)
    bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
    bess_mw = solar_ac/1000 * bess_pct / 100
    wind_mw = selected['Contract Demand (kVA)']/1000

    st.subheader("\U0001F527 Extension Options Capacity")
    st.markdown(f"- **Solar to CD**: {solar_to_cd_mw:.2f} MW")
    st.markdown(f"- **Solar to SL**: {solar_to_sl_mw:.2f} MW")
    st.markdown(f"- **BESS ({bess_pct}%)**: {bess_mw:.2f} MW")
    st.markdown(f"- **Wind**: {wind_mw:.2f} MW")
except Exception as e:
    st.error(f"Error calculating capacities: {e}")

# --- ROI Analysis Section ---
st.title("\U0001F4C8 ROI Analysis")

waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5  # kWh/year
bess_impact_rate = 0.91 + 0.74  # ₹/unit baseline charge
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

try:
    base_tariff = selected['Base Tariff']

    solar_to_cd_roi = ((solar_to_cd_mw * solar_gen_per_mw * base_tariff) / (solar_to_cd_mw * capex_solar_per_mw)) if solar_to_cd_mw > 0 else 0
    solar_to_sl_roi = ((solar_to_sl_mw * solar_gen_per_mw * base_tariff) / (solar_to_sl_mw * capex_solar_per_mw)) if solar_to_sl_mw > 0 else 0

    bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    wind_saving = wind_mw * wind_gen_per_mw * base_tariff
    wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)

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

except KeyError as e:
    st.error(f"Missing required data column: {e}")
    st.stop()
