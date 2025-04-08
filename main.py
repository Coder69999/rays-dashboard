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

# Color styling function
def red_text(val):
    return f"<span style='color:#d62728;font-weight:bold'>{val}</span>"

# Main display
st.title(f"\U0001F4CA Client Overview: {client}")

# Create two main columns with a vertical separator
col1, col_sep, col2 = st.columns([5, 0.5, 5])

with col1:
    st.subheader("\u26A1 Basic Load Information")
    st.markdown(f"**Voltage Level:** {red_text(f'{selected['Voltage Level']} kV')}", unsafe_allow_html=True)
    st.markdown(f"**Sanctioned Load:** {red_text(f'{selected['Sanctioned Load (kVA)']:,.0f} kVA')}", unsafe_allow_html=True)
    st.markdown(f"**Contract Demand:** {red_text(f'{selected['Contract Demand (kVA)']:,.0f} kVA')}", unsafe_allow_html=True)
    st.markdown(f"**Average Load Factor:** {red_text(f'{get_percentage(selected['Average Load Factor']*100):.2f}%')}", unsafe_allow_html=True)
    st.markdown(f"**Annual Consumption:** {red_text(f'{selected['Annual Consumption']:,.0f} kWh')}", unsafe_allow_html=True)

    try:
        pm = get_percentage(selected['6-10 PM Consumption']) * 100
        am = get_percentage(selected['6-8 AM Consumption']) * 100
        total_peak = f"{pm + am:.2f}%"
        st.markdown(f"**Peak Hour Consumption:** {red_text(f'{total_peak} (6-10 PM + 6-8 AM)')}", unsafe_allow_html=True)
    except:
        st.warning("Peak hour data not available.")

with col_sep:
    st.markdown(
        """
        <div style='height:100%; border-left: 3px solid #bbb;'></div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.subheader("\U0001F31E Existing Solar Setup & Setoff")
    st.markdown(f"**Installed Solar Capacity:** {red_text(f'{selected['Installed Solar Capacity (DC)']:,.0f} kW')}", unsafe_allow_html=True)
    st.markdown(f"**Annual Setoff:** {red_text(f'{selected['Annual Setoff']:,.0f} kWh')}", unsafe_allow_html=True)
    st.markdown(f"**Green Energy Contribution:** {red_text(f'{get_percentage(selected['Percent Green Consumption'])*100:.2f}%')}", unsafe_allow_html=True)

    if 'Solar Utilization' in selected:
        st.markdown(f"**Solar Utilization:** {red_text(f'{get_percentage(selected['Solar Utilization']):.2f}%')}", unsafe_allow_html=True)

    if 'Solar ROI' in selected:
        st.markdown(f"**Solar ROI:** {red_text(f'{get_percentage(selected['Solar ROI']):.2f}%')}", unsafe_allow_html=True)

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

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
