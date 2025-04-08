import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# --- Load and clean data ---
@st.cache_data
def load_data():
    data_path = "D-V1.xlsx"  # Replace with your actual Excel file
    sheet_name = "Sheet1"
    df = pd.read_excel(data_path, sheet_name=sheet_name)

    df.columns = df.columns.str.strip()  # Remove extra spaces

    # Convert percentage columns from string to float (remove % symbol)
    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption',
                    'Percent Green Consumption', 'Solar Utilization', 'Solar ROI']
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '', regex=False).astype(float)

    return df

df = load_data()

# --- Sidebar: Client selection ---
client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

def get_percentage(val):
    try:
        return float(val)
    except:
        return 0.0

# --- Main Layout ---
st.title(f"📊 Client Dashboard: {client}")

col1, col_sep, col2 = st.columns([5, 1, 5])

# --- Left Column: Basic Load Info ---
with col1:
    st.subheader("🔌 Basic Load Information")
    st.write(f"**Voltage Level**  \n{selected['Voltage Level']} kV")
    st.write(f"**Sanctioned Load**  \n{selected['Sanctioned Load (kVA)']:,.0f} kVA")
    st.write(f"**Contract Demand**  \n{selected['Contract Demand (kVA)']:,.0f} kVA")

    st.markdown("---")

    st.subheader("📉 Load Profile")
    st.write(f"**Average Load Factor**  \n{get_percentage(selected['Average Load Factor']):.0f}%")
    st.write(f"**Annual Consumption**  \n{selected['Annual Consumption']:,.0f} kWh")

    try:
        pm = get_percentage(selected['6-10 PM Consumption'])
        am = get_percentage(selected['6-8 AM Consumption'])
        peak = pm + am
        st.write(f"**Peak Hour Consumption**  \n{peak:.0f}% (6-10 PM + 6-8 AM)")
    except KeyError as e:
        st.error(f"Missing data column: {e}")

# --- Vertical Separator ---
with col_sep:
    st.markdown(
        """
        <style>
        .vertical-line {
            border-left: 2px solid #666;
            height: 100%;
            margin-top: 30px;
        }
        </style>
        <div class="vertical-line"></div>
        """,
        unsafe_allow_html=True
    )

# --- Right Column: Solar Info ---
with col2:
    st.subheader("☀️ Solar Setup & Performance")
    st.write(f"**Installed Solar Capacity**  \n{selected['Installed Solar Capacity (DC)']} kW")
    st.write(f"**Annual Solar Generation (Setoff)**  \n{selected['Annual Setoff']:,.0f} kWh")
    st.write(f"**Green Energy %**  \n{get_percentage(selected['Percent Green Consumption']):.0f}%")

    st.markdown("<br>", unsafe_allow_html=True)

    if 'Solar Utilization' in selected:
        st.write(f"**Solar Utilization**  \n{get_percentage(selected['Solar Utilization']):.0f}%")
    if 'Solar ROI' in selected:
        st.write(f"**Solar ROI**  \n{get_percentage(selected['Solar ROI']):.0f}%")

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# --- ROI Analysis Section ---
st.title("📈 ROI Analysis")

# Sidebar controls
bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)
waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)

# Constants
capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5  # kWh/year
bess_impact_rate = 0.91 + 0.74  # ₹/unit
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5  # kWh/year

# ROI Calculations
try:
    cd = selected['Contract Demand (kVA)'] / 1000
    sl = selected['Sanctioned Load (kVA)'] / 1000
    solar = selected['Installed Solar Capacity (DC)'] / 1000
    base_tariff = selected['Base Tariff']

    solar_cd = max(cd - solar, 0)
    solar_sl = max(sl - solar, 0)

    roi_cd = ((solar_cd * solar_gen_per_mw * base_tariff) / (solar_cd * capex_solar_per_mw)) if solar_cd > 0 else 0
    roi_sl = ((solar_sl * solar_gen_per_mw * base_tariff) / (solar_sl * capex_solar_per_mw)) if solar_sl > 0 else 0

    # BESS
    bess_mw = solar * bess_pct / 100
    bess_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    roi_bess = (bess_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    # Wind
    wind_mw = cd
    wind_saving = wind_mw * wind_gen_per_mw * base_tariff
    roi_wind = wind_saving / (wind_mw * wind_capex_per_mw)

except KeyError as e:
    st.error(f"Missing required data column: {e}")
    st.stop()

# ROI DataFrame
roi_df = pd.DataFrame({
    'Option': ['Solar to CD', 'Solar to SL', f'BESS ({bess_pct}%)', 'Wind'],
    'ROI (%)': [roi_cd * 100, roi_sl * 100, roi_bess * 100, roi_wind * 100]
})

# ROI Chart
chart = alt.Chart(roi_df).mark_bar().encode(
    x=alt.X('Option', sort=None),
    y='ROI (%)',
    color='Option'
).properties(height=400)

st.altair_chart(chart, use_container_width=True)

# Recommendation
best = roi_df.loc[roi_df['ROI (%)'].idxmax()]
st.success(f"✅ Recommended Option: {best['Option']} (ROI: {best['ROI (%)']:.2f}%)")
