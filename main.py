import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# Load and clean data
@st.cache_data
def load_data():
    data_path = "D-V4.xlsx"
    sheet_name = "Sheet1"
    df = pd.read_excel(data_path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()

    percent_cols = ['Average Load Factor', '6-10 PM Consumption', '6-8 AM Consumption', 'Percent Green Consumption']
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '').astype(float)

    return df

df = load_data()

# Sidebar for client selection
client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
selected = df[df['Client Name'] == client].iloc[0]

def get_percentage(value):
    try:
        return float(str(value).replace('%', ''))
    except:
        return 0.0

st.title(f"\U0001F4CA Client Overview: {client}")

# Load Info
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
        f"{selected['Sanctioned Load (mVA)']:,.2f} mVA",
        f"{selected['Contract Demand (mVA)']:,.2f} mVA",
        f"{get_percentage(selected['Average Load Factor']*100):.2f}%",
        f"{selected['Annual Consumption']:,.0f} kWh",
        f"{get_percentage(selected['6-10 PM Consumption'])*100 + get_percentage(selected['6-8 AM Consumption'])*100:.2f}% (6-10 PM + 6-8 AM)"
    ]
}

# Solar Info
solar_info = {
    "Parameter": [
        "Solar Capacity (AC)",
        "Solar Capacity (DC)",
        "Annual Setoff",
        "Green Energy Contribution"
    ],
    
    "Value": [
        f"{selected.get('Installed Solar Capacity (AC)', 0):,.2f} MW",
        f"{selected.get('Installed Solar Capacity (DC)', 0):,.2f} MWp",
        f"{selected['Annual Setoff']*100:,.0f} kWh",
        f"{get_percentage(selected['Percent Green Consumption']*100):.2f}%"
    ]
}

df_load = pd.DataFrame(load_info)
df_solar = pd.DataFrame(solar_info)

# Display side-by-side tables
col1, col_sep, col2 = st.columns([6, 0.1, 6])

with col1:
    st.subheader("\u26A1 Basic Load Information")
    st.markdown(
        df_load.to_html(index=False, escape=False, formatters={
            "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
        }).replace(
            '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
        ),
        unsafe_allow_html=True
    )

with col_sep:
    st.markdown("<div style='height:100%; border-left: 3px solid #bbb;'></div>", unsafe_allow_html=True)

with col2:
    st.subheader("\U0001F31E Existing Solar Setup")
    st.markdown(
        df_solar.to_html(index=False, escape=False, formatters={
            "Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"
        }).replace(
            '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
        ),
        unsafe_allow_html=True
    )

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# --- New Opportunities Table Section ---
st.title("\U0001F4A1 Available Opportunities")

# Get solar capacities (default to 0 if not available)
solar_ac = selected.get('Installed Solar Capacity (AC)', 0)
solar_dc = selected.get('Installed Solar Capacity (DC)', 0)
contract_demand = selected['Contract Demand (mVA)']
sanctioned_load = selected['Sanctioned Load (mVA)']

# Calculate available contract demand (Opportunity 1)
available_cd_ac = contract_demand - solar_ac
available_cd_pct = (available_cd_ac / contract_demand) * 100 if contract_demand > 0 else 0

# Calculate available sanctioned load capacity (Opportunity 2)
available_sl_ac = sanctioned_load - solar_ac
available_sl_pct = (available_sl_ac / contract_demand) * 100 if contract_demand > 0 else 0

# Prepare opportunities data
opportunities = []

# Opportunity 1: Solar to Contract Demand
if available_cd_pct >= 20:  # Only show if at least 20% available
    opportunity1_ac = available_cd_ac
    opportunity1_dc = available_cd_ac * 1.4  # DC is 1.4 times AC
    opportunities.append({
        "Opportunity": "Solar to Contract Demand",
        "Available AC Capacity (kW)": f"{opportunity1_ac:,.2f}",
        "Recommended DC Capacity (kW)": f"{opportunity1_dc:,.2f}",
        "Status": "Available" if opportunity1_ac > 0 else "Not Available"
    })
else:
    opportunities.append({
        "Opportunity": "Solar to Contract Demand",
        "Available AC Capacity (kW)": f"{available_cd_ac:,.2f}",
        "Recommended DC Capacity (kW)": "N/A (Less than 20% available)",
        "Status": "Not Viable"
    })

# Opportunity 2: Increase Contract Demand to Sanctioned Load + Solar
if (sanctioned_load > contract_demand) and (available_sl_pct >= 20):
    opportunity2_ac = available_sl_ac
    opportunity2_dc = available_sl_ac * 1.4
    opportunities.append({
        "Opportunity": "Increase CD to SL + Solar",
        "Available AC Capacity (kW)": f"{opportunity2_ac:,.2f}",
        "Recommended DC Capacity (kW)": f"{opportunity2_dc:,.2f}",
        "Status": "Available" if opportunity2_ac > 0 else "Not Available",
        "CD Increase Required": f"{(sanctioned_load - contract_demand):,.2f} mVA"
    })
else:
    reason = ""
    if sanctioned_load <= contract_demand:
        reason = "Sanctioned load ≤ Current CD"
    else:
        reason = "Available capacity < 20% of CD"
    
    opportunities.append({
        "Opportunity": "Increase CD to SL + Solar",
        "Available AC Capacity (kW)": f"{available_sl_ac:,.2f}",
        "Recommended DC Capacity (kW)": f"N/A ({reason})",
        "Status": "Not Viable",
        "CD Increase Required": f"{(sanctioned_load - contract_demand):,.2f} mVA" if sanctioned_load > contract_demand else "N/A"
    })

# Add other opportunities (placeholders for now)
opportunities.append({
    "Opportunity": "BESS Installation",
    "Available AC Capacity (kW)": "Based on solar",
    "Recommended DC Capacity (kW)": "Based on % selected",
    "Status": "Pending",
    "CD Increase Required": "N/A"
})

opportunities.append({
    "Opportunity": "Wind Installation",
    "Available AC Capacity (kW)": "Based on contract demand",
    "Recommended DC Capacity (kW)": "N/A",
    "Status": "Pending",
    "CD Increase Required": "N/A"
})

# Create and display opportunities table
df_opportunities = pd.DataFrame(opportunities)

# Reorder columns to make CD Increase Required appear after Status
df_opportunities = df_opportunities[['Opportunity', 'Available AC Capacity (kW)', 
                                   'Recommended DC Capacity (kW)', 'Status', 
                                   'CD Increase Required']]

st.markdown(
    df_opportunities.to_html(index=False, escape=False, formatters={
        "Available AC Capacity (kW)": lambda x: f"<span style='color:#0066cc;font-weight:bold'>{x}</span>",
        "Recommended DC Capacity (kW)": lambda x: f"<span style='color:#009933;font-weight:bold'>{x}</span>",
        "Status": lambda x: f"<span style='color:{"green" if x in ["Available"] else "orange" if x == "Pending" else "red"};font-weight:bold'>{x}</span>",
        "CD Increase Required": lambda x: f"<span style='color:#9900cc;font-weight:bold'>{x}</span>"
    }).replace(
        '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
    ),
    unsafe_allow_html=True
)

st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)

# ROI Analysis
st.title("\U0001F4C8 ROI Analysis")

# Updated BESS capacity slider starting from 0%
bess_pct = st.sidebar.select_slider(
    "Select BESS Size (% of Solar)",
    options=[0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100],
    value=10
)

# Updated waiver calculation including 0% case
if bess_pct == 0:
    waiver_pct = 0
elif bess_pct < 30:
    waiver_pct = 75 + (bess_pct // 5 - 1) * 5  # 5% steps from 75% to 100%
else:
    waiver_pct = 100  # 100% for 30% and above

# Updated waiver display label
st.sidebar.markdown(f"""
**Transmission and Wheeling Charges Waiver**  
<span style='font-size:24px; color:#4CAF50'>{waiver_pct}%</span>
""", unsafe_allow_html=True)

capex_solar_per_mw = 3.5e6
capex_bess_per_mw = 4.0e6
solar_gen_per_mw = 16.5e5
bess_impact_rate = 0.91 + 0.74
wind_capex_per_mw = 6.5e6
wind_gen_per_mw = 26.0e5

try:
    available_cd = contract_demand/1000 - solar_ac/1000
    available_sl = selected['Sanctioned Load (mVA)']/1000 - solar_ac/1000

    base_tariff = selected['Base Tariff']

    solar_to_cd_roi = ((available_cd * solar_gen_per_mw * base_tariff) / (available_cd * capex_solar_per_mw)) if available_cd > 0 else 0
    solar_to_sl_roi = ((available_sl * solar_gen_per_mw * base_tariff) / (available_sl * capex_solar_per_mw)) if available_sl > 0 else 0

    bess_mw = solar_dc / 1000 * bess_pct / 100
    bess_waiver_saving = selected['Annual Consumption'] * (bess_impact_rate * waiver_pct / 100)
    bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw > 0 else 0

    wind_mw = contract_demand/1000
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
