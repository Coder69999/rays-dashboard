import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

def format_indian(t):
    """Indian number formatting with commas (e.g., 2,42,30,450)"""
    try:
        t = float(t)
        if t.is_integer():
            t = int(t)
        
        s = str(t).split('.')
        if len(s) == 1:
            # Integer formatting
            s = s[0]
            if len(s) > 3:
                last_three = s[-3:]
                other_numbers = s[:-3]
                res = ""
                while len(other_numbers) > 2:
                    res = "," + other_numbers[-2:] + res
                    other_numbers = other_numbers[:-2]
                if other_numbers:
                    res = other_numbers + res
                return res + "," + last_three
            return s
        else:
            # Float formatting
            before_decimal = s[0]
            after_decimal = s[1][:2]  # Take 2 decimal places
            if len(before_decimal) > 3:
                last_three = before_decimal[-3:]
                other_numbers = before_decimal[:-3]
                res = ""
                while len(other_numbers) > 2:
                    res = "," + other_numbers[-2:] + res
                    other_numbers = other_numbers[:-2]
                if other_numbers:
                    res = other_numbers + res
                return res + "," + last_three + "." + after_decimal
            return before_decimal + "." + after_decimal
    except:
        return str(t)

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

def main():
    df = load_data()
    client = st.sidebar.selectbox("Select Client", df['Client Name'].unique())
    selected = df[df['Client Name'] == client].iloc[0]

    def get_percentage(value):
        try:
            return float(str(value).replace('%', ''))
        except:
            return 0.0

    st.title(f"\U0001F4CA Client Overview: {client}")

    # Load Info Table
    load_info = {
        "Parameter": ["Voltage Level", "Sanctioned Load", "Contract Demand", 
                     "Average Load Factor", "Annual Consumption", "Peak Hour Consumption"],
        "Value": [
            f"{selected['Voltage Level']} kV",
            f"{format_indian(selected['Sanctioned Load (mVA)'])} mVA",
            f"{format_indian(selected['Contract Demand (mVA)'])} mVA",
            f"{get_percentage(selected['Average Load Factor']*100):.2f}%",
            f"{format_indian(selected['Annual Consumption'])} kWh",
            f"{get_percentage(selected['6-10 PM Consumption'])*100 + get_percentage(selected['6-8 AM Consumption'])*100:.2f}% (6-10 PM + 6-8 AM)"
        ]
    }

    # Solar Info Table
    solar_info = {
        "Parameter": ["Solar Capacity (AC)", "Solar Capacity (DC)", 
                     "Annual Setoff", "Green Energy Contribution"],
        "Value": [
            f"{format_indian(selected.get('Installed Solar Capacity (AC)', 0))} MW",
            f"{format_indian(selected.get('Installed Solar Capacity (DC)', 0))} MWp",
            f"{format_indian(selected['Annual Setoff'])} kWh",
            f"{get_percentage(selected['Percent Green Consumption']*100):.2f}%"
        ]
    }

    # Display Tables
    col1, col_sep, col2 = st.columns([6, 0.1, 6])
    with col1:
        st.subheader("\u26A1 Basic Load Information")
        st.markdown(
            pd.DataFrame(load_info).to_html(
                index=False, 
                escape=False,
                formatters={"Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"}
            ).replace(
                '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
            ),
            unsafe_allow_html=True
        )
    
    with col_sep:
        st.markdown("<div style='height:100%; border-left: 3px solid #bbb;'></div>", unsafe_allow_html=True)

    with col2:
        st.subheader("\U0001F31E Existing Solar Setup")
        st.markdown(
            pd.DataFrame(solar_info).to_html(
                index=False, 
                escape=False,
                formatters={"Value": lambda x: f"<span style='color:#e67300;font-weight:bold'>{x}</span>"}
            ).replace(
                '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
            ),
            unsafe_allow_html=True
        )

    st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" />""", unsafe_allow_html=True)

    # --- Opportunities Section ---
    st.title("\U0001F4A1 Available Opportunities")
    solar_ac = selected.get('Installed Solar Capacity (AC)', 0)
    solar_dc = selected.get('Installed Solar Capacity (DC)', 0)
    contract_demand = selected['Contract Demand (mVA)']
    sanctioned_load = selected['Sanctioned Load (mVA)']

    # Opportunity Calculations
    opportunities = []
    
    # Opportunity 1: Solar to Contract Demand
    available_cd_ac = contract_demand - solar_ac
    available_cd_pct = (available_cd_ac / contract_demand) * 100 if contract_demand > 0 else 0
    if available_cd_pct >= 20:
        opportunities.append({
            "Opportunity": "Solar to Contract Demand",
            "Available AC Capacity (kW)": f"{format_indian(available_cd_ac)}",
            "Recommended DC Capacity (kW)": f"{format_indian(available_cd_ac * 1.4)}",
            "Status": "Available",
            "CD Increase Required": "N/A"
        })
    else:
        opportunities.append({
            "Opportunity": "Solar to Contract Demand",
            "Available AC Capacity (kW)": f"{format_indian(available_cd_ac)}",
            "Recommended DC Capacity (kW)": "N/A (Less than 20%)",
            "Status": "Not Viable",
            "CD Increase Required": "N/A"
        })

    # Opportunity 2: Increase CD to SL + Solar
    available_sl_ac = sanctioned_load - solar_ac
    available_sl_pct = (available_sl_ac / contract_demand) * 100 if contract_demand > 0 else 0
    if (sanctioned_load > contract_demand) and (available_sl_pct >= 20):
        opportunities.append({
            "Opportunity": "Increase CD to SL + Solar",
            "Available AC Capacity (kW)": f"{format_indian(available_sl_ac)}",
            "Recommended DC Capacity (kW)": f"{format_indian(available_sl_ac * 1.4)}",
            "Status": "Available",
            "CD Increase Required": f"{format_indian(sanctioned_load - contract_demand)} mVA"
        })
    else:
        reason = "Sanctioned load ≤ Current CD" if sanctioned_load <= contract_demand else "Available capacity < 20%"
        opportunities.append({
            "Opportunity": "Increase CD to SL + Solar",
            "Available AC Capacity (kW)": f"{format_indian(available_sl_ac)}",
            "Recommended DC Capacity (kW)": f"N/A ({reason})",
            "Status": "Not Viable",
            "CD Increase Required": f"{format_indian(sanctioned_load - contract_demand)} mVA" if sanctioned_load > contract_demand else "N/A"
        })

    # Other Opportunities
    opportunities.extend([
        {
            "Opportunity": "BESS Installation",
            "Available AC Capacity (kW)": "Based on solar",
            "Recommended DC Capacity (kW)": "Based on % selected",
            "Status": "Pending",
            "CD Increase Required": "N/A"
        },
        {
            "Opportunity": "Wind Installation",
            "Available AC Capacity (kW)": "Based on contract demand",
            "Recommended DC Capacity (kW)": "N/A",
            "Status": "Pending",
            "CD Increase Required": "N/A"
        }
    ])

    # Display Opportunities Table
    df_opportunities = pd.DataFrame(opportunities)[['Opportunity', 'Available AC Capacity (kW)', 
                                                  'Recommended DC Capacity (kW)', 'Status', 
                                                  'CD Increase Required']]

    st.markdown(
        df_opportunities.to_html(
            index=False, 
            escape=False,
            formatters={
                "Available AC Capacity (kW)": lambda x: f"<span style='color:#0066cc;font-weight:bold'>{x}</span>",
                "Recommended DC Capacity (kW)": lambda x: f"<span style='color:#009933;font-weight:bold'>{x}</span>",
                "Status": lambda x: f"<span style='color:{"green" if x=="Available" else "orange" if x=="Pending" else "red"};font-weight:bold'>{x}</span>",
                "CD Increase Required": lambda x: f"<span style='color:#9900cc;font-weight:bold'>{x}</span>"
            }
        ).replace(
            '<th>', '<th style="text-align:center; background-color:#f0f0f0; color:#000;">'
        ),
        unsafe_allow_html=True
    )

    st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" />""", unsafe_allow_html=True)

    # --- ROI Analysis ---
    st.title("\U0001F4C8 ROI Analysis")
    
    # BESS Configuration
    bess_pct = st.sidebar.select_slider(
        "Select BESS Size (% of Solar)",
        options=list(range(0, 101, 5)),
        value=10
    )
    waiver_pct = 0 if bess_pct == 0 else (75 + (bess_pct//5 - 1)*5 if bess_pct < 30 else 100)
    st.sidebar.markdown(f"""
    **Transmission and Wheeling Charges Waiver**  
    <span style='font-size:24px; color:#4CAF50'>{waiver_pct}%</span>
    """, unsafe_allow_html=True)

    # ROI Calculations
    capex = {
        'solar': 3.5e6,
        'bess': 4.0e6,
        'wind': 6.5e6
    }
    generation = {
        'solar': 16.5e5,
        'wind': 26.0e5
    }
    
    try:
        available_cd = contract_demand/1000 - solar_ac/1000
        available_sl = selected['Sanctioned Load (mVA)']/1000 - solar_ac/1000
        base_tariff = selected['Base Tariff']

        # Solar ROI Calculations
        solar_to_cd_roi = ((available_cd * generation['solar'] * base_tariff) / 
                         (available_cd * capex['solar'])) if available_cd > 0 else 0
        solar_to_sl_roi = ((available_sl * generation['solar'] * base_tariff) / 
                         (available_sl * capex['solar'])) if available_sl > 0 else 0
        
        # BESS ROI Calculation
        bess_mw = solar_dc/1000 * bess_pct/100
        bess_roi = (selected['Annual Consumption'] * (1.65 * waiver_pct/100) / 
                   (bess_mw * capex['bess'])) if bess_mw > 0 else 0
        
        # Wind ROI Calculation
        wind_roi = (contract_demand/1000 * generation['wind'] * base_tariff / 
                   (contract_demand/1000 * capex['wind']))
        
        # Display Results
        roi_data = pd.DataFrame({
            'Option': ['Solar to CD', 'Solar to SL', f'BESS ({bess_pct}%)', 'Wind'],
            'ROI (%)': [solar_to_cd_roi*100, solar_to_sl_roi*100, bess_roi*100, wind_roi*100]
        })
        
        st.altair_chart(alt.Chart(roi_data).mark_bar().encode(
            x=alt.X('Option', sort=None),
            y='ROI (%)',
            color='Option'
        ).properties(height=400), use_container_width=True)
        
        best_option = roi_data.loc[roi_data['ROI (%)'].idxmax()]
        st.success(f"Recommended Option: {best_option['Option']} (ROI: {best_option['ROI (%)']:.2f}%)")
        
    except KeyError as e:
        st.error(f"Missing required data column: {e}")

if __name__ == "__main__":
    main()
