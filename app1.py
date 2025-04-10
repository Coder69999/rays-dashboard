bplist00�_WebMainResource�	
_WebResourceMIMEType_WebResourceFrameName^WebResourceURL_WebResourceTextEncodingName_WebResourceDataYtext/htmlP_file:///index.htmlUutf-8Ol<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta http-equiv="Content-Style-Type" content="text/css">
  <title></title>
  <meta name="Generator" content="Cocoa HTML Writer">
  <meta name="CocoaVersion" content="2575.4">
  <style type="text/css">
    p.p1 {margin: 0.0px 0.0px 0.0px 0.0px; font: 12.0px Helvetica}
    p.p2 {margin: 0.0px 0.0px 0.0px 0.0px; font: 12.0px Helvetica; min-height: 14.0px}
  </style>
</head>
<body>
<p class="p1">import streamlit as st</p>
<p class="p1">import pandas as pd</p>
<p class="p1">import numpy as np</p>
<p class="p1">import altair as alt</p>
<p class="p2"><br></p>
<p class="p1"># Load processed Excel file</p>
<p class="p1">data_path = "ChatgPT-1.xlsx"<span class="Apple-converted-space">  </span># Replace with uploaded file path</p>
<p class="p1">sheet_name = "Sheet1"</p>
<p class="p1">df = pd.read_excel(data_path, sheet_name=sheet_name)</p>
<p class="p2"><br></p>
<p class="p1"># Rename columns and clean up</p>
<p class="p1">df = df.rename(columns={</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Client Name': 'Client',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Contract Demand (kVA)': 'CD',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Sanctioned Load (kVA)': 'SL',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Installed Solar Capacity (kW)': 'Installed_kW',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Monthly Consumption (kWh)': 'Monthly_kWh',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Tariff (₹/unit)': 'Tariff',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Voltage Level': 'Voltage',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Evening Load %': 'Evening_Load_Pct',</p>
<p class="p1"><span class="Apple-converted-space">    </span>'ToD Details (if any)': 'ToD_kWh'</p>
<p class="p1">})</p>
<p class="p2"><br></p>
<p class="p1"># Derived columns</p>
<p class="p1">df['Installed_MW'] = df['Installed_kW'] / 1000</p>
<p class="p1">df['Annual_kWh'] = df['Monthly_kWh'] * 12</p>
<p class="p1">df['Evening_Load_Pct'] = df['Evening_Load_Pct'].fillna(0)</p>
<p class="p1">df['ToD_kWh'] = df['ToD_kWh'].fillna(0)</p>
<p class="p2"><br></p>
<p class="p1"># Sidebar filters</p>
<p class="p1">client = st.sidebar.selectbox("Select Client", df['Client'].unique())</p>
<p class="p1">bess_pct = st.sidebar.slider("Select BESS Size (% of Solar)", 5, 30, 10)</p>
<p class="p1">waiver_pct = st.sidebar.slider("Charge Waiver (%) with BESS", 75, 100, 75)</p>
<p class="p2"><br></p>
<p class="p1">selected = df[df['Client'] == client].iloc[0]</p>
<p class="p2"><br></p>
<p class="p1"># --- LOGIC ---</p>
<p class="p1">capex_solar_per_mw = 3.5e6</p>
<p class="p1">capex_bess_per_mw = 4.0e6</p>
<p class="p1">solar_gen_per_mw = 16.5e5<span class="Apple-converted-space">  </span># kWh/year</p>
<p class="p1">bess_impact_rate = 0.91 + 0.74<span class="Apple-converted-space">  </span># ₹/unit baseline charge</p>
<p class="p1">wind_capex_per_mw = 6.5e6</p>
<p class="p1">wind_gen_per_mw = 26.0e5</p>
<p class="p2"><br></p>
<p class="p1"># Calculations</p>
<p class="p1">available_cd = selected['CD'] - selected['Installed_kW'] / 1000</p>
<p class="p1">available_sl = selected['SL'] - selected['Installed_kW'] / 1000</p>
<p class="p1">solar_to_cd_roi = ((available_cd * solar_gen_per_mw * selected['Tariff']) / (available_cd * capex_solar_per_mw)) if available_cd &gt; 0 else 0</p>
<p class="p1">solar_to_sl_roi = ((available_sl * solar_gen_per_mw * selected['Tariff']) / (available_sl * capex_solar_per_mw)) if available_sl &gt; 0 else 0</p>
<p class="p2"><br></p>
<p class="p1"># BESS</p>
<p class="p1">bess_mw = selected['Installed_MW'] * bess_pct / 100</p>
<p class="p1">bess_waiver_saving = selected['Annual_kWh'] * (bess_impact_rate * waiver_pct / 100)</p>
<p class="p1">bess_roi = (bess_waiver_saving / (bess_mw * capex_bess_per_mw)) if bess_mw &gt; 0 else 0</p>
<p class="p2"><br></p>
<p class="p1"># Wind</p>
<p class="p1">wind_mw = selected['CD']</p>
<p class="p1">wind_saving = wind_mw * wind_gen_per_mw * selected['Tariff']</p>
<p class="p1">wind_roi = wind_saving / (wind_mw * wind_capex_per_mw)</p>
<p class="p2"><br></p>
<p class="p1"># Display Metrics</p>
<p class="p1">st.title(f"📊 Extension Opportunities for {client}")</p>
<p class="p1">st.metric("Installed Capacity (MW)", f"{selected['Installed_MW']:.2f}")</p>
<p class="p1">st.metric("Monthly Usage (lakh kWh)", f"{selected['Monthly_kWh'] / 100000:.2f}")</p>
<p class="p1">st.metric("Annual Usage (Cr kWh)", f"{selected['Annual_kWh'] / 1e7:.2f}")</p>
<p class="p2"><br></p>
<p class="p1"># ROI Chart</p>
<p class="p1">roi_data = pd.DataFrame({</p>
<p class="p1"><span class="Apple-converted-space">    </span>'Option': ['Solar to CD', 'Solar to SL', f'BESS ({bess_pct}%)', 'Wind'],</p>
<p class="p1"><span class="Apple-converted-space">    </span>'ROI (%)': [solar_to_cd_roi * 100, solar_to_sl_roi * 100, bess_roi * 100, wind_roi * 100]</p>
<p class="p1">})</p>
<p class="p2"><br></p>
<p class="p1">st.subheader("📈 ROI Comparison")</p>
<p class="p1">chart = alt.Chart(roi_data).mark_bar().encode(</p>
<p class="p1"><span class="Apple-converted-space">    </span>x=alt.X('Option', sort=None),</p>
<p class="p1"><span class="Apple-converted-space">    </span>y='ROI (%)',</p>
<p class="p1"><span class="Apple-converted-space">    </span>color='Option'</p>
<p class="p1">).properties(height=400)</p>
<p class="p2"><br></p>
<p class="p1">st.altair_chart(chart, use_container_width=True)</p>
<p class="p2"><br></p>
<p class="p1"># Recommendation</p>
<p class="p1">best_option = roi_data.loc[roi_data['ROI (%)'].idxmax()]</p>
<p class="p1">st.success(f"Recommended Option: {best_option['Option']} (ROI: {best_option['ROI (%)']:.2f}%)")</p>
</body>
</html>
    ( > U d � � � � � �                           *