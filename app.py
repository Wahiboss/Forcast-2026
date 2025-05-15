import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from pandas import ExcelWriter

st.set_page_config(page_title="2026 Forecast Simulator", layout="wide")

# --- RENAME COLUMNS MAP ---
month_map = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
    '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}

# --- LOAD HISTORICAL FILE ---
try:
    df_hist = pd.read_excel("Classeur2.xlsx")
    df_hist.rename(columns=month_map, inplace=True)
    df_hist['Total'] = df_hist[month_map.values()].sum(axis=1)
    df_hist['YearClean'] = df_hist['Year'].astype(str).str.extract(r'(\d{4})').astype(int)

    product_trends = df_hist.groupby(['YearClean', 'Product Line2'])['Total'].sum().reset_index()
    pivot = product_trends.pivot(index='Product Line2', columns='YearClean', values='Total').fillna(0)

    if 2024 in pivot.columns and 2025 in pivot.columns:
        pivot['Growth (%)'] = ((pivot[2025] - pivot[2024]) / pivot[2024].replace(0, 1)) * 100

        def classify(g):
            if g > 5:
                return '🟢 Evolving'
            elif g < -5:
                return '🔴 Declining'
            else:
                return '🟡 Stable'

        pivot['Trend'] = pivot['Growth (%)'].apply(classify)
        st.sidebar.header("📊 Product Line Trends (2024 → 2025)")
        st.sidebar.dataframe(pivot[['Growth (%)', 'Trend']])
    else:
        st.sidebar.warning("📌 Not enough data to compute 2024 → 2025 trend.")

except Exception as e:
    st.sidebar.error(f"❌ Failed to analyze Classeur2.xlsx: {e}")

# --- Load Current Forecast File ---
df_2025 = pd.read_excel("Classeur1.xlsx", sheet_name=0)
df_2025.rename(columns=month_map, inplace=True)

# --- Define markets and product lines ---
markets = df_2025['Sales District'].unique()
product_lines = df_2025['Product Line2'].unique()

# --- Evolution Inputs ---
evo_dict = {}
for market in markets:
    for line in product_lines:
        key = (market, line)
        trend = pivot['Trend'][line] if line in pivot.index and 'Trend' in pivot.columns else ''
        if 'Evolving' in trend:
            default_value = 8.0
        elif 'Declining' in trend:
            default_value = -6.0
        else:
            default_value = 0.0
        evo_dict[key] = default_value

# --- Seasonality Inputs ---
seasonality_defaults = {
    'EMFRAN':  [9.0, 8.0, 9.0, 8.0, 9.0, 8.0, 9.0, 6.0, 9.0, 9.0, 9.0, 7.0],
    'EMMORO':  [6.0, 8.0, 8.0, 8.0, 7.0, 6.0, 10.0, 8.0, 9.0, 9.0, 9.0, 14.0],
    'EMSPAI': [9.0, 5.0, 9.0, 5.0, 5.0, 7.0, 9.0, 7.0, 9.0, 8.0, 11.0, 16.0]
}
seasonality = seasonality_defaults.copy()
# Default fallback if market not defined
for market in markets:
    if market not in seasonality:
        seasonality[market] = [1/12.0]*12

# --- Evolution Inputs (with sliders for manual review) ---
st.sidebar.header("📈 Evolution per Product Line & Market")
evo_dict = {}
for market in markets:
    with st.sidebar.expander(f"📊 {market} - Configure Growth", expanded=False):
        for line in product_lines:
            key = (market, line)
            trend = pivot['Trend'][line] if line in pivot.index and 'Trend' in pivot.columns else ''
            if 'Evolving' in trend:
                default_value = 8.0
            elif 'Declining' in trend:
                default_value = -6.0
            else:
                default_value = 0.0
            evo_dict[key] = st.number_input(
                f"{line} ({trend})", value=default_value, step=0.5,
                help="Annual evolution as % vs 2025", key=f"evo_{market}_{line}"
            )
st.sidebar.header("📆 Monthly Seasonality (%)")
# use previously defined EMFRAN/EMMORO/EMSPAI seasonality_defaults
seasonality = seasonality_defaults.copy()
months = list(month_map.values())
for market in markets:
    with st.sidebar.expander(f"📅 {market} - Monthly Seasonality", expanded=False):
        if market in seasonality_defaults:
            base_values = seasonality_defaults[market]
            seasonality[market] = [st.number_input(f"{month}", min_value=0.0, max_value=100.0, value=base_values[i], step=0.1, key=f"{market}_{month}") for i, month in enumerate(months)]
        else:
            seasonality[market] = [st.number_input(f"{month}", min_value=0.0, max_value=100.0, value=8.3, step=0.1, key=f"{market}_{month}") for month in months]

# --- Normalize Seasonality ---
for market in markets:
    total = sum(seasonality[market])
    if total > 0:
        seasonality[market] = [s / total for s in seasonality[market]]

# --- Forecast Calculation ---
st.header("🔄 Forecasted Values for 2026")
forecast_rows = []
for _, row in df_2025.iterrows():
    market = row['Sales District']
    product_line = row.get('Product Line2', 'UNKNOWN')
    sku = row.get('Product', 'SKU_UNKNOWN')
    sales_2025 = row[months].sum()
    growth_factor = 1 + evo_dict.get((market, product_line), 0) / 100
    forecast_2026 = [round(sales_2025 * growth_factor * s, 2) for s in seasonality[market]]
    forecast_rows.append([market, product_line, sku, sales_2025, sum(forecast_2026)] + forecast_2026)

df_forecast = pd.DataFrame(forecast_rows, columns=['Market', 'Product Line', 'SKU', 'Total_2025', 'Total_2026'] + months)
st.dataframe(df_forecast)

# --- KPI Tiles ---
total_2025 = df_forecast['Total_2025'].sum()
total_2026 = df_forecast['Total_2026'].sum()
variation = ((total_2026 - total_2025) / total_2025) * 100 if total_2025 > 0 else 0
col1, col2, col3 = st.columns(3)
col1.metric("📦 Total 2025", f"{total_2025:,.0f}")
col2.metric("🚀 Total 2026", f"{total_2026:,.0f}")
col3.metric("📊 Variation %", f"{variation:.2f}%")

# --- Product Line Comparison Chart ---
st.header("📌 KPI Comparison: 2025 vs 2026")
selected_line = st.selectbox("Choose Product Line", options=df_forecast['Product Line'].unique())
df_selected = df_forecast[df_forecast['Product Line'] == selected_line]
chart_data = df_selected.groupby("Market")[['Total_2025', 'Total_2026']].sum().reset_index()
fig_selected = go.Figure()
fig_selected.add_trace(go.Bar(x=chart_data['Market'], y=chart_data['Total_2025'], name='2025', marker_color='lightblue'))
fig_selected.add_trace(go.Bar(x=chart_data['Market'], y=chart_data['Total_2026'], name='2026', marker_color='orange'))
fig_selected.update_layout(barmode='group', title=f"Sales by Market for {selected_line}", xaxis_title="Market", yaxis_title="Sales")
st.plotly_chart(fig_selected, use_container_width=True)

# --- Evolution by Product Line ---
st.header("📊 Product Line Evolution by Market")
selected_market = st.selectbox("Choose Market to View", options=markets)
df_market = df_forecast[df_forecast['Market'] == selected_market]
df_line = df_market.groupby("Product Line")[months].sum().T.reset_index()
df_line = pd.melt(df_line, id_vars='index', var_name='Product Line', value_name='Forecast')
df_line.rename(columns={'index': 'Month'}, inplace=True)
fig_line = px.line(df_line, x='Month', y='Forecast', color='Product Line', markers=True,
                   title=f"2026 Forecast by Product Line - {selected_market}")
st.plotly_chart(fig_line, use_container_width=True)

# --- Export Button ---
buffer = io.BytesIO()
with ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df_forecast.to_excel(writer, index=False, sheet_name='Forecast2026')
buffer.seek(0)
st.download_button(
    label="📥 Download Forecast 2026 as Excel",
    data=buffer,
    file_name="Forecast_2026.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
