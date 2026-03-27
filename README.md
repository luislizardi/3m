import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
from datetime import datetime

# Read the Excel file
df = pd.read_excel('Mercury 1.0.xlsx', sheet_name='Sheet1')

# Convert date column to datetime
df['Date'] = pd.to_datetime(df['Date'])

# Sort by date
df = df.sort_values('Date')

# Define column mappings for easier reference
columns = {
    'Date': 'Date',
    'S&P 500': 'S&P 500',
    'Nasdaq 100': 'Nasdaq 100',
    'Dow Jones': 'Dow Jones',
    'MSCI world': 'MSCI World',
    'US 10YT': 'US 10YT',
    'US2YT (%)': 'US 2YT',
    '10-2 years spread': '10-2 Spread',
    'VIX riesgo de los mercados': 'VIX',
    'DXY index': 'DXY',
    'USD/EUR (': 'USD/EUR',
    'Petroleo WTI $/Bar': 'Oil WTI',
    'Oro $/oz': 'Gold',
    'Bs/Usd ( tasa oficial)': 'Bs/USD',
    'SIlver USD/Oz': 'Silver',
}

# Rename columns for easier access
df.columns = [columns.get(col, col) for col in df.columns]

# Create the dashboard with subplots
fig = make_subplots(
    rows=5, cols=2,
    subplot_titles=('S&P 500', 'US Treasury Yields', 
                    'DXY Index', 'VIX Volatility',
                    'Oil, Gold & Silver', '10-2 Year Spread',
                    'USD/EUR Exchange Rate', 'MSCI World',
                    'Dow Jones', 'Nasdaq 100'),
    vertical_spacing=0.08,
    horizontal_spacing=0.12,
    row_heights=[0.2, 0.2, 0.2, 0.2, 0.2]
)

# Row 1, Col 1: S&P 500
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['S&P 500'], mode='lines', name='S&P 500',
               line=dict(color='#1f77b4', width=2)),
    row=1, col=1
)

# Row 1, Col 2: US Treasury Yields (10yr and 2yr)
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['US 10YT'], mode='lines', name='US 10YR',
               line=dict(color='#ff7f0e', width=2)),
    row=1, col=2
)
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['US 2YT'], mode='lines', name='US 2YR',
               line=dict(color='#2ca02c', width=2)),
    row=1, col=2
)

# Row 2, Col 1: DXY Index
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['DXY'], mode='lines', name='DXY',
               line=dict(color='#d62728', width=2)),
    row=2, col=1
)

# Row 2, Col 2: VIX
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['VIX'], mode='lines', name='VIX',
               line=dict(color='#9467bd', width=2)),
    row=2, col=2
)

# Row 3, Col 1: Oil, Gold & Silver
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['Oil WTI'], mode='lines', name='Oil WTI',
               line=dict(color='#8c564b', width=2)),
    row=3, col=1
)
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['Gold'], mode='lines', name='Gold',
               line=dict(color='#e377c2', width=2)),
    row=3, col=1
)
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['Silver'], mode='lines', name='Silver',
               line=dict(color='#7f7f7f', width=2)),
    row=3, col=1
)

# Row 3, Col 2: 10-2 Year Spread
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['10-2 Spread'], mode='lines', name='10-2 Spread',
               line=dict(color='#bcbd22', width=2, dash='dash')),
    row=3, col=2
)

# Row 4, Col 1: USD/EUR
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['USD/EUR'], mode='lines', name='USD/EUR',
               line=dict(color='#17becf', width=2)),
    row=4, col=1
)

# Row 4, Col 2: MSCI World
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['MSCI World'], mode='lines', name='MSCI World',
               line=dict(color='#ff9896', width=2)),
    row=4, col=2
)

# Row 5, Col 1: Dow Jones
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['Dow Jones'], mode='lines', name='Dow Jones',
               line=dict(color='#c5b0d5', width=2)),
    row=5, col=1
)

# Row 5, Col 2: Nasdaq 100
fig.add_trace(
    go.Scatter(x=df['Date'], y=df['Nasdaq 100'], mode='lines', name='Nasdaq 100',
               line=dict(color='#f7b6d2', width=2)),
    row=5, col=2
)

# Update layout
fig.update_layout(
    title_text="Financial Markets Dashboard",
    title_font_size=24,
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    hovermode='x unified',
    template='plotly_white',
    height=1200,
    width=1400
)

# Update axes labels
fig.update_xaxes(title_text="Date", row=5, col=1)
fig.update_xaxes(title_text="Date", row=5, col=2)

# Add range slider at the bottom
fig.update_xaxes(rangeslider_visible=True, row=5, col=1)
fig.update_xaxes(rangeslider_visible=True, row=5, col=2)

# Show the dashboard
fig.show()

# Create a summary statistics table
print("\n" + "="*80)
print("SUMMARY STATISTICS")
print("="*80)

summary_stats = df[['S&P 500', 'US 10YT', 'DXY', 'VIX', 'Gold', 'Oil WTI']].describe()
print(summary_stats)

# Create correlation matrix
print("\n" + "="*80)
print("CORRELATION MATRIX (Selected Variables)")
print("="*80)

correlation_vars = ['S&P 500', 'US 10YT', 'DXY', 'VIX', 'Gold', 'Oil WTI', '10-2 Spread']
correlation_matrix = df[correlation_vars].corr()
print(correlation_matrix)

# Calculate key performance metrics
print("\n" + "="*80)
print("PERFORMANCE METRICS (Period to Date)")
print("="*80)

latest_data = df.iloc[-1]
first_data = df.iloc[0]

print(f"Period: {first_data['Date'].strftime('%Y-%m-%d')} to {latest_data['Date'].strftime('%Y-%m-%d')}")
print(f"\nS&P 500: {first_data['S&P 500']:.2f} → {latest_data['S&P 500']:.2f} "
      f"({((latest_data['S&P 500']/first_data['S&P 500'])-1)*100:.2f}%)")
print(f"US 10YR: {first_data['US 10YT']:.2f}% → {latest_data['US 10YT']:.2f}% "
      f"({latest_data['US 10YT'] - first_data['US 10YT']:.2f}bps)")
print(f"DXY: {first_data['DXY']:.2f} → {latest_data['DXY']:.2f} "
      f"({((latest_data['DXY']/first_data['DXY'])-1)*100:.2f}%)")
print(f"Gold: ${first_data['Gold']:.2f} → ${latest_data['Gold']:.2f} "
      f"({((latest_data['Gold']/first_data['Gold'])-1)*100:.2f}%)")
print(f"Oil: ${first_data['Oil WTI']:.2f} → ${latest_data['Oil WTI']:.2f} "
      f"({((latest_data['Oil WTI']/first_data['Oil WTI'])-1)*100:.2f}%)")
