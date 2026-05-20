"""
OSA Analytics — On-Shelf Availability visuals
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import GLOBAL_CSS, THEME, get_theme, get_plotly_layout, osa_color, process_osa, load_pressure_targets, inject_sidebar_toggle, inject_theme_toggle
from auth import require_login, logout

st.set_page_config(page_title="OSA Analytics", page_icon="🟢", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

# ── Logout button ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_logout_osa", use_container_width=True):
        logout()


# ─── helpers ────────────────────────────────────────────────────────────────

def apply_layout(fig, title="", height=380):
    T = get_theme()
    layout = get_plotly_layout(height=height, title=title)
    # Also fix hover label and colorbar text for dark/light mode
    layout["hoverlabel"] = dict(
        bgcolor=T["surface"],
        font_color=T["text"],
        bordercolor=T["border"],
    )
    fig.update_layout(**layout)
    # Make sure all axes use correct text color (multi-axis charts)
    fig.update_xaxes(tickfont_color=T["text"], title_font_color=T["text"])
    fig.update_yaxes(tickfont_color=T["text"], title_font_color=T["text"])
    return fig

def color_bars(values):
    return [osa_color(v) for v in values]

# ─── page header ────────────────────────────────────────────────────────────

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>🟢 OSA Analytics</h1>
  <p>On-Shelf Availability — upload your OSA <code>.xls</code> file to explore performance.</p>
</div>
""", unsafe_allow_html=True)

# ─── file upload ────────────────────────────────────────────────────────────

col_main, col_pt = st.columns([3, 2])

with col_main:
    st.markdown('<div class="sec-label">OSA Data File</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload OSA file (.xls, .xlsx or .csv)", type=["xls","xlsx","csv"],
                                 label_visibility="collapsed")

with col_pt:
    st.markdown('<div class="sec-label">Pressure Targets <span style="color:var(--muted);font-size:10px">(optional override)</span></div>', unsafe_allow_html=True)
    pt_file = st.file_uploader("Upload Pressure Targets .xlsx (optional)", type=["xlsx"],
                                key="pt_upload", label_visibility="collapsed")

if uploaded is None:
    st.markdown('<div class="result-box info">⬆️  Upload your OSA data file (CSV or Excel) to see the analytics dashboard. Optionally upload a Pressure Targets file to override SKU targets.</div>',
                unsafe_allow_html=True)
    st.stop()

# ─── process ────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Processing data…")
def load_osa(file_bytes, filename="", pt_bytes=None):
    from engines import load_pressure_targets
    targets = load_pressure_targets(pt_bytes) if pt_bytes else None
    return process_osa(file_bytes, filename, targets_override=targets)

pt_bytes = pt_file.read() if pt_file else None

try:
    df = load_osa(uploaded.read(), uploaded.name, pt_bytes)
except Exception as e:
    st.markdown(f'<div class="result-box error">❌ {e}</div>', unsafe_allow_html=True)
    st.stop()

if pt_file:
    from engines import load_pressure_targets
    n_targets = len(load_pressure_targets(pt_file.getvalue()))
    st.markdown(f'<div class="result-box success">✅ Pressure targets loaded — {n_targets} SKU overrides applied.</div>', unsafe_allow_html=True)

months  = sorted(df['Month'].dropna().unique(), key=lambda m: pd.to_datetime(m, format='%B').month)
brands  = sorted(df['BRAND NAME'].dropna().unique())
cats    = sorted(df['PRODUCT CATEGORY'].dropna().unique())
accts   = sorted(df['ACCOUNT'].dropna().unique())

# ─── filters ────────────────────────────────────────────────────────────────

st.markdown('<div class="sec-label">Filters</div>', unsafe_allow_html=True)
fc1, fc2, fc3 = st.columns(3)
sel_months = fc1.multiselect("Month", months, default=months)
sel_brands = fc2.multiselect("Brand", brands, default=brands)
sel_cats   = fc3.multiselect("Category", cats, default=cats)

dff = df[
    df['Month'].isin(sel_months) &
    df['BRAND NAME'].isin(sel_brands) &
    df['PRODUCT CATEGORY'].isin(sel_cats)
].copy()

if dff.empty:
    st.markdown('<div class="result-box error">No data matches the selected filters.</div>', unsafe_allow_html=True)
    st.stop()

# ─── KPI cards ──────────────────────────────────────────────────────────────

overall_osa   = dff['OSA'].mean()
n_brands      = dff['BRAND NAME'].nunique()
n_accts       = dff['ACCOUNT'].nunique()
n_skus        = len(dff)
pct_on_target = (dff['OSA'] == 100).mean() * 100

T = get_theme()  # live theme — respects light/dark mode
kpi_color = osa_color(overall_osa)
st.markdown(f"""
<div class="metric-row">
  <div class="metric-card">
    <div class="m-label">Overall OSA</div>
    <div class="m-value" style="color:{kpi_color}">{overall_osa:.1f}%</div>
    <div class="m-sub">mean across selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">On Target SKUs</div>
    <div class="m-value" style="color:{T['green']}">{pct_on_target:.1f}%</div>
    <div class="m-sub">qty ≥ pressure target</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Brands</div>
    <div class="m-value">{n_brands}</div>
    <div class="m-sub">in selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Accounts</div>
    <div class="m-value">{n_accts}</div>
    <div class="m-sub">outlet groups</div>
  </div>
  <div class="metric-card">
    <div class="m-label">SKU Records</div>
    <div class="m-value">{n_skus:,}</div>
    <div class="m-sub">after cleaning</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ─── Row 1: Brand OSA bar  +  Category OSA bar ──────────────────────────────

r1c1, r1c2 = st.columns(2)

with r1c1:
    brand_osa = (dff.groupby('BRAND NAME')['OSA'].mean().round(1)
                   .sort_values(ascending=True).reset_index())
    brand_osa.columns = ['Brand','OSA']
    fig = go.Figure(go.Bar(
        x=brand_osa['OSA'], y=brand_osa['Brand'],
        orientation='h',
        marker_color=color_bars(brand_osa['OSA']),
        text=brand_osa['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    fig.add_vline(x=95, line_dash="dot", line_color=T['green'],
                  annotation_text="95% target", annotation_font_color=T['green'],
                  annotation_font_size=10)
    apply_layout(fig, "OSA by Brand", height=max(320, len(brand_osa)*32))
    fig.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig, use_container_width=True)

with r1c2:
    cat_osa = (dff.groupby('PRODUCT CATEGORY')['OSA'].mean().round(1)
                 .sort_values(ascending=True).reset_index())
    cat_osa.columns = ['Category','OSA']
    fig2 = go.Figure(go.Bar(
        x=cat_osa['OSA'], y=cat_osa['Category'],
        orientation='h',
        marker_color=color_bars(cat_osa['OSA']),
        text=cat_osa['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    fig2.add_vline(x=95, line_dash="dot", line_color=T['green'],
                   annotation_text="95% target", annotation_font_color=T['green'],
                   annotation_font_size=10)
    apply_layout(fig2, "OSA by Category", height=max(320, len(cat_osa)*32))
    fig2.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig2, use_container_width=True)

# ─── Row 2: Weekly trend by brand (line chart) ──────────────────────────────

st.markdown('<div class="sec-label">Weekly Trend by Brand</div>', unsafe_allow_html=True)

week_brand = (dff.groupby(['WEEK_LABEL','BRAND NAME'])['OSA']
               .mean().round(1).reset_index())
week_brand.columns = ['Week','Brand','OSA']
week_order = [w for w in [f'Week {i}' for i in range(1,7)] if w in week_brand['Week'].values]

palette = [T['blue'], T['green'], T['amber'],
           T['purple'], T['cyan'], T['red'],
           '#fb923c','#e879f9','#a3e635','#38bdf8']

fig3 = go.Figure()
for i, brand in enumerate(week_brand['Brand'].unique()):
    sub = week_brand[week_brand['Brand']==brand].set_index('Week').reindex(week_order).reset_index()
    fig3.add_trace(go.Scatter(
        x=sub['Week'], y=sub['OSA'],
        name=brand, mode='lines+markers',
        line=dict(color=palette[i % len(palette)], width=2),
        marker=dict(size=7),
        hovertemplate=f"<b>{brand}</b><br>%{{x}}: %{{y:.1f}}%<extra></extra>",
    ))
fig3.add_hline(y=95, line_dash="dot", line_color=T['green'],
               annotation_text="95%", annotation_font_color=T['green'])
fig3.add_hrect(y0=80, y1=95, fillcolor=T['amber'], opacity=0.05, line_width=0)
apply_layout(fig3, "", height=360)
fig3.update_yaxes(range=[0,110], ticksuffix='%')
st.plotly_chart(fig3, use_container_width=True)

# ─── Row 3: Account heatmap ──────────────────────────────────────────────────

st.markdown('<div class="sec-label">Account × Brand Heatmap</div>', unsafe_allow_html=True)

heat = (dff.groupby(['ACCOUNT','BRAND NAME'])['OSA']
          .mean().round(1).reset_index())
heat_pivot = heat.pivot(index='ACCOUNT', columns='BRAND NAME', values='OSA').fillna(np.nan)

# Custom discrete colorscale: red / amber / green
colorscale = [
    [0.0,  '#f87171'],
    [0.8,  '#fbbf24'],
    [0.95, '#34c97b'],
    [1.0,  '#34c97b'],
]

fig4 = go.Figure(go.Heatmap(
    z=heat_pivot.values,
    x=list(heat_pivot.columns),
    y=list(heat_pivot.index),
    colorscale=colorscale,
    zmin=0, zmax=100,
    text=np.where(np.isnan(heat_pivot.values), '',
                  np.round(heat_pivot.values, 0).astype(int).astype(str) + '%'),
    texttemplate='%{text}',
    textfont=dict(size=9, color='#0f1117'),
    hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>OSA: %{z:.1f}%<extra></extra>',
    showscale=True,
    colorbar=dict(
        ticksuffix='%', tickfont=dict(color=T['text']),
        outlinecolor=T['border'], outlinewidth=1,
    ),
))
h4 = max(400, len(heat_pivot)*28)
apply_layout(fig4, "Account × Brand OSA (%)", height=h4)
fig4.update_xaxes(tickangle=-35, tickfont=dict(size=10))
fig4.update_yaxes(tickfont=dict(size=10))
st.plotly_chart(fig4, use_container_width=True)

# ─── Row 4: OSA distribution + month comparison ─────────────────────────────

r4c1, r4c2 = st.columns(2)

with r4c1:
    st.markdown('<div class="sec-label">OSA Score Distribution</div>', unsafe_allow_html=True)
    on  = (dff['OSA'] == 100).sum()
    off = (dff['OSA'] == 0).sum()
    fig5 = go.Figure(go.Pie(
        labels=['On Target (100%)', 'Off Target (0%)'],
        values=[on, off],
        hole=0.55,
        marker_colors=[T['green'], T['red']],
        textfont=dict(size=12, color=T['text']),
        hovertemplate='%{label}: %{value:,} SKUs (%{percent})<extra></extra>',
    ))
    layout5 = get_plotly_layout(height=300)
    layout5['annotations'] = [dict(text=f"{on/(on+off)*100:.1f}%", x=0.5, y=0.5,
                             font=dict(size=22, color=T['green']), showarrow=False)]
    fig5.update_layout(**layout5)
    st.plotly_chart(fig5, use_container_width=True)

with r4c2:
    st.markdown('<div class="sec-label">OSA by Month</div>', unsafe_allow_html=True)
    if len(sel_months) > 1:
        month_osa = (dff.groupby('Month')['OSA'].mean().round(1)
                       .reindex(months).dropna().reset_index())
        month_osa.columns = ['Month','OSA']
        fig6 = go.Figure(go.Bar(
            x=month_osa['Month'], y=month_osa['OSA'],
            marker_color=color_bars(month_osa['OSA']),
            text=month_osa['OSA'].map(lambda v: f"{v:.1f}%"),
            textposition='outside', textfont=dict(size=11, color=T['text']),
        ))
        fig6.add_hline(y=95, line_dash="dot", line_color=T['green'])
        apply_layout(fig6, "", height=300)
        fig6.update_yaxes(range=[0,115], ticksuffix='%')
        st.plotly_chart(fig6, use_container_width=True)
    else:
        st.markdown('<div class="result-box info">Select multiple months to compare.</div>',
                    unsafe_allow_html=True)

# ─── Row 5: Top / bottom accounts ────────────────────────────────────────────

st.markdown('<div class="sec-label">Account Performance Ranking</div>', unsafe_allow_html=True)

acct_osa = (dff.groupby('ACCOUNT')['OSA'].mean().round(1)
              .sort_values(ascending=False).reset_index())
acct_osa.columns = ['Account','OSA']

ra1, ra2 = st.columns(2)

with ra1:
    top10 = acct_osa.head(10).sort_values('OSA')
    fig7 = go.Figure(go.Bar(
        x=top10['OSA'], y=top10['Account'], orientation='h',
        marker_color=[T['green']]*len(top10),
        text=top10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    apply_layout(fig7, "🏆 Top 10 Accounts", height=340)
    fig7.update_xaxes(range=[0,115], ticksuffix='%')
    st.plotly_chart(fig7, use_container_width=True)

with ra2:
    bot10 = acct_osa.tail(10).sort_values('OSA')
    fig8 = go.Figure(go.Bar(
        x=bot10['OSA'], y=bot10['Account'], orientation='h',
        marker_color=color_bars(bot10['OSA']),
        text=bot10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    apply_layout(fig8, "⚠️ Bottom 10 Accounts", height=340)
    fig8.update_xaxes(range=[0,115], ticksuffix='%')
    st.plotly_chart(fig8, use_container_width=True)

st.markdown('<div class="footer">OSA Analytics — Uganda Field Data</div>', unsafe_allow_html=True)
