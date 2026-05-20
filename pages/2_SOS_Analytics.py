"""
SOS Analytics — Share of Shelf visuals
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import GLOBAL_CSS, THEME, get_theme, get_plotly_layout, sos_color, process_sos, inject_sidebar_toggle, inject_theme_toggle
from auth import require_login, logout

st.set_page_config(page_title="SOS Analytics", page_icon="📦", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

# ── Logout button ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_logout_sos", use_container_width=True):
        logout()


# ─── helper ─────────────────────────────────────────────────────────────────

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
    return [sos_color(v) for v in values]

# ─── page header ────────────────────────────────────────────────────────────

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>📦 SOS Analytics</h1>
  <p>Share of Shelf — upload your SOS <code>.xlsx</code> file to explore performance.</p>
</div>
""", unsafe_allow_html=True)

# ─── file upload ────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Upload SOS .xlsx file", type=["xlsx"],
                             label_visibility="collapsed")

if uploaded is None:
    st.markdown('<div class="result-box info">⬆️  Upload your SOS <code>.xlsx</code> file above to see the analytics dashboard.</div>',
                unsafe_allow_html=True)
    st.stop()

# ─── process ────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Processing data…")
def load_sos(file_bytes):
    return process_sos(file_bytes)

try:
    df = load_sos(uploaded.read())
except Exception as e:
    st.markdown(f'<div class="result-box error">❌ {e}</div>', unsafe_allow_html=True)
    st.stop()

months  = (df[['MONTH','MONTH_NUM']].drop_duplicates()
           .sort_values('MONTH_NUM')['MONTH'].tolist())
brands  = sorted(df['PRODUCT_NAME'].dropna().unique())
accts   = sorted(df['ACCOUNT'].dropna().unique())

# ─── filters ────────────────────────────────────────────────────────────────

st.markdown('<div class="sec-label">Filters</div>', unsafe_allow_html=True)
fc1, fc2 = st.columns(2)
sel_months = fc1.multiselect("Month", months, default=months)
sel_accts  = fc2.multiselect("Account", accts, default=accts)

dff = df[df['MONTH'].isin(sel_months) & df['ACCOUNT'].isin(sel_accts)].copy()

if dff.empty:
    st.markdown('<div class="result-box error">No data matches the selected filters.</div>', unsafe_allow_html=True)
    st.stop()

# ─── KPI cards ──────────────────────────────────────────────────────────────

overall_sos  = dff['FACINGS SOS%'].mean()
n_brands     = dff['PRODUCT_NAME'].nunique()
n_accts_sel  = dff['ACCOUNT'].nunique()
pct_present  = (dff['FACINGS SOS%'] > 0).mean() * 100
avg_pos      = dff[dff['POSITION'] > 0]['POSITION'].mean()

T = get_theme()  # live theme — respects light/dark mode
kc = sos_color(overall_sos)
st.markdown(f"""
<div class="metric-row">
  <div class="metric-card">
    <div class="m-label">Overall SOS%</div>
    <div class="m-value" style="color:{kc}">{overall_sos:.1f}%</div>
    <div class="m-sub">mean across selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Present on Shelf</div>
    <div class="m-value" style="color:{T['green']}">{pct_present:.1f}%</div>
    <div class="m-sub">SOS% &gt; 0</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Avg Shelf Position</div>
    <div class="m-value">{avg_pos:.1f}</div>
    <div class="m-sub">1 = best / eye level</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Brands</div>
    <div class="m-value">{n_brands}</div>
    <div class="m-sub">in selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Accounts</div>
    <div class="m-value">{n_accts_sel}</div>
    <div class="m-sub">key account groups</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ─── Row 1: SOS by Brand bar + SOS by Account bar ───────────────────────────

r1c1, r1c2 = st.columns(2)

with r1c1:
    brand_sos = (dff.groupby('PRODUCT_NAME')['FACINGS SOS%']
                    .mean().round(1).sort_values(ascending=True).reset_index())
    brand_sos.columns = ['Brand','SOS']
    fig1 = go.Figure(go.Bar(
        x=brand_sos['SOS'], y=brand_sos['Brand'], orientation='h',
        marker_color=color_bars(brand_sos['SOS']),
        text=brand_sos['SOS'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    fig1.add_vline(x=20, line_dash="dot", line_color=T['green'],
                   annotation_text="20% target", annotation_font_color=T['green'],
                   annotation_font_size=10)
    apply_layout(fig1, "SOS% by Brand", height=max(300, len(brand_sos)*32))
    fig1.update_xaxes(range=[0, max(brand_sos['SOS'].max()*1.25, 28)], ticksuffix='%')
    st.plotly_chart(fig1, use_container_width=True)

with r1c2:
    acct_sos = (dff.groupby('ACCOUNT')['FACINGS SOS%']
                   .mean().round(1).sort_values(ascending=True).reset_index())
    acct_sos.columns = ['Account','SOS']
    fig2 = go.Figure(go.Bar(
        x=acct_sos['SOS'], y=acct_sos['Account'], orientation='h',
        marker_color=color_bars(acct_sos['SOS']),
        text=acct_sos['SOS'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=T['text']),
    ))
    fig2.add_vline(x=20, line_dash="dot", line_color=T['green'],
                   annotation_text="20% target", annotation_font_color=T['green'],
                   annotation_font_size=10)
    apply_layout(fig2, "SOS% by Account", height=max(300, len(acct_sos)*32))
    fig2.update_xaxes(range=[0, max(acct_sos['SOS'].max()*1.25, 28)], ticksuffix='%')
    st.plotly_chart(fig2, use_container_width=True)

# ─── Row 2: Account × Brand SOS heatmap ─────────────────────────────────────

st.markdown('<div class="sec-label">Account × Brand SOS% Heatmap</div>', unsafe_allow_html=True)

heat = (dff.groupby(['ACCOUNT','PRODUCT_NAME'])['FACINGS SOS%']
           .mean().round(1).reset_index())
heat_pivot = heat.pivot(index='ACCOUNT', columns='PRODUCT_NAME', values='FACINGS SOS%').fillna(np.nan)

colorscale_sos = [
    [0.0,  '#f87171'],
    [0.01, '#fbbf24'],
    [0.2,  '#34c97b'],
    [1.0,  '#34c97b'],
]

fig3 = go.Figure(go.Heatmap(
    z=heat_pivot.values,
    x=list(heat_pivot.columns),
    y=list(heat_pivot.index),
    colorscale=colorscale_sos,
    zmin=0, zmax=max(50, float(np.nanmax(heat_pivot.values)) if not np.all(np.isnan(heat_pivot.values)) else 50),
    text=np.where(np.isnan(heat_pivot.values), '',
                  np.round(heat_pivot.values, 0).astype(int).astype(str) + '%'),
    texttemplate='%{text}',
    textfont=dict(size=9, color='#0f1117'),
    hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>SOS: %{z:.1f}%<extra></extra>',
    showscale=True,
    colorbar=dict(ticksuffix='%', tickfont=dict(color=T['text']),
                  outlinecolor=T['border'], outlinewidth=1),
))
apply_layout(fig3, "Account × Brand SOS (%)", height=max(380, len(heat_pivot)*28))
fig3.update_xaxes(tickangle=-35, tickfont=dict(size=10))
fig3.update_yaxes(tickfont=dict(size=10))
st.plotly_chart(fig3, use_container_width=True)

# ─── Row 3: SOS% by month trend (line)  +  Shelf Position bubble ────────────

r3c1, r3c2 = st.columns(2)

with r3c1:
    st.markdown('<div class="sec-label">SOS% Trend by Brand</div>', unsafe_allow_html=True)
    if len(sel_months) > 1:
        trend = (dff.groupby(['MONTH','PRODUCT_NAME'])['FACINGS SOS%']
                    .mean().round(1).reset_index())
        trend.columns = ['Month','Brand','SOS']
        month_order = [m for m in months if m in trend['Month'].values]

        palette = [T['blue'], T['green'], T['amber'],
                   T['purple'], T['cyan'], T['red'],
                   '#fb923c','#e879f9','#a3e635','#38bdf8']

        fig4 = go.Figure()
        for i, brand in enumerate(trend['Brand'].unique()):
            sub = trend[trend['Brand']==brand].set_index('Month').reindex(month_order).reset_index()
            fig4.add_trace(go.Scatter(
                x=sub['Month'], y=sub['SOS'],
                name=brand, mode='lines+markers',
                line=dict(color=palette[i % len(palette)], width=2),
                marker=dict(size=7),
                hovertemplate=f"<b>{brand}</b><br>%{{x}}: %{{y:.1f}}%<extra></extra>",
            ))
        fig4.add_hline(y=20, line_dash="dot", line_color=T['green'],
                       annotation_text="20% target", annotation_font_color=T['green'])
        apply_layout(fig4, "", height=340)
        fig4.update_yaxes(ticksuffix='%')
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.markdown('<div class="result-box info">Select multiple months to see trend.</div>',
                    unsafe_allow_html=True)

with r3c2:
    st.markdown('<div class="sec-label">Shelf Position by Brand & Account</div>', unsafe_allow_html=True)
    pos_data = (dff[dff['POSITION'] > 0]
                .groupby(['PRODUCT_NAME','ACCOUNT'])['POSITION']
                .mean().round(1).reset_index())
    pos_data.columns = ['Brand','Account','Position']

    if not pos_data.empty:
        # Bubble: x=brand, y=account, size=inverse position (better pos = bigger bubble)
        pos_data['Size'] = (10 / pos_data['Position']).clip(upper=10) * 8

        palette2 = [T['blue'], T['green'], T['amber'],
                    T['purple'], T['cyan'], T['red'],
                    '#fb923c','#e879f9']

        fig5 = go.Figure()
        for i, brand in enumerate(pos_data['Brand'].unique()):
            sub = pos_data[pos_data['Brand'] == brand]
            fig5.add_trace(go.Scatter(
                x=[brand]*len(sub),
                y=sub['Account'],
                mode='markers',
                name=brand,
                marker=dict(
                    size=sub['Size'],
                    color=palette2[i % len(palette2)],
                    opacity=0.8,
                    line=dict(width=1, color=T['border'])
                ),
                text=sub['Position'].map(lambda v: f"Pos {v:.1f}"),
                hovertemplate=f"<b>{brand}</b><br>%{{y}}<br>Avg position: %{{text}}<extra></extra>",
            ))
        apply_layout(fig5, "Avg Shelf Position (bubble size = better position)", height=340)
        fig5.update_yaxes(autorange='reversed')
        fig5.update_xaxes(tickangle=-35)
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.markdown('<div class="result-box info">No position data available.</div>',
                    unsafe_allow_html=True)

# ─── Row 4: Presence vs Absence stacked bar per account ─────────────────────

st.markdown('<div class="sec-label">Brand Presence per Account (% of records)</div>',
            unsafe_allow_html=True)

presence = (dff.groupby(['ACCOUNT','PRODUCT_NAME'])
              .apply(lambda g: (g['FACINGS SOS%'] > 0).mean() * 100)
              .round(1).reset_index())
presence.columns = ['Account','Brand','PresencePct']
piv = presence.pivot(index='Account', columns='Brand', values='PresencePct').fillna(0)

palette3 = [T['blue'], T['green'], T['amber'],
            T['purple'], T['cyan'], T['red'],
            '#fb923c','#e879f9','#a3e635','#38bdf8']

fig6 = go.Figure()
for i, brand in enumerate(piv.columns):
    fig6.add_trace(go.Bar(
        name=brand, x=piv.index, y=piv[brand],
        marker_color=palette3[i % len(palette3)],
        hovertemplate=f"<b>{brand}</b><br>%{{x}}<br>Present: %{{y:.1f}}%<extra></extra>",
    ))
fig6.update_layout(barmode='group')
apply_layout(fig6, "", height=380)
fig6.update_yaxes(range=[0,115], ticksuffix='%')
fig6.update_xaxes(tickangle=-35)
st.plotly_chart(fig6, use_container_width=True)

st.markdown('<div class="footer">SOS Analytics — Uganda Field Data</div>', unsafe_allow_html=True)
