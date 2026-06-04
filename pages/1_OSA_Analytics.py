"""OSA Analytics — On-Shelf Availability"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, get_theme, get_plotly_layout, osa_color,
    process_osa, load_pressure_targets, mhsku_load,
    UG_FOCUS_CATEGORIES, inject_sidebar_toggle, inject_theme_toggle,
    osa_data_save, osa_data_load, osa_data_clear,
)
from auth import require_login, logout

st.set_page_config(page_title="OSA Analytics", page_icon="🟢", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_osa", use_container_width=True):
        logout()

# ── helpers ───────────────────────────────────────────────────────────────────
def T(): return get_theme()

def apply_layout(fig, title="", height=380):
    th = T()
    fig.update_layout(**get_plotly_layout(height=height, title=title))
    fig.update_layout(hoverlabel=dict(bgcolor=th["surface"], font_color=th["text"], bordercolor=th["border"]))
    fig.update_xaxes(tickfont_color=th["text"], title_font_color=th["text"])
    fig.update_yaxes(tickfont_color=th["text"], title_font_color=th["text"])
    return fig

def color_bars(values): return [osa_color(v) for v in values]

PALETTE = lambda th: [th['blue'], th['green'], th['amber'], th['purple'],
                       th['cyan'], th['red'], '#fb923c', '#e879f9', '#a3e635', '#38bdf8']

# ── page header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>🟢 OSA Analytics</h1>
  <p>On-Shelf Availability — upload your OSA file or use previously saved data.</p>
</div>
""", unsafe_allow_html=True)

# ── persistent data banner + clear ───────────────────────────────────────────
saved_df, saved_meta = osa_data_load()

with st.expander("💾  Saved Data", expanded=(saved_df is None)):
    if saved_df is not None:
        st.markdown(
            f"<div style='font-size:13px;color:#34c97b;background:rgba(52,201,123,.08);"
            f"border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
            f"✅ <strong>{saved_meta.get('filename','file')}</strong> — "
            f"{saved_meta.get('rows',0):,} rows · saved {saved_meta.get('saved_at','')[:16].replace('T',' ')}"
            f" by <strong>{saved_meta.get('uploaded_by','?')}</strong></div>",
            unsafe_allow_html=True,
        )
        col_use, col_clear = st.columns([3, 1])
        use_saved = col_use.checkbox("Use saved data (skip upload)", value=True, key="osa_use_saved")
        if col_clear.button("🗑️ Clear saved data", key="osa_clear"):
            osa_data_clear()
            st.rerun()
    else:
        st.info("No saved data yet. Upload a file below and it will be saved automatically.")
        use_saved = False

# ── file upload ───────────────────────────────────────────────────────────────
col_main, col_pt = st.columns([3, 2])
with col_main:
    st.markdown('<div class="sec-label">OSA Data File</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload OSA file (.xls, .xlsx or .csv)",
                                 type=["xls","xlsx","csv"], label_visibility="collapsed")
with col_pt:
    st.markdown('<div class="sec-label">Pressure Targets <span style="color:var(--muted);font-size:10px">(optional)</span></div>', unsafe_allow_html=True)
    pt_file = st.file_uploader("Upload Pressure Targets .xlsx (optional)",
                                type=["xlsx"], key="pt_upload", label_visibility="collapsed")

# ── load data ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Processing OSA data…", max_entries=3, ttl=3600)
def load_osa(file_bytes: bytes, filename: str = "", pt_bytes: bytes | None = None):
    from engines import load_pressure_targets, mhsku_load
    if pt_bytes:
        targets = load_pressure_targets(pt_bytes)
    else:
        records, _, _ = mhsku_load()
        targets = {r["code"]: r["pressure_target"] for r in records} if records else None
    return process_osa(file_bytes, filename, targets_override=targets)

if uploaded is not None:
    pt_bytes = pt_file.read() if pt_file else None
    try:
        raw = uploaded.read()
        df = load_osa(raw, uploaded.name, pt_bytes)
        del raw
        osa_data_save(df, filename=uploaded.name,
                      uploaded_by=st.session_state.get("username","user"))
        st.success(f"✅ **{uploaded.name}** loaded and saved — {len(df):,} rows.")
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()
elif saved_df is not None and use_saved:
    df = saved_df
else:
    st.markdown('<div class="result-box info">⬆️  Upload your OSA data file to continue, or enable saved data above.</div>', unsafe_allow_html=True)
    st.stop()

if pt_file and uploaded is not None:
    from engines import load_pressure_targets as _lpt
    n_t = len(_lpt(pt_file.getvalue()))
    st.markdown(f'<div class="result-box success">✅ Pressure targets loaded — {n_t} SKU overrides applied.</div>', unsafe_allow_html=True)

# ── brand / MHSKU banner ──────────────────────────────────────────────────────
ug_brands_osa = sorted(df['BRAND NAME'].dropna().unique()) if 'BRAND NAME' in df.columns else []
st.markdown(
    "<div style='font-size:12px;color:#34c97b;background:rgba(52,201,123,.08);"
    "border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
    f"🇺🇬 <strong>Uganda brands only</strong> — {', '.join(ug_brands_osa)}. "
    f"<strong>{len(df):,}</strong> rows loaded.</div>",
    unsafe_allow_html=True,
)

# MHSKU toggle
mhsku_records, mhsku_ts, mhsku_by = mhsku_load()
_mhsku_codes = {r['code'] for r in mhsku_records} if mhsku_records else set()
_mhsku_skus  = {r['sku'].upper().strip() for r in mhsku_records} if mhsku_records else set()

mhsku_col, _ = st.columns([2, 5])
with mhsku_col:
    mhsku_only = st.toggle(
        f"📋 Must Have SKUs only ({len(_mhsku_codes)} SKUs)" if _mhsku_codes
        else "📋 Must Have SKUs (none loaded)",
        value=False, key="osa_mhsku_toggle", disabled=not bool(_mhsku_codes),
    )
if mhsku_only and _mhsku_codes:
    _before = len(df)
    if 'PRODUCT CODE' in df.columns:
        df = df[df['PRODUCT CODE'].isin(_mhsku_codes)].copy()
    elif 'DESCRIPTION' in df.columns:
        df = df[df['DESCRIPTION'].str.upper().str.strip().isin(_mhsku_skus)].copy()
    if not df.empty:
        st.markdown(
            f"<div style='font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);"
            f"border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:6px 14px;margin-bottom:8px'>"
            f"📋 <strong>MHSKU filter active</strong> — {len(df):,} rows ({_before-len(df):,} hidden)</div>",
            unsafe_allow_html=True,
        )

# ── filter lists ──────────────────────────────────────────────────────────────
months  = sorted(df['Month'].dropna().unique(), key=lambda m: pd.to_datetime(m, format='%B').month)
brands  = sorted(df['BRAND NAME'].dropna().unique())
cats_all = sorted(df['PRODUCT CATEGORY'].dropna().unique())
cats    = sorted([c for c in cats_all if c in UG_FOCUS_CATEGORIES]) or cats_all
accts   = sorted(df['ACCOUNT'].dropna().unique())
outlets = sorted(df['CUSTOMER NAME'].dropna().unique()) if 'CUSTOMER NAME' in df.columns else []

# ── filters ───────────────────────────────────────────────────────────────────
with st.expander("🔍  Filters", expanded=True):
    fc1, fc2, fc3 = st.columns(3)
    sel_months  = fc1.multiselect("Month",           months,  default=months,  key="osa_months")
    sel_brands  = fc2.multiselect("Brand",           brands,  default=brands,  key="osa_brands")
    sel_accts   = fc3.multiselect("Chain / Account", accts,   default=accts,   key="osa_accts")
    fc4, fc5, _ = st.columns(3)
    sel_cats    = fc4.multiselect("Category",        cats,    default=cats,    key="osa_cats")
    sel_outlets = fc5.multiselect("Outlet",          outlets, default=outlets, key="osa_outlets") if outlets else []

# ── apply filters ─────────────────────────────────────────────────────────────
dff = df[
    df['Month'].isin(sel_months) &
    df['BRAND NAME'].isin(sel_brands) &
    df['ACCOUNT'].isin(sel_accts) &
    df['PRODUCT CATEGORY'].isin(sel_cats)
].copy()
if sel_outlets and 'CUSTOMER NAME' in dff.columns:
    dff = dff[dff['CUSTOMER NAME'].isin(sel_outlets)]

if dff.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

th = T()

# ── KPIs ──────────────────────────────────────────────────────────────────────
overall_osa = dff['OSA'].mean()
kc = osa_color(overall_osa)
st.markdown(f"""
<div class="metric-row">
  <div class="metric-card">
    <div class="m-label">Overall OSA%</div>
    <div class="m-value" style="color:{kc}">{overall_osa:.1f}%</div>
    <div class="m-sub">mean across selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">On Target (100%)</div>
    <div class="m-value" style="color:{th['green']}">{(dff['OSA']==100).mean()*100:.1f}%</div>
    <div class="m-sub">of records</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Brands</div>
    <div class="m-value">{dff['BRAND NAME'].nunique()}</div>
    <div class="m-sub">in selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Accounts</div>
    <div class="m-value">{dff['ACCOUNT'].nunique()}</div>
    <div class="m-sub">key groups</div>
  </div>
  <div class="metric-card">
    <div class="m-label">SKUs</div>
    <div class="m-value">{dff['DESCRIPTION'].nunique() if 'DESCRIPTION' in dff.columns else '—'}</div>
    <div class="m-sub">unique products</div>
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ── Row 1: OSA by Brand  +  OSA by Category ──────────────────────────────────
r1a, r1b = st.columns(2)
with r1a:
    brand_osa = (dff.groupby('BRAND NAME')['OSA'].mean().round(1)
                    .sort_values().reset_index())
    brand_osa.columns = ['Brand','OSA']
    fig1 = go.Figure(go.Bar(
        x=brand_osa['OSA'], y=brand_osa['Brand'], orientation='h',
        marker_color=color_bars(brand_osa['OSA']),
        text=brand_osa['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    fig1.add_vline(x=95, line_dash="dot", line_color=th['green'],
                   annotation_text="95% target", annotation_font_color=th['green'],
                   annotation_font_size=10)
    apply_layout(fig1, "OSA% by Brand", height=max(300, len(brand_osa)*40))
    fig1.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig1, use_container_width=True)

with r1b:
    cat_osa = (dff.groupby('PRODUCT CATEGORY')['OSA'].mean().round(1)
                  .sort_values().reset_index())
    cat_osa.columns = ['Category','OSA']
    fig2 = go.Figure(go.Bar(
        x=cat_osa['OSA'], y=cat_osa['Category'], orientation='h',
        marker_color=color_bars(cat_osa['OSA']),
        text=cat_osa['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    fig2.add_vline(x=95, line_dash="dot", line_color=th['green'],
                   annotation_text="95% target", annotation_font_color=th['green'],
                   annotation_font_size=10)
    apply_layout(fig2, "OSA% by Category", height=max(300, len(cat_osa)*40))
    fig2.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig2, use_container_width=True)

# ── Row 2: Weekly trend — Brand  +  Weekly trend — Category ──────────────────
st.markdown("---")
st.markdown('<div class="sec-label">Weekly Trend</div>', unsafe_allow_html=True)
t2a, t2b = st.columns(2)

week_order = [f'Week {i}' for i in range(1, 7)]
pal = PALETTE(th)

with t2a:
    st.markdown('<div class="sec-label" style="font-size:11px;margin-bottom:6px">by Brand</div>', unsafe_allow_html=True)
    week_brand = (dff.groupby(['WEEK_LABEL','BRAND NAME'])['OSA']
                     .mean().round(1).reset_index())
    week_brand.columns = ['Week','Brand','OSA']
    wo = [w for w in week_order if w in week_brand['Week'].values]
    fig3 = go.Figure()
    for i, brand in enumerate(sorted(week_brand['Brand'].unique())):
        sub = week_brand[week_brand['Brand']==brand].set_index('Week').reindex(wo).reset_index()
        fig3.add_trace(go.Scatter(
            x=sub['Week'], y=sub['OSA'], name=brand, mode='lines+markers',
            line=dict(color=pal[i % len(pal)], width=2), marker=dict(size=7),
            hovertemplate=f"<b>{brand}</b><br>%{{x}}: %{{y:.1f}}%<extra></extra>",
        ))
    fig3.add_hline(y=95, line_dash="dot", line_color=th['green'],
                   annotation_text="95%", annotation_font_color=th['green'])
    fig3.add_hrect(y0=80, y1=95, fillcolor=th['amber'], opacity=0.05, line_width=0)
    apply_layout(fig3, "", height=340)
    fig3.update_yaxes(range=[0, 110], ticksuffix='%')
    st.plotly_chart(fig3, use_container_width=True)

with t2b:
    st.markdown('<div class="sec-label" style="font-size:11px;margin-bottom:6px">by Category</div>', unsafe_allow_html=True)
    week_cat = (dff.groupby(['WEEK_LABEL','PRODUCT CATEGORY'])['OSA']
                   .mean().round(1).reset_index())
    week_cat.columns = ['Week','Category','OSA']
    wo_c = [w for w in week_order if w in week_cat['Week'].values]
    fig3b = go.Figure()
    for i, cat in enumerate(sorted(week_cat['Category'].unique())):
        sub = week_cat[week_cat['Category']==cat].set_index('Week').reindex(wo_c).reset_index()
        fig3b.add_trace(go.Scatter(
            x=sub['Week'], y=sub['OSA'], name=cat, mode='lines+markers',
            line=dict(color=pal[i % len(pal)], width=2), marker=dict(size=7),
            hovertemplate=f"<b>{cat}</b><br>%{{x}}: %{{y:.1f}}%<extra></extra>",
        ))
    fig3b.add_hline(y=95, line_dash="dot", line_color=th['green'],
                    annotation_text="95%", annotation_font_color=th['green'])
    fig3b.add_hrect(y0=80, y1=95, fillcolor=th['amber'], opacity=0.05, line_width=0)
    apply_layout(fig3b, "", height=340)
    fig3b.update_yaxes(range=[0, 110], ticksuffix='%')
    fig3b.update_layout(legend=dict(orientation='h', yanchor='top', y=-0.15, xanchor='left', x=0))
    st.plotly_chart(fig3b, use_container_width=True)

# ── Row 3: Heatmaps — Account×Brand  +  Account×Category ────────────────────
st.markdown("---")
h3a, h3b = st.columns(2)

colorscale = [[0.0,'#f87171'],[0.8,'#fbbf24'],[0.95,'#34c97b'],[1.0,'#34c97b']]

def _heatmap(pivot, title):
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
        colorscale=colorscale, zmin=0, zmax=100,
        text=np.where(np.isnan(pivot.values), '',
                      np.round(pivot.values, 0).astype(int).astype(str) + '%'),
        texttemplate='%{text}', textfont=dict(size=9, color='#0f1117'),
        hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>OSA: %{z:.1f}%<extra></extra>',
        showscale=True,
        colorbar=dict(ticksuffix='%', tickfont=dict(color=T()['text']),
                      outlinecolor=T()['border'], outlinewidth=1),
    ))
    apply_layout(fig, title, height=max(380, len(pivot)*26))
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=9))
    fig.update_yaxes(tickfont=dict(size=9))
    return fig

with h3a:
    st.markdown('<div class="sec-label">Account × Brand Heatmap</div>', unsafe_allow_html=True)
    heat = (dff.groupby(['ACCOUNT','BRAND NAME'])['OSA']
               .mean().round(1).reset_index()
               .pivot(index='ACCOUNT', columns='BRAND NAME', values='OSA'))
    st.plotly_chart(_heatmap(heat, "Account × Brand OSA (%)"), use_container_width=True)

with h3b:
    st.markdown('<div class="sec-label">Account × Category Heatmap</div>', unsafe_allow_html=True)
    heat_cat = (dff.groupby(['ACCOUNT','PRODUCT CATEGORY'])['OSA']
                   .mean().round(1).reset_index()
                   .pivot(index='ACCOUNT', columns='PRODUCT CATEGORY', values='OSA'))
    st.plotly_chart(_heatmap(heat_cat, "Account × Category OSA (%)"), use_container_width=True)

# ── Row 4: Distribution + Month comparison ───────────────────────────────────
st.markdown("---")
r4c1, r4c2 = st.columns(2)

with r4c1:
    st.markdown('<div class="sec-label">OSA Score Distribution</div>', unsafe_allow_html=True)
    on  = (dff['OSA'] == 100).sum()
    off = (dff['OSA'] == 0).sum()
    tot = on + off
    fig5 = go.Figure(go.Pie(
        labels=['On Target (100%)', 'Off Target (0%)'],
        values=[on, off], hole=0.55,
        marker_colors=[th['green'], th['red']],
        textfont=dict(size=12, color=th['text']),
        hovertemplate='%{label}: %{value:,} SKUs (%{percent})<extra></extra>',
    ))
    layout5 = get_plotly_layout(height=300)
    layout5['annotations'] = [dict(text=f"{on/tot*100:.1f}%" if tot else "0%",
                                   x=0.5, y=0.5, font=dict(size=22, color=th['green']), showarrow=False)]
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
            textposition='outside', textfont=dict(size=11, color=th['text']),
        ))
        fig6.add_hline(y=95, line_dash="dot", line_color=th['green'])
        apply_layout(fig6, "", height=300)
        fig6.update_yaxes(range=[0, 115], ticksuffix='%')
        st.plotly_chart(fig6, use_container_width=True)
    else:
        st.info("Select multiple months to compare.")

# ── Row 5: Top/Bottom accounts ────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="sec-label">Account Performance Ranking</div>', unsafe_allow_html=True)
acct_osa = (dff.groupby('ACCOUNT')['OSA'].mean().round(1)
               .sort_values(ascending=False).reset_index())
acct_osa.columns = ['Account','OSA']

ra1, ra2 = st.columns(2)
with ra1:
    top10 = acct_osa.head(10).sort_values('OSA')
    fig7 = go.Figure(go.Bar(
        x=top10['OSA'], y=top10['Account'], orientation='h',
        marker_color=[th['green']]*len(top10),
        text=top10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    apply_layout(fig7, "🏆 Top 10 Accounts", height=340)
    fig7.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig7, use_container_width=True)

with ra2:
    bot10 = acct_osa.tail(10).sort_values('OSA')
    fig8 = go.Figure(go.Bar(
        x=bot10['OSA'], y=bot10['Account'], orientation='h',
        marker_color=color_bars(bot10['OSA']),
        text=bot10['OSA'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    apply_layout(fig8, "⚠️ Bottom 10 Accounts", height=340)
    fig8.update_xaxes(range=[0, 115], ticksuffix='%')
    st.plotly_chart(fig8, use_container_width=True)

st.markdown('<div class="footer">OSA Analytics — Uganda Field Data</div>', unsafe_allow_html=True)
