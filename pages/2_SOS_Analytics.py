"""
SOS Analytics — rebuilt from scratch.
Brands and categories are hardcoded. The page receives a clean DataFrame
from process_sos() that already contains ONLY Uganda brand rows.
No additional brand filtering is needed here — process_sos guarantees it.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, get_theme, get_plotly_layout, sos_color,
    process_sos, build_sos_excel, merge_sos_chunks,
    UG_FOCUS_CATEGORIES, UG_SOS_TARGETS, UG_BRANDS,
    inject_sidebar_toggle, inject_theme_toggle, get_sos_target,
    UG_CATEGORY_CANONICAL, is_ug_brand,
    sos_data_save, sos_data_load, sos_data_clear,
)
from auth import require_login, logout

st.set_page_config(page_title="SOS Analytics", page_icon="📦", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(
        f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>"
        f"Signed in as <strong style='color:var(--text)'>{user}</strong></div>",
        unsafe_allow_html=True,
    )
    if st.button("🚪  Sign Out", key="logout_sos", use_container_width=True):
        logout()

# ── helpers ───────────────────────────────────────────────────────────────────

def T():
    return get_theme()

def apply_layout(fig, title="", height=380):
    th = T()
    fig.update_layout(**get_plotly_layout(height=height, title=title))
    fig.update_layout(hoverlabel=dict(bgcolor=th["surface"], font_color=th["text"], bordercolor=th["border"]))
    fig.update_xaxes(tickfont_color=th["text"], title_font_color=th["text"])
    fig.update_yaxes(tickfont_color=th["text"], title_font_color=th["text"])
    return fig

def bar_colors(vals):
    return [sos_color(v) for v in vals]

# ── page header ───────────────────────────────────────────────────────────────

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>📦 SOS Analytics</h1>
  <p>Share of Shelf — upload your SOS <code>.xlsx</code> file or use previously saved data.</p>
</div>
""", unsafe_allow_html=True)

# ── persistent data banner + clear ───────────────────────────────────────────
saved_df, saved_meta = sos_data_load()

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
        use_saved = col_use.checkbox("Use saved data (skip upload)", value=True, key="sos_use_saved")
        if col_clear.button("🗑️ Clear saved data", key="sos_clear"):
            sos_data_clear()
            st.rerun()
    else:
        st.info("No saved data yet. Upload a file below and it will be saved automatically.")
        use_saved = False

# ── file upload ───────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload SOS .xlsx file(s)",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

@st.cache_data(show_spinner="Processing SOS data…", max_entries=3, ttl=1800)
def load_sos(file_bytes_list: list[bytes], filenames: list[str]) -> pd.DataFrame:
    if len(file_bytes_list) == 1:
        return process_sos(file_bytes_list[0], filenames[0])
    merged = merge_sos_chunks(file_bytes_list)
    return process_sos(merged, "merged.xlsx")

if uploaded_files:
    if len(uploaded_files) > 1:
        st.markdown(
            f'<div style="font-size:12px;color:var(--muted);background:var(--surface);'
            f'border:1px solid var(--border);border-radius:8px;padding:8px 14px;margin-bottom:12px">'
            f'📂 <strong>{len(uploaded_files)} files</strong> — will be merged.</div>',
            unsafe_allow_html=True,
        )
    try:
        all_bytes = [f.read() for f in uploaded_files]
        all_names = [f.name for f in uploaded_files]
        df = load_sos(all_bytes, all_names)
        del all_bytes
        sos_data_save(df, filename=all_names[0] if len(all_names)==1 else f"{len(all_names)} files",
                      uploaded_by=st.session_state.get("username","user"))
        st.success(f"✅ Loaded and saved — {len(df):,} rows.")
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()
elif saved_df is not None and use_saved:
    df = saved_df
else:
    st.markdown(
        '<div class="result-box info">⬆️  Upload your SOS <code>.xlsx</code> file above, '
        'or enable saved data above.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ── At this point df contains ONLY Uganda brand rows (guaranteed by process_sos)
# ── Brands and categories displayed are EXACTLY what's in these hardcoded lists:

# Hardcoded brand list — only these will ever appear in any chart
ALLOWED_BRANDS = set(UG_BRANDS) | {'COSY POA', 'TISHU POA', 'ULTRA CLEAN'}

# Double-check — paranoia filter (should never trigger given process_sos rebuild)
from engines import is_ug_brand
df = df[df['PRODUCT_NAME'].apply(is_ug_brand)].copy()

if df.empty:
    st.error("❌ No Uganda brand data found in this file. Please upload the correct SOS export.")
    st.stop()

# ── summary badge ─────────────────────────────────────────────────────────────

brands_found  = sorted(df['PRODUCT_NAME'].dropna().unique())
n_rows        = len(df)
months_series = df[['MONTH','MONTH_NUM']].drop_duplicates().sort_values('MONTH_NUM')
months        = months_series['MONTH'].tolist()

st.markdown(
    "<div style='font-size:12px;color:#34c97b;background:rgba(52,201,123,.08);"
    "border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
    f"🇺🇬 <strong>Uganda brands only</strong> — "
    f"{', '.join(brands_found)}. "
    f"<strong>{n_rows:,}</strong> rows · "
    f"<strong>{len(months)}</strong> month(s) loaded.</div>",
    unsafe_allow_html=True,
)

# ── filters ───────────────────────────────────────────────────────────────────

accts   = sorted(df['ACCOUNT'].dropna().unique())
cats    = sorted(df['PRODUCT_CATEGORY'].dropna().unique())
outlets = sorted(df['CUSTOMER NAME'].dropna().unique()) if 'CUSTOMER NAME' in df.columns else []
regions = sorted(df['REGION'].dropna().unique()) if 'REGION' in df.columns else []

dates_parsed = df['DATE_PARSED'].dropna()
date_min = dates_parsed.min().date() if not dates_parsed.empty else None
date_max = dates_parsed.max().date() if not dates_parsed.empty else None

with st.expander("🔍  Filters", expanded=True):
    fc1, fc2, fc3 = st.columns(3)
    sel_months = fc1.multiselect("Month",           months,      default=months,      key="sos_months")
    sel_accts  = fc2.multiselect("Chain / Account", accts,       default=accts,       key="sos_accts")
    sel_brands = fc3.multiselect("Brand",           brands_found, default=brands_found, key="sos_brands")

    fc4, fc5, fc6 = st.columns(3)
    sel_cats    = fc4.multiselect("Category",  cats,    default=cats,    key="sos_cats")
    sel_outlets = fc5.multiselect("Outlet",    outlets, default=outlets, key="sos_outlets") if outlets else []
    sel_regions = fc6.multiselect("Region",    regions, default=regions, key="sos_regions") if regions else []

    if date_min and date_max:
        sel_dates = fc6.date_input(
            "Date range", value=(date_min, date_max),
            min_value=date_min, max_value=date_max,
            key="sos_dates", label_visibility="collapsed",
        )
        d_from, d_to = (sel_dates[0], sel_dates[1]) if isinstance(sel_dates, (list,tuple)) and len(sel_dates)==2 else (date_min, date_max)
    else:
        d_from, d_to = date_min, date_max

# ── apply filters ─────────────────────────────────────────────────────────────

dff = df[
    df['MONTH'].isin(sel_months) &
    df['ACCOUNT'].isin(sel_accts) &
    df['PRODUCT_NAME'].isin(sel_brands) &
    df['PRODUCT_CATEGORY'].isin(sel_cats)
].copy()

if sel_outlets and 'CUSTOMER NAME' in dff.columns:
    dff = dff[dff['CUSTOMER NAME'].isin(sel_outlets)]
if sel_regions and 'REGION' in dff.columns:
    dff = dff[dff['REGION'].isin(sel_regions)]
if d_from and d_to:
    dff = dff[(dff['DATE_PARSED'].dt.date >= d_from) & (dff['DATE_PARSED'].dt.date <= d_to)]

if dff.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

# ── KPI cards ─────────────────────────────────────────────────────────────────

th = T()
overall_sos = dff['FACINGS SOS%'].mean()
pct_present = (dff['FACINGS SOS%'] > 0).mean() * 100
avg_pos     = dff[dff['POSITION'] > 0]['POSITION'].mean() if (dff['POSITION'] > 0).any() else 0
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
    <div class="m-value" style="color:{th['green']}">{pct_present:.1f}%</div>
    <div class="m-sub">SOS% &gt; 0</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Avg Shelf Position</div>
    <div class="m-value">{avg_pos:.1f}</div>
    <div class="m-sub">1 = best / eye level</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Brands</div>
    <div class="m-value">{dff['PRODUCT_NAME'].nunique()}</div>
    <div class="m-sub">in selection</div>
  </div>
  <div class="metric-card">
    <div class="m-label">Accounts</div>
    <div class="m-value">{dff['ACCOUNT'].nunique()}</div>
    <div class="m-sub">key account groups</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ── Row 1: SOS by Brand  +  SOS by Account ────────────────────────────────────

r1a, r1b = st.columns(2)

with r1a:
    brand_sos = (dff.groupby('PRODUCT_NAME')['FACINGS SOS%']
                    .mean().round(1).sort_values().reset_index())
    brand_sos.columns = ['Brand', 'SOS']
    avg_tgt = int(dff['SOS_TARGET'].mean()) if 'SOS_TARGET' in dff.columns else 20
    fig = go.Figure(go.Bar(
        x=brand_sos['SOS'], y=brand_sos['Brand'], orientation='h',
        marker_color=bar_colors(brand_sos['SOS']),
        text=brand_sos['SOS'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    fig.add_vline(x=avg_tgt, line_dash="dot", line_color=th['green'],
                  annotation_text=f"{avg_tgt}% avg target",
                  annotation_font_color=th['green'], annotation_font_size=10)
    apply_layout(fig, "SOS% by Brand", height=max(300, len(brand_sos)*36))
    fig.update_xaxes(range=[0, max(brand_sos['SOS'].max()*1.3, avg_tgt*1.5, 30)], ticksuffix='%')
    st.plotly_chart(fig, use_container_width=True)

with r1b:
    acct_sos = (dff.groupby('ACCOUNT')['FACINGS SOS%']
                   .mean().round(1).sort_values().reset_index())
    acct_sos.columns = ['Account', 'SOS']
    fig2 = go.Figure(go.Bar(
        x=acct_sos['SOS'], y=acct_sos['Account'], orientation='h',
        marker_color=bar_colors(acct_sos['SOS']),
        text=acct_sos['SOS'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
    ))
    fig2.add_vline(x=avg_tgt, line_dash="dot", line_color=th['green'],
                   annotation_text=f"{avg_tgt}% avg target",
                   annotation_font_color=th['green'], annotation_font_size=10)
    apply_layout(fig2, "SOS% by Account", height=max(300, len(acct_sos)*36))
    fig2.update_xaxes(range=[0, max(acct_sos['SOS'].max()*1.3, avg_tgt*1.5, 30)], ticksuffix='%')
    st.plotly_chart(fig2, use_container_width=True)

# ── Row 2: Account × Brand heatmap ───────────────────────────────────────────

st.markdown('<div class="sec-label">Account × Brand SOS% Heatmap</div>', unsafe_allow_html=True)

heat_pivot = (dff.groupby(['ACCOUNT','PRODUCT_NAME'])['FACINGS SOS%']
                 .mean().round(1).reset_index()
                 .pivot(index='ACCOUNT', columns='PRODUCT_NAME', values='FACINGS SOS%'))

fig3 = go.Figure(go.Heatmap(
    z=heat_pivot.values,
    x=list(heat_pivot.columns),
    y=list(heat_pivot.index),
    colorscale=[[0.0,'#f87171'],[0.15,'#fbbf24'],[0.3,'#34c97b'],[1.0,'#34c97b']],
    zmin=0,
    zmax=max(50, float(np.nanmax(heat_pivot.values)) if not np.all(np.isnan(heat_pivot.values)) else 50),
    text=np.where(np.isnan(heat_pivot.values), '',
                  np.char.add(np.round(np.where(np.isnan(heat_pivot.values), 0, heat_pivot.values), 0)
                                .astype(int).astype(str), '%')),
    texttemplate='%{text}',
    textfont=dict(size=9, color='#0f1117'),
    hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>SOS: %{z:.1f}%<extra></extra>',
    showscale=True,
    colorbar=dict(ticksuffix='%', tickfont=dict(color=th['text']),
                  outlinecolor=th['border'], outlinewidth=1),
))
apply_layout(fig3, "Account × Brand SOS (%)", height=max(380, len(heat_pivot)*30))
fig3.update_xaxes(tickangle=-35, tickfont=dict(size=10))
st.plotly_chart(fig3, use_container_width=True)

# ── Row 3: SOS by Category ────────────────────────────────────────────────────

st.markdown("---")
st.markdown('<div class="sec-label">SOS% by Category vs MT Target</div>', unsafe_allow_html=True)

dff_cat = dff[dff['PRODUCT_CATEGORY'].isin(UG_FOCUS_CATEGORIES)].copy()
if dff_cat.empty:
    dff_cat = dff.copy()

cat_sos = (dff_cat.groupby('PRODUCT_CATEGORY')
               .agg(SOS=('FACINGS SOS%', 'mean')).round(1).reset_index())
cat_sos['Target']     = cat_sos['PRODUCT_CATEGORY'].apply(get_sos_target)
cat_sos['vs_target']  = cat_sos['SOS'] - cat_sos['Target']
cat_sos = cat_sos.sort_values('SOS')

def _cat_color(row):
    t = T()
    if row['SOS'] >= row['Target']:           return t['green']
    if row['SOS'] >= row['Target'] * 0.5:     return t['amber']
    return t['red']

r3a, r3b = st.columns([3, 2])

with r3a:
    bar_cols = cat_sos.apply(_cat_color, axis=1).tolist()
    fig_cat = go.Figure()
    fig_cat.add_trace(go.Bar(
        x=cat_sos['SOS'], y=cat_sos['PRODUCT_CATEGORY'], orientation='h',
        marker_color=bar_cols,
        text=cat_sos['SOS'].map(lambda v: f"{v:.1f}%"),
        textposition='outside', textfont=dict(size=10, color=th['text']),
        name='Actual SOS%',
    ))
    fig_cat.add_trace(go.Scatter(
        x=cat_sos['Target'], y=cat_sos['PRODUCT_CATEGORY'], mode='markers',
        marker=dict(symbol='line-ns', size=16, line=dict(width=2.5, color=th['green'])),
        name='MT Target',
    ))
    apply_layout(fig_cat, "SOS% by Category vs MT Target", height=max(320, len(cat_sos)*38))
    fig_cat.update_xaxes(range=[0, max(cat_sos['SOS'].max()*1.3, cat_sos['Target'].max()*1.2)], ticksuffix='%')
    fig_cat.update_layout(legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
    st.plotly_chart(fig_cat, use_container_width=True)

with r3b:
    st.markdown('<div class="sec-label">vs MT Target</div>', unsafe_allow_html=True)
    for _, row in cat_sos.sort_values('vs_target').iterrows():
        gap = row['vs_target']
        col_hex = th['green'] if gap >= 0 else (th['amber'] if gap >= -row['Target']*0.5 else th['red'])
        arrow   = '▲' if gap >= 0 else '▼'
        st.markdown(
            f"<div style='display:flex;justify-content:space-between;align-items:center;"
            f"padding:6px 12px;margin-bottom:4px;background:var(--surface);"
            f"border:1px solid var(--border);border-radius:8px;font-size:13px'>"
            f"<span style='color:var(--text);font-weight:500'>{row['PRODUCT_CATEGORY']}</span>"
            f"<span style='color:var(--muted);font-size:11px'>tgt {int(row['Target'])}%</span>"
            f"<span style='color:{col_hex};font-weight:700'>{row['SOS']:.1f}%"
            f" <span style='font-size:10px'>{arrow} {abs(gap):.1f}pp</span></span>"
            f"</div>",
            unsafe_allow_html=True,
        )

# ── Category × Account heatmap ────────────────────────────────────────────────

st.markdown('<div class="sec-label" style="margin-top:16px">Category × Account SOS% Heatmap</div>',
            unsafe_allow_html=True)

cat_pivot = (dff_cat.groupby(['ACCOUNT','PRODUCT_CATEGORY'])['FACINGS SOS%']
                .mean().round(1).reset_index()
                .pivot(index='ACCOUNT', columns='PRODUCT_CATEGORY', values='FACINGS SOS%'))

z_vals  = cat_pivot.values
z_color = np.full_like(z_vals, np.nan, dtype=float)
for ci, col_name in enumerate(cat_pivot.columns):
    tgt = get_sos_target(str(col_name))
    for ri in range(len(cat_pivot.index)):
        v = z_vals[ri, ci]
        if np.isnan(v):    z_color[ri, ci] = np.nan
        elif v >= tgt:     z_color[ri, ci] = 1.0
        elif v >= tgt*0.5: z_color[ri, ci] = 0.5
        else:              z_color[ri, ci] = 0.0

fig_ch = go.Figure(go.Heatmap(
    z=z_color, x=list(cat_pivot.columns), y=list(cat_pivot.index),
    colorscale=[[0.0,'#f87171'],[0.5,'#fbbf24'],[1.0,'#34c97b']],
    zmin=0, zmax=1,
    text=np.where(np.isnan(z_vals), '',
                  np.char.add(np.round(np.where(np.isnan(z_vals), 0, z_vals), 0)
                                .astype(int).astype(str), '%')),
    texttemplate='%{text}', textfont=dict(size=9, color='#0f1117'),
    hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>SOS: %{text}<extra></extra>',
    showscale=False,
))
apply_layout(fig_ch, "Category × Account (🟢 ≥ target  🟡 50–99%  🔴 < 50%)",
             height=max(300, len(cat_pivot)*30))
fig_ch.update_xaxes(tickangle=-35, tickfont=dict(size=10))
st.plotly_chart(fig_ch, use_container_width=True)

st.markdown("---")

# ── Row 4: Trend lines ────────────────────────────────────────────────────────

st.markdown("---")
st.markdown('<div class="sec-label">Monthly Trend</div>', unsafe_allow_html=True)
r4a, r4b = st.columns(2)

palette = lambda t: [t['blue'], t['green'], t['amber'], t['purple'],
                     t['cyan'], t['red'], '#fb923c', '#e879f9', '#a3e635', '#38bdf8']

with r4a:
    st.markdown('<div class="sec-label" style="font-size:11px;margin-bottom:6px">by Brand</div>', unsafe_allow_html=True)
    if len(sel_months) > 1:
        trend = (dff.groupby(['MONTH','PRODUCT_NAME'])['FACINGS SOS%']
                    .mean().round(1).reset_index())
        month_order = [m for m in months if m in trend['MONTH'].values]
        pal = palette(th)
        fig4 = go.Figure()
        for i, brand in enumerate(sorted(trend['PRODUCT_NAME'].unique())):
            sub = trend[trend['PRODUCT_NAME']==brand].set_index('MONTH').reindex(month_order).reset_index()
            fig4.add_trace(go.Scatter(
                x=sub['MONTH'], y=sub['FACINGS SOS%'], name=brand,
                mode='lines+markers',
                line=dict(color=pal[i % len(pal)], width=2), marker=dict(size=7),
            ))
        fig4.add_hline(y=avg_tgt, line_dash="dot", line_color=th['green'],
                       annotation_text=f"{avg_tgt}% avg target",
                       annotation_font_color=th['green'])
        apply_layout(fig4, "", height=340)
        fig4.update_yaxes(ticksuffix='%')
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Select multiple months to see trend.")

with r4b:
    st.markdown('<div class="sec-label" style="font-size:11px;margin-bottom:6px">by Category</div>', unsafe_allow_html=True)
    if len(sel_months) > 1 and not dff_cat.empty:
        cat_trend = (dff_cat.groupby(['MONTH','PRODUCT_CATEGORY'])['FACINGS SOS%']
                        .mean().round(1).reset_index())
        month_order_c = [m for m in months if m in cat_trend['MONTH'].values]
        pal = palette(th)
        fig5 = go.Figure()
        for i, cat in enumerate(sorted(cat_trend['PRODUCT_CATEGORY'].unique())):
            sub = cat_trend[cat_trend['PRODUCT_CATEGORY']==cat].set_index('MONTH').reindex(month_order_c).reset_index()
            tgt = get_sos_target(cat)
            fig5.add_trace(go.Scatter(
                x=sub['MONTH'], y=sub['FACINGS SOS%'], name=cat,
                mode='lines+markers',
                line=dict(color=pal[i % len(pal)], width=2), marker=dict(size=7),
                hovertemplate=f"<b>{cat}</b><br>%{{x}}: %{{y:.1f}}% (target {tgt}%)<extra></extra>",
            ))
        apply_layout(fig5, "", height=340)
        fig5.update_yaxes(ticksuffix='%')
        fig5.update_layout(legend=dict(orientation='h', yanchor='top', y=-0.15, xanchor='left', x=0))
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.info("Select multiple months to see category trend.")

# ── Row 4b: Heatmaps — Account×Brand  +  Account×Category ───────────────────

st.markdown("---")
h4a, h4b = st.columns(2)

_sos_cs = [[0.0,'#f87171'],[0.15,'#fbbf24'],[0.3,'#34c97b'],[1.0,'#34c97b']]

def _sos_heatmap(pivot, title):
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
        colorscale=_sos_cs, zmin=0,
        zmax=max(50, float(np.nanmax(pivot.values)) if not np.all(np.isnan(pivot.values)) else 50),
        text=np.where(np.isnan(pivot.values), '',
                      np.char.add(np.round(np.where(np.isnan(pivot.values), 0, pivot.values), 0)
                                    .astype(int).astype(str), '%')),
        texttemplate='%{text}', textfont=dict(size=9, color='#0f1117'),
        hovertemplate='<b>%{y}</b> × <b>%{x}</b><br>SOS: %{z:.1f}%<extra></extra>',
        showscale=True,
        colorbar=dict(ticksuffix='%', tickfont=dict(color=th['text']),
                      outlinecolor=th['border'], outlinewidth=1),
    ))
    apply_layout(fig, title, height=max(380, len(pivot)*26))
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=9))
    return fig

with h4a:
    st.markdown('<div class="sec-label">Account × Brand Heatmap</div>', unsafe_allow_html=True)
    hp = (dff.groupby(['ACCOUNT','PRODUCT_NAME'])['FACINGS SOS%']
             .mean().round(1).reset_index()
             .pivot(index='ACCOUNT', columns='PRODUCT_NAME', values='FACINGS SOS%'))
    st.plotly_chart(_sos_heatmap(hp, "Account × Brand SOS (%)"), use_container_width=True)

with h4b:
    st.markdown('<div class="sec-label">Account × Category Heatmap</div>', unsafe_allow_html=True)
    hp_cat = (dff.groupby(['ACCOUNT','PRODUCT_CATEGORY'])['FACINGS SOS%']
                 .mean().round(1).reset_index()
                 .pivot(index='ACCOUNT', columns='PRODUCT_CATEGORY', values='FACINGS SOS%'))
    st.plotly_chart(_sos_heatmap(hp_cat, "Account × Category SOS (%)"), use_container_width=True)

# ── Row 5: Shelf Position  +  Brand Presence ─────────────────────────────────

st.markdown("---")
r5a, r5b = st.columns(2)

with r5a:
    st.markdown('<div class="sec-label">Shelf Position by Brand & Account</div>', unsafe_allow_html=True)
    pos_data = (dff[dff['POSITION'] > 0]
                .groupby(['PRODUCT_NAME','ACCOUNT'])['POSITION']
                .mean().round(1).reset_index())
    pos_data.columns = ['Brand','Account','Position']
    if not pos_data.empty:
        pos_data['Size'] = (10 / pos_data['Position']).clip(1, 10) * 8
        pal = palette(th)
        fig6 = go.Figure()
        for i, brand in enumerate(sorted(pos_data['Brand'].unique())):
            sub = pos_data[pos_data['Brand']==brand]
            fig6.add_trace(go.Scatter(
                x=[brand]*len(sub), y=sub['Account'], mode='markers', name=brand,
                marker=dict(size=sub['Size'], color=pal[i % len(pal)], opacity=0.8,
                            line=dict(width=1, color=th['border'])),
                text=sub['Position'].map(lambda v: f"Pos {v:.1f}"),
                hovertemplate=f"<b>{brand}</b><br>%{{y}}<br>Avg position: %{{text}}<extra></extra>",
            ))
        apply_layout(fig6, "Avg Shelf Position (bigger = better)", height=340)
        fig6.update_yaxes(autorange='reversed')
        fig6.update_xaxes(tickangle=-35)
        st.plotly_chart(fig6, use_container_width=True)
    else:
        st.info("No shelf position data available.")

with r5b:
    st.markdown('<div class="sec-label">Brand Presence per Account (% of records)</div>',
                unsafe_allow_html=True)
    presence = (dff.groupby(['ACCOUNT','PRODUCT_NAME'])
                   .apply(lambda g: (g['FACINGS SOS%'] > 0).mean() * 100)
                   .round(1).reset_index())
    presence.columns = ['Account','Brand','PresencePct']
    piv = presence.pivot(index='Account', columns='Brand', values='PresencePct').fillna(0)
    pal = palette(th)
    fig7 = go.Figure()
    for i, brand in enumerate(piv.columns):
        fig7.add_trace(go.Bar(
            name=brand, x=piv.index, y=piv[brand],
            marker_color=pal[i % len(pal)],
            hovertemplate=f"<b>{brand}</b><br>%{{x}}<br>Present: %{{y:.1f}}%<extra></extra>",
        ))
    fig7.update_layout(barmode='group')
    apply_layout(fig7, "", height=340)
    fig7.update_yaxes(range=[0, 115], ticksuffix='%')
    fig7.update_xaxes(tickangle=-35)
    st.plotly_chart(fig7, use_container_width=True)

st.markdown('<div class="footer">SOS Analytics — Uganda Field Data</div>', unsafe_allow_html=True)
