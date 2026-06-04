"""
Must Have SKU Reference — Uganda Report Generator
==================================================
Persistent store for MHSKU list and SOS targets.
• Any logged-in user can VIEW the reference.
• Any logged-in user can UPLOAD a new / updated file to add/refresh SKUs.
• Only ADMIN users can CLEAR the entire store.
"""

import streamlit as st
import sys, os, io
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, THEME, get_theme,
    inject_sidebar_toggle, inject_theme_toggle,
    parse_mhsku_file, mhsku_save, mhsku_load, mhsku_clear,
)
from auth import require_login, logout, is_admin

st.set_page_config(page_title="MHSKU Reference", page_icon="📋", layout="wide")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    role_badge = "🔑 Admin" if is_admin() else "👤 Analyst"
    st.markdown(
        f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>"
        f"Signed in as <strong style='color:var(--text)'>{user}</strong> "
        f"<span style='color:#4f8ef7;font-size:11px'>({role_badge})</span></div>",
        unsafe_allow_html=True,
    )
    if st.button("🚪  Sign Out", key="logout_mhsku", use_container_width=True):
        logout()

# ── Page CSS ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.sku-header { display:flex; align-items:center; justify-content:space-between; margin-bottom:4px; }
.sku-badge  {
    display:inline-block; background:rgba(79,142,247,.15); border:1px solid rgba(79,142,247,.35);
    color:#4f8ef7; border-radius:20px; font-size:11px; font-family:'DM Mono',monospace;
    letter-spacing:1px; padding:3px 12px; text-transform:uppercase;
}
.sku-meta   { font-size:12px; color:var(--muted); margin-bottom:24px; }
.admin-zone {
    background:rgba(248,113,113,.06); border:1px solid rgba(248,113,113,.2);
    border-radius:12px; padding:20px 24px; margin-top:32px;
}
.admin-zone h4 { color:#f87171; margin:0 0 6px 0; font-size:14px; }
.admin-zone p  { color:#9ca3af; font-size:13px; margin:0 0 16px 0; }
.empty-state {
    text-align:center; padding:60px 20px;
    color:var(--muted); border:1.5px dashed var(--border);
    border-radius:16px; margin-top:12px;
}
.empty-state .icon { font-size:40px; margin-bottom:12px; }
.empty-state h3    { font-size:16px; font-weight:600; color:var(--text); margin-bottom:6px; }
.upload-card {
    background:var(--surface); border:1px solid var(--border);
    border-radius:16px; padding:24px 28px; margin-bottom:24px;
}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>📋 Must Have SKU Reference</h1>
  <p>Upload the latest MHSKU file to persist it as a shared reference — used across OSA & SOS analysis.</p>
</div>
""", unsafe_allow_html=True)

# ── Load current store ────────────────────────────────────────────────────────
records, last_updated, last_updated_by = mhsku_load()
df_store = pd.DataFrame(records) if records else pd.DataFrame(
    columns=["code", "sku", "category", "pressure_target"]
)

# ── Upload card ───────────────────────────────────────────────────────────────
st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown("#### 📤 Update MHSKU Reference")
st.markdown(
    "<p style='font-size:13px;color:var(--muted);margin-bottom:16px'>"
    "Upload the latest MHSKU + SOS Targets .xlsx to refresh the shared reference. "
    "Existing SKUs are merged — new codes are added, matching codes are updated.</p>",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Upload MHSKU .xlsx file",
    type=["xlsx"],
    key="mhsku_upload",
    label_visibility="collapsed",
    help="Must contain a 'MHSKU' sheet with columns: Code | SKU | CATEGORIZED | PRESSURE TARGET",
)

if uploaded:
    file_bytes = uploaded.read()
    try:
        parsed = parse_mhsku_file(file_bytes)
        st.success(f"✅  Parsed **{len(parsed)} SKUs** from `{uploaded.name}` — ready to save.")

        # Preview table
        preview_df = pd.DataFrame(parsed).rename(columns={
            "code": "Code", "sku": "SKU", "category": "Category", "pressure_target": "Pressure Target"
        })
        with st.expander(f"Preview ({len(parsed)} rows)", expanded=False):
            st.dataframe(preview_df, use_container_width=True, hide_index=True)

        if st.button("💾  Save to Reference Memory", key="save_mhsku", use_container_width=True):
            mhsku_save(parsed, uploaded_by=st.session_state.get("username", "unknown"))
            st.success("✅  Reference saved! All pages will now use this MHSKU list.")
            st.rerun()
    except Exception as e:
        st.error(f"❌  Could not parse file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# ── Current reference view ────────────────────────────────────────────────────
st.markdown("---")

if last_updated:
    try:
        ts = datetime.fromisoformat(last_updated.replace("Z", "+00:00"))
        ts_fmt = ts.strftime("%d %b %Y, %H:%M UTC")
    except Exception:
        ts_fmt = last_updated
    st.markdown(
        f"<div class='sku-meta'>Last updated: <strong style='color:var(--text)'>{ts_fmt}</strong>"
        f" &nbsp;·&nbsp; by <strong style='color:var(--text)'>{last_updated_by or 'unknown'}</strong>"
        f" &nbsp;·&nbsp; <strong style='color:var(--text)'>{len(df_store)}</strong> SKUs stored</div>",
        unsafe_allow_html=True,
    )

if df_store.empty:
    st.markdown("""
    <div class="empty-state">
      <div class="icon">📭</div>
      <h3>No MHSKU reference uploaded yet</h3>
      <p>Upload the MHSKU .xlsx file above to create the shared reference.</p>
    </div>
    """, unsafe_allow_html=True)
else:
    # Category tabs
    categories = ["All"] + sorted(df_store["category"].dropna().unique().tolist())
    tab_labels = [f"{c} ({len(df_store[df_store['category']==c])})" if c != "All" else f"All ({len(df_store)})" for c in categories]
    tabs = st.tabs(tab_labels)

    display_df = df_store.rename(columns={
        "code": "Code", "sku": "SKU", "category": "Category", "pressure_target": "Pressure Target (PCS)"
    })

    for tab, cat in zip(tabs, categories):
        with tab:
            filtered = display_df if cat == "All" else display_df[display_df["Category"] == cat]
            st.dataframe(
                filtered.reset_index(drop=True),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Pressure Target (PCS)": st.column_config.NumberColumn(format="%d PCS"),
                },
            )
            st.caption(f"{len(filtered)} SKUs")

    # Download current reference as Excel
    st.markdown("<br>", unsafe_allow_html=True)
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="MHSKU Reference")
    out_buf.seek(0)
    st.download_button(
        "⬇️  Download Current Reference (.xlsx)",
        data=out_buf.getvalue(),
        file_name=f"MHSKU_Reference_{datetime.utcnow().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

# ── Admin zone — clear ────────────────────────────────────────────────────────
if is_admin():
    st.markdown("""
    <div class="admin-zone">
      <h4>🔑 Admin Controls</h4>
      <p>Clear the entire MHSKU reference store. This cannot be undone — you will need to re-upload the file.</p>
    </div>
    """, unsafe_allow_html=True)

    col_confirm, col_btn = st.columns([3, 1])
    with col_confirm:
        confirm_text = st.text_input(
            "Type CLEAR to confirm",
            placeholder="Type CLEAR to enable the button",
            key="clear_confirm",
            label_visibility="collapsed",
        )
    with col_btn:
        can_clear = confirm_text.strip().upper() == "CLEAR"
        if st.button("🗑️  Clear Memory", key="clear_mhsku", disabled=not can_clear, use_container_width=True):
            mhsku_clear()
            st.success("✅  MHSKU reference cleared.")
            st.rerun()
    if not df_store.empty:
        st.caption(f"⚠️  This will delete all {len(df_store)} stored SKUs.")
else:
    st.markdown(
        "<div style='margin-top:24px;font-size:12px;color:var(--muted);text-align:center'>"
        "🔒 Clearing the reference requires admin access.</div>",
        unsafe_allow_html=True,
    )
