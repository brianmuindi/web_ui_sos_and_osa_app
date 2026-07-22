"""
Report Generator — rebuilt.
Calls process_osa / process_sos from engines and streams the result as a download.
All NaN-mask bugs are handled inside engines.py (v26+). This page is intentionally
thin — it just wires upload → process → download.
"""

import streamlit as st
import pandas as pd
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import (
    GLOBAL_CSS, process_osa, build_osa_excel,
    process_sos, build_sos_excel,
    load_pressure_targets, merge_sos_chunks,
    inject_sidebar_toggle, inject_theme_toggle,
    mhsku_load, osa_data_load, sos_data_load,
)
from auth import require_login, logout
from datetime import datetime

st.set_page_config(page_title="Report Generator", page_icon="⬇️", layout="centered")
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
    if st.button("🚪  Sign Out", key="logout_rpt", use_container_width=True):
        logout()

st.markdown("""
<style>
.block-container { max-width: 720px !important; }
.rg-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 16px; padding: 24px 28px 20px; margin-bottom: 20px;
}
.rg-card h3 { font-size:16px; font-weight:600; color:var(--text); margin:0 0 4px 0; }
.rg-card p  { font-size:13px; color:var(--muted); margin:0 0 18px 0; line-height:1.5; }
.info-grid  { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-top:16px; }
.info-box   { background:var(--bg); border:1px solid var(--border); border-radius:10px; padding:14px; }
.info-box .lbl { font-family:'DM Mono',monospace; font-size:10px; letter-spacing:1.5px;
                 text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
.info-box .val { font-size:13px; color:var(--text); line-height:1.55; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>⬇️ Report Generator</h1>
  <p>Upload your data file and download a fully formatted Excel report instantly.</p>
</div>
""", unsafe_allow_html=True)

# ── helper: safe process wrapper ─────────────────────────────────────────────

def _safe_process_osa(file_bytes, filename, targets):
    """Run process_osa with every known NaN-safe guard applied."""
    df = process_osa(file_bytes, filename, targets_override=targets)
    # Belt-and-braces: fill any boolean mask columns that may still have NaN
    for col in ('COUNTRY', 'REGION'):
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '')
    return df

# ═══════════════════════════════════════════════════════════════════════════════
#  OSA REPORT
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="rg-card">
  <h3>🟢 OSA Report — On-Shelf Availability</h3>
  <p>One tab per month: Brand OSA · Category OSA · SKU OSA · Account × Brand · Account × Category · Account × SKU · Legend</p>
</div>
""", unsafe_allow_html=True)

# ── MHSKU memory targets ──────────────────────────────────────────────────────
mhsku_records, _, mhsku_by = mhsku_load()
_mem_targets = {r["code"]: r["pressure_target"] for r in mhsku_records} if mhsku_records else None

if _mem_targets:
    st.markdown(
        f"<div style='font-size:12px;color:#34c97b;background:rgba(52,201,123,.08);"
        f"border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:12px'>"
        f"✅ <strong>MHSKU Reference loaded from memory</strong> — {len(mhsku_records)} SKUs "
        f"(last updated by <strong>{mhsku_by or 'admin'}</strong>). "
        f"These pressure targets will be applied automatically if no file is uploaded above.</div>",
        unsafe_allow_html=True,
    )

# ── Saved data option ─────────────────────────────────────────────────────────
osa_saved_df, osa_saved_meta = osa_data_load()
_use_osa_saved = False

if osa_saved_df is not None:
    saved_at = osa_saved_meta.get('saved_at', '')[:16].replace('T', ' ')
    st.markdown(
        f"<div style='font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);"
        f"border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
        f"💾 Saved data available: <strong>{osa_saved_meta.get('filename','file')}</strong> — "
        f"{osa_saved_meta.get('rows',0):,} rows · {saved_at}</div>",
        unsafe_allow_html=True,
    )
    _use_osa_saved = st.checkbox("Use saved OSA data (skip upload)", value=False, key="rg_osa_saved")

# ── File upload ───────────────────────────────────────────────────────────────
osa_col, pt_col = st.columns([3, 2])
with osa_col:
    st.markdown('<div class="sec-label">OSA Data File</div>', unsafe_allow_html=True)
    osa_file = st.file_uploader(
        "Upload OSA file (.xls, .xlsx or .csv)", type=["xls","xlsx","csv"],
        key="osa_upload", label_visibility="collapsed",
        disabled=_use_osa_saved,
    )
with pt_col:
    st.markdown(
        '<div class="sec-label">Pressure Targets '
        '<span style="color:var(--muted);font-size:10px">(optional)</span></div>',
        unsafe_allow_html=True,
    )
    osa_pt_file = st.file_uploader(
        "Upload Pressure Targets .xlsx", type=["xlsx"],
        key="osa_pt_upload", label_visibility="collapsed",
    )

_osa_ready = _use_osa_saved or (osa_file is not None)

if _osa_ready:
    if st.button("⚡  Generate OSA Report", key="osa_btn", use_container_width=True,
                 type="primary"):
        try:
            # ── Resolve targets ───────────────────────────────────────────────
            if osa_pt_file:
                with st.spinner("Loading pressure targets…"):
                    targets = load_pressure_targets(osa_pt_file.read())
            elif _mem_targets:
                targets = _mem_targets
            else:
                targets = None

            # ── Process data ──────────────────────────────────────────────────
            if _use_osa_saved:
                with st.spinner("Building OSA report from saved data…"):
                    df_osa = osa_saved_df.copy()
                    xlsx_bytes = build_osa_excel(df_osa)
            else:
                with st.spinner("Reading file…"):
                    file_bytes = osa_file.read()
                with st.spinner("Processing OSA data…"):
                    df_osa = _safe_process_osa(file_bytes, osa_file.name, targets)
                    del file_bytes
                with st.spinner("Building Excel report…"):
                    xlsx_bytes = build_osa_excel(df_osa)

            months_found = sorted(df_osa['Month'].dropna().unique()) if 'Month' in df_osa.columns else []
            rows_found   = len(df_osa)
            del df_osa

            fname = f"OSA_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.session_state.update({
                "osa_xlsx": xlsx_bytes, "osa_fname": fname,
                "osa_error": None, "osa_rows": rows_found,
                "osa_months": months_found,
                "osa_overrides": len(targets) if targets else 0,
            })

        except Exception as e:
            import traceback
            st.session_state.update({"osa_xlsx": None, "osa_error": str(e)})
            with st.expander("🔍 Error details"):
                st.code(traceback.format_exc())

    # ── Result ────────────────────────────────────────────────────────────────
    if st.session_state.get("osa_xlsx"):
        months_str = ", ".join(st.session_state.get("osa_months", []))
        n_ov = st.session_state.get("osa_overrides", 0)
        st.markdown(
            f"<div class='result-box success'>✅ <strong>OSA report ready</strong> — "
            f"{st.session_state.get('osa_rows',0):,} rows · months: <strong>{months_str}</strong>"
            + (f" · {n_ov} SKU target overrides" if n_ov else "") +
            "</div>",
            unsafe_allow_html=True,
        )
        st.download_button(
            label=f"⬇️  Download  {st.session_state['osa_fname']}",
            data=st.session_state["osa_xlsx"],
            file_name=st.session_state["osa_fname"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="osa_dl", use_container_width=True,
        )
    elif st.session_state.get("osa_error"):
        st.error(f"❌ {st.session_state['osa_error']}")

else:
    st.markdown(
        '<div class="result-box info">⬆️  Upload your OSA data file (CSV or Excel) above, '
        'or enable saved data.</div>',
        unsafe_allow_html=True,
    )

st.markdown("<hr>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  SOS REPORT
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="rg-card">
  <h3>📦 SOS Report — Share of Shelf</h3>
  <p>Summary grid of all Key Accounts · one detail tab per account · Legend sheet.</p>
</div>
""", unsafe_allow_html=True)

# ── Saved data option ─────────────────────────────────────────────────────────
sos_saved_df, sos_saved_meta = sos_data_load()
_use_sos_saved = False

if sos_saved_df is not None:
    saved_at_sos = sos_saved_meta.get('saved_at', '')[:16].replace('T', ' ')
    st.markdown(
        f"<div style='font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);"
        f"border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:8px 14px;margin-bottom:8px'>"
        f"💾 Saved data available: <strong>{sos_saved_meta.get('filename','file')}</strong> — "
        f"{sos_saved_meta.get('rows',0):,} rows · {saved_at_sos}</div>",
        unsafe_allow_html=True,
    )
    _use_sos_saved = st.checkbox("Use saved SOS data (skip upload)", value=False, key="rg_sos_saved")

sos_files = st.file_uploader(
    "Upload SOS .xlsx file(s)",
    type=["xlsx"], key="sos_upload", label_visibility="collapsed",
    accept_multiple_files=True,
    disabled=_use_sos_saved,
)

if sos_files and len(sos_files) > 1:
    st.markdown(
        f"<div style='font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);"
        f"border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:6px 14px;margin-bottom:8px'>"
        f"📂 <strong>{len(sos_files)} files</strong> selected — will be merged.</div>",
        unsafe_allow_html=True,
    )

_sos_ready = _use_sos_saved or bool(sos_files)

if _sos_ready:
    if st.button("⚡  Generate SOS Report", key="sos_btn", use_container_width=True,
                 type="primary"):
        try:
            if _use_sos_saved:
                with st.spinner("Building SOS report from saved data…"):
                    df_sos = sos_saved_df.copy()
                    xlsx_bytes = build_sos_excel(df_sos)
            else:
                with st.spinner("Reading file(s)…"):
                    all_bytes = [f.read() for f in sos_files]
                    all_names = [f.name for f in sos_files]
                with st.spinner("Processing SOS data…"):
                    if len(all_bytes) > 1:
                        merged = merge_sos_chunks(all_bytes)
                        del all_bytes
                        df_sos = process_sos(merged, "merged_sos.xlsx")
                        del merged
                    else:
                        df_sos = process_sos(all_bytes[0], all_names[0])
                        del all_bytes
                with st.spinner("Building Excel report…"):
                    xlsx_bytes = build_sos_excel(df_sos)

            months_found = sorted(df_sos['MONTH'].dropna().unique()) if 'MONTH' in df_sos.columns else []
            brands_found = sorted(df_sos['PRODUCT_NAME'].dropna().unique()) if 'PRODUCT_NAME' in df_sos.columns else []
            rows_found   = len(df_sos)
            del df_sos

            fname = f"SOS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.session_state.update({
                "sos_xlsx": xlsx_bytes, "sos_fname": fname,
                "sos_error": None, "sos_rows": rows_found,
                "sos_months": months_found, "sos_brands": brands_found,
            })

        except Exception as e:
            import traceback
            st.session_state.update({"sos_xlsx": None, "sos_error": str(e)})
            with st.expander("🔍 Error details"):
                st.code(traceback.format_exc())

    # ── Result ────────────────────────────────────────────────────────────────
    if st.session_state.get("sos_xlsx"):
        months_str = ", ".join(st.session_state.get("sos_months", []))
        brands_str = ", ".join(st.session_state.get("sos_brands", []))
        st.markdown(
            f"<div class='result-box success'>✅ <strong>SOS report ready</strong> — "
            f"{st.session_state.get('sos_rows',0):,} rows · "
            f"months: <strong>{months_str}</strong> · "
            f"brands: <strong>{brands_str}</strong></div>",
            unsafe_allow_html=True,
        )
        st.download_button(
            label=f"⬇️  Download  {st.session_state['sos_fname']}",
            data=st.session_state["sos_xlsx"],
            file_name=st.session_state["sos_fname"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sos_dl", use_container_width=True,
        )
    elif st.session_state.get("sos_error"):
        st.error(f"❌ {st.session_state['sos_error']}")

else:
    st.markdown(
        '<div class="result-box info">⬆️  Upload your SOS <code>.xlsx</code> file above, '
        'or enable saved data.</div>',
        unsafe_allow_html=True,
    )

st.markdown("<hr>", unsafe_allow_html=True)

# ── Info cards ────────────────────────────────────────────────────────────────

st.markdown("""
<div class="info-grid">
  <div class="info-box">
    <div class="lbl">OSA Color Coding</div>
    <div class="val">🟢 ≥ 95% &nbsp; 🟡 80–94% &nbsp; 🔴 &lt; 80%<br>Applied to every data cell automatically.</div>
  </div>
  <div class="info-box">
    <div class="lbl">SOS Color Coding</div>
    <div class="val">🟢 ≥ target &nbsp; 🟡 50–99% &nbsp; 🔴 &lt; 50%<br>Targets set per category.</div>
  </div>
  <div class="info-box">
    <div class="lbl">OSA Logic</div>
    <div class="val">Score = 100 if Qty ≥ Pressure Target, else 0. Missing targets default to 6.</div>
  </div>
  <div class="info-box">
    <div class="lbl">SOS Logic</div>
    <div class="val">Mean Facings SOS% per brand per account per month. Uganda brands only.</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="footer">All processing is server-side. Files are not stored permanently.</div>',
    unsafe_allow_html=True,
)
