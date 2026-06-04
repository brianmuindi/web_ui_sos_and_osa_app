"""
Report Generator — download formatted Excel reports
"""

import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from engines import GLOBAL_CSS, THEME, process_osa, build_osa_excel, process_sos, build_sos_excel, load_pressure_targets, merge_sos_chunks, inject_sidebar_toggle, inject_theme_toggle, mhsku_load
from auth import require_login, logout
from datetime import datetime

st.set_page_config(page_title="Report Generator", page_icon="⬇️", layout="centered")
require_login()
st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
inject_sidebar_toggle()
inject_theme_toggle()

# ── Logout button ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    user = st.session_state.get("username", "user")
    st.markdown(f"<div style='font-size:12px;color:var(--muted);padding:0 8px 4px;'>Signed in as <strong style='color:var(--text)'>{user}</strong></div>", unsafe_allow_html=True)
    if st.button("🚪  Sign Out", key="logout_logout_rpt", use_container_width=True):
        logout()


st.markdown("""
<style>
.block-container { max-width: 680px !important; }
.rg-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 16px; padding: 28px 28px 20px;
    margin-bottom: 20px;
}
.rg-card h3 { font-size: 16px; font-weight: 600; color: var(--text); margin: 0 0 6px 0; }
.rg-card p  { font-size: 13px; color: var(--muted); margin: 0 0 20px 0; }
.info-grid {
    display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 16px;
}
.info-box { background: var(--bg); border: 1px solid var(--border); border-radius: 10px; padding: 14px; }
.info-box .lbl { font-family:'DM Mono',monospace; font-size:10px; letter-spacing:1.5px;
                  text-transform:uppercase; color: var(--muted); margin-bottom:6px; }
.info-box .val { font-size:13px; color: var(--text); line-height:1.55; }
</style>
""", unsafe_allow_html=True)

# ─── Header ─────────────────────────────────────────────────────────────────

st.markdown("""
<div class="pg-title">
  <div class="eyebrow">Uganda Field Analytics</div>
  <h1>⬇️ Report Generator</h1>
  <p>Upload your data file and download a fully formatted Excel report instantly.</p>
</div>
""", unsafe_allow_html=True)

# ─── OSA card ───────────────────────────────────────────────────────────────

st.markdown("""
<div class="rg-card">
  <h3>🟢 OSA Report — On-Shelf Availability</h3>
  <p>Generates tabs per month: Brand OSA, Category OSA, Account × Brand, Account × Category — plus a Legend sheet.</p>
</div>
""", unsafe_allow_html=True)

osa_col, pt_col = st.columns([3, 2])
with osa_col:
    st.markdown('<div class="sec-label">OSA Data File</div>', unsafe_allow_html=True)
    osa_file = st.file_uploader("Upload OSA file (.xls, .xlsx or .csv)", type=["xls","xlsx","csv"],
                                 key="osa_upload", label_visibility="collapsed")
with pt_col:
    st.markdown('<div class="sec-label">Pressure Targets <span style="color:var(--muted);font-size:10px">(optional)</span></div>', unsafe_allow_html=True)
    osa_pt_file = st.file_uploader("Upload Pressure Targets .xlsx", type=["xlsx"],
                                    key="osa_pt_upload", label_visibility="collapsed")

# ── Auto-load MHSKU memory as pressure targets fallback ──────────────────────
mhsku_records, mhsku_ts, mhsku_by = mhsku_load()
if mhsku_records:
    st.markdown(
        f"<div style='font-size:12px;color:#34c97b;background:rgba(52,201,123,.08);"
        f"border:1px solid rgba(52,201,123,.2);border-radius:8px;padding:8px 14px;margin-bottom:12px'>"
        f"✅ <strong>MHSKU Reference loaded from memory</strong> — {len(mhsku_records)} SKUs "
        f"(last updated by {mhsku_by or 'unknown'}). "
        f"These pressure targets will be applied automatically if no file is uploaded above.</div>",
        unsafe_allow_html=True,
    )
    _mem_targets = {r["code"]: r["pressure_target"] for r in mhsku_records}
else:
    _mem_targets = None

if osa_file:
    if st.button("⚡  Generate OSA Report", key="osa_btn", use_container_width=True):
        try:
            with st.spinner("Reading file…"):
                file_bytes = osa_file.read()
                pt_bytes   = osa_pt_file.read() if osa_pt_file else None
            with st.spinner("Building OSA report…"):
                # Targets: uploaded file > MHSKU memory > none
                if pt_bytes:
                    targets = load_pressure_targets(pt_bytes)
                    del pt_bytes
                elif _mem_targets:
                    targets = _mem_targets
                else:
                    targets = None
                df_osa     = process_osa(file_bytes, osa_file.name, targets_override=targets)
                del file_bytes  # release raw bytes once processed
                xlsx_bytes = build_osa_excel(df_osa)
                del df_osa      # release processed df once Excel built
                fname      = f"OSA_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                n_overrides = len(targets) if targets else 0
            st.session_state["osa_xlsx"]      = xlsx_bytes
            st.session_state["osa_fname"]     = fname
            st.session_state["osa_error"]     = None
            st.session_state["osa_overrides"] = n_overrides
        except Exception as e:
            st.session_state["osa_xlsx"]  = None
            st.session_state["osa_error"] = str(e)

    if st.session_state.get("osa_xlsx"):
        n = st.session_state.get("osa_overrides", 0)
        note = f" ({n} SKU target overrides applied)" if n else ""
        st.markdown(f'<div class="result-box success">✅ OSA report ready{note} — click below to download.</div>',
                    unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️  Download  {st.session_state['osa_fname']}",
            data=st.session_state["osa_xlsx"],
            file_name=st.session_state["osa_fname"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="osa_dl", use_container_width=True,
        )
    elif st.session_state.get("osa_error"):
        st.markdown(f'<div class="result-box error">❌ {st.session_state["osa_error"]}</div>',
                    unsafe_allow_html=True)
else:
    st.markdown('<div class="result-box info">⬆️  Upload your OSA data file (CSV or Excel) above.</div>',
                unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ─── SOS card ───────────────────────────────────────────────────────────────

st.markdown("""
<div class="rg-card">
  <h3>📦 SOS Report — Share of Shelf</h3>
  <p>Generates a Summary grid of all Key Accounts, one detail tab per account, and a Legend sheet.</p>
</div>
""", unsafe_allow_html=True)

sos_files = st.file_uploader(
    "Upload SOS .xlsx file(s) — select multiple if your export is split into chunks",
    type=["xlsx"], key="sos_upload", label_visibility="collapsed",
    accept_multiple_files=True,
)

if sos_files:
    if len(sos_files) > 1:
        st.markdown(
            f'<div style="font-size:12px;color:#4f8ef7;background:rgba(79,142,247,.08);border:1px solid rgba(79,142,247,.2);border-radius:8px;padding:6px 14px;margin-bottom:8px">' +
            f'📂 {len(sos_files)} files selected — will be merged before building report.</div>',
            unsafe_allow_html=True,
        )
    if st.button("⚡  Generate SOS Report", key="sos_btn", use_container_width=True):
        try:
            with st.spinner("Reading file(s)…"):
                all_bytes = [f.read() for f in sos_files]
                all_names = [f.name for f in sos_files]
            with st.spinner("Merging & building SOS report…" if len(sos_files) > 1 else "Building SOS report…"):
                if len(all_bytes) > 1:
                    merged_bytes = merge_sos_chunks(all_bytes)
                    del all_bytes
                    df_sos = process_sos(merged_bytes, "merged_sos.xlsx")
                    del merged_bytes
                else:
                    df_sos = process_sos(all_bytes[0], all_names[0])
                    del all_bytes
                xlsx_bytes = build_sos_excel(df_sos)
                del df_sos
                fname = f"SOS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.session_state["sos_xlsx"]  = xlsx_bytes
            st.session_state["sos_fname"] = fname
            st.session_state["sos_error"] = None
        except Exception as e:
            st.session_state["sos_xlsx"]  = None
            st.session_state["sos_error"] = str(e)

    if st.session_state.get("sos_xlsx"):
        st.markdown('<div class="result-box success">✅ SOS report ready — click below to download.</div>',
                    unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️  Download  {st.session_state['sos_fname']}",
            data=st.session_state["sos_xlsx"],
            file_name=st.session_state["sos_fname"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sos_dl", use_container_width=True,
        )
    elif st.session_state.get("sos_error"):
        st.markdown(f'<div class="result-box error">❌ {st.session_state["sos_error"]}</div>',
                    unsafe_allow_html=True)
else:
    st.markdown('<div class="result-box info">⬆️  Upload your SOS <code>.xlsx</code> file above.</div>',
                unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ─── Info cards ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="info-grid">
  <div class="info-box">
    <div class="lbl">OSA Color Coding</div>
    <div class="val">🟢 ≥ 95% &nbsp; 🟡 80–94% &nbsp; 🔴 &lt; 80%<br>Applied to every data cell automatically.</div>
  </div>
  <div class="info-box">
    <div class="lbl">SOS Color Coding</div>
    <div class="val">🟢 ≥ 20% &nbsp; 🟡 1–19% &nbsp; 🔴 = 0%<br>Applied to every SOS% cell automatically.</div>
  </div>
  <div class="info-box">
    <div class="lbl">OSA Logic</div>
    <div class="val">Not-listed SKUs dropped. Missing targets defaulted to 6. Score = 100 if Qty ≥ PT, else 0.</div>
  </div>
  <div class="info-box">
    <div class="lbl">SOS Logic</div>
    <div class="val">Mean Facings SOS% per brand per account per month. Position averaged where &gt; 0.</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="footer">All processing is server-side. Files are not stored.</div>',
            unsafe_allow_html=True)
