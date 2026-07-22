"""
Shared data processing engines and chart theme for OSA & SOS.
Imported by all pages.
"""

import io
import itertools
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
#  CHART THEME
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
#  THEME  (dark default, light mode toggled via session_state)
# ─────────────────────────────────────────────────────────────────────────────

THEME_DARK = {
    "bg": "#0f1117", "surface": "#1a1d27", "border": "#2a2d3a",
    "text": "#e8eaf0", "muted": "#6b7280",
    "green": "#34c97b", "amber": "#f59e0b", "red": "#f87171",
    "blue": "#4f8ef7", "purple": "#a78bfa", "cyan": "#22d3ee",
}
THEME_LIGHT = {
    "bg": "#f0f4ff", "surface": "#ffffff", "border": "#c7d7f5",
    "text": "#0f1f4b", "muted": "#5a6a8a",
    "green": "#16a34a", "amber": "#d97706", "red": "#dc2626",
    "blue": "#1d4ed8", "purple": "#7c3aed", "cyan": "#0891b2",
}

# Default chart theme (dark) — pages call get_plotly_layout() at render time
THEME = THEME_DARK

def get_theme():
    try:
        import streamlit as st
        if st.session_state.get("light_mode", False):
            return THEME_LIGHT
        return THEME_DARK
    except Exception:
        return THEME_DARK

def get_plotly_layout(height=380, title=""):
    t = get_theme()
    layout = dict(
        paper_bgcolor=t["surface"],
        plot_bgcolor=t["surface"],
        height=height,
        font=dict(family="DM Sans, sans-serif", color=t["text"], size=12),
        margin=dict(l=20, r=20, t=40 if title else 20, b=20),
        legend=dict(
            bgcolor=t["bg"],
            bordercolor=t["border"],
            borderwidth=1,
            font=dict(size=11, color=t["text"]),   # explicit text color
            title_font=dict(color=t["text"]),
        ),
        xaxis=dict(
            gridcolor=t["border"],
            linecolor=t["border"],
            tickfont=dict(size=11, color=t["text"]),
            title_font=dict(color=t["text"]),
        ),
        yaxis=dict(
            gridcolor=t["border"],
            linecolor=t["border"],
            tickfont=dict(size=11, color=t["text"]),
            title_font=dict(color=t["text"]),
        ),
    )
    # Only add title key if there's actually a title — avoids "undefined" rendering
    if title:
        layout["title"] = dict(
            text=title,
            font=dict(size=13, color=t["muted"]),
            x=0, pad=dict(l=0)
        )
    return layout

def osa_color(val):
    t = get_theme()
    if __import__('pandas').isna(val): return t["muted"]
    if val >= 95: return t["green"]
    if val >= 80: return t["amber"]
    return t["red"]

def sos_color(val, target: int = None):
    """Return theme colour for an SOS value vs its category target."""
    t = get_theme()
    thr = target if target is not None else SOS_GREEN_THR
    if __import__('pandas').isna(val) or val == 0: return t["red"]
    if val >= thr:            return t["green"]
    if val >= thr * 0.5:      return t["amber"]
    return t["red"]

# Keep for backward compat
PLOTLY_LAYOUT = dict(
    paper_bgcolor=THEME_DARK["surface"], plot_bgcolor=THEME_DARK["surface"],
    font=dict(family="DM Sans, sans-serif", color=THEME_DARK["text"], size=12),
    margin=dict(l=20, r=20, t=40, b=20),
    legend=dict(bgcolor=THEME_DARK["bg"], bordercolor=THEME_DARK["border"], borderwidth=1, font=dict(size=11)),
    xaxis=dict(gridcolor=THEME_DARK["border"], linecolor=THEME_DARK["border"], tickfont=dict(size=11)),
    yaxis=dict(gridcolor=THEME_DARK["border"], linecolor=THEME_DARK["border"], tickfont=dict(size=11)),
)

# ─────────────────────────────────────────────────────────────────────────────
#  GLOBAL CSS  — uses CSS variables so one class on <body> flips the theme
# ─────────────────────────────────────────────────────────────────────────────

GLOBAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

/* ── CSS variables: dark (default) ── */
:root {
  --bg:      #0f1117;
  --surface: #1a1d27;
  --border:  #2a2d3a;
  --text:    #e8eaf0;
  --muted:   #6b7280;
  --blue:    #4f8ef7;
  --blue2:   #6b9ff8;
  --green:   #34c97b;
  --amber:   #f59e0b;
  --red:     #f87171;
  --sidebar-bg: #1a1d27;
}

/* ── CSS variables: light mode ── */
body.light-mode {
  --bg:      #f0f4ff;
  --surface: #ffffff;
  --border:  #c7d7f5;
  --text:    #0f1f4b;
  --muted:   #5a6a8a;
  --blue:    #1d4ed8;
  --blue2:   #2563eb;
  --green:   #16a34a;
  --amber:   #d97706;
  --red:     #dc2626;
  --sidebar-bg: #e8eeff;
}

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; }
#MainMenu, footer, header { visibility: hidden; }

/* ── app background ── */
.stApp { background: var(--bg) !important; }
.stApp > div { background: var(--bg) !important; }
.block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1100px; }

/* ── sidebar ── */
section[data-testid="stSidebar"] { background: var(--sidebar-bg) !important; }
section[data-testid="stSidebar"] * { color: var(--text) !important; }
section[data-testid="stSidebar"] a { color: var(--blue) !important; }

/* ── main text ── */
.stMarkdown, .stMarkdown p, div[data-testid="stMarkdownContainer"] p { color: var(--text) !important; }

/* ── page title ── */
.pg-title { margin-bottom: 1.5rem; }
.pg-title .eyebrow { font-family:'DM Mono',monospace; font-size:10px; letter-spacing:3px; color:var(--blue); text-transform:uppercase; margin-bottom:4px; }
.pg-title h1 { font-size:26px; font-weight:600; color:var(--text); margin:0 0 4px 0; }
.pg-title p  { font-size:13px; color:var(--muted); margin:0; }

/* ── metric cards ── */
.metric-row { display:flex; gap:12px; margin-bottom:20px; flex-wrap:wrap; }
.metric-card { flex:1; min-width:130px; background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:16px 18px; }
.metric-card .m-label { font-family:'DM Mono',monospace; font-size:9px; letter-spacing:2px; text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
.metric-card .m-value { font-size:28px; font-weight:600; color:var(--text); line-height:1; }
.metric-card .m-sub   { font-size:11px; color:var(--muted); margin-top:4px; }

/* ── section label ── */
.sec-label { font-family:'DM Mono',monospace; font-size:10px; letter-spacing:2px; text-transform:uppercase; color:var(--muted); margin:20px 0 8px 0; }

/* ── file uploader ── */
div[data-testid="stFileUploader"] { border:2px dashed var(--border) !important; border-radius:12px; background:var(--surface); padding:12px; }
div[data-testid="stFileUploader"]:hover { border-color:var(--blue) !important; }
div[data-testid="stFileUploader"] label { color:var(--text) !important; }
div[data-testid="stFileUploader"] small { color:var(--muted) !important; }

/* ── primary button ── */
div[data-testid="stButton"] > button { background:var(--blue) !important; color:#fff !important; border:none !important; border-radius:10px !important; padding:12px 20px !important; font-size:14px !important; font-weight:600 !important; font-family:'DM Sans',sans-serif !important; width:100%; }
div[data-testid="stButton"] > button:hover { background:var(--blue2) !important; }

/* ── download button ── */
div[data-testid="stDownloadButton"] > button { background:var(--green) !important; color:#fff !important; border:none !important; border-radius:10px !important; padding:12px 20px !important; font-size:14px !important; font-weight:600 !important; width:100%; }

/* ── select / multiselect ── */
div[data-testid="stSelectbox"] > div, div[data-testid="stMultiSelect"] > div { background:var(--surface) !important; border-color:var(--border) !important; color:var(--text) !important; }

/* ── tabs ── */
button[data-baseweb="tab"] { color:var(--muted) !important; font-size:13px !important; }
button[data-baseweb="tab"][aria-selected="true"] { color:var(--blue) !important; }
div[data-baseweb="tab-highlight"] { background-color:var(--blue) !important; }
div[data-baseweb="tab-border"]    { background-color:var(--border) !important; }

/* ── alerts ── */
.result-box { border-radius:10px; padding:12px 16px; font-size:13px; margin:8px 0; display:flex; gap:10px; align-items:flex-start; }
.result-box.success { background:rgba(52,201,123,.12); border:1px solid rgba(52,201,123,.4); color:var(--green); }
.result-box.error   { background:rgba(220,38,38,.08);  border:1px solid rgba(220,38,38,.3);  color:var(--red); }
.result-box.info    { background:rgba(29,78,216,.08);  border:1px solid rgba(29,78,216,.3);  color:var(--blue); }

hr { border-color:var(--border) !important; margin:20px 0 !important; }
.footer { text-align:center; font-size:11px; color:var(--muted); margin-top:12px; }

/* ── multiselect tags ── */
span[data-baseweb="tag"] { background:var(--blue) !important; }

/* ── light mode: extra overrides ── */
body.light-mode .stApp,
body.light-mode .stApp > div { background: var(--bg) !important; }
body.light-mode section[data-testid="stSidebar"] { background: var(--sidebar-bg) !important; box-shadow: 2px 0 12px rgba(29,78,216,.08); }
body.light-mode .metric-card { box-shadow: 0 1px 4px rgba(29,78,216,.08); }
body.light-mode div[data-testid="stTextInput"] input,
body.light-mode div[data-testid="stSelectbox"] > div,
body.light-mode div[data-testid="stMultiSelect"] > div { background: #f8faff !important; color: var(--text) !important; border-color: var(--border) !important; }
body.light-mode hr { border-color: var(--border) !important; }

/* ── light mode: force ALL text dark ── */
body.light-mode p,
body.light-mode h1, body.light-mode h2, body.light-mode h3,
body.light-mode label, body.light-mode span,
body.light-mode div[data-testid="stMarkdownContainer"],
body.light-mode div[data-testid="stMarkdownContainer"] * { color: var(--text) !important; }

/* ── light mode: nav card buttons — dark text on white/blue ── */
body.light-mode section[data-testid="column"] div[data-testid="stButton"] > button {
    background: #ffffff !important;
    border: 2px solid var(--blue) !important;
    color: var(--text) !important;
}
body.light-mode section[data-testid="column"] div[data-testid="stButton"] > button:hover {
    background: var(--blue) !important;
    color: #ffffff !important;
}

/* ── light mode: primary action buttons stay blue with white text ── */
body.light-mode div[data-testid="stButton"] > button {
    color: #ffffff !important;
}
/* but nav card buttons override to dark text (above rule is more specific) */

/* ── light mode: legend pills ── */
body.light-mode .legend-pill { background: #ffffff; border-color: var(--border); color: var(--text) !important; }

/* ── light mode: home title ── */
body.light-mode .home-hero h1 { color: var(--text) !important; }
body.light-mode .home-hero p  { color: var(--muted) !important; }
body.light-mode .home-hero .eyebrow { color: var(--blue) !important; }

/* ── light mode: pg-title ── */
body.light-mode .pg-title h1 { color: var(--text) !important; }
body.light-mode .pg-title p  { color: var(--muted) !important; }

/* ── light mode: metric cards ── */
body.light-mode .metric-card .m-value { color: var(--text) !important; }
body.light-mode .metric-card .m-label { color: var(--muted) !important; }
body.light-mode .metric-card .m-sub   { color: var(--muted) !important; }

/* ── light mode: result boxes ── */
body.light-mode .result-box.info    { background: rgba(29,78,216,.06) !important; color: #1d4ed8 !important; }
body.light-mode .result-box.success { background: rgba(22,163,74,.06) !important; color: #15803d !important; }
body.light-mode .result-box.error   { background: rgba(220,38,38,.06) !important; color: #b91c1c !important; }
</style>
"""



# ── Sidebar toggle injected via components.v1.html (bypasses Streamlit's HTML sanitiser)
SIDEBAR_TOGGLE_JS = """
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body { background: transparent; overflow: hidden; }
  #btn {
    position: fixed; top: 0; left: 0;
    width: 38px; height: 38px;
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px; cursor: pointer;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center; gap: 5px;
    transition: border-color .2s, background .2s;
  }
  #btn:hover { border-color: #4f8ef7; background: rgba(79,142,247,.15); }
  .bar {
    display: block; width: 16px; height: 2px;
    background: #e8eaf0; border-radius: 2px;
    transition: transform .25s, opacity .25s;
  }
  #btn.closed .b1 { transform: rotate(45deg) translate(5px, 5px); }
  #btn.closed .b2 { opacity: 0; }
  #btn.closed .b3 { transform: rotate(-45deg) translate(5px, -5px); }
</style>

<div id="btn" title="Toggle sidebar">
  <span class="bar b1"></span>
  <span class="bar b2"></span>
  <span class="bar b3"></span>
</div>

<script>
(function() {
  var btn = document.getElementById('btn');

  // Click Streamlit's own internal collapse button
  function clickStreamlitToggle() {
    var p = window.parent.document;
    var collapseBtn = p.querySelector('[data-testid="stSidebarCollapseButton"] button');
    if (!collapseBtn) {
        collapseBtn = p.querySelector('button[data-testid="stSidebarCollapseButton"]');
    }
    if (collapseBtn) {
      collapseBtn.click();
      return true;
    }
    return false;
  }

  function isSidebarCollapsed() {
    var p = window.parent.document;
    var sidebar = p.querySelector('section[data-testid="stSidebar"]');
    if (!sidebar) return false;
    // Streamlit adds data-collapsed or aria-expanded, or we check transform
    var style = window.parent.getComputedStyle(sidebar);
    var rect = sidebar.getBoundingClientRect();
    return rect.right <= 10 || sidebar.getAttribute('aria-expanded') === 'false';
  }

  function syncIcon() {
    if (isSidebarCollapsed()) {
      btn.classList.add('closed');
    } else {
      btn.classList.remove('closed');
    }
  }

  btn.addEventListener('click', function() {
    var didClick = clickStreamlitToggle();
    // Update icon after short delay to let Streamlit animate
    setTimeout(syncIcon, 350);
  });

  // Keep icon in sync with actual sidebar state
  function watchSidebar() {
    var p = window.parent.document;
    var sidebar = p.querySelector('section[data-testid="stSidebar"]');
    if (!sidebar) {
      setTimeout(watchSidebar, 300);
      return;
    }
    syncIcon();
    // Watch for attribute/style changes on the sidebar
    var observer = new MutationObserver(function() { syncIcon(); });
    observer.observe(sidebar, { attributes: true, attributeFilter: ['style', 'class', 'aria-expanded'] });
    // Also watch the parent for transform changes
    observer.observe(sidebar.parentElement || p.body, { childList: true, subtree: false });
  }

  // Wait for parent page to fully render
  setTimeout(watchSidebar, 500);
})();
</script>
"""

def inject_sidebar_toggle():
    """
    Injects a toggle button OUTSIDE the sidebar (in the main area, top-left),
    so it stays visible even when the sidebar is collapsed.
    Uses components.html with JS that targets Streamlit's collapse button.
    The iframe is hidden by collapsing the Streamlit component container via JS.
    """
    import streamlit as st
    import streamlit.components.v1 as components

    components.html("""
<!DOCTYPE html>
<html>
<head>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  html, body { background:transparent; overflow:hidden; width:42px; height:42px; }
  #btn {
    width:42px; height:42px;
    background:#1a1d27; border:1.5px solid #2a2d3a; border-radius:8px;
    cursor:pointer; display:flex; flex-direction:column;
    align-items:center; justify-content:center; gap:5px;
    transition:border-color .2s, background .2s;
  }
  #btn:hover { border-color:#4f8ef7; background:rgba(79,142,247,.15); }
  .bar { display:block; width:16px; height:2px; background:#e8eaf0; border-radius:2px; transition:transform .25s, opacity .25s; }
  #btn.closed .b1 { transform:rotate(45deg) translate(5px,5px); }
  #btn.closed .b2 { opacity:0; }
  #btn.closed .b3 { transform:rotate(-45deg) translate(5px,-5px); }
</style>
</head>
<body>
<div id="btn"><span class="bar b1"></span><span class="bar b2"></span><span class="bar b3"></span></div>
<script>
(function(){
  var btn = document.getElementById('btn');

  function getCollapseBtn() {
    var p = window.parent.document;
    var b = p.querySelector('[data-testid="stSidebarCollapseButton"] button');
    if (!b) b = p.querySelector('button[data-testid="stSidebarCollapseButton"]');
    return b;
  }

  function isSidebarCollapsed() {
    var p = window.parent.document;
    var sidebar = p.querySelector('section[data-testid="stSidebar"]');
    if (!sidebar) return false;
    var rect = sidebar.getBoundingClientRect();
    return rect.right < 20;
  }

  function syncIcon() {
    if (isSidebarCollapsed()) btn.classList.add('closed');
    else btn.classList.remove('closed');
  }

  btn.addEventListener('click', function(){
    var cb = getCollapseBtn();
    if (cb) {
      cb.click();
      setTimeout(syncIcon, 400);
    }
  });

  // Also: resize this iframe container to 0 height in parent, then position fixed
  function positionSelf() {
    try {
      var p = window.parent.document;
      // Find our iframe in parent
      var frames = p.querySelectorAll('iframe');
      var me = null;
      for (var i=0; i<frames.length; i++) {
        try {
          if (frames[i].contentWindow === window) { me = frames[i]; break; }
        } catch(e) {}
      }
      if (me) {
        // Position the iframe fixed top-left
        me.style.cssText = 'position:fixed!important;top:10px!important;left:10px!important;width:42px!important;height:42px!important;border:none!important;z-index:99999!important;background:transparent!important;';
        // Collapse the wrapper div height to 0
        var wrapper = me.parentElement;
        while (wrapper && wrapper !== p.body) {
          if (wrapper.style !== undefined) {
            var h = wrapper.getBoundingClientRect().height;
            if (h > 42 && h < 200) {
              wrapper.style.height = '0';
              wrapper.style.overflow = 'visible';
              wrapper.style.marginTop = '0';
              wrapper.style.marginBottom = '0';
              wrapper.style.paddingTop = '0';
              wrapper.style.paddingBottom = '0';
            }
          }
          wrapper = wrapper.parentElement;
        }
      }
    } catch(e) {}
  }

  function init() {
    positionSelf();
    syncIcon();
    // Watch sidebar for collapse/expand
    var p = window.parent.document;
    var sidebar = p.querySelector('section[data-testid="stSidebar"]');
    if (sidebar) {
      var obs = new MutationObserver(syncIcon);
      obs.observe(sidebar, {attributes:true, attributeFilter:['style','class']});
      obs.observe(sidebar.parentElement || p.body, {childList:true, subtree:true, attributeFilter:['style']});
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', function(){ setTimeout(init,300); });
  else setTimeout(init, 300);
  setTimeout(init, 800);
  setTimeout(init, 1500);
})();
</script>
</body>
</html>
""", height=50, scrolling=False)


def inject_theme_toggle():
    """
    Renders a 🌙/☀️ toggle in the sidebar that flips between dark and light mode.
    Applies/removes the 'light-mode' class on <body> via JS, which triggers all
    CSS variable overrides in GLOBAL_CSS. Also re-runs the page so Plotly charts
    re-render with the correct colours.
    """
    import streamlit as st
    import streamlit.components.v1 as components

    # Initialise session state
    if "light_mode" not in st.session_state:
        st.session_state["light_mode"] = False

    is_light = st.session_state["light_mode"]
    label    = "☀️  Light mode" if not is_light else "🌙  Dark mode"
    tooltip  = "Switch to light mode" if not is_light else "Switch to dark mode"

    with st.sidebar:
        if st.button(label, key="_theme_toggle", help=tooltip, use_container_width=True):
            st.session_state["light_mode"] = not st.session_state["light_mode"]
            st.rerun()

    # Apply the CSS class to <body> in the parent page
    mode_class = "light-mode" if st.session_state["light_mode"] else ""
    components.html(f"""
<script>
(function() {{
  function apply() {{
    var body = window.parent.document.body;
    if (!body) return;
    if ("{mode_class}" === "light-mode") {{
      body.classList.add("light-mode");
    }} else {{
      body.classList.remove("light-mode");
    }}
  }}
  apply();
  setTimeout(apply, 200);
  setTimeout(apply, 600);
}})();
</script>
""", height=0, scrolling=False)


OSA_CHAIN_KEYWORDS = [
    'CARREFOUR','NAIVAS','QUICKMART','CHANDARANA','CLEANSHELF',
    'DEFCO','MAGUNAS','EASTMATT','MATHAI','KHETIAS','POWERSTAR',
    'JAZA','KASSMATT','LEESTAR','SKYMATT',
]
OSA_DEFAULT_PT = 6

HDR_BLUE   = PatternFill('solid', fgColor='1F4E79')
HDR_MED    = PatternFill('solid', fgColor='2E75B6')
AVG_FILL   = PatternFill('solid', fgColor='FFF2CC')
GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
RED_FILL   = PatternFill('solid', fgColor='FFC7CE')
AMBER_FILL = PatternFill('solid', fgColor='FFEB9C')
GRAY_FILL  = PatternFill('solid', fgColor='F2F2F2')
_thin      = Side(style='thin', color='BFBFBF')

def _bdr():
    return Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def _osa_fill_xl(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return None
    return GREEN_FILL if val >= 95 else (AMBER_FILL if val >= 80 else RED_FILL)

OSA_WEEKS = ['Week 1','Week 2','Week 3','Week 4','Week 5','Week 6']


def load_pressure_targets(file_bytes: bytes) -> dict:
    """
    Parse the Pressure Targets reference file (.xlsx).
    Returns {PRODUCT_CODE: numeric_target} dict.
    Expects columns: Code | SKU | Pressure Target
    Target values like '6PCS' are parsed to extract the number (6).
    """
    import re as _re
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws  = wb.active
    targets = {}
    for row in ws.iter_rows(values_only=True):
        code = row[1]; pt = row[3]
        if not code or not pt:
            continue
        code = str(code).strip()
        if code.lower() in ('code', 'mt 1', 'mt 2', 'mt 3', 'mt 4'):
            continue
        nums = _re.findall(r'\d+', str(pt).strip())
        if nums:
            targets[code] = int(nums[0])
    return targets


def process_osa(file_bytes: bytes, filename: str = "",
                targets_override: dict = None) -> pd.DataFrame:
    """Run the full OSA pipeline and return the enriched DataFrame.
    Accepts .csv, .xls, or .xlsx — auto-detected from content.

    targets_override: optional {PRODUCT_CODE: int} from load_pressure_targets().
    When provided, overrides PRESSURE TARGET for matching SKU codes.
    """
    fname = filename.lower()

    # ── Load (large-file optimised — handles up to 300 MB) ───────────────────
    if fname.endswith(".csv"):
        df = read_large_csv(file_bytes)
    else:
        # Try Excel with OSA sheet first, fall back to first sheet, then CSV
        try:
            try:
                df = read_large_excel(file_bytes, sheet_name='OSA')
            except Exception:
                df = read_large_excel(file_bytes, sheet_name=0)
        except Exception:
            df = read_large_csv(file_bytes)

    # ── Normalise column names ────────────────────────────────────────────────
    _OSA_DATE_ALIASES = [
        "date reported", "date_reported", "visit date", "date created",
        "date_created", "survey date", "submission date", "date", "created date",
    ]
    col_lower = {c.lower().strip(): c for c in df.columns}
    if 'DATE REPORTED' not in df.columns:
        for alias in _OSA_DATE_ALIASES:
            if alias in col_lower:
                df = df.rename(columns={col_lower[alias]: 'DATE REPORTED'})
                break
        else:
            available = ", ".join(f"'{c}'" for c in df.columns[:20])
            raise ValueError(
                f"Could not find a date column in your OSA file. "
                f"Expected 'DATE REPORTED' or similar. "
                f"Columns found: {available}"
            )

    # Remap other common OSA column variants
    _OSA_COL_ALIASES = {
        "CUSTOMER NAME":   ["customer name", "customer_name", "outlet", "outlet name", "store name", "store"],
        "BRAND NAME":      ["brand name", "brand_name", "brand", "product brand"],
        "PRODUCT CATEGORY":["product category", "product_category", "category", "cat"],
        "PRODUCT CODE":    ["product code", "product_code", "sku code", "item code", "code"],
        "DESCRIPTION":     ["description", "product description", "product_description",
                             "product name", "product_name", "sku name", "sku_name",
                             "item description", "item name", "item_name", "sku description"],
        "PRESSURE TARGET": ["pressure target", "pressure_target", "target", "pt"],
        "QUANTITY":        ["quantity", "qty", "stock qty", "count"],
        "STOCK LEVEL":     ["stock level", "stock_level", "availability", "available"],
        "COUNTRY":         ["country", "country name", "country_name", "nation"],
        "REGION":          ["region", "territory", "area", "district"],
    }
    col_lower = {c.lower().strip(): c for c in df.columns}
    renames = {}
    for canonical, aliases in _OSA_COL_ALIASES.items():
        if canonical in df.columns:
            continue
        for alias in aliases:
            if alias in col_lower:
                renames[col_lower[alias]] = canonical
                break
    if renames:
        df = df.rename(columns=renames)

    # ── Early filters — shrink the frame BEFORE expensive operations ───────────
    # Filter Uganda brands first (biggest row reduction on multi-country exports)
    _UG_REGION_KEYWORDS = [
        'kampala', 'uganda', 'entebbe', 'jinja', 'mbarara', 'gulu',
        'wakiso', 'mukono', 'lira', 'mbale', 'arua', 'masaka',
    ]
    if 'BRAND NAME' in df.columns:
        df = df[df['BRAND NAME'].astype(str).apply(is_ug_osa_brand)].copy()

    # Filter Uganda outlets (drop Kenya / other-country rows early)
    if 'COUNTRY' in df.columns:
        mask_ug = df['COUNTRY'].astype(str).str.strip().str.lower().isin(['uganda', 'ug'])
        df = df[mask_ug].copy()
    elif 'REGION' in df.columns:
        mask_ug = df['REGION'].astype(str).str.strip().str.lower().apply(
            lambda r: any(kw in r for kw in _UG_REGION_KEYWORDS) if r not in ('nan','none','') else False
        )
        df = df[mask_ug].copy()

    # ── Parse dates ───────────────────────────────────────────────────────────
    # DATE REPORTED may be:
    #   a) Excel serial number (float like 46172.49)  → convert via Excel epoch
    #   b) ISO string "2026-05-30"                    → dayfirst=False
    #   c) DMY string "30/05/2026"                    → dayfirst=True
    _date_col = df['DATE REPORTED'].dropna()
    _is_numeric = pd.to_numeric(_date_col, errors='coerce').notna().mean() > 0.8

    if _is_numeric:
        # Excel serial: days since 1899-12-30
        _serials = pd.to_numeric(df['DATE REPORTED'], errors='coerce')
        df['DATE_PARSED'] = pd.to_datetime(
            _serials.apply(lambda x: pd.Timestamp('1899-12-30') + pd.Timedelta(days=x)
                           if pd.notna(x) else pd.NaT)
        )
    else:
        _sample = _date_col.astype(str).head(20)
        _is_iso = _sample.str.match(r'^\d{4}[-/]\d').any()
        df['DATE_PARSED'] = pd.to_datetime(df['DATE REPORTED'], errors='coerce',
                                           dayfirst=not _is_iso)

    df = df[df['DATE_PARSED'].notna()].copy()

    # ── Derive Month if missing ───────────────────────────────────────────────
    if 'Month' not in df.columns:
        df['Month'] = df['DATE_PARSED'].dt.strftime('%B')

    # ── Week labels ───────────────────────────────────────────────────────────
    df['WEEK_NUM']   = df['DATE_PARSED'].dt.isocalendar().week.astype('Int64')
    df['WEEK_LABEL'] = None
    for month, group in df.groupby('Month'):
        sw = sorted(group['WEEK_NUM'].dropna().unique())
        wm = {w: f'Week {i+1}' for i, w in enumerate(sw)}
        df.loc[df['Month'] == month, 'WEEK_LABEL'] = df.loc[df['Month'] == month, 'WEEK_NUM'].map(wm)

    # ── Apply pressure targets override (from reference file) ────────────────
    df['PRESSURE TARGET'] = pd.to_numeric(df['PRESSURE TARGET'], errors='coerce').fillna(0)
    df['QUANTITY']        = pd.to_numeric(df['QUANTITY'],        errors='coerce').fillna(0)

    if targets_override:
        # For each row, if the product code is in the reference file, use that target
        # Only override where the data has 0 (unlisted/missing) — never reduce a non-zero target
        if 'PRODUCT CODE' in df.columns:
            code_col = 'PRODUCT CODE'
        else:
            code_col = None

        if code_col:
            def apply_override(row):
                code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                if code in targets_override:
                    # Use reference target if data target is 0 or missing
                    if row['PRESSURE TARGET'] == 0:
                        return float(targets_override[code])
                return row['PRESSURE TARGET']
            df['PRESSURE TARGET'] = df.apply(apply_override, axis=1)

    # ── OSA logic ─────────────────────────────────────────────────────────────
    df = df[~((df['PRESSURE TARGET'] == 0) & (df['STOCK LEVEL'] == 'Not available'))].copy()
    df.loc[(df['PRESSURE TARGET'] == 0) & (df['STOCK LEVEL'] == 'Available'),
           'PRESSURE TARGET'] = OSA_DEFAULT_PT
    df['OSA'] = (df['QUANTITY'] >= df['PRESSURE TARGET']).astype(int) * 100

    # ── Accounts ──────────────────────────────────────────────────────────────
    def get_account(name):
        if pd.isna(name): return 'UNKNOWN'
        n = str(name).upper().strip()
        for c in OSA_CHAIN_KEYWORDS:
            if c in n: return c
        return name
    df['ACCOUNT'] = df['CUSTOMER NAME'].apply(get_account)

    # ── Canonicalise PRODUCT CATEGORY to focus categories ────────────────────
    if 'PRODUCT CATEGORY' in df.columns:
        df['PRODUCT CATEGORY'] = (df['PRODUCT CATEGORY']
                                   .str.upper().str.strip()
                                   .map(lambda c: UG_CATEGORY_CANONICAL.get(c, c)))
        df = df[~df['PRODUCT CATEGORY'].isin(UG_EXCLUDED_CATEGORIES)].copy()

    return df


def build_osa_excel(df: pd.DataFrame) -> bytes:
    """Build the formatted OSA Excel workbook from a processed DataFrame."""
    if 'PRODUCT CATEGORY' in df.columns:
        df = df[~df['PRODUCT CATEGORY'].astype(str).str.upper().str.strip()
                .isin(UG_EXCLUDED_CATEGORIES)].copy()
    def make_pivot(data, rows):
        wp = [w for w in OSA_WEEKS if w in data['WEEK_LABEL'].values]
        pt = data.pivot_table(index=rows, columns='WEEK_LABEL', values='OSA', aggfunc='mean')
        pt = pt.reindex(columns=[w for w in wp if w in pt.columns])
        pt['Average'] = pt.mean(axis=1)
        return pt.round(1)

    def write_sheet(ws, pivot_df, title, row_labels):
        ws.freeze_panes = f'{get_column_letter(len(row_labels)+1)}3'
        all_cols   = list(pivot_df.columns)
        n_rc       = len(row_labels)
        total_cols = n_rc + len(all_cols)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        tc = ws.cell(1, 1, title)
        tc.font = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
        tc.fill = HDR_BLUE; tc.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 22
        for ci, lbl in enumerate(row_labels, 1):
            c = ws.cell(2, ci, lbl)
            c.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
            c.fill = HDR_MED; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = _bdr()
        for ci, col in enumerate(all_cols, n_rc+1):
            is_avg = col == 'Average'
            c = ws.cell(2, ci, col)
            c.font = Font(name='Calibri', bold=True, color='7F3F00' if is_avg else 'FFFFFF', size=10)
            c.fill = AVG_FILL if is_avg else HDR_MED
            c.alignment = Alignment(horizontal='center'); c.border = _bdr()
        ws.row_dimensions[2].height = 18
        pr = pivot_df.reset_index(); pf = None
        for ri, row in pr.iterrows():
            er = ri+3; bg = GRAY_FILL if ri%2==0 else None
            fv = row[row_labels[0]] if row_labels[0] in row.index else row.iloc[0]
            is_new = fv != pf
            for ci, lbl in enumerate(row_labels, 1):
                val = row[lbl] if lbl in row.index else row.iloc[ci-1]
                disp = (str(val) if pd.notna(val) else '') if (ci>1 or is_new) else ''
                c = ws.cell(er, ci, disp)
                c.font = Font(name='Calibri', size=10, bold=(ci==1))
                c.fill = bg or PatternFill()
                c.alignment = Alignment(vertical='center', horizontal='left', indent=1); c.border = _bdr()
            for ci, col in enumerate(all_cols, n_rc+1):
                val = row.get(col, np.nan)
                if pd.isna(val):
                    c = ws.cell(er, ci, '-')
                    c.font = Font(name='Calibri', size=10, color='BFBFBF'); c.fill = bg or PatternFill()
                else:
                    c = ws.cell(er, ci, round(val,1)/100); c.number_format = '0%'
                    f = _osa_fill_xl(val)
                    c.fill = f if f else (AVG_FILL if col=='Average' else (bg or PatternFill()))
                    c.font = Font(name='Calibri', size=10, bold=(col=='Average'))
                c.alignment = Alignment(horizontal='center', vertical='center'); c.border = _bdr()
            pf = fv
        for i in range(n_rc):
            lbl = row_labels[i]
            default_w = 40 if lbl == 'SKU' else (28 if i == 0 else 22)
            ws.column_dimensions[get_column_letter(i+1)].width = default_w
        for i in range(len(all_cols)):
            ws.column_dimensions[get_column_letter(n_rc+i+1)].width = 10

    wb    = openpyxl.Workbook(); first = True
    months = sorted(df['Month'].dropna().unique(), key=lambda m: datetime.strptime(m,'%B').month)

    # ── Build SKU label: pick a human-friendly description column when present,
    #     falling back to product code. Detect common description column variants
    #     so previously-saved frames (e.g. containing 'PRODUCT DESCRIPTION') are handled.
    df = df.copy()
    col_map = {c.upper().strip(): c for c in df.columns}
    desc_candidates = [
        'DESCRIPTION', 'PRODUCT DESCRIPTION', 'PRODUCT_DESCRIPTION',
        'PRODUCT NAME', 'PRODUCT_NAME', 'SKU', 'SKU NAME', 'SKU_NAME',
        'ITEM DESCRIPTION', 'ITEM_DESCRIPTION'
    ]
    desc_col = None
    for cand in desc_candidates:
        if cand in col_map:
            desc_col = col_map[cand]
            break
    if desc_col:
        df['SKU'] = df[desc_col].astype(str).str.strip()
    elif 'PRODUCT CODE' in df.columns:
        df['SKU'] = df['PRODUCT CODE'].astype(str).str.strip()

    configs = [
        ('Brand OSA',       ['BRAND NAME'],                 'OSA — BY BRAND'),
        ('Category OSA',    ['PRODUCT CATEGORY'],           'OSA — BY PRODUCT CATEGORY'),
        ('Account x Brand', ['ACCOUNT','BRAND NAME'],       'OSA — BY ACCOUNT & BRAND'),
        ('Account x Cat',   ['ACCOUNT','PRODUCT CATEGORY'], 'OSA — BY ACCOUNT & CATEGORY'),
    ]
    if 'SKU' in df.columns:
        configs += [
            ('SKU OSA',       ['SKU'],            'OSA — BY SKU'),
            ('Account x SKU', ['ACCOUNT','SKU'],  'OSA — BY ACCOUNT & SKU'),
        ]
    ms = []
    for month in months:
        md = df[df['Month']==month]
        wk = sorted(md['WEEK_LABEL'].dropna().unique(), key=lambda w: int(w.split()[-1]))
        ms.append((month, f"{len(md):,} rows | {len(wk)} week(s): {', '.join(wk)}"))
        for tb, rl, tt in configs:
            tn = f"{month[:3]} - {tb}"
            ws = wb.active if first else wb.create_sheet(tn)
            if first: ws.title = tn; first = False
            write_sheet(ws, make_pivot(md, rl), f"{tt} — {month.upper()} (Uganda)", rl)

    wl = wb.create_sheet('Legend')
    wl.merge_cells('A1:D1'); wl['A1'] = 'OSA REPORT — LEGEND & METHODOLOGY'
    wl['A1'].font = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
    wl['A1'].fill = HDR_BLUE; wl['A1'].alignment = Alignment(horizontal='center')
    notes = [('',''),('COLOR CODING',''),
        ('Green  ≥ 95%','On-shelf availability at or above target'),
        ('Amber  80–94%','Near target – monitor closely'),
        ('Red  < 80%','Below target – requires attention'),
        ('',''),('METHODOLOGY',''),
        ('Step 1','Only rows with a valid date in DATE REPORTED are included.'),
        ('Step 2','Items where Pressure Target=0 AND Stock=Not Available are dropped.'),
        ('Step 3',f'Listed items with Pressure Target=0 → default target set to {OSA_DEFAULT_PT}.'),
        ('Step 4','Quantity ≥ Pressure Target → 100%  |  Quantity < Pressure Target → 0%'),
        ('Step 5','Weeks numbered 1,2,3… in order within each month.'),
        ('Step 6','Mean OSA score per Brand / Category / SKU / Account grouping.'),
        ('',''),('DATA SUMMARY',''),
    ] + [(f'  {m}', s) for m, s in ms]
    for ri, (k, v) in enumerate(notes, 3):
        wl[f'A{ri}']=k; wl[f'B{ri}']=v
        wl[f'A{ri}'].font = Font(name='Calibri', bold=any(x in k for x in ('CODING','METHODOLOGY','DATA','Step')), size=10)
        wl[f'B{ri}'].font = Font(name='Calibri', size=10)
    wl.column_dimensions['A'].width = 32; wl.column_dimensions['B'].width = 80

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
#  SOS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

SOS_KEY_ACCOUNTS = [
    ('CARREFOUR',        ['CARREFOUR']),
    ('CAPITAL SHOPPERS', ['CAPITAL']),
    ('CAC',              ['CAC']),
    ('CHINA TOWN',       ['CHINA TOWN','CHINATOWN']),
    ('CITY JOY',         ['CITY JOY']),
    ('CYNIBEL',          ['CYNIBEL','CYNABEL']),
    ('CROWN STORES',     ['CROWN']),
    ('ECOMART',          ['ECOMART','ECO MART']),
    ('FRAINE',           ['FRAINE']),
    ('KENJOY',           ['KENJOY']),
    ('LEXORMART',        ['LEXORMART','LEXOR MART']),
    ('MASTERS',          ['MASTERS','MASTER']),
    ('MEGA',             ['MEGA']),
    ('PEOPLES',          ["PEOPLE'S",'PEOPLES']),
    ('QUALITY',          ['QUALITY']),
    ('QUALIWORTH',       ['QUALIWORTH']),
    ('QUICK PICK',       ['QUICK PICK']),
    ('S&S',              ['S&S']),
    ('SAVE MORE',        ['SAVE MORE','SAVE SUPERMARKET']),
    ('SHELL',            ['SHELL']),
    ('SHOPWISE',         ['SHOPWISE']),
    ('STANDARD',         ['STANDARD']),
    ('TMT',              ['TMT']),
    ('TOTAL',            ['TOTAL']),
]
SOS_GREEN_THR = 20
BLOCKS_PER_ROW = 5

# ─────────────────────────────────────────────────────────────────────────────
#  UGANDA BRAND & SOS TARGET CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# ── Definitive Uganda brand list — sourced from Items_Range_Uganda_March_26.xlsx ──
# Only these 5 brands are tracked. Anything else is excluded from all charts and filters.
# Prefixes cover all known SFA export variants (e.g. "Tishu Poa", "TISHUPOA", "TISHU POA").
UG_BRANDS = [
    'FAY',       # 53 SKUs — widest range across all categories
    'COSY',      # 1 SKU  — Serviettes
    'SIFA',      # 3 SKUs — Toilet Paper
    'TISHU POA', # 4 SKUs — Toilet Paper
    'ULTRA',     # 12 SKUs — Scouring Pads / Cleaners
]

# SFA sometimes exports brand names with variants — map all to canonical names above
UG_BRAND_ALIASES = {
    'TISHU':       'TISHU POA',
    'TISHUPOA':    'TISHU POA',
    'SIFA TP':     'SIFA',
    'COSY POA':    'COSY',
    'ULTRA CLEAN': 'ULTRA',
}

# SOS targets per category — exactly as per the MT SOS Target table
UG_SOS_TARGETS = {
    # Primary names (as shown in table)
    'TOILET PAPER':            30,
    'SERVIETTES':              30,
    'KITCHEN TOWELS':          40,
    'ALUMINUM FOILS':          40,
    'CLING FILM':              40,
    'FACIALS':                 40,
    'MULTIFOLDS HAND TOWELS':  50,
    'ULTRA SCOURERS':          20,
    'BAKING PAPER':            50,
    'WIPES':                   20,
    # Aliases that appear in SFA data exports
    'ALUMINIUM FOIL':          40,
    'ALLUMINIUM FOIL':         40,
    'HAND TOWEL':              50,
    'HAND TOWELS':             50,
    'MULTIFOLD HAND TOWELS':   50,
    'SCOURERS':                20,
    'SCOURING PADS':           20,
    'LARGE WIPES':             20,
    'POCKET WIPES':            20,
}

# The 10 focus categories shown in charts — matches the MT SOS Target table exactly.
# Any category NOT in this set is excluded from the category-level SOS charts.
UG_FOCUS_CATEGORIES = {
    'TOILET PAPER',
    'SERVIETTES',
    'KITCHEN TOWELS',
    'ALUMINUM FOILS',
    'CLING FILM',
    'FACIALS',
    'MULTIFOLDS HAND TOWELS',
    'ULTRA SCOURERS',
    'BAKING PAPER',
    'WIPES',
}

# Mapping from SFA export variant → canonical focus category name
UG_CATEGORY_CANONICAL = {
    'TOILET PAPER':           'TOILET PAPER',
    'SERVIETTES':             'SERVIETTES',
    'KITCHEN TOWELS':         'KITCHEN TOWELS',
    'ALUMINUM FOILS':         'ALUMINUM FOILS',
    'ALUMINIUM FOIL':         'ALUMINUM FOILS',
    'ALLUMINIUM FOIL':        'ALUMINUM FOILS',
    'CLING FILM':             'CLING FILM',
    'FACIALS':                'FACIALS',
    'FACIAL':                 'FACIALS',
    'MULTIFOLDS HAND TOWELS': 'MULTIFOLDS HAND TOWELS',
    'MULTIFOLD HAND TOWELS':  'MULTIFOLDS HAND TOWELS',
    'HAND TOWEL':             'MULTIFOLDS HAND TOWELS',
    'HAND TOWELS':            'MULTIFOLDS HAND TOWELS',
    'ULTRA SCOURERS':         'ULTRA SCOURERS',
    'SCOURERS':               'ULTRA SCOURERS',
    'SCOURING PADS':          'ULTRA SCOURERS',
    'BAKING PAPER':           'BAKING PAPER',
    'WIPES':                  'WIPES',
    'LARGE WIPES':            'WIPES',
    'POCKET WIPES':           'WIPES',
}

# Valid brand/category combinations within the tracked SOS category range.
# This prevents zero/competitor rows in wide SFA exports from creating invalid
# combinations such as ULTRA under SERVIETTES or COSY under SPONGE.
UG_BRAND_CATEGORY_RULES = {
    'FAY': {
        'ALUMINUM FOILS', 'BABY WIPES', 'BAKING PAPER', 'CLING FILM',
        'FACIALS', 'KITCHEN TOWELS', 'MULTIFOLDS HAND TOWELS',
        'POCKET TISSUE', 'POCKET TISSUES', 'SERVIETTES',
        'TOILET PAPER', 'WIPES',
    },
    'COSY':      {'SERVIETTES'},
    'SIFA':      {'TOILET PAPER'},
    'TISHU POA': {'TOILET PAPER'},
    'ULTRA':     {
        'ALL PURPOSE CLEANER', 'SPONGE', 'ULTRA SCOURERS',
        'WASHROOM HYGIENE PRODUCTS',
    },
}

UG_EXCLUDED_CATEGORIES = {'BAR SOAP'}

def is_valid_sos_brand_category(brand: str, category: str) -> bool:
    """Validate tracked SOS combinations; retain unknown/non-focus categories."""
    b = UG_BRAND_ALIASES.get(str(brand).upper().strip(), str(brand).upper().strip())
    raw_cat = str(category).upper().strip()
    cat = UG_CATEGORY_CANONICAL.get(raw_cat, raw_cat)
    if cat in UG_EXCLUDED_CATEGORIES:
        return False
    if b in UG_BRAND_CATEGORY_RULES:
        return cat in UG_BRAND_CATEGORY_RULES[b]
    return False

# ── Brand matching — used by both OSA ('BRAND NAME') and SOS ('PRODUCT_NAME') ──
# Exact prefixes derived from Items_Range_Uganda_March_26.xlsx
_UG_BRAND_PREFIXES = [b.upper() for b in UG_BRANDS]  # ['FAY', 'COSY', 'SIFA', 'TISHU POA', 'ULTRA']
_UG_ALIAS_KEYS     = [k.upper() for k in UG_BRAND_ALIASES]

def _matches_ug_brand(n: str) -> bool:
    """Core match: True if normalised uppercase name n starts with any UG brand prefix
    or exactly equals any known alias."""
    # Check aliases first (exact match on full name)
    if n in _UG_ALIAS_KEYS:
        return True
    # Check prefixes — name must START with a brand prefix
    return any(n.startswith(b) for b in _UG_BRAND_PREFIXES)

def is_ug_osa_brand(name: str) -> bool:
    """Return True if the OSA BRAND NAME belongs to a Uganda-tracked brand."""
    if not name or not isinstance(name, str) or name.lower() in ('nan', 'none', ''):
        return False
    return _matches_ug_brand(name.upper().strip())

def is_ug_brand(name: str) -> bool:
    """Return True if the SOS PRODUCT_NAME belongs to a Uganda-tracked brand."""
    if not name or not isinstance(name, str) or name.lower() in ('nan', 'none', ''):
        return False
        return False
    return _matches_ug_brand(name.upper().strip())

def get_sos_target(category: str) -> int:
    """Return the MT SOS target % for a given category."""
    if not category:
        return 20
    cat = str(category).upper().strip()
    return UG_SOS_TARGETS.get(cat, 20)  # default 20% if not in table

def sos_color_for_target(val: float, target: int) -> str:
    """Return 'green', 'amber', or 'red' based on actual vs category target."""
    if val >= target:        return 'green'
    if val >= target * 0.5:  return 'amber'
    return 'red'

GAP_COLS       = 1
ACCT_COLORS    = [
    ('1F4E79','FFFFFF'),('375623','FFFFFF'),('7B2C2C','FFFFFF'),
    ('4A3069','FFFFFF'),('7F4F1F','FFFFFF'),
]
_thin2 = Side(style='thin', color='BFBFBF')
def _sos_bdr(): return Border(left=_thin2, right=_thin2, top=_thin2, bottom=_thin2)

def _sos_fill_xl(val, target: int = None):
    """Colour-code an SOS value against its category target.
    Falls back to SOS_GREEN_THR (20%) when no target supplied.
    """
    t = target if target is not None else SOS_GREEN_THR
    if pd.isna(val) or val == 0:  return PatternFill('solid', fgColor='FFC7CE')  # red
    if val >= t:                   return PatternFill('solid', fgColor='C6EFCE')  # green
    return PatternFill('solid', fgColor='FFEB9C')                                 # amber


# ─────────────────────────────────────────────────────────────────────────────
#  SOS FORMAT DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def _detect_sos_format(df: pd.DataFrame) -> str:
    """
    Return 'survey' for the standard DATE CREATED/FACINGS SOS% format,
    or 'negotiated' for the SHOP NAME/BRAND GROUP/FACING SOS planogram format.
    """
    cols_lower = {c.lower().strip() for c in df.columns}
    # Negotiated format has these distinctive columns
    negotiated_signals = {"brand group", "facing sos", "length sos", "area sos", "shop name"}
    if len(negotiated_signals & cols_lower) >= 3:
        return "negotiated"
    return "survey"


def _map_chain(name: str) -> str | None:
    """Map a shop/outlet name to its parent chain using SOS_KEY_ACCOUNTS keywords."""
    if not name or (isinstance(name, float)):
        return None
    n = str(name).upper().strip()
    for label, kws in SOS_KEY_ACCOUNTS:
        if any(k in n for k in kws):
            return label
    # Also try common chains not in SOS_KEY_ACCOUNTS
    for chain in ['CARREFOUR','NAIVAS','QUICKMART','CHANDARANA','KHETIAS',
                  'CLEANSHELF','MAGUNAS','EASTMATT','DEFCO','MATHAI',
                  'POWERSTAR','KASSMATT','LEESTAR','SKYMATT','JAZA']:
        if chain in n:
            return chain
    return None  # unrecognised outlet — exclude from analysis


def merge_sos_chunks(chunk_bytes_list: list[bytes]) -> bytes:
    """
    Merge multiple SOS export files (same column structure) into one.
    Returns combined bytes as an in-memory xlsx so process_sos() can
    handle it with a single call.

    De-duplicates rows by all columns to handle overlapping exports.
    """
    dfs = []
    for fb in chunk_bytes_list:
        wb = openpyxl.load_workbook(io.BytesIO(fb), read_only=True, data_only=True)
        ws = wb.active
        row_iter = ws.iter_rows(values_only=True)
        header = next(row_iter, None)
        if header is None:
            wb.close()
            continue
        chunk_df = pd.DataFrame(row_iter, columns=header)
        wb.close()
        dfs.append(chunk_df)

    if not dfs:
        raise ValueError("No valid data found in uploaded files.")

    combined = pd.concat(dfs, ignore_index=True)
    # Drop exact duplicates that can appear in overlapping exports
    combined = combined.drop_duplicates()

    # Write back to xlsx bytes so process_sos() can consume it normally
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Worksheet")
    buf.seek(0)
    return buf.getvalue()


def process_sos_negotiated(file_bytes: bytes) -> pd.DataFrame:
    """
    Process the negotiated/planogram SOS format.
    Columns: ACCOUNT | SHOP NAME | REGION | CATEGORY | DISPLAY TYPE |
             BRAND GROUP | TARGET SOS | LENGTH SOS | FACING SOS | AREA SOS |
             LENGTH VARIANCE | FACING VARIANCE | AREA VARIANCE

    Returns a normalised DataFrame compatible with the analytics pages,
    with synthetic MONTH/MONTH_NUM set to 'Snapshot' / 0.
    """
    df = read_large_excel(file_bytes, sheet_name=0)

    # Normalise column names
    renames = {}
    col_lower = {c.lower().strip(): c for c in df.columns}
    alias_map = {
        'SHOP NAME':    ['shop name','outlet name','store name','customer name'],
        'BRAND GROUP':  ['brand group','brand','product_name','product name','sku'],
        'CATEGORY':     ['category','product category','cat'],
        'REGION':       ['region','territory','area'],
        'COUNTRY':      ['country','country name','country_name','nation'],
        'FACING SOS':   ['facing sos','facings sos','facings sos%','sos%','sos','facing sos%'],
        'TARGET SOS':   ['target sos','target','sos target'],
        'FACING VARIANCE': ['facing variance','variance'],
    }
    for canonical, aliases in alias_map.items():
        if canonical not in df.columns:
            for a in aliases:
                if a in col_lower:
                    renames[col_lower[a]] = canonical
                    break
    if renames:
        df = df.rename(columns=renames)

    # Ensure required columns exist
    for col in ('FACING SOS', 'TARGET SOS', 'FACING VARIANCE'):
        if col not in df.columns:
            df[col] = np.nan

    # Map shop to chain
    df['ACCOUNT'] = df['SHOP NAME'].apply(_map_chain)

    # Add synthetic time columns so analytics pages don't break
    df['DATE_PARSED'] = pd.NaT
    df['MONTH']       = 'Snapshot'
    df['MONTH_NUM']   = 0

    # Rename to canonical SOS analytics column names
    df = df.rename(columns={
        'BRAND GROUP': 'PRODUCT_NAME',
        'CATEGORY':    'PRODUCT_CATEGORY',
        'FACING SOS':  'FACINGS SOS%',
        'SHOP NAME':   'CUSTOMER NAME',
    })

    # Add POSITION as NaN (not in this format)
    if 'POSITION' not in df.columns:
        df['POSITION'] = np.nan

    # ── Coerce FACINGS SOS% to numeric ────────────────────────────────────────
    if 'FACINGS SOS%' in df.columns:
        sos_raw = df['FACINGS SOS%'].astype(str).str.strip().str.rstrip('%')
        df['FACINGS SOS%'] = pd.to_numeric(sos_raw, errors='coerce')
        frac_mask = df['FACINGS SOS%'].notna() & (df['FACINGS SOS%'] <= 1.0) & (df['FACINGS SOS%'] > 0)
        df.loc[frac_mask, 'FACINGS SOS%'] = df.loc[frac_mask, 'FACINGS SOS%'] * 100

    # ── Filter to Uganda brands only ──────────────────────────────────────────
    ug_mask = df['PRODUCT_NAME'].astype(str).apply(is_ug_brand)
    df = df[ug_mask].copy()

    # ── Filter to Uganda outlets only ─────────────────────────────────────────
    _UG_REGION_KEYWORDS = [
        'kampala', 'uganda', 'entebbe', 'jinja', 'mbarara', 'gulu',
        'wakiso', 'mukono', 'lira', 'mbale', 'arua', 'masaka',
    ]
    if 'COUNTRY' in df.columns:
        mask_ug = df['COUNTRY'].astype(str).str.strip().str.lower().isin(['uganda', 'ug'])
        df = df[mask_ug].copy()
    elif 'REGION' in df.columns:
        mask_ug = df['REGION'].astype(str).str.strip().str.lower().apply(
            lambda r: any(kw in r for kw in _UG_REGION_KEYWORDS) if r not in ('nan','none','') else False
        )
        df = df[mask_ug].copy()

    # ── Canonicalise category names ───────────────────────────────────────────
    if 'PRODUCT_CATEGORY' in df.columns:
        df['PRODUCT_CATEGORY'] = (df['PRODUCT_CATEGORY']
                                   .str.upper().str.strip()
                                   .map(lambda c: UG_CATEGORY_CANONICAL.get(c, c)))

    # ── Attach category SOS target ────────────────────────────────────────────
    if 'PRODUCT_CATEGORY' in df.columns:
        df['SOS_TARGET'] = df['PRODUCT_CATEGORY'].apply(get_sos_target)
    else:
        df['SOS_TARGET'] = 20

    return df


# ── SOS column aliases ────────────────────────────────────────────────────────
# Maps canonical name → list of possible names found in different SFA exports.
# Lower-cased for matching; original header case is preserved in the file.
_SOS_COL_ALIASES: dict[str, list[str]] = {
    "DATE CREATED":   ["date created", "date_created", "visit date", "visitdate",
                       "created date", "created_date", "date", "survey date",
                       "date reported", "date_reported", "submission date"],
    "CUSTOMER NAME":  ["customer name", "customer_name", "outlet name", "outlet",
                       "account name", "store name", "store", "shop name",
                       "customername", "client name", "outlet_name"],
    "PRODUCT_NAME":   ["product_name", "product name", "sku name", "sku",
                       "brand name", "brand", "item name", "item"],
    "PRODUCT_CATEGORY": ["product category", "product_category", "category",
                          "product cat", "prod category", "item category", "cat"],
    "FACINGS SOS%":   ["facings sos%", "facings sos", "sos%", "sos %", "sos",
                       "share of shelf", "share of shelf %", "facings%",
                       "facing sos%", "facing sos", "facing%"],
    "POSITION":       ["position", "shelf position", "shelf_position",
                       "pos", "shelf pos", "shelving position"],
    "COUNTRY":        ["country", "country name", "country_name", "nation"],
    "REGION":         ["region", "territory", "area", "district"],
}

def _resolve_sos_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename columns in the SOS dataframe to canonical names.
    Works case-insensitively and handles common export variants.
    Raises a clear ValueError listing available columns if a required
    column cannot be matched.
    """
    col_lower = {c.lower().strip(): c for c in df.columns}
    renames = {}
    missing = []

    for canonical, aliases in _SOS_COL_ALIASES.items():
        # Already present under canonical name — nothing to do
        if canonical in df.columns:
            continue
        # Try aliases in priority order
        matched = None
        for alias in aliases:
            if alias in col_lower:
                matched = col_lower[alias]
                break
        if matched:
            if matched != canonical:
                renames[matched] = canonical
        else:
            # POSITION, FACINGS SOS%, COUNTRY, REGION, and PRODUCT_CATEGORY are non-fatal — we can continue without them
            if canonical not in ("POSITION", "FACINGS SOS%", "COUNTRY", "REGION", "PRODUCT_CATEGORY"):
                missing.append(canonical)

    if missing:
        available = ", ".join(f"'{c}'" for c in df.columns[:20])
        raise ValueError(
            f"Could not find required column(s): {', '.join(missing)}. "
            f"Columns in your file: {available}"
        )

    if renames:
        df = df.rename(columns=renames)

    # Ensure optional columns exist as NaN if absent
    for col in ("POSITION", "FACINGS SOS%", "PRODUCT CATEGORY"):
        if col not in df.columns:
            df[col] = np.nan

    return df


def process_sos(file_bytes: bytes, filename: str = "") -> pd.DataFrame:
    """
    Rebuilt — content-aware column detection.

    Instead of trusting column headers (which vary across SFA exports), we scan
    the actual cell VALUES to identify each column. Brands and categories are
    hardcoded constants — nothing from outside those lists can appear in charts.

    Steps:
      1. Find BRAND column  — whichever column has the most Uganda brand matches
      2. Find SOS%  column  — numeric column with 'sos/share/facing' in header, or heuristic
      3. Find DATE  column  — parseable datetime column
      4. Find CUSTOMER col  — column whose values map to known key accounts
      5. Find CATEGORY col  — optional, by header keyword
      6. Find POSITION col  — optional
      7. Find REGION/COUNTRY — optional, for filter UI
    """
    df = read_large_excel(file_bytes, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    # ── STEP 1: Find BRAND column by scanning cell values ────────────────────
    _best_col, _best_score = None, 0
    for col in df.columns:
        try:
            vals = df[col].dropna().astype(str).str.upper().str.strip()
            score = int(vals.apply(is_ug_brand).sum())
            if score > _best_score:
                _best_score, _best_col = score, col
        except Exception:
            continue

    if _best_col is None or _best_score == 0:
        raise ValueError(
            "No Uganda brands found (Fay, Cosy, Sifa, Tishu Poa, Ultra). "
            "Please upload the correct SOS .xlsx file."
        )

    df = df.rename(columns={_best_col: 'PRODUCT_NAME'})
    df['PRODUCT_NAME'] = df['PRODUCT_NAME'].astype(str).str.strip()
    # HARD FILTER — only Uganda brands, no exceptions
    df = df[df['PRODUCT_NAME'].astype(str).apply(is_ug_brand)].copy()
    if df.empty:
        raise ValueError("No Uganda brand rows found after filtering.")
    # Normalise aliases (e.g. "Tishu" → "Tishu Poa")
    df['PRODUCT_NAME'] = df['PRODUCT_NAME'].apply(
        lambda x: UG_BRAND_ALIASES.get(x.upper().strip(), x.strip())
    )

    # ── STEP 2: Find SOS% column ─────────────────────────────────────────────
    _SOS_KW = ['sos', 'share', 'facing', 'shelf']
    _sos_col_found = False
    # Try header keyword first
    for col in df.columns:
        if col == 'PRODUCT_NAME': continue
        if any(kw in col.lower() for kw in _SOS_KW):
            cleaned = pd.to_numeric(df[col].astype(str).str.strip().str.rstrip('%'), errors='coerce')
            if cleaned.notna().sum() > 0:
                df['FACINGS SOS%'] = cleaned
                _sos_col_found = True
                break
    # Fallback: first numeric col whose values look like percentages
    if not _sos_col_found:
        for col in df.columns:
            if col in ('PRODUCT_NAME', 'FACINGS SOS%'): continue
            cleaned = pd.to_numeric(df[col].astype(str).str.strip().str.rstrip('%'), errors='coerce')
            valid = cleaned.dropna()
            if len(valid) > 0 and valid.between(0, 100).mean() > 0.7:
                df['FACINGS SOS%'] = cleaned
                break
    if 'FACINGS SOS%' not in df.columns:
        df['FACINGS SOS%'] = np.nan
    # Normalise 0–1 fractions to 0–100
    frac = df['FACINGS SOS%'].notna() & (df['FACINGS SOS%'] > 0) & (df['FACINGS SOS%'] <= 1)
    df.loc[frac, 'FACINGS SOS%'] = df.loc[frac, 'FACINGS SOS%'] * 100
    df = df[df['FACINGS SOS%'].notna()].copy()

    # ── STEP 3: Find DATE column ─────────────────────────────────────────────
    # Use smart date parsing: ISO format (YYYY-MM-DD...) must use dayfirst=False
    # to avoid swapping month/day (e.g. 2026-06-04 wrongly parsed as April).
    def _smart_parse_dates(series):
        sample = series.dropna().astype(str).head(20)
        # ISO format starts with 4-digit year
        is_iso = sample.str.match(r'^\d{4}[-/]\d').any()
        return pd.to_datetime(series, errors='coerce', dayfirst=not is_iso)

    _DATE_KW = ['date', 'created', 'visit', 'survey', 'submission', 'reported']
    _date_found = False
    for col in df.columns:
        if col in ('PRODUCT_NAME', 'FACINGS SOS%'): continue
        if any(kw in col.lower() for kw in _DATE_KW):
            parsed = _smart_parse_dates(df[col])
            if parsed.notna().sum() > len(df) * 0.3:
                df['DATE_PARSED'] = parsed
                _date_found = True
                break
    if not _date_found:
        for col in df.columns:
            if col in ('PRODUCT_NAME', 'FACINGS SOS%', 'DATE_PARSED'): continue
            parsed = _smart_parse_dates(df[col])
            if parsed.notna().sum() > len(df) * 0.3:
                df['DATE_PARSED'] = parsed
                break
    if 'DATE_PARSED' not in df.columns:
        df['DATE_PARSED'] = pd.NaT
    df = df[df['DATE_PARSED'].notna()].copy()
    df['MONTH']     = df['DATE_PARSED'].dt.strftime('%B')
    df['MONTH_NUM'] = df['DATE_PARSED'].dt.month

    # ── STEP 4: Find CUSTOMER column and map to key accounts ─────────────────
    def _get_account(name):
        if pd.isna(name): return None
        n = str(name).upper().strip()
        for label, kws in SOS_KEY_ACCOUNTS:
            if any(k in n for k in kws): return label
        return None

    _CUST_KW = ['customer', 'outlet', 'store', 'shop', 'client']
    _cust_found = False
    for col in df.columns:
        if col in ('PRODUCT_NAME', 'FACINGS SOS%', 'DATE_PARSED', 'MONTH', 'MONTH_NUM'): continue
        if any(kw in col.lower() for kw in _CUST_KW):
            accts = df[col].apply(_get_account)
            if accts.notna().sum() > 0:
                df['CUSTOMER NAME'] = df[col]
                df['ACCOUNT'] = accts
                _cust_found = True
                break
    # Fallback: scan all columns for account name matches
    if not _cust_found:
        for col in df.columns:
            if col in ('PRODUCT_NAME', 'FACINGS SOS%', 'DATE_PARSED', 'MONTH', 'MONTH_NUM'): continue
            try:
                accts = df[col].apply(_get_account)
                if accts.notna().sum() > len(df) * 0.05:
                    df['CUSTOMER NAME'] = df[col]
                    df['ACCOUNT'] = accts
                    break
            except Exception:
                continue

    if 'ACCOUNT' not in df.columns:
        raise ValueError(
            "Could not find an outlet/customer column matching Uganda key accounts "
            "(Carrefour, Shopwise, Masters, Fraine, etc.)."
        )
    df = df[df['ACCOUNT'].notna()].copy()

    # ── STEP 5: Find CATEGORY column (optional) ───────────────────────────────
    # Prefer columns with 'product' in name; skip anything with 'customer'
    _SKIP_COLS = {'PRODUCT_NAME', 'FACINGS SOS%', 'DATE_PARSED', 'MONTH',
                  'MONTH_NUM', 'CUSTOMER NAME', 'ACCOUNT'}
    _cat_col_found = False
    # First pass: look for 'product' + 'category' in header
    for col in df.columns:
        if col in _SKIP_COLS: continue
        if 'customer' in col.lower(): continue  # skip customer category
        cl = col.lower()
        if 'product' in cl and 'cat' in cl:
            df['PRODUCT_CATEGORY'] = (
                df[col].astype(str).str.upper().str.strip()
                .map(lambda c: UG_CATEGORY_CANONICAL.get(c, c))
            )
            _cat_col_found = True
            break
    # Second pass: any column with 'category' but not 'customer'
    if not _cat_col_found:
        for col in df.columns:
            if col in _SKIP_COLS: continue
            if 'customer' in col.lower(): continue
            if 'cat' in col.lower():
                df['PRODUCT_CATEGORY'] = (
                    df[col].astype(str).str.upper().str.strip()
                    .map(lambda c: UG_CATEGORY_CANONICAL.get(c, c))
                )
                break
    if 'PRODUCT_CATEGORY' not in df.columns:
        df['PRODUCT_CATEGORY'] = 'UNKNOWN'

    # Remove impossible brand/category pairs from wide SOS exports.
    df = df[
        df.apply(
            lambda row: is_valid_sos_brand_category(
                row['PRODUCT_NAME'], row['PRODUCT_CATEGORY']
            ),
            axis=1,
        )
    ].copy()

    # ── STEP 6: Find POSITION column (optional) ───────────────────────────────
    for col in df.columns:
        if 'position' in col.lower() or col.lower() in ('pos',):
            df['POSITION'] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            break
    if 'POSITION' not in df.columns:
        df['POSITION'] = 0

    # ── STEP 7: Find REGION / COUNTRY columns (optional) ─────────────────────
    for col in df.columns:
        if 'region' in col.lower() or 'territory' in col.lower():
            if 'REGION' not in df.columns: df['REGION'] = df[col]
    for col in df.columns:
        if 'country' in col.lower():
            if 'COUNTRY' not in df.columns: df['COUNTRY'] = df[col]

    # ── STEP 8: Attach SOS target per row ────────────────────────────────────
    df['SOS_TARGET'] = df['PRODUCT_CATEGORY'].apply(get_sos_target)

    # ── Keep only the columns the analytics page needs ────────────────────────
    keep = ['PRODUCT_NAME', 'FACINGS SOS%', 'POSITION', 'DATE_PARSED',
            'MONTH', 'MONTH_NUM', 'CUSTOMER NAME', 'ACCOUNT',
            'PRODUCT_CATEGORY', 'SOS_TARGET']
    if 'REGION'  in df.columns: keep.append('REGION')
    if 'COUNTRY' in df.columns: keep.append('COUNTRY')
    return df[[c for c in keep if c in df.columns]].copy()

def _build_sos_negotiated_excel(df: pd.DataFrame) -> bytes:
    """
    Build a formatted Excel report for the negotiated/planogram SOS format.
    Structure:
      - Summary sheet: Brand Group × Chain, FACING SOS% heatmap
      - Category sheets: one per PRODUCT_CATEGORY showing all shops
      - Region sheet: SOS by region
    """
    wb = openpyxl.Workbook()
    first = True

    brands    = sorted(df['PRODUCT_NAME'].dropna().unique())
    accounts  = sorted(df['ACCOUNT'].dropna().unique())
    cats      = sorted(df['PRODUCT_CATEGORY'].dropna().unique()) if 'PRODUCT_CATEGORY' in df.columns else []
    regions   = sorted(df['REGION'].dropna().unique()) if 'REGION' in df.columns else []

    def _sos_fill(val, target=20):
        if val is None or (isinstance(val, float) and np.isnan(val)): return None
        return GREEN_FILL if val >= target else (AMBER_FILL if val > 0 else RED_FILL)

    def _write_hdr(ws, row, col, val, bold=True, fill=None, sz=10):
        c = ws.cell(row, col, val)
        c.font = Font(name='Calibri', bold=bold, size=sz, color='FFFFFF' if fill else '1F4E79')
        if fill: c.fill = fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = _bdr()
        return c

    def _write_val(ws, row, col, val):
        c = ws.cell(row, col, round(float(val), 1) if isinstance(val, (int, float)) and not np.isnan(val) else '')
        fill = _sos_fill(val) if isinstance(val, (int, float)) and not np.isnan(val) else None
        if fill: c.fill = fill
        c.font = Font(name='Calibri', size=9)
        c.alignment = Alignment(horizontal='center')
        c.border = _bdr()
        return c

    # ── Sheet 1: Brand × Account summary ─────────────────────────────────────
    ws1 = wb.active; ws1.title = 'SOS Summary'
    ws1.freeze_panes = 'B2'
    _write_hdr(ws1, 1, 1, 'BRAND GROUP', fill=HDR_BLUE, sz=11)
    ws1.column_dimensions['A'].width = 28
    for ci, acct in enumerate(accounts, 2):
        _write_hdr(ws1, 1, ci, acct, fill=HDR_MED, sz=9)
        ws1.column_dimensions[get_column_letter(ci)].width = 14

    for ri, brand in enumerate(brands, 2):
        ws1.cell(ri, 1, brand).font = Font(name='Calibri', bold=True, size=9)
        ws1.cell(ri, 1).border = _bdr()
        for ci, acct in enumerate(accounts, 2):
            sub = df[(df['PRODUCT_NAME'] == brand) & (df['ACCOUNT'] == acct)]
            val = sub['FACINGS SOS%'].mean() if not sub.empty else np.nan
            _write_val(ws1, ri, ci, val)
        ws1.row_dimensions[ri].height = 14

    # ── Sheet 2: Category × Account ──────────────────────────────────────────
    ws2 = wb.create_sheet('SOS by Category')
    ws2.freeze_panes = 'B2'
    _write_hdr(ws2, 1, 1, 'CATEGORY', fill=HDR_BLUE, sz=11)
    ws2.column_dimensions['A'].width = 26
    for ci, acct in enumerate(accounts, 2):
        _write_hdr(ws2, 1, ci, acct, fill=HDR_MED, sz=9)
        ws2.column_dimensions[get_column_letter(ci)].width = 14
    for ri, cat in enumerate(cats, 2):
        ws2.cell(ri, 1, cat).font = Font(name='Calibri', bold=True, size=9)
        ws2.cell(ri, 1).border = _bdr()
        for ci, acct in enumerate(accounts, 2):
            sub = df[(df['PRODUCT_CATEGORY'] == cat) & (df['ACCOUNT'] == acct)] if 'PRODUCT_CATEGORY' in df.columns else pd.DataFrame()
            val = sub['FACINGS SOS%'].mean() if not sub.empty else np.nan
            _write_val(ws2, ri, ci, val)
        ws2.row_dimensions[ri].height = 14

    # ── Sheet 3: Shop-level detail ────────────────────────────────────────────
    ws3 = wb.create_sheet('Shop Detail')
    shops = sorted(df['CUSTOMER NAME'].dropna().unique()) if 'CUSTOMER NAME' in df.columns else []
    ws3.freeze_panes = 'B2'
    _write_hdr(ws3, 1, 1, 'BRAND GROUP', fill=HDR_BLUE, sz=11)
    ws3.column_dimensions['A'].width = 28
    for ci, shop in enumerate(shops, 2):
        _write_hdr(ws3, 1, ci, shop, fill=HDR_MED, sz=9)
        ws3.column_dimensions[get_column_letter(ci)].width = 16
    for ri, brand in enumerate(brands, 2):
        ws3.cell(ri, 1, brand).font = Font(name='Calibri', bold=True, size=9)
        ws3.cell(ri, 1).border = _bdr()
        for ci, shop in enumerate(shops, 2):
            sub = df[(df['PRODUCT_NAME'] == brand) & (df['CUSTOMER NAME'] == shop)] if 'CUSTOMER NAME' in df.columns else pd.DataFrame()
            val = sub['FACINGS SOS%'].mean() if not sub.empty else np.nan
            _write_val(ws3, ri, ci, val)
        ws3.row_dimensions[ri].height = 14

    # ── Sheet 4: Region summary ───────────────────────────────────────────────
    if regions:
        ws4 = wb.create_sheet('SOS by Region')
        ws4.freeze_panes = 'B2'
        _write_hdr(ws4, 1, 1, 'BRAND GROUP', fill=HDR_BLUE, sz=11)
        ws4.column_dimensions['A'].width = 28
        for ci, reg in enumerate(regions, 2):
            _write_hdr(ws4, 1, ci, reg, fill=HDR_MED, sz=9)
            ws4.column_dimensions[get_column_letter(ci)].width = 16
        for ri, brand in enumerate(brands, 2):
            ws4.cell(ri, 1, brand).font = Font(name='Calibri', bold=True, size=9)
            ws4.cell(ri, 1).border = _bdr()
            for ci, reg in enumerate(regions, 2):
                sub = df[(df['PRODUCT_NAME'] == brand) & (df['REGION'] == reg)] if 'REGION' in df.columns else pd.DataFrame()
                val = sub['FACINGS SOS%'].mean() if not sub.empty else np.nan
                _write_val(ws4, ri, ci, val)
            ws4.row_dimensions[ri].height = 14

    # ── Legend ────────────────────────────────────────────────────────────────
    wl = wb.create_sheet('Legend')
    wl.merge_cells('A1:C1')
    wl['A1'] = 'SOS NEGOTIATED REPORT — LEGEND'
    wl['A1'].font = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
    wl['A1'].fill = HDR_BLUE; wl['A1'].alignment = Alignment(horizontal='center')
    for ri, (k, v) in enumerate([
        ('',''), ('COLOR CODING',''),
        ('Green  ≥ 20%', 'Strong share of shelf'),
        ('Amber  > 0%',  'Present but below target'),
        ('Red = 0%',     'Not present / zero facing'),
        ('',''), ('SHEETS',''),
        ('SOS Summary',    'Mean FACING SOS% per Brand Group × Chain'),
        ('SOS by Category','Mean FACING SOS% per Category × Chain'),
        ('Shop Detail',    'FACING SOS% per Brand Group × individual outlet'),
        ('SOS by Region',  'Mean FACING SOS% per Brand Group × Region'),
    ], 3):
        wl[f'A{ri}'] = k; wl[f'B{ri}'] = v
        wl[f'A{ri}'].font = Font(name='Calibri', bold=any(x in k for x in ('COLOR','SHEETS')), size=10)
        wl[f'B{ri}'].font = Font(name='Calibri', size=10)
    wl.column_dimensions['A'].width = 28; wl.column_dimensions['B'].width = 60

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


def build_sos_excel(df: pd.DataFrame) -> bytes:
    """Build the formatted SOS Excel workbook from a processed DataFrame.
    Handles both survey format (has real months) and negotiated/snapshot format.
    """
    # Apply the validation here too so previously saved data cannot regenerate
    # invalid combinations after the processing rule has been corrected.
    if {'PRODUCT_NAME', 'PRODUCT_CATEGORY'}.issubset(df.columns):
        df = df[
            df.apply(
                lambda row: is_valid_sos_brand_category(
                    row['PRODUCT_NAME'], row['PRODUCT_CATEGORY']
                ),
                axis=1,
            )
        ].copy()
    is_negotiated = df['MONTH'].eq('Snapshot').all() if 'MONTH' in df.columns else False
    if is_negotiated:
        return _build_sos_negotiated_excel(df)

    months_ordered = (df[['MONTH','MONTH_NUM']].drop_duplicates()
                      .sort_values('MONTH_NUM')['MONTH'].tolist())
    brands = sorted(df['PRODUCT_NAME'].dropna().unique())
    accounts_with_data = [a for a,_ in SOS_KEY_ACCOUNTS if a in df['ACCOUNT'].values]

    # Build brand → SOS target lookup using category targets
    brand_sos_target = {}
    if 'SOS_TARGET' in df.columns and 'PRODUCT_CATEGORY' in df.columns:
        for brand in brands:
            sub = df[df['PRODUCT_NAME'] == brand]
            if not sub.empty and 'SOS_TARGET' in sub.columns:
                brand_sos_target[brand] = int(sub['SOS_TARGET'].mode()[0])
    # fallback
    for b in brands:
        brand_sos_target.setdefault(b, SOS_GREEN_THR)

    def catbrand_block_data(account):
        sub = df[df['ACCOUNT'] == account]
        if sub.empty: return None, []
        # (category, brand) pairs that actually exist for this account, category-ordered
        pairs = (sub[['PRODUCT_CATEGORY', 'PRODUCT_NAME']].drop_duplicates())
        pairs = pairs[pairs['PRODUCT_CATEGORY'].isin(cats)]
        pairs = pairs.sort_values(['PRODUCT_CATEGORY', 'PRODUCT_NAME'])
        row_order = list(pairs.itertuples(index=False, name=None))
        sp = sub.groupby(['PRODUCT_CATEGORY', 'PRODUCT_NAME', 'MONTH'])['FACINGS SOS%'].mean().round(1)
        r = {}
        for cat, brand in row_order:
            r[(cat, brand)] = {m: sp.get((cat, brand, m), np.nan) for m in months_ordered}
        return r, row_order

    def write_catbrand_block(ws, sr, sc, account, cbd, row_order):
        nm = len(months_ordered); bc = 2 + nm + 1  # Category, Brand, months..., Avg
        ws.merge_cells(start_row=sr, start_column=sc, end_row=sr, end_column=sc+bc-1)
        c = ws.cell(sr, sc, f'{account} — Category × Brand')
        c.font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F4E79')
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = _sos_bdr()
        ws.row_dimensions[sr].height = 18
        headers = ['Category', 'Brand'] + [m[:3] for m in months_ordered] + ['Avg']
        for hi, h in enumerate(headers):
            c = ws.cell(sr+1, sc+hi, h)
            c.font = Font(name='Calibri', bold=True, size=9, color='1F4E79' if hi < 2 else ('375623' if hi==len(headers)-1 else '1F4E79'))
            c.fill = PatternFill('solid', fgColor='E2EFDA' if hi==len(headers)-1 else 'D6E4F0')
            c.alignment = Alignment(horizontal='center'); c.border = _sos_bdr()
        ws.row_dimensions[sr+1].height = 16

        row = sr + 2
        prev_cat = None
        cat_start_row = row
        for cat, brand in row_order:
            alt = PatternFill('solid', fgColor='F2F2F2') if row % 2 == 0 else PatternFill()
            tgt = get_sos_target(cat)
            if cat != prev_cat:
                if prev_cat is not None and row - 1 > cat_start_row:
                    ws.merge_cells(start_row=cat_start_row, start_column=sc, end_row=row-1, end_column=sc)
                cat_start_row = row
                prev_cat = cat
            cc = ws.cell(row, sc, cat if row == cat_start_row else '')
            cc.font = Font(name='Calibri', bold=True, size=9)
            cc.fill = PatternFill('solid', fgColor='D6E4F0')
            cc.alignment = Alignment(horizontal='left', indent=1, vertical='center'); cc.border = _sos_bdr()

            bc2 = ws.cell(row, sc+1, brand)
            bc2.font = Font(name='Calibri', size=9); bc2.fill = alt
            bc2.alignment = Alignment(horizontal='left', indent=1, vertical='center'); bc2.border = _sos_bdr()

            vals = []
            for mi, month in enumerate(months_ordered):
                sv = cbd[(cat, brand)][month]
                c2 = ws.cell(row, sc+2+mi)
                if pd.isna(sv):
                    c2.value = '-'; c2.font = Font(name='Calibri', size=9, color='BFBFBF'); c2.fill = alt
                else:
                    c2.value = sv/100; c2.number_format = '0%'
                    c2.fill = _sos_fill_xl(sv, tgt); c2.font = Font(name='Calibri', size=9)
                    vals.append(sv)
                c2.alignment = Alignment(horizontal='center', vertical='center'); c2.border = _sos_bdr()
            av = round(np.mean(vals), 1) if vals else np.nan
            c3 = ws.cell(row, sc+2+nm)
            if pd.isna(av): c3.value = '-'; c3.font = Font(name='Calibri', size=9, color='BFBFBF'); c3.fill = alt
            else:
                c3.value = av/100; c3.number_format = '0%'
                c3.fill = _sos_fill_xl(av, tgt); c3.font = Font(name='Calibri', size=9, bold=True)
            c3.alignment = Alignment(horizontal='center', vertical='center'); c3.border = _sos_bdr()
            row += 1
        if row - 1 > cat_start_row:
            ws.merge_cells(start_row=cat_start_row, start_column=sc, end_row=row-1, end_column=sc)
        return row  # next free row

    def block_data(account):
        sub = df[df['ACCOUNT']==account]
        if sub.empty: return None
        sp = sub.groupby(['PRODUCT_NAME','MONTH'])['FACINGS SOS%'].mean().round(1)
        pp = sub[sub['POSITION']>0].groupby(['PRODUCT_NAME','MONTH'])['POSITION'].mean().round(1)
        r = {}
        for b in brands:
            r[b] = {m: (sp.get((b,m), np.nan), pp.get((b,m), np.nan)) for m in months_ordered}
        return r

    def write_block(ws, sr, sc, account, bd, color_idx):
        bg, fg = ACCT_COLORS[color_idx % len(ACCT_COLORS)]
        nm = len(months_ordered); bc = 1+nm+1
        ws.merge_cells(start_row=sr, start_column=sc, end_row=sr, end_column=sc+bc-1)
        c=ws.cell(sr,sc,account); c.font=Font(name='Calibri',bold=True,size=11,color=fg)
        c.fill=PatternFill('solid',fgColor=bg); c.alignment=Alignment(horizontal='center',vertical='center'); c.border=_sos_bdr()
        ws.row_dimensions[sr].height=18
        c=ws.cell(sr+1,sc,'Brand'); c.font=Font(name='Calibri',bold=True,size=9,color='1F4E79')
        c.fill=PatternFill('solid',fgColor='BDD7EE'); c.alignment=Alignment(horizontal='center'); c.border=_sos_bdr()
        if nm>1: ws.merge_cells(start_row=sr+1,start_column=sc+1,end_row=sr+1,end_column=sc+nm)
        c2=ws.cell(sr+1,sc+1,'SOS %'); c2.font=Font(name='Calibri',bold=True,size=9,color='1F4E79')
        c2.fill=PatternFill('solid',fgColor='BDD7EE'); c2.alignment=Alignment(horizontal='center'); c2.border=_sos_bdr()
        c3=ws.cell(sr+1,sc+nm+1,'Position'); c3.font=Font(name='Calibri',bold=True,size=9,color='375623')
        c3.fill=PatternFill('solid',fgColor='E2EFDA'); c3.alignment=Alignment(horizontal='center'); c3.border=_sos_bdr()
        c=ws.cell(sr+2,sc,''); c.fill=PatternFill('solid',fgColor='BDD7EE'); c.border=_sos_bdr()
        for mi,m in enumerate(months_ordered):
            c=ws.cell(sr+2,sc+1+mi,m[:3]); c.font=Font(name='Calibri',bold=True,size=8,color='1F4E79')
            c.fill=PatternFill('solid',fgColor='BDD7EE'); c.alignment=Alignment(horizontal='center'); c.border=_sos_bdr()
        c=ws.cell(sr+2,sc+nm+1,'Avg'); c.font=Font(name='Calibri',bold=True,size=8,color='375623')
        c.fill=PatternFill('solid',fgColor='E2EFDA'); c.alignment=Alignment(horizontal='center'); c.border=_sos_bdr()
        for bi,brand in enumerate(brands):
            dr=sr+3+bi; alt=PatternFill('solid',fgColor='F2F2F2') if bi%2==0 else PatternFill()
            c=ws.cell(dr,sc,brand); c.font=Font(name='Calibri',size=9); c.fill=alt
            c.alignment=Alignment(horizontal='left',indent=1,vertical='center'); c.border=_sos_bdr()
            pvs=[]
            for mi,month in enumerate(months_ordered):
                sv,pv=bd[brand][month]
                c2=ws.cell(dr,sc+1+mi)
                if pd.isna(sv): c2.value='-'; c2.font=Font(name='Calibri',size=9,color='BFBFBF'); c2.fill=alt
                else:
                    c2.value=sv/100; c2.number_format='0%'; c2.fill=_sos_fill_xl(sv, brand_sos_target.get(brand, SOS_GREEN_THR)); c2.font=Font(name='Calibri',size=9)
                c2.alignment=Alignment(horizontal='center',vertical='center'); c2.border=_sos_bdr()
                if not pd.isna(pv): pvs.append(pv)
            ap=round(np.mean(pvs),1) if pvs else np.nan
            c3=ws.cell(dr,sc+nm+1)
            if pd.isna(ap): c3.value='-'; c3.font=Font(name='Calibri',size=9,color='BFBFBF'); c3.fill=alt
            else: c3.value=ap; c3.fill=PatternFill('solid',fgColor='E2EFDA'); c3.font=Font(name='Calibri',size=9,bold=True)
            c3.alignment=Alignment(horizontal='center',vertical='center'); c3.border=_sos_bdr()

    wb=openpyxl.Workbook(); nc=1+len(months_ordered)+1
    ws_sum=wb.active; ws_sum.title='Summary'; ws_sum.sheet_view.showGridLines=False
    tw=BLOCKS_PER_ROW*(nc+GAP_COLS)
    ws_sum.merge_cells(start_row=1,start_column=1,end_row=1,end_column=tw)
    tc=ws_sum.cell(1,1,f'SHARE OF SHELF (SOS) — KEY ACCOUNTS — {", ".join(months_ordered).upper()}')
    tc.font=Font(name='Calibri',bold=True,size=14,color='FFFFFF')
    tc.fill=PatternFill('solid',fgColor='1F4E79'); tc.alignment=Alignment(horizontal='center',vertical='center')
    ws_sum.row_dimensions[1].height=24
    ndr=len(brands)+3; bh=ndr+2; BS=3
    for idx,account in enumerate(accounts_with_data):
        rb=idx//BLOCKS_PER_ROW; cb=idx%BLOCKS_PER_ROW
        sr=BS+rb*bh; sc=1+cb*(nc+GAP_COLS)
        bd=block_data(account)
        if bd is None: continue
        write_block(ws_sum,sr,sc,account,bd,idx)
    for band in range(BLOCKS_PER_ROW):
        base=1+band*(nc+GAP_COLS)
        ws_sum.column_dimensions[get_column_letter(base)].width=16
        for i in range(len(months_ordered)+1):
            ws_sum.column_dimensions[get_column_letter(base+1+i)].width=6
        if GAP_COLS: ws_sum.column_dimensions[get_column_letter(base+nc)].width=2
    # ── Category setup (moved earlier so it's available for per-account sheets) ─
    cats = sorted(df['PRODUCT_CATEGORY'].dropna().unique())
    cats = [c for c in cats if c not in ('UNKNOWN', 'nan', '')]

    for idx,account in enumerate(accounts_with_data):
        ws=wb.create_sheet(account[:25]); ws.sheet_view.showGridLines=False
        bd=block_data(account)
        if bd is None: continue
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=nc)
        abg,afg=ACCT_COLORS[idx%len(ACCT_COLORS)]
        tc=ws.cell(1,1,f'SOS REPORT — {account} — {", ".join(months_ordered).upper()}')
        tc.font=Font(name='Calibri',bold=True,size=13,color=afg)
        tc.fill=PatternFill('solid',fgColor=abg); tc.alignment=Alignment(horizontal='center',vertical='center')
        ws.row_dimensions[1].height=22
        write_block(ws,2,1,account,bd,idx)
        ws.column_dimensions[get_column_letter(1)].width=24
        for i in range(len(months_ordered)+1): ws.column_dimensions[get_column_letter(2+i)].width=12

        # ── Category × Brand section, below the Brand section on the same tab ──
        cbd, row_order = catbrand_block_data(account)
        if cbd and row_order:
            cb_start_row = 2 + (len(brands) + 3) + 2  # brand block height + gap
            write_catbrand_block(ws, cb_start_row, 1, account, cbd, row_order)
            ws.column_dimensions[get_column_letter(2)].width = 16
    if cats:
        HDR_CAT  = PatternFill('solid', fgColor='1F4E79')
        HDR_MED2 = PatternFill('solid', fgColor='2E75B6')

        ws_cat = wb.create_sheet('SOS by Category')
        ws_cat.sheet_view.showGridLines = False
        ws_cat.freeze_panes = 'B3'

        title_cols = max(len(accounts_with_data) * len(months_ordered), 4)
        ws_cat.merge_cells(start_row=1, start_column=1, end_row=1,
                           end_column=1 + title_cols)
        tc = ws_cat.cell(1, 1,
            f'SOS% BY CATEGORY — {", ".join(months_ordered).upper()}')
        tc.font = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
        tc.fill = HDR_CAT
        tc.alignment = Alignment(horizontal='center', vertical='center')
        ws_cat.row_dimensions[1].height = 22

        # Header row: Category | Acct1/Mo1 | Acct1/Mo2 | … | Acct2/Mo1 | …
        ws_cat.cell(2, 1, 'CATEGORY').font = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
        ws_cat.cell(2, 1).fill = HDR_CAT
        ws_cat.cell(2, 1).border = _sos_bdr()
        ws_cat.cell(2, 1).alignment = Alignment(horizontal='center')
        ws_cat.column_dimensions['A'].width = 28
        col_idx = 2
        acct_month_cols = []  # (account, month, col)
        for ai, acct in enumerate(accounts_with_data):
            for mi, month in enumerate(months_ordered):
                lbl = f'{acct[:12]}\n{month[:3]}'
                c = ws_cat.cell(2, col_idx, lbl)
                c.font = Font(name='Calibri', bold=True, size=8, color='FFFFFF')
                c.fill = HDR_MED2 if ai % 2 == 0 else HDR_CAT
                c.alignment = Alignment(horizontal='center', wrap_text=True)
                c.border = _sos_bdr()
                ws_cat.column_dimensions[get_column_letter(col_idx)].width = 9
                acct_month_cols.append((acct, month, col_idx))
                col_idx += 1
        ws_cat.row_dimensions[2].height = 28

        for ri, cat in enumerate(cats, 3):
            alt = PatternFill('solid', fgColor='EBF3FB') if ri % 2 == 0 else PatternFill()
            c = ws_cat.cell(ri, 1, cat)
            c.font = Font(name='Calibri', bold=True, size=9)
            c.fill = alt; c.border = _sos_bdr()
            c.alignment = Alignment(horizontal='left', indent=1)
            for acct, month, ci in acct_month_cols:
                sub = df[(df['PRODUCT_CATEGORY'] == cat) &
                         (df['ACCOUNT'] == acct) &
                         (df['MONTH'] == month)]
                val = sub['FACINGS SOS%'].mean() if not sub.empty else np.nan
                cv = ws_cat.cell(ri, ci)
                if pd.isna(val):
                    cv.value = '-'
                    cv.font = Font(name='Calibri', size=9, color='BFBFBF')
                    cv.fill = alt
                else:
                    tgt = get_sos_target(cat)
                    cv.value = val / 100
                    cv.number_format = '0%'
                    cv.fill = _sos_fill_xl(val, tgt)
                    cv.font = Font(name='Calibri', size=9)
                cv.alignment = Alignment(horizontal='center', vertical='center')
                cv.border = _sos_bdr()
            ws_cat.row_dimensions[ri].height = 14

    # ── Sheet: Category × Brand detail (per month) ───────────────────────────
    if cats:
        ws_cbd = wb.create_sheet('Category × Brand')
        ws_cbd.sheet_view.showGridLines = False
        ws_cbd.freeze_panes = 'C3'

        ws_cbd.merge_cells(start_row=1, start_column=1, end_row=1,
                           end_column=2 + len(accounts_with_data))
        tc2 = ws_cbd.cell(1, 1,
            f'SOS% — CATEGORY × BRAND × ACCOUNT — {", ".join(months_ordered).upper()}')
        tc2.font = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
        tc2.fill = PatternFill('solid', fgColor='1F4E79')
        tc2.alignment = Alignment(horizontal='center', vertical='center')
        ws_cbd.row_dimensions[1].height = 22

        # Header: Category | Brand | Acct1 | Acct2 | …
        for ci2, hdr in enumerate(['CATEGORY', 'BRAND'] + accounts_with_data, 1):
            c = ws_cbd.cell(2, ci2, hdr)
            c.font = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F4E79' if ci2 <= 2 else '2E75B6')
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            c.border = _sos_bdr()
        ws_cbd.column_dimensions['A'].width = 26
        ws_cbd.column_dimensions['B'].width = 14
        for ci2 in range(3, 3 + len(accounts_with_data)):
            ws_cbd.column_dimensions[get_column_letter(ci2)].width = 11
        ws_cbd.row_dimensions[2].height = 26

        row = 3
        for cat in cats:
            cat_brands = sorted(df[df['PRODUCT_CATEGORY'] == cat]['PRODUCT_NAME'].unique())
            if not cat_brands:
                continue
            first = True
            for brand in cat_brands:
                alt = PatternFill('solid', fgColor='EBF3FB') if row % 2 == 0 else PatternFill()
                # Category cell — merge vertically on first brand
                cc = ws_cbd.cell(row, 1, cat if first else '')
                cc.font = Font(name='Calibri', bold=True, size=9)
                cc.fill = PatternFill('solid', fgColor='D6E4F0')
                cc.alignment = Alignment(horizontal='left', indent=1, vertical='center')
                cc.border = _sos_bdr()
                first = False

                bc = ws_cbd.cell(row, 2, brand)
                bc.font = Font(name='Calibri', size=9)
                bc.fill = alt; bc.border = _sos_bdr()
                bc.alignment = Alignment(horizontal='left', indent=1)

                for ai2, acct in enumerate(accounts_with_data, 3):
                    sub2 = df[(df['PRODUCT_CATEGORY'] == cat) &
                              (df['PRODUCT_NAME'] == brand) &
                              (df['ACCOUNT'] == acct)]
                    val2 = sub2['FACINGS SOS%'].mean() if not sub2.empty else np.nan
                    cv2 = ws_cbd.cell(row, ai2)
                    if pd.isna(val2):
                        cv2.value = '-'
                        cv2.font = Font(name='Calibri', size=9, color='BFBFBF')
                        cv2.fill = alt
                    else:
                        tgt2 = get_sos_target(cat)
                        cv2.value = val2 / 100
                        cv2.number_format = '0%'
                        cv2.fill = _sos_fill_xl(val2, tgt2)
                        cv2.font = Font(name='Calibri', size=9)
                    cv2.alignment = Alignment(horizontal='center', vertical='center')
                    cv2.border = _sos_bdr()
                ws_cbd.row_dimensions[row].height = 14
                row += 1


    wl=wb.create_sheet('Legend'); wl.merge_cells('A1:B1')
    wl['A1']='SOS REPORT — LEGEND & METHODOLOGY'
    wl['A1'].font=Font(name='Calibri',bold=True,size=13,color='FFFFFF')
    wl['A1'].fill=PatternFill('solid',fgColor='1F4E79'); wl['A1'].alignment=Alignment(horizontal='center')
    for ri,(k,v) in enumerate([('',''),('METRICS',''),
        ('SOS%','Facings SOS% = Brand Facings ÷ Total Facings. Averaged across outlets per account per month.'),
        ('Position','Average shelf position (1=best). Only where brand has at least one facing.'),
        ('',''),('COLOR CODING (SOS%)',''),
        ('Green  ≥ 20%','Strong share of shelf'),('Amber   > 0%','Present but low share'),('Red = 0%','Not present'),
        ('',''),('DATA',''),
        ('Months covered',', '.join(months_ordered)),('Accounts included',', '.join(accounts_with_data)),
    ], 3):
        wl[f'A{ri}']=k; wl[f'B{ri}']=v
        wl[f'A{ri}'].font=Font(name='Calibri',bold=any(x in k for x in ('METRICS','COLOR','DATA')),size=10)
        wl[f'B{ri}'].font=Font(name='Calibri',size=10)
    wl.column_dimensions['A'].width=28; wl.column_dimensions['B'].width=82
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
#  MHSKU PERSISTENT REFERENCE  (file-backed JSON store)
# ─────────────────────────────────────────────────────────────────────────────

import json, os, pathlib, tempfile

_MHSKU_STORE = pathlib.Path(tempfile.gettempdir()) / "ug_mhsku_store.json"


def _load_store() -> dict:
    if _MHSKU_STORE.exists():
        try:
            return json.loads(_MHSKU_STORE.read_text())
        except Exception:
            pass
    return {}


def _save_store(data: dict):
    _MHSKU_STORE.write_text(json.dumps(data, ensure_ascii=False, indent=2))


def mhsku_save(records: list[dict], uploaded_by: str = ""):
    """
    Persist a list of MHSKU records.
    Each record: {code, sku, category, pressure_target}
    Merges on code — new upload wins for matching codes.
    """
    store = _load_store()
    existing = {r["code"]: r for r in store.get("records", [])}
    for r in records:
        existing[r["code"]] = r
    store["records"] = list(existing.values())
    store["last_updated"] = datetime.utcnow().isoformat() + "Z"
    store["last_updated_by"] = uploaded_by
    _save_store(store)


def mhsku_load() -> tuple[list[dict], str, str]:
    """Returns (records, last_updated, last_updated_by)."""
    store = _load_store()
    return (
        store.get("records", []),
        store.get("last_updated", ""),
        store.get("last_updated_by", ""),
    )


def mhsku_clear():
    """Wipe the entire MHSKU store. Admin only — caller must gate access."""
    _save_store({})


# ─────────────────────────────────────────────────────────────────────────────
#  UPLOADED DATA PERSISTENCE  (file-backed parquet store)
#  Stores the last-uploaded OSA and SOS DataFrames so the app remembers data
#  across page navigations and browser refreshes.
# ─────────────────────────────────────────────────────────────────────────────

_DATA_STORE_DIR = pathlib.Path(tempfile.gettempdir()) / "ug_analytics_store"
_DATA_STORE_DIR.mkdir(exist_ok=True)

_OSA_STORE_PATH  = _DATA_STORE_DIR / "osa_data.pkl"
_SOS_STORE_PATH  = _DATA_STORE_DIR / "sos_data.pkl"
_OSA_META_PATH   = _DATA_STORE_DIR / "osa_meta.json"
_SOS_META_PATH   = _DATA_STORE_DIR / "sos_meta.json"


def _save_df(df: pd.DataFrame, path: pathlib.Path, meta: dict, meta_path: pathlib.Path):
    import pickle
    try:
        with open(str(path), 'wb') as _f:
            pickle.dump(df, _f, protocol=4)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, default=str))
    except Exception:
        pass


def _load_df(path: pathlib.Path, meta_path: pathlib.Path):
    """Returns (df, meta) or (None, {})."""
    import pickle
    if not path.exists():
        return None, {}
    try:
        with open(str(path), 'rb') as _f:
            df = pickle.load(_f)
        meta = json.loads(meta_path.read_text()) if meta_path.exists() else {}
        return df, meta
    except Exception:
        return None, {}


def osa_data_save(df: pd.DataFrame, filename: str = "", uploaded_by: str = ""):
    from datetime import datetime as _dt
    meta = {"filename": filename, "uploaded_by": uploaded_by,
            "saved_at": _dt.now().isoformat(), "rows": len(df)}
    _save_df(df, _OSA_STORE_PATH, meta, _OSA_META_PATH)


def osa_data_load():
    """Returns (df, meta) — df is None if no data stored."""
    return _load_df(_OSA_STORE_PATH, _OSA_META_PATH)


def osa_data_clear():
    for p in (_OSA_STORE_PATH, _OSA_META_PATH):
        if p.exists(): p.unlink()


def sos_data_save(df: pd.DataFrame, filename: str = "", uploaded_by: str = ""):
    from datetime import datetime as _dt
    meta = {"filename": filename, "uploaded_by": uploaded_by,
            "saved_at": _dt.now().isoformat(), "rows": len(df)}
    _save_df(df, _SOS_STORE_PATH, meta, _SOS_META_PATH)


def sos_data_load():
    """Returns (df, meta) — df is None if no data stored."""
    return _load_df(_SOS_STORE_PATH, _SOS_META_PATH)


def sos_data_clear():
    for p in (_SOS_STORE_PATH, _SOS_META_PATH):
        if p.exists(): p.unlink()


def parse_mhsku_file(file_bytes: bytes) -> list[dict]:
    """
    Parse a MHSKU reference .xlsx file.
    Expects columns: Code | SKU | CATEGORIZED | PRESSURE TARGET
    Returns list of dicts: {code, sku, category, pressure_target}
    """
    import re as _re
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    # Prefer 'MHSKU' sheet, fall back to first sheet
    ws = wb["MHSKU"] if "MHSKU" in wb.sheetnames else wb.active
    records = []
    for row in ws.iter_rows(values_only=True):
        code, sku, cat, pt = (row[i] if len(row) > i else None for i in range(4))
        if not sku:
            continue
        sku_str = str(sku).strip()
        if sku_str.lower() in ("sku", ""):
            continue
        code_str = str(code).strip() if code else ""
        cat_str  = str(cat).strip()  if cat  else "MHSKU"
        nums = _re.findall(r"\d+", str(pt).strip()) if pt else []
        pt_val = int(nums[0]) if nums else 6
        records.append({"code": code_str, "sku": sku_str, "category": cat_str, "pressure_target": pt_val})
    return records


# ─────────────────────────────────────────────────────────────────────────────
#  LARGE-FILE OPTIMISED LOADERS  (handles up to ~300 MB)
# ─────────────────────────────────────────────────────────────────────────────

# Always stream — no size threshold needed. Streaming is always safer on cloud.

def _optimise_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Downcast numeric columns and convert low-cardinality object columns to
    category dtype.  Typically cuts RAM by 50-70% on wide survey exports.
    """
    for col in df.columns:
        col_dtype = df[col].dtype
        if col_dtype == object:
            n_unique = df[col].nunique(dropna=False)
            if n_unique / max(len(df), 1) < 0.5:
                try:
                    df[col] = df[col].astype('category')
                except Exception:
                    pass
        elif col_dtype == np.float64:
            try: df[col] = df[col].astype(np.float32)
            except Exception: pass
        elif col_dtype == np.int64:
            try: df[col] = df[col].astype(np.int32)
            except Exception: pass
    return df


def read_large_csv(file_bytes: bytes, chunksize: int = 50_000) -> pd.DataFrame:
    """Read a CSV in chunks to avoid peak-RAM spikes on large files."""
    chunks = []
    for chunk in pd.read_csv(
        io.BytesIO(file_bytes), chunksize=chunksize,
        low_memory=True, encoding_errors="replace"
    ):
        chunks.append(_optimise_dtypes(chunk))
    return pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame()


def _bytes_to_tempfile(file_bytes: bytes, suffix: str = ".xlsx") -> str:
    """Write bytes to a named temp file and return its path.

    Using a real file path avoids keeping a second io.BytesIO copy of the
    data in RAM while openpyxl streams through it.
    """
    import tempfile as _tf
    fd, path = _tf.mkstemp(suffix=suffix)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(file_bytes)
    except Exception:
        os.close(fd)
        raise
    return path


def read_large_excel(file_bytes: bytes, sheet_name=0) -> pd.DataFrame:
    """
    Stream-read an Excel file using openpyxl read_only mode.

    Key RAM savings vs naïve pd.read_excel():
    • Writes bytes to a temp file so openpyxl reads from disk, not a
      second in-memory BytesIO buffer.
    • Yields rows through a generator into pd.DataFrame() — never builds
      a full Python list of all rows before constructing the frame.
    • Calls wb.close() immediately after the header pass to release the
      workbook object before row iteration.

    This is safe up to ~300 MB on Streamlit Cloud's 1 GB RAM limit.
    """
    tmp = _bytes_to_tempfile(file_bytes, suffix=".xlsx")
    try:
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        # Resolve sheet
        if isinstance(sheet_name, str):
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        else:
            ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb.active

        row_iter = ws.iter_rows(values_only=True)
        header = next(row_iter, None)
        if header is None:
            wb.close()
            return pd.DataFrame()

        # Stream rows in chunks to avoid building a full Python list in RAM
        CHUNK = 50_000
        chunks = []
        while True:
            batch = list(itertools.islice(row_iter, CHUNK))
            if not batch:
                break
            chunks.append(_optimise_dtypes(pd.DataFrame(batch, columns=header)))
        wb.close()
        return pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame(columns=header)
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def smart_load(file_bytes: bytes, filename: str, sheet_name=0) -> pd.DataFrame:
    """
    Auto-detect format and load efficiently — handles CSV, XLS, XLSX up to 300 MB.
    For Excel files it first tries the named sheet ('OSA' or 'SOS'), then falls back.
    """
    fname = filename.lower()
    if fname.endswith(".csv"):
        return read_large_csv(file_bytes)
    try:
        return read_large_excel(file_bytes, sheet_name=sheet_name)
    except Exception:
        # Final fallback — treat as CSV (some SFA exports have wrong extension)
        return read_large_csv(file_bytes)
