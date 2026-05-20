"""
Shared data processing engines and chart theme for OSA & SOS.
Imported by all pages.
"""

import io
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

def sos_color(val):
    t = get_theme()
    if __import__('pandas').isna(val) or val == 0: return t["red"]
    if val >= 20: return t["green"]
    return t["amber"]

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

    # ── Load ──────────────────────────────────────────────────────────────────
    if fname.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        # Try Excel with OSA sheet first, fall back to first sheet, then CSV
        try:
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='OSA')
            except Exception:
                df = pd.read_excel(io.BytesIO(file_bytes))
        except Exception:
            # Last resort: try as CSV (some exports have wrong extension)
            df = pd.read_csv(io.BytesIO(file_bytes))

    # ── Parse dates ───────────────────────────────────────────────────────────
    df['DATE_PARSED'] = pd.to_datetime(df['DATE REPORTED'], errors='coerce')
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
    return df


def build_osa_excel(df: pd.DataFrame) -> bytes:
    """Build the formatted OSA Excel workbook from a processed DataFrame."""
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
            ws.column_dimensions[get_column_letter(i+1)].width = 28 if i==0 else 22
        for i in range(len(all_cols)):
            ws.column_dimensions[get_column_letter(n_rc+i+1)].width = 10

    wb    = openpyxl.Workbook(); first = True
    months = sorted(df['Month'].dropna().unique(), key=lambda m: datetime.strptime(m,'%B').month)
    configs = [
        ('Brand OSA',       ['BRAND NAME'],                 'OSA — BY BRAND'),
        ('Category OSA',    ['PRODUCT CATEGORY'],           'OSA — BY PRODUCT CATEGORY'),
        ('Account x Brand', ['ACCOUNT','BRAND NAME'],       'OSA — BY ACCOUNT & BRAND'),
        ('Account x Cat',   ['ACCOUNT','PRODUCT CATEGORY'], 'OSA — BY ACCOUNT & CATEGORY'),
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
        ('Step 6','Mean OSA score per Brand / Category / Account grouping.'),
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
GAP_COLS       = 1
ACCT_COLORS    = [
    ('1F4E79','FFFFFF'),('375623','FFFFFF'),('7B2C2C','FFFFFF'),
    ('4A3069','FFFFFF'),('7F4F1F','FFFFFF'),
]
_thin2 = Side(style='thin', color='BFBFBF')
def _sos_bdr(): return Border(left=_thin2, right=_thin2, top=_thin2, bottom=_thin2)

def _sos_fill_xl(val):
    if pd.isna(val) or val==0: return PatternFill('solid', fgColor='FFC7CE')
    if val >= SOS_GREEN_THR:   return PatternFill('solid', fgColor='C6EFCE')
    return PatternFill('solid', fgColor='FFEB9C')


def process_sos(file_bytes: bytes) -> pd.DataFrame:
    """Run the full SOS pipeline and return the enriched DataFrame."""
    df = pd.read_excel(io.BytesIO(file_bytes))
    df['DATE_PARSED'] = pd.to_datetime(df['DATE CREATED'], errors='coerce')
    df = df[df['DATE_PARSED'].notna()].copy()
    df['MONTH']     = df['DATE_PARSED'].dt.strftime('%B')
    df['MONTH_NUM'] = df['DATE_PARSED'].dt.month

    def get_account(name):
        if pd.isna(name): return None
        n = str(name).upper().strip()
        for label, kws in SOS_KEY_ACCOUNTS:
            if any(k in n for k in kws): return label
        return None
    df['ACCOUNT'] = df['CUSTOMER NAME'].apply(get_account)
    df = df[df['ACCOUNT'].notna()].copy()
    return df


def build_sos_excel(df: pd.DataFrame) -> bytes:
    """Build the formatted SOS Excel workbook from a processed DataFrame."""
    months_ordered = (df[['MONTH','MONTH_NUM']].drop_duplicates()
                      .sort_values('MONTH_NUM')['MONTH'].tolist())
    brands = sorted(df['PRODUCT_NAME'].dropna().unique())
    accounts_with_data = [a for a,_ in SOS_KEY_ACCOUNTS if a in df['ACCOUNT'].values]

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
                    c2.value=sv/100; c2.number_format='0%'; c2.fill=_sos_fill_xl(sv); c2.font=Font(name='Calibri',size=9)
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
