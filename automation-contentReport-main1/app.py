import streamlit as st
from pathlib import Path
import pandas as pd

st.set_page_config(
    page_title="Mercedes-Benz | Content Overview Generator",
    page_icon="automation-contentReport-main1/mercedes-logo.png",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
html, body { background:transparent !important; }
.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] { background-color:var(--background-color,#0e1116) !important; color:var(--text-color,#f3f5f6) !important; }
[data-testid="stHeader"] { background:var(--background-color,#0e1116); }
[data-testid="stToolbar"] { right:1.5rem; }
.block-container { max-width:1280px; padding-top:5rem; padding-bottom:4rem; }
header { border-bottom:0; }
.brandbar { min-height:74px; display:flex; align-items:center; gap:17px; border-bottom:1px solid color-mix(in srgb,var(--text-color,#f3f5f6) 16%,transparent); margin-bottom:60px; padding:8px 0; }
.brandbar img { width:46px; height:46px; object-fit:contain; }
.brandname { font-family:'Manrope',sans-serif; font-size:14px; font-weight:700; letter-spacing:.13em; text-transform:uppercase; }

.eyebrow { color:color-mix(in srgb,var(--text-color,#f3f5f6) 72%,transparent); font-size:11px; font-weight:700; letter-spacing:.18em; text-transform:uppercase; margin-bottom:12px; }
h1 { font-family:'Manrope',sans-serif !important; font-size:clamp(2.25rem,4vw,3.5rem) !important; line-height:1.08 !important; letter-spacing:-.055em !important; font-weight:600 !important; color:var(--text-color,#f3f5f6) !important; margin:0 0 16px !important; }

[data-testid="stVerticalBlockBorderWrapper"] { background:var(--secondary-background-color,#171b21); border:1px solid color-mix(in srgb,var(--text-color,#f3f5f6) 16%,transparent); border-radius:4px; padding:20px 28px 8px; box-shadow:0 12px 36px rgba(0,0,0,.12); }
label, [data-testid="stWidgetLabel"] p { color:var(--text-color,#f3f5f6) !important; font-weight:600 !important; font-size:13px !important; }
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea, [data-testid="stSelectbox"] [data-baseweb="select"] > div { background:var(--background-color,#22272e) !important; border:1px solid color-mix(in srgb,var(--text-color,#f3f5f6) 16%,transparent) !important; border-radius:3px !important; color:var(--text-color,#f3f5f6) !important; }
[data-testid="stTextInput"] input::placeholder, [data-testid="stTextArea"] textarea::placeholder { color:color-mix(in srgb,var(--text-color,#f3f5f6) 62%,transparent) !important; opacity:1; }
[data-testid="stTextInput"] input:focus, [data-testid="stTextArea"] textarea:focus, [data-testid="stSelectbox"] [data-baseweb="select"]:focus-within > div { border-color:var(--text-color,#b8c0c5) !important; box-shadow:0 0 0 1px var(--text-color,#b8c0c5) !important; }
.stButton button { background:var(--text-color,#111820) !important; border:1px solid var(--text-color,#111820) !important; color:var(--background-color,#fff) !important; border-radius:2px !important; padding:.62rem 1.15rem !important; font-weight:600 !important; letter-spacing:.01em; }
.stButton button:hover { opacity:.82; }
.footerline { border-top:1px solid color-mix(in srgb,var(--text-color,#f3f5f6) 16%,transparent); margin-top:70px; padding-top:18px; color:color-mix(in srgb,var(--text-color,#f3f5f6) 62%,transparent); font-size:11px; letter-spacing:.08em; text-transform:uppercase; }
@media(max-width:700px) { .brandbar { margin-bottom:38px; } [data-testid="stVerticalBlockBorderWrapper"] { padding:16px 14px 6px; } }
</style>
""", unsafe_allow_html=True)

@st.cache_data

def load_campaigns(excel_path):
    frame = pd.read_excel(excel_path, usecols=["Campaign"])
    values = frame["Campaign"].dropna().astype(str).str.strip()
    return sorted((value for value in values.unique() if value), key=str.casefold)


excel_path = Path(__file__).resolve().parents[1] / "input" / "Content_Overview_Enriched.xlsx"
try:
    campaigns = load_campaigns(excel_path)
except Exception as exc:
    campaigns = []
    campaign_load_error = str(exc)
else:
    campaign_load_error = None

logo_path = Path(__file__).with_name("mercedes-logo.png")
logo_data = logo_path.read_bytes()
import base64
logo_uri = "data:image/png;base64," + base64.b64encode(logo_data).decode("ascii")

st.markdown(f"""
<div class="brandbar">
  <img src="{logo_uri}" alt="Mercedes-Benz star">
  <span class="brandname">Mercedes-Benz</span>
  </div>
<div class="eyebrow">Campaign reporting</div>
<h1>Content Overview<br>Generator</h1>
""", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("**Campaign Name**")
    if campaigns:
        campaign = st.selectbox(
            "Select a campaign",
            options=campaigns,
            label_visibility="collapsed",
        )
    else:
        campaign = ""
        st.error(f"Could not load campaign names from the Excel file{': ' + campaign_load_error if campaign_load_error else '.'}")
    asset_ids = st.text_area(
        "Symphony Asset IDs",
        placeholder="Paste one or more asset IDs, separated by commas or new lines",
        height=130,
    )

    if st.button("Generate PowerPoint", type="primary"):        #test
        st.success("PowerPoint generated!")
        st.write("Campaign:", campaign)
        st.write("Asset IDs:", asset_ids)

st.markdown("<div class='footerline'>Mercedes-Benz &middot; Content Overview Generator</div>", unsafe_allow_html=True)


