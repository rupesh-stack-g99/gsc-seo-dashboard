import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# Ensure dependencies are available
try:
    import xlsxwriter
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])
    import xlsxwriter

# =========================================================================
# 1. STYLE ENGINE & STYLING CONFIG
# =========================================================================
st.set_page_config(
    page_title="GSC Performance Deficit",
    page_icon="🎯",
    layout="wide"
)

st.markdown("""
    <style>
    .stApp { background-color: #fafafa; }
    h1, h2, h3, h4 { color: #0f172a !important; font-family: monospace; }
    
    .alert-banner {
        background-color: #0f172a;
        color: #ffffff;
        padding: 20px;
        border-radius: 8px;
        font-family: monospace;
        margin-bottom: 25px;
        border-left: 6px solid #ef4444;
    }
    .alert-banner h2 { color: #ffffff !important; margin: 0 0 8px 0; }
    .alert-banner p { color: #94a3b8; margin: 0; font-size: 0.95rem; }
    
    .metric-panel {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-val { font-size: 1.8rem; font-weight: 700; color: #10b981; font-family: monospace; }
    .metric-val.red-val { color: #ef4444; }
    .metric-lbl { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 4px; }
    </style>
""", unsafe_allow_html=True)

st.title("🎯 GSC Deficit Engine (Lightweight)")
st.write("Clean SEO Data Processing | Positions $\le$ 30 Only | Excludes UTM Parameters")

# =========================================================================
# 2. DATA HARVESTING & CLEANING PIPELINE
# =========================================================================
def clean_and_normalize(df, dim_name):
    df.columns = [col.strip() for col in df.columns]
    target_col = next((col for col in df.columns if col.lower() in [
        dim_name.lower(), 'query', 'page', 'device', 'country', 'top ' + dim_name.lower()
    ]), None)
    
    if not target_col:
        return None
        
    norm = pd.DataFrame()
    norm[dim_name] = df[target_col].astype(str).str.strip()
    
    # Scrub UTM tracked strings instantly
    if dim_name == 'Pages':
        norm = norm[~norm['Pages'].str.lower().str.contains('utm_|_utm|utm=', na=False)]
        
    def find_and_parse(keywords, default=0.0):
        col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
        diff_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' in c.lower()), None)
        
        val = pd.to_numeric(df[col], errors='coerce').fillna(default) if col else pd.Series(default, index=df.index)
        val_delta = pd.to_numeric(df[diff_col], errors='coerce').fillna(0.0) if diff_col else pd.Series(0.0, index=df.index)
        return val, val_delta

    norm['Clicks'], norm['Clicks_Delta'] = find_and_parse(['click'])
    norm['Impressions'], norm['Impressions_Delta'] = find_and_parse(['impression'])
    norm['Position'], norm['Position_Delta'] = find_and_parse(['position'], default=99.0)
    
    # Strict Position Cap
    norm = norm[norm['Position'] <= 30.0]
    return norm

def extract_gsc_zip(uploaded_file):
    parsed = {}
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            names = z.namelist()
            mappings = {'Queries': 'queries.csv', 'Pages': 'pages.csv'}
            for key, pattern in mappings.items():
                matched = next((n for n in names if pattern in n.lower()), None)
                if matched:
                    with z.open(matched) as f:
                        df = pd.read_csv(f)
                        norm_df = clean_and_normalize(df, key)
                        if norm_df is not None:
                            parsed[key] = norm_df
        return parsed
    except Exception:
        return None

# =========================================================================
# 3. INTERACTIVE LAYOUT & RENDER
# =========================================================================
uploaded_file = st.file_uploader("Upload GSC ZIP export:", type=["zip"])

if uploaded_file is not None:
    gsc = extract_gsc_zip(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q = gsc['Queries']
        df_p = gsc['Pages']
        
        # Simple KPIs
        total_clicks = int(df_q['Clicks'].sum())
        total_impr = int(df_q['Impressions'].sum())
        avg_pos = round(df_q['Position'].mean(), 1)
        
        # Render Metrics
        st.markdown(f"""
        <div class="alert-banner">
            <h2>📊 SITE VISIBILITY SNAPSHOT ($\le$ Position 30)</h2>
            <p>Database processed. Campaign parameters removed automatically.</p>
        </div>
        """, unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{total_clicks:,}</div><div class="metric-lbl">Total Clicks</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{total_impr:,}</div><div class="metric-lbl">Total Impressions</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{avg_pos}</div><div class="metric-lbl">Average Position</div></div>', unsafe_allow_html=True)
            
        st.markdown("<br/>", unsafe_allow_html=True)
        
        # Tabs for Raw Clean Data
        tab_queries, tab_pages = st.tabs(["🔑 Active Queries", "📄 Clean Landing Pages"])
        
        with tab_queries:
            st.subheader("Queries Ranking up to Position 30")
            st.dataframe(df_q, use_container_width=True)
            
        with tab_pages:
            st.subheader("Pages Ranking up to Position 30 (No UTMs)")
            st.dataframe(df_p, use_container_width=True)
            
    else:
        st.error("❌ ZIP file must contain both 'queries.csv' and 'pages.csv'.")
