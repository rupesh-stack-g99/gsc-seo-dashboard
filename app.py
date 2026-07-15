import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# CONFIGURATION & CSS
# =========================================================================
st.set_page_config(
    page_title="SEO Deficit Engine",
    page_icon="⚠️",
    layout="wide"
)

st.markdown("""
    <style>
    /* Slate Minimalist Styling */
    .stApp { background-color: #fafafa; }
    h1, h2, h3, h4 { color: #0f172a !important; font-family: monospace; }
    
    /* Status Headers */
    .alert-banner {
        background-color: #0f172a;
        color: #ffffff;
        padding: 24px;
        border-radius: 8px;
        font-family: monospace;
        margin-bottom: 25px;
        border-left: 6px solid #ef4444;
    }
    .alert-banner h2 { color: #ffffff !important; margin: 0 0 8px 0; }
    .alert-banner p { color: #94a3b8; margin: 0; font-size: 0.95rem; }
    
    /* Metrics panel */
    .metric-panel {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-val { font-size: 1.8rem; font-weight: 700; color: #ef4444; font-family: monospace; }
    .metric-lbl { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 4px; }
    </style>
""", unsafe_allow_html=True)

st.title("🚨 SEO Deficit & Defect Engine")
st.write("Upload your standard GSC raw ZIP export. The engine automatically filters, calculates, and generates clean action files of your underperforming data.")

# =========================================================================
# SYSTEM FILTERS
# =========================================================================
st.sidebar.header("🔧 Deficit Parameters")
BRAND_KEYWORD = st.sidebar.text_input("Brand Filter (Exclude)", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Minimum Impressions", min_value=1, value=100)
MAX_CANNIBAL_GAP = st.sidebar.slider("Overlap Rank Proximity", 1, 20, 10)

CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}

# =========================================================================
# UTILITIES: DATA PARSING
# =========================================================================
def parse_and_normalize(df, dimension_name):
    df.columns = [col.strip() for col in df.columns]
    target_col = next((col for col in df.columns if col.lower() in [dimension_name.lower(), 'query', 'page', 'device', 'country', 'search appearance', 'top ' + dimension_name.lower()]), None)
    if not target_col:
        return None
    
    normalized = pd.DataFrame()
    normalized[dimension_name] = df[target_col].astype(str).str.strip()
    
    clicks_col = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
    clicks_diff = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' in col.lower()), None)
    normalized['Clicks'] = pd.to_numeric(df[clicks_col], errors='coerce').fillna(0) if clicks_col else 0
    normalized['Clicks_Delta'] = pd.to_numeric(df[clicks_diff], errors='coerce').fillna(0) if clicks_diff else 0
    
    impr_col = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
    impr_diff = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' in col.lower()), None)
    normalized['Impressions'] = pd.to_numeric(df[impr_col], errors='coerce').fillna(0) if impr_col else 0
    normalized['Impressions_Delta'] = pd.to_numeric(df[impr_diff], errors='coerce').fillna(0) if impr_diff else 0

    ctr_col = next((col for col in df.columns if 'ctr' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
    if ctr_col:
        normalized['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
        normalized['CTR'] = pd.to_numeric(normalized['CTR'], errors='coerce').fillna(0.0)
    else:
        normalized['CTR'] = (normalized['Clicks'] / normalized['Impressions'] * 100).fillna(0.0)
    
    pos_col = next((col for col in df.columns if 'position' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
    pos_diff = next((col for col in df.columns if 'position' in col.lower() and 'difference' in col.lower()), None)
    normalized['Position'] = pd.to_numeric(df[pos_col], errors='coerce').fillna(99.0) if pos_col else 99.0
    normalized['Position_Delta'] = pd.to_numeric(df[pos_diff], errors='coerce').fillna(0.0) if pos_diff else 0.0
    
    return normalized

@st.cache_data
def process_gsc_zip(uploaded_file):
    extracted_dfs = {}
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            file_list = z.namelist()
            file_targets = {
                'Queries': next((f for f in file_list if "queries.csv" in f.lower()), None),
                'Pages': next((f for f in file_list if "pages.csv" in f.lower()), None)
            }
            for key, filename in file_targets.items():
                if filename:
                    with z.open(filename) as f:
                        df = pd.read_csv(f)
                        normalized = parse_and_normalize(df, key)
                        if normalized is not None:
                            extracted_dfs[key] = normalized
        return extracted_dfs
    except Exception as e:
        st.error(f"ZIP Unpacking Error: {e}")
        return None

def convert_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# =========================================================================
# APP EXECUTION & PIPELINES
# =========================================================================
uploaded_file = st.file_uploader("Drop GSC Export Zip here to run diagnostics:", type=["zip"])

if uploaded_file is not None:
    with st.spinner("Extracting parameters and running database joins..."):
        gsc_data = process_gsc_zip(uploaded_file)
        
    if gsc_data and 'Queries' in gsc_data and 'Pages' in gsc_data:
        df_queries = gsc_data['Queries']
        df_pages = gsc_data['Pages']
        
        # Apply brand filters
        if BRAND_KEYWORD:
            df_queries = df_queries[~df_queries['Queries'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
            
        # ---------------------------------------------------------------------
        # 1. OUTCOME DATASET: TRAFFIC LOSSES (Bleeders)
        # ---------------------------------------------------------------------
        df_bleed = df_queries[df_queries['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).copy()
        df_bleed['Clicks_Lost'] = df_bleed['Clicks_Delta'].abs().astype(int)
        df_bleed_clean = df_bleed[['Queries', 'Clicks', 'Clicks_Lost', 'Impressions', 'CTR', 'Position', 'Position_Delta']].copy()
        df_bleed_clean.columns = ['Query', 'Current Clicks', 'Clicks Lost vs Last Period', 'Impressions', 'CTR %', 'Current Rank', 'Rank Change']
        df_bleed_clean.index = np.arange(1, len(df_bleed_clean) + 1)

        # ---------------------------------------------------------------------
        # 2. OUTCOME DATASET: KEYWORD CLASHES (Cannibalization Mapping)
        # ---------------------------------------------------------------------
        cannibal_list = []
        queries_sorted = df_queries[df_queries['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False)
        for idx, q_row in queries_sorted.head(200).iterrows():
            query_txt = q_row['Queries']
            q_pos = q_row['Position']
            
            matching_urls = df_pages[
                (df_pages['Position'] >= q_pos - MAX_CANNIBAL_GAP) & 
                (df_pages['Position'] <= q_pos + MAX_CANNIBAL_GAP) &
                (df_pages['Impressions'] >= MIN_IMPRESSIONS / 2)
            ].sort_values(by=['Clicks', 'Impressions'], ascending=[False, False])
            
            if len(matching_urls) > 1:
                primary_url = matching_urls.iloc[0]['Pages']
                primary_pos = round(matching_urls.iloc[0]['Position'], 1)
                primary_clicks = int(matching_urls.iloc[0]['Clicks'])
                
                for sub_idx in range(1, min(len(matching_urls), 3)):
                    sub_row = matching_urls.iloc[sub_idx]
                    cannibal_url = sub_row['Pages']
                    cannibal_pos = round(sub_row['Position'], 1)
                    cannibal_clicks = int(sub_row['Clicks'])
                    
                    if cannibal_url != primary_url:
                        cannibal_list.append({
                            "Conflicting Query": query_txt,
                            "Primary Authority URL (KEEP)": primary_url,
                            "Primary Authority Rank": primary_pos,
                            "Primary Clicks": primary_clicks,
                            "Cannibal Competing URL (FIX)": cannibal_url,
                            "Cannibal Rank": cannibal_pos,
                            "Cannibal Clicks": cannibal_clicks,
                            "Rank Gap Offset": abs(primary_pos - cannibal_pos)
                        })
        df_clashes = pd.DataFrame(cannibal_list).drop_duplicates(subset=['Conflicting Query', 'Cannibal Competing URL (FIX)']) if cannibal_list else pd.DataFrame()
        if not df_clashes.empty:
            df_clashes.index = np.arange(1, len(df_clashes) + 1)

        # ---------------------------------------------------------------------
        # 3. OUTCOME DATASET: CLICK DEFICITS (Page 1 Underperforming CTR)
        # ---------------------------------------------------------------------
        ctr_gaps = []
        page1_queries = df_queries[(df_queries['Position'] <= 10) & (df_queries['Impressions'] >= MIN_IMPRESSIONS)]
        for _, row in page1_queries.iterrows():
            kw = row['Queries']
            clicks = row['Clicks']
            impr = row['Impressions']
            actual_ctr = row['CTR']
            pos = round(row['Position'])
            
            benchmark = CTR_BENCHMARKS.get(pos, 2.0)
            if actual_ctr < (benchmark * 0.7):
                lost_clicks = int((impr * (benchmark / 100)) - clicks)
                if lost_clicks > 0:
                    ctr_gaps.append({
                        "Keyword": kw,
                        "Rank": pos,
                        "Actual CTR": f"{round(actual_ctr, 1)}%",
                        "Expected Benchmark CTR": f"{benchmark}%",
                        "Estimated Click Deficit": lost_clicks,
                        "Total Impressions": int(impr)
                    })
        df_gaps = pd.DataFrame(ctr_gaps).sort_values(by='Estimated Click Deficit', ascending=False) if ctr_gaps else pd.DataFrame()
        if not df_gaps.empty:
            df_gaps.index = np.arange(1, len(df_gaps) + 1)

        # ---------------------------------------------------------------------
        # 4. OUTCOME DATASET: PAGE 2 LAGGARDS (Stuck in Positions 11-20)
        # ---------------------------------------------------------------------
        df_laggards = df_queries[
            (df_queries['Position'] >= 11.0) & 
            (df_queries['Position'] <= 20.0) & 
            (df_queries['Impressions'] >= MIN_IMPRESSIONS)
        ].sort_values(by='Impressions', ascending=False).copy()
        
        df_lag_clean = df_laggards[['Queries', 'Position', 'Impressions', 'Clicks', 'CTR']].copy()
        df_lag_clean.columns = ['Keyword Near Page 1', 'Current Rank', 'Lost Impressions Opportunity', 'Current Clicks Received', 'CTR %']
        df_lag_clean.index = np.arange(1, len(df_lag_clean) + 1)

        # =========================================================================
        # EXECUTIVE BRIEFING PANEL
        # =========================================================================
        st.markdown(f"""
        <div class="alert-banner">
            <h2>🚨 OUTCOME REPORT: PERFORMANCE LEAKS DISCOVERED</h2>
            <p>Your raw GSC data has been processed. We bypassed the positive performance indicators and extracted only your organic issues.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # High-level Damage Metrics
        c_left, c_mid1, c_mid2, c_right = st.columns(4)
        with c_left:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{len(df_bleed_clean)}</div><div class="metric-lbl">Bleeding Keywords</div></div>', unsafe_allow_html=True)
        with c_mid1:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{len(df_clashes)}</div><div class="metric-lbl">URL Ranking Fights</div></div>', unsafe_allow_html=True)
        with c_mid2:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{len(df_gaps)}</div><div class="metric-lbl">CTR Deficit Gaps</div></div>', unsafe_allow_html=True)
        with c_right:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{len(df_lag_clean)}</div><div class="metric-lbl">Page 2 Laggards</div></div>', unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)

        # =========================================================================
        # DIRECT TABLES WITH IMMEDIATE CSV DOWNLOADS
        # =========================================================================
        
        # Section 1: Traffic Bleeders
        st.subheader("📉 1. Traffic Bleeders (Largest Click Losses)")
        st.write("These keywords lost the most traffic over the selected time range compared to the prior period.")
        if not df_bleed_clean.empty:
            st.dataframe(df_bleed_clean, use_container_width=True)
            st.download_button("💾 Download Traffic Loss Data (CSV)", data=convert_to_csv(df_bleed_clean), file_name="gsc_traffic_losses.csv", mime="text/csv")
        else:
            st.info("No negative click trends detected.")

        st.markdown("---")

        # Section 2: Keyword Clashes
        st.subheader("🎯 2. Authority Clashes (Keyword Cannibalization Map)")
        st.write("These conflicting URLs are competing with each other in Google's SERP indices for identical search terms.")
        if not df_clashes.empty:
            st.dataframe(df_clashes, use_container_width=True)
            st.download_button("💾 Download Clash Data (CSV)", data=convert_to_csv(df_clashes), file_name="gsc_keyword_clashes.csv", mime="text/csv")
        else:
            st.info("No active URL cannibalization mapped.")

        st.markdown("---")

        # Section 3: CTR Underachievers
        st.subheader("📈 3. Click Efficiency Gaps (Underperforming Page-1 CTR)")
        st.write("These keywords rank on Page 1 but perform below search-layout standards. Their Meta titles require immediate optimization.")
        if not df_gaps.empty:
            st.dataframe(df_gaps, use_container_width=True)
            st.download_button("💾 Download CTR Deficit Data (CSV)", data=convert_to_csv(df_gaps), file_name="gsc_ctr_deficits.csv", mime="text/csv")
        else:
            st.info("No structural CTR issues found on Page 1.")

        st.markdown("---")

        # Section 4: Page 2 Opportunities
        st.subheader("🚀 4. Page 2 Laggards (Keywords Stuck in Positions 11–20)")
        st.write("These keywords are trapped on Page 2 with massive impression scale, leaving traffic on the table.")
        if not df_lag_clean.empty:
            st.dataframe(df_lag_clean, use_container_width=True)
            st.download_button("💾 Download Page-2 Laggards (CSV)", data=convert_to_csv(df_lag_clean), file_name="gsc_page2_laggards.csv", mime="text/csv")
        else:
            st.info("No high-impression Page-2 opportunities detected.")

    else:
        st.error("❌ The uploaded GSC ZIP does not contain clean 'queries.csv' or 'pages.csv' data sets.")
