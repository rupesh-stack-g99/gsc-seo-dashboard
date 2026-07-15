import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# PAGE CONFIGURATION & LAYOUT
# =========================================================================
st.set_page_config(
    page_title="Ultimate SEO Expert Strategy Hub",
    page_icon="🚀",
    layout="wide"
)

# Custom Styling for SEO Agency-Grade Reports
st.markdown("""
    <style>
    .expert-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        margin-bottom: 25px;
    }
    .action-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        border-left: 6px solid #ff4b4b;
        margin-bottom: 15px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04);
    }
    .win-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        border-left: 6px solid #10b981;
        margin-bottom: 15px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04);
    }
    .ctr-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        border-left: 6px solid #f59e0b;
        margin-bottom: 15px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04);
    }
    .badge-label {
        font-weight: bold;
        text-transform: uppercase;
        font-size: 0.75rem;
        padding: 3px 8px;
        border-radius: 4px;
        margin-bottom: 8px;
        display: inline-block;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚀 Ultimate GSC SEO Growth & Diagnostics Engine")
st.write("Upload your GSC raw ZIP export. The engine will synthesize a custom, high-priority work queue.")

# =========================================================================
# SIDEBAR CONTROLS
# =========================================================================
st.sidebar.header("🛠️ Diagnostic Parameters")
BRAND_KEYWORD = st.sidebar.text_input("Brand Keyword to Exclude", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Min Impressions for Analysis", min_value=1, value=100)
MAX_POSITION_GAP = st.sidebar.number_input("Max Position Difference (Proximity Limit)", min_value=1, max_value=20, value=10)
MAX_POSITION_LIMIT = st.sidebar.number_input("Max Allowed Position (Filter Boundary)", min_value=10, max_value=100, value=50)

# CTR Benchmarks for organic positions
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}

# =========================================================================
# HELPER: NORMALIZE METRICS
# =========================================================================
def normalize_gsc_df(df, dimension_name):
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

# =========================================================================
# GSC ZIP UNPACKER
# =========================================================================
@st.cache_data
def unpack_and_analyze_zip(uploaded_file):
    extracted_dfs = {}
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            file_list = z.namelist()
            
            file_targets = {
                'Queries': next((f for f in file_list if "queries.csv" in f.lower()), None),
                'Pages': next((f for f in file_list if "pages.csv" in f.lower()), None),
                'Devices': next((f for f in file_list if "devices.csv" in f.lower()), None),
                'Countries': next((f for f in file_list if "countries.csv" in f.lower()), None),
                'SearchAppearance': next((f for f in file_list if "search_appearance.csv" in f.lower() or "searchappearance" in f.lower()), None)
            }
            
            for key, filename in file_targets.items():
                if filename:
                    with z.open(filename) as f:
                        df = pd.read_csv(f)
                        normalized = normalize_gsc_df(df, key if key != 'SearchAppearance' else 'SearchAppearance')
                        if normalized is not None:
                            extracted_dfs[key] = normalized
        return extracted_dfs
    except Exception as e:
        st.error(f"Error extracting ZIP files: {e}")
        return None

# =========================================================================
# MAIN APP FLOW
# =========================================================================
uploaded_file = st.file_uploader("Upload your raw GSC ZIP File:", type=["zip"])

if uploaded_file is None:
    st.info("""
    💡 **Ready to go?** Download your Performance ZIP from GSC, drop it here, and get a tailored workflow immediately.
    """)
else:
    with st.spinner("Decoding GSC package and mapping opportunities..."):
        gsc_data = unpack_and_analyze_zip(uploaded_file)
        
    if gsc_data and 'Queries' in gsc_data and 'Pages' in gsc_data:
        df_queries = gsc_data['Queries']
        df_pages = gsc_data['Pages']
        
        # Exclude brand keywords
        if BRAND_KEYWORD:
            df_queries = df_queries[~df_queries['Queries'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
        
        # =========================================================================
        # BACKGROUND AUDIT ENGINE (Pre-computing findings)
        # =========================================================================
        
        # 1. Map Cannibalization Conflicts
        conflict_results = []
        queries_sorted = df_queries[df_queries['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False)
        
        for idx, q_row in queries_sorted.head(100).iterrows():
            query_txt = q_row['Queries']
            q_pos = q_row['Position']
            q_clicks_delta = q_row['Clicks_Delta']
            
            matching_urls = df_pages[
                (df_pages['Position'] >= q_pos - MAX_POSITION_GAP) & 
                (df_pages['Position'] <= q_pos + MAX_POSITION_GAP) &
                (df_pages['Impressions'] >= MIN_IMPRESSIONS / 2)
            ].sort_values(by=['Clicks', 'Impressions'], ascending=[False, False])
            
            if len(matching_urls) > 1:
                primary_url = matching_urls.iloc[0]['Pages']
                primary_clicks = int(matching_urls.iloc[0]['Clicks'])
                primary_pos = round(matching_urls.iloc[0]['Position'], 1)
                
                for sub_idx in range(1, min(len(matching_urls), 3)):
                    sub_row = matching_urls.iloc[sub_idx]
                    cannibal_url = sub_row['Pages']
                    cannibal_clicks = int(sub_row['Clicks'])
                    cannibal_pos = round(sub_row['Position'], 1)
                    
                    pos_gap = abs(primary_pos - cannibal_pos)
                    
                    if pos_gap <= MAX_POSITION_GAP and cannibal_url != primary_url:
                        threat = "🔴 High Threat (Rank Loss)" if q_clicks_delta < 0 else "🟡 Moderate Competition"
                        conflict_results.append({
                            "Threat Level": threat,
                            "Target Keyword": query_txt,
                            "Primary Authority URL": primary_url,
                            "Primary Position": primary_pos,
                            "Primary Clicks": primary_clicks,
                            "Cannibal Target URL": cannibal_url,
                            "Cannibal Position": cannibal_pos,
                            "Cannibal Clicks": cannibal_clicks,
                            "Position Gap": pos_gap
                        })
        
        df_conflicts = pd.DataFrame(conflict_results).drop_duplicates(subset=['Target Keyword', 'Cannibal Target URL']) if conflict_results else pd.DataFrame()
        
        # 2. Striking Distance Page-2 Wins
        striking_df = df_queries[
            (df_queries['Position'] >= 11.0) & 
            (df_queries['Position'] <= 20.0) & 
            (df_queries['Impressions'] >= MIN_IMPRESSIONS)
        ].sort_values(by='Impressions', ascending=False)
        
        # 3. CTR Underperformers
        ctr_boosters = []
        top_rank_queries = df_queries[(df_queries['Position'] <= 10) & (df_queries['Impressions'] >= MIN_IMPRESSIONS)]
        
        for _, row in top_rank_queries.iterrows():
            kw = row['Queries']
            clicks = row['Clicks']
            impr = row['Impressions']
            actual_ctr = row['CTR']
            pos = round(row['Position'])
            
            benchmark_ctr = CTR_BENCHMARKS.get(pos, 2.0)
            if actual_ctr < (benchmark_ctr * 0.7):
                ctr_boosters.append({
                    "Keyword": kw,
                    "Ranking Position": pos,
                    "Actual CTR": round(actual_ctr, 2),
                    "Benchmark CTR": benchmark_ctr,
                    "Click Deficit": int((impr * (benchmark_ctr / 100)) - clicks)
                })
        df_ctr = pd.DataFrame(ctr_boosters).sort_values(by='Click Deficit', ascending=False) if ctr_boosters else pd.DataFrame()

        # =========================================================================
        # AGENCY-GRADE EXECUTIVE SUMMARY HEADER
        # =========================================================================
        st.markdown(f"""
        <div class="expert-header">
            <h2>📋 SEO Expert Director's Briefing</h2>
            <p>Our GSC Diagnostic scan has run successfully. We have analyzed cannibalization footprints, metadata benchmarks, and search positioning trends. Below is your prioritized strategic roadmap.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # =========================================================================
        # WORKFLOW CONTAINER
        # =========================================================================
        col_left, col_right = st.columns([1.1, 0.9])
        
        with col_left:
            st.subheader("🚨 Priority 1: Top 5 Critical Cannibalization Wins")
            st.write("These primary URLs are bleeding ranking power and organic clicks because they are fighting with secondary pages over the exact same searches.")
            
            if not df_conflicts.empty:
                # Group by primary URL to prioritize major pages
                top_priority_pages = df_conflicts.groupby('Primary Authority URL').agg({
                    'Target Keyword': 'count',
                    'Primary Clicks': 'sum',
                    'Cannibal Target URL': 'nunique'
                }).rename(columns={
                    'Target Keyword': 'Keyword Conflicts',
                    'Cannibal Target URL': 'Competing Pages'
                }).sort_values(by='Keyword Conflicts', ascending=False).head(5).reset_index()
                
                for idx, row in top_priority_pages.iterrows():
                    st.markdown(f"""
                    <div class="action-card">
                        <span class="badge-label" style="background-color: #ffebe9; color: #ff3b30;">Priority #{idx+1} — De-Cannibalize URL</span>
                        <h4 style="margin: 4px 0;">🎯 Core Target URL: {row['Primary Authority URL']}</h4>
                        <p style="margin: 6px 0; font-size: 0.9rem; color: #555;">
                            This URL is under attack by <b>{row['Competing Pages']} competing page(s)</b> across <b>{row['Keyword Conflicts']} different keywords</b>, diluting ranking power.
                        </p>
                        <p style="margin: 4px 0; font-weight: bold; font-size: 0.85rem; color: #ff3b30;">
                            👉 Immediate Action: Find the competing URLs in the tables below. Re-anchor internal links with core keywords pointing exclusively to this Primary URL. Strip exact core keyword matches from the metadata of competing URLs.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("✅ Perfect Authority Flow! No cannibalization pages detected on your domain.")
                
            # Underperforming CTR Section
            st.markdown("<br/>", unsafe_allow_html=True)
            st.subheader("📈 Priority 2: Underperforming CTR Targets")
            st.write("These keywords rank on **Page 1**, but their Title Tags or Meta Descriptions are failing to earn their fair share of traffic.")
            
            if not df_ctr.empty:
                for idx, row in df_ctr.head(3).iterrows():
                    st.markdown(f"""
                    <div class="ctr-card">
                        <span class="badge-label" style="background-color: #fef3c7; color: #d97706;">CTR Title Optimization Target</span>
                        <h4 style="margin: 4px 0;">🔑 Keyword: "{row['Keyword']}" (Position {row['Ranking Position']})</h4>
                        <p style="margin: 6px 0; font-size: 0.9rem; color: #555;">
                            Actual CTR: <b>{row['Actual CTR']}%</b> vs Benchmark CTR: <b>{row['Benchmark CTR']}%</b>. You lost approximately <b>{row['Click Deficit']} clicks</b> this period simply due to poor title styling.
                        </p>
                        <p style="margin: 4px 0; font-weight: bold; font-size: 0.85rem; color: #d97706;">
                            👉 Immediate Action: rewrite Title Tag. Inject bracketed text (e.g., [2026 Guide]), numbers, or strong action verbs to increase layout stand-out and drive click engagement.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("✅ Outstanding CTR! Your organic listings are out-clicking standard industry benchmarks.")
                
        with col_right:
            st.subheader("🚀 Priority 3: Page-2 'Striking Distance' Quick Wins")
            st.write("These keywords are sitting on Page 2 (Positions 11–20). They are highly relevant and just need a small structural bump to cross into high-traffic territory.")
            
            if not striking_df.empty:
                for idx, row in striking_df.head(4).iterrows():
                    st.markdown(f"""
                    <div class="win-card">
                        <span class="badge-label" style="background-color: #ecfdf5; color: #059669;">Striking Distance Page-2 Target</span>
                        <h4 style="margin: 4px 0;">🔑 Keyword: "{row['Queries']}"</h4>
                        <p style="margin: 6px 0; font-size: 0.9rem; color: #555;">
                            Currently ranking at Position <b>{round(row['Position'], 1)}</b> with <b>{int(row['Impressions'])} Search Impressions</b>.
                        </p>
                        <p style="margin: 4px 0; font-weight: bold; font-size: 0.85rem; color: #059669;">
                            👉 Immediate Action: Locate the page ranking for this keyword. Add 1-2 new, helpful paragraphs targeting this term. Additionally, link to it from higher authority blog pages using exact-match anchor text.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No striking distance page 2 keywords found in this impression tier.")

        # =========================================================================
        # TABBED VIEW FOR RAW DATA TABLES (FIXED 1-BASED INDEXES)
        # =========================================================================
        st.markdown("---")
        st.subheader("📊 Deep-Dive Audit Data Tables")
        
        tab_c, tab_s, tab_ctr, tab_d = st.tabs([
            "🎯 All Cannibalization Mappings",
            "🚀 Striking Distance Terms",
            "📈 CTR Underperformers List",
            "📱 Device Usability"
        ])
        
        with tab_c:
            if not df_conflicts.empty:
                df_conflicts_disp = df_conflicts.copy()
                df_conflicts_disp.index = np.arange(1, len(df_conflicts_disp) + 1)
                st.dataframe(df_conflicts_disp, use_container_width=True)
            else:
                st.write("No cannibalizations resolved.")
                
        with tab_s:
            if not striking_df.empty:
                striking_df_disp = striking_df[['Queries', 'Clicks', 'Impressions', 'CTR', 'Position']].copy()
                striking_df_disp.index = np.arange(1, len(striking_df_disp) + 1)
                st.dataframe(striking_df_disp, use_container_width=True)
            else:
                st.write("No Page 2 metrics captured.")
                
        with tab_ctr:
            if not df_ctr.empty:
                df_ctr_disp = df_ctr.copy()
                df_ctr_disp.index = np.arange(1, len(df_ctr_disp) + 1)
                st.dataframe(df_ctr_disp, use_container_width=True)
            else:
                st.write("No CTR optimizations required.")
                
        with tab_d:
            if 'Devices' in gsc_data:
                df_dev_disp = gsc_data['Devices'].copy()
                df_dev_disp.index = np.arange(1, len(df_dev_disp) + 1)
                st.dataframe(df_dev_disp, use_container_width=True)
            else:
                st.write("Device statistics not found in GSC ZIP.")
                
    else:
        st.error("❌ Invalid GSC ZIP File: Please upload the unaltered ZIP exported directly from Google Search Console.")
