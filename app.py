import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# PAGE CONFIGURATION & LAYOUT
# =========================================================================
st.set_page_config(
    page_title="SEO Damage Control & Fix Engine",
    page_icon="🚨",
    layout="wide"
)

# Dark, ultra-clean "Fix-First" Dashboard UI styling
st.markdown("""
    <style>
    .stApp {
        background-color: #fafafa;
    }
    h1, h2, h3 {
        color: #0f172a !important;
        font-family: 'Inter', sans-serif;
    }
    .fix-banner {
        background: linear-gradient(135deg, #7f1d1d 0%, #450a0a 100%);
        color: #fef2f2;
        padding: 30px;
        border-radius: 12px;
        margin-bottom: 30px;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }
    .fix-banner h2 {
        color: #ffffff !important;
        margin-top: 0;
    }
    .fix-banner p {
        color: #fca5a5;
        margin-bottom: 0;
        font-size: 1.05rem;
    }
    .error-card {
        background: #ffffff;
        border: 1px solid #fee2e2;
        border-left: 5px solid #ef4444;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .warning-card {
        background: #ffffff;
        border: 1px solid #fef3c7;
        border-left: 5px solid #f59e0b;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .code-box {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 8px 12px;
        border-radius: 6px;
        font-family: 'Courier New', Courier, monospace;
        font-size: 0.85rem;
        color: #334155;
        word-break: break-all;
        margin: 6px 0 12px 0;
    }
    .fix-badge {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        padding: 3px 8px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 8px;
    }
    .badge-red { background-color: #fee2e2; color: #991b1b; }
    .badge-orange { background-color: #fef3c7; color: #92400e; }
    </style>
""", unsafe_allow_html=True)

st.title("🚨 SEO Damage Control & Leak Fixer")
st.write("Upload your GSC ZIP package. The engine will skip the vanity metrics and pull only the leaks, drops, errors, and optimization deficits.")

# =========================================================================
# SIDEBAR CONTROLS
# =========================================================================
st.sidebar.header("⚙️ Filter Rules")
BRAND_KEYWORD = st.sidebar.text_input("Exclude Branded Searches", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Minimum Impressions Threshold", min_value=1, value=100)
MAX_CANNIBAL_GAP = st.sidebar.slider("Cannibalization Proximity (Positions)", 1, 20, 10)

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
                'Devices': next((f for f in file_list if "devices.csv" in f.lower()), None)
            }
            for key, filename in file_targets.items():
                if filename:
                    with z.open(filename) as f:
                        df = pd.read_csv(f)
                        normalized = normalize_gsc_df(df, key)
                        if normalized is not None:
                            extracted_dfs[key] = normalized
        return extracted_dfs
    except Exception as e:
        st.error(f"Error reading GSC Archive: {e}")
        return None

# =========================================================================
# ANALYSIS AND RENDERING
# =========================================================================
uploaded_file = st.file_uploader("Upload GSC ZIP Export:", type=["zip"])

if uploaded_file is None:
    st.info("📂 Drop your exported Search Console ZIP here. We will instantly map every leaking or underperforming keyword & URL relationship.")
else:
    with st.spinner("Compiling negative SEO signals..."):
        gsc_data = unpack_and_analyze_zip(uploaded_file)
        
    if gsc_data and 'Queries' in gsc_data and 'Pages' in gsc_data:
        df_queries = gsc_data['Queries']
        df_pages = gsc_data['Pages']
        
        if BRAND_KEYWORD:
            df_queries = df_queries[~df_queries['Queries'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
            
        # 1. TRAFFIC BLEEDERS (Greatest Click Drop)
        traffic_bleeders = df_queries[df_queries['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(10)
        
        # 2. CANNIBALIZATION MAP
        cannibal_list = []
        queries_sorted = df_queries[df_queries['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False)
        for idx, q_row in queries_sorted.head(100).iterrows():
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
                
                for sub_idx in range(1, min(len(matching_urls), 3)):
                    sub_row = matching_urls.iloc[sub_idx]
                    cannibal_url = sub_row['Pages']
                    cannibal_pos = round(sub_row['Position'], 1)
                    
                    if cannibal_url != primary_url:
                        cannibal_list.append({
                            "Query": query_txt,
                            "Primary Page (Keep)": primary_url,
                            "Primary Pos": primary_pos,
                            "Cannibal Page (Fix)": cannibal_url,
                            "Cannibal Pos": cannibal_pos
                        })
        df_cannibal = pd.DataFrame(cannibal_list).drop_duplicates(subset=['Query', 'Cannibal Page (Fix)']) if cannibal_list else pd.DataFrame()
        
        # 3. UNDERPERFORMING CTR (Page 1 Underachievers)
        ctr_leaks = []
        page1_queries = df_queries[(df_queries['Position'] <= 10) & (df_queries['Impressions'] >= MIN_IMPRESSIONS)]
        for _, row in page1_queries.iterrows():
            kw = row['Queries']
            clicks = row['Clicks']
            impr = row['Impressions']
            actual_ctr = row['CTR']
            pos = round(row['Position'])
            
            benchmark = CTR_BENCHMARKS.get(pos, 2.0)
            if actual_ctr < (benchmark * 0.7): # 30% below expected standard
                ctr_leaks.append({
                    "Keyword": kw,
                    "Rank": pos,
                    "Actual CTR": f"{round(actual_ctr, 1)}%",
                    "Expected CTR": f"{benchmark}%",
                    "Lost Clicks": int((impr * (benchmark / 100)) - clicks)
                })
        df_ctr_leaks = pd.DataFrame(ctr_leaks).sort_values(by='Lost Clicks', ascending=False) if ctr_leaks else pd.DataFrame()

        # =========================================================================
        # RENDER ACTION BOARD
        # =========================================================================
        st.markdown("""
        <div class="fix-banner">
            <h2>🚨 Organic Performance Deficit & Fix Roadmap</h2>
            <p>Every item listed below represents missed traffic, ranking loss, or poor click efficiency. Fix these issues to recover your performance.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.subheader("🔴 Structural & Authority Leaks")
            
            # Show Cannibalization
            st.write("#### 1. Keyword Cannibalization (Internal Page Fights)")
            if not df_cannibal.empty:
                for idx, row in df_cannibal.head(3).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="error-card">
                        <span class="fix-badge badge-red">Cannibalization Deficit #{idx+1}</span>
                        <h5 style="margin:0 0 4px 0;">Target Search: "{row['Query']}"</h5>
                        <p style="margin: 4px 0; font-size: 0.85rem; font-weight: bold; color: #1e293b;">Primary Authority URL (Keep and strengthen):</p>
                        <div class="code-box">{row['Primary Page (Keep)']} (Rank: {row['Primary Pos']})</div>
                        <p style="margin: 4px 0; font-size: 0.85rem; font-weight: bold; color: #991b1b;">Conflicting URL (Diluting authority):</p>
                        <div class="code-box">{row['Cannibal Page (Fix)']} (Rank: {row['Cannibal Pos']})</div>
                        <p style="margin: 6px 0 0 0; font-size: 0.85rem; color: #475569; line-height: 1.4;">
                            <b>Fix Action:</b> Open the Conflicting URL page and link exact-match anchor text ("{row['Query']}") directly to your Primary Authority URL. Consider trimming matching keyword variations from the title/H1 headers of the Conflicting page.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("No critical cannibalization trends found!")
                
            # Show CTR Underachievers
            st.write("#### 2. Click Efficiency Deficits (Page-1 CTR Drops)")
            if not df_ctr_leaks.empty:
                for idx, row in df_ctr_leaks.head(3).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="warning-card">
                        <span class="fix-badge badge-orange">CTR Optimization Needed #{idx+1}</span>
                        <h5 style="margin:0 0 4px 0;">Keyword: "{row['Keyword']}"</h5>
                        <p style="margin: 4px 0; font-size: 0.85rem; color: #475569;">
                            Ranks at Position <b>{row['Rank']}</b>, but CTR is only <b>{row['Actual CTR']}</b> (vs. <b>{row['Expected CTR']}</b> benchmark). 
                            You leaked <b>{row['Lost Clicks']} potential clicks</b> simply due to low-impact presentation.
                        </p>
                        <p style="margin: 6px 0 0 0; font-size: 0.85rem; font-weight: bold; color: #92400e; line-height: 1.4;">
                            <b>Fix Action:</b> Rewrite the metadata for the page ranking for this query. Use brackets, clear action-oriented modifiers, or numbers in the meta title to capture search intent more effectively.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("All Page-1 rankings are winning healthy click volume!")
                
        with col_right:
            st.subheader("📉 Traffic Loss & Visibility Deficits")
            
            # Show Dropping Keywords
            st.write("#### 3. Traffic Bleeders (Largest Click Losses)")
            if not traffic_bleeders.empty:
                for idx, row in traffic_bleeders.head(4).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="error-card">
                        <span class="fix-badge badge-red">Decline Signal #{idx+1}</span>
                        <h5 style="margin:0 0 4px 0;">Keyword: "{row['Queries']}"</h5>
                        <p style="margin: 4px 0; font-size: 0.85rem; color: #475569;">
                            This keyword lost <b>{int(abs(row['Clicks_Delta']))} clicks</b> over the previous period. Current position: <b>{round(row['Position'], 1)}</b>.
                        </p>
                        <p style="margin: 6px 0 0 0; font-size: 0.85rem; font-weight: bold; color: #b91c1c; line-height: 1.4;">
                            <b>Fix Action:</b> Check the historical ranking trend. If position has slipped, update the page with updated context, clear subheadings, and verify that the target content is still serving the core search intent.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No negative click trend detected in your dataset.")
                
            # Show Device issues
            st.write("#### 4. Mobile Layout Performance Gaps")
            if 'Devices' in gsc_data:
                df_dev = gsc_data['Devices']
                mob_row = df_dev[df_dev['Devices'].str.lower() == 'mobile']
                desk_row = df_dev[df_dev['Devices'].str.lower() == 'desktop']
                
                if not mob_row.empty and not desk_row.empty:
                    m_ctr = mob_row.iloc[0]['CTR']
                    d_ctr = desk_row.iloc[0]['CTR']
                    
                    if m_ctr < (d_ctr * 0.8):
                        st.markdown(f"""
                        <div class="error-card">
                            <span class="fix-badge badge-red">Mobile Deficit Alert</span>
                            <h5 style="margin:0 0 4px 0;">Mobile CTR Underperforming Desktop</h5>
                            <p style="margin: 4px 0; font-size: 0.85rem; color: #475569;">
                                Mobile CTR: <b>{round(m_ctr, 2)}%</b> | Desktop CTR: <b>{round(d_ctr, 2)}%</b>. 
                                Mobile listings are under-converting desktop clicks by over 20%.
                            </p>
                            <p style="margin: 6px 0 0 0; font-size: 0.85rem; font-weight: bold; color: #b91c1c; line-height: 1.4;">
                                <b>Fix Action:</b> Run mobile-friendliness or Core Web Vitals checks. Verify viewports, ensure button layout elements are not jumping (LCP/CLS issues), and confirm dynamic content displays properly on smaller screens.
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.success("Mobile and desktop click ratios are healthy and closely aligned!")
            else:
                st.info("Upload `devices.csv` to diagnostic tools to check for cross-device visibility drops.")

        # =========================================================================
        # 1-INDEXED NEGATIVE DEEP DIVE TABLES
        # =========================================================================
        st.markdown("---")
        st.subheader("📊 Supplementary Deficit Tables (Negative Focus)")
        
        tab_bleeder, tab_c_raw, tab_ctr_raw = st.tabs([
            "📉 Complete Traffic Bleeder List", 
            "🎯 Detailed Cannibalization Overlaps", 
            "📈 Complete CTR Deficit List"
        ])
        
        with tab_bleeder:
            if not traffic_bleeders.empty:
                disp_bleed = traffic_bleeders[['Queries', 'Clicks', 'Impressions', 'CTR', 'Position', 'Clicks_Delta']].copy()
                disp_bleed.index = np.arange(1, len(disp_bleed) + 1)
                st.dataframe(disp_bleed, use_container_width=True)
                
        with tab_c_raw:
            if not df_cannibal.empty:
                disp_cannibal = df_cannibal.copy()
                disp_cannibal.index = np.arange(1, len(disp_cannibal) + 1)
                st.dataframe(disp_cannibal, use_container_width=True)
                
        with tab_ctr_raw:
            if not df_ctr_leaks.empty:
                disp_ctr = df_ctr_leaks.copy()
                disp_ctr.index = np.arange(1, len(disp_ctr) + 1)
                st.dataframe(disp_ctr, use_container_width=True)
                
    else:
        st.error("❌ Invalid GSC ZIP Format. Please upload a direct, unmodified ZIP archive from Google Search Console.")
