import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# CONFIGURATION & STYLE ENGINE
# =========================================================================
st.set_page_config(
    page_title="The Ultimate SEO Deficit & Defect Engine",
    page_icon="🚨",
    layout="wide"
)

st.markdown("""
    <style>
    .stApp { background-color: #fafafa; }
    h1, h2, h3, h4 { color: #0f172a !important; font-family: monospace; }
    
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
    
    .section-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚨 Ultimate SEO Deficit, Leak & Opportunity Engine")
st.write("Calculates all structural optimization points up to **Rank 30**, excluding UTM tracking tags.")

# =========================================================================
# SIDEBAR CONTROL PARAMETERS
# =========================================================================
st.sidebar.header("🔧 Engine Settings")
BRAND_KEYWORD = st.sidebar.text_input("Brand Identifier (To Filter)", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Minimum Impressions Threshold", min_value=1, value=100)
MAX_CANNIBAL_GAP = st.sidebar.slider("Cannibalization Position Proximity", 1, 15, 8)

# CTR expectations up to Position 30
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 31):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# UTILITY FUNCTIONS: DATA PARSING
# =========================================================================
def parse_and_clean_df(df, dimension_name):
    df.columns = [col.strip() for col in df.columns]
    target_col = next((col for col in df.columns if col.lower() in [dimension_name.lower(), 'query', 'page', 'device', 'country', 'search appearance', 'top ' + dimension_name.lower()]), None)
    if not target_col:
        return None
    
    normalized = pd.DataFrame()
    normalized[dimension_name] = df[target_col].astype(str).str.strip()
    
    # Strip UTM Tracked URL Strings Immediately
    if dimension_name == 'Pages':
        normalized = normalized[~normalized['Pages'].str.lower().str.contains('utm_|_utm', na=False)]
        
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
def load_all_zip_components(uploaded_file):
    extracted_dfs = {}
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            file_list = z.namelist()
            targets = {
                'Queries': 'queries.csv',
                'Pages': 'pages.csv',
                'Devices': 'devices.csv',
                'Countries': 'countries.csv',
                'SearchAppearance': 'search_appearance.csv'
            }
            for key, pattern in targets.items():
                found_file = next((f for f in file_list if pattern in f.lower()), None)
                if found_file:
                    with z.open(found_file) as f:
                        df = pd.read_csv(f)
                        normalized = parse_and_clean_df(df, key if key in ['Queries', 'Pages'] else 'Name')
                        if normalized is not None:
                            extracted_dfs[key] = normalized
        return extracted_dfs
    except Exception as e:
        st.error(f"Error reading archive components: {e}")
        return None

def download_csv_btn(df, filename, label="💾 Download CSV"):
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button(label, data=csv_data, file_name=filename, mime="text/csv")

# =========================================================================
# APP ANALYSIS PIPELINE
# =========================================================================
uploaded_file = st.file_uploader("Upload raw export ZIP package from GSC:", type=["zip"])

if uploaded_file is not None:
    g_data = load_all_zip_components(uploaded_file)
    
    if g_data and 'Queries' in g_data and 'Pages' in g_data:
        # Get raw base queries and pages
        df_q = g_data['Queries'].copy()
        df_p = g_data['Pages'].copy()
        
        # Apply strict position limits (Rank <= 30) across all findings
        df_q = df_q[df_q['Position'] <= 30.0]
        df_p = df_p[df_p['Position'] <= 30.0]
        
        # Segregate Branded vs Non-Branded Queries
        is_brand_mask = df_q['Queries'].str.lower().str.contains(BRAND_KEYWORD, na=False) if BRAND_KEYWORD else pd.Series(False, index=df_q.index)
        df_brand = df_q[is_brand_mask].copy()
        df_nonbrand = df_q[~is_brand_mask].copy()
        
        # For overall reporting, default queries base is nonbrand (if specified) or all
        df_working_queries = df_nonbrand if BRAND_KEYWORD else df_q
        
        # =========================================================================
        # EXECUTIVE SUMMARY COMPUTATION
        # =========================================================================
        total_clicks_curr = df_working_queries['Clicks'].sum()
        total_clicks_diff = df_working_queries['Clicks_Delta'].sum()
        total_clicks_prev = max(1, total_clicks_curr - total_clicks_diff)
        pct_clicks_change = round((total_clicks_diff / total_clicks_prev) * 100, 2)
        
        total_impr_curr = df_working_queries['Impressions'].sum()
        total_impr_diff = df_working_queries['Impressions_Delta'].sum()
        total_impr_prev = max(1, total_impr_curr - total_impr_diff)
        pct_impr_change = round((total_impr_diff / total_impr_prev) * 100, 2)
        
        avg_pos_change = round(df_working_queries['Position_Delta'].mean(), 2)
        avg_ctr_change = round(df_working_queries['CTR'].mean() - (df_working_queries['Clicks'] - df_working_queries['Clicks_Delta']).sum() / max(1, (df_working_queries['Impressions'] - df_working_queries['Impressions_Delta']).sum()) * 100, 2)
        
        winning_kw_count = len(df_working_queries[df_working_queries['Clicks_Delta'] > 0])
        losing_kw_count = len(df_working_queries[df_working_queries['Clicks_Delta'] < 0])
        winning_pages_count = len(df_p[df_p['Clicks_Delta'] > 0])
        losing_pages_count = len(df_p[df_p['Clicks_Delta'] < 0])
        
        # Simple SEO Health Score formula based on winning vs losing ratios:
        health_score = int(max(10, min(100, 100 - (losing_kw_count / max(1, winning_kw_count + losing_kw_count) * 80))))

        # Render Banner
        st.markdown(f"""
        <div class="alert-banner">
            <h2>🚨 MASTER REPORT: GSC RAW ACTION FILE ($\le$ POSITION 30)</h2>
            <p>System scrubbed UTM records and locked all evaluation modules strictly to Position 30 margins.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Executive Summary Tab
        st.subheader("📊 GSC Executive Summary Dashboard")
        sum_col1, sum_col2, sum_col3, sum_col4, sum_col5 = st.columns(5)
        with sum_col1:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{pct_clicks_change}%</div><div class="metric-lbl">Clicks Change</div></div>', unsafe_allow_html=True)
        with sum_col2:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{pct_impr_change}%</div><div class="metric-lbl">Impressions Change</div></div>', unsafe_allow_html=True)
        with sum_col3:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{avg_pos_change}</div><div class="metric-lbl">Avg Pos Delta</div></div>', unsafe_allow_html=True)
        with sum_col4:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{health_score}/100</div><div class="metric-lbl">Core Health Score</div></div>', unsafe_allow_html=True)
        with sum_col5:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{losing_kw_count}</div><div class="metric-lbl">Losing Keywords</div></div>', unsafe_allow_html=True)
        
        st.markdown("<br/>", unsafe_allow_html=True)

        # Tabs for Grouping
        tab_keywords, tab_pages, tab_ctr_rank, tab_countries_devices, tab_opportunities = st.tabs([
            "🔑 Keyword Analysis", 
            "📄 Page Performance", 
            "📉 CTR & Rankings", 
            "🌍 Countries & Devices", 
            "🚀 Optimization Opportunities"
        ])

        # =========================================================================
        # TAB 1: KEYWORDS
        # =========================================================================
        with tab_keywords:
            st.markdown("### 🔑 Comprehensive Keyword Analysis ($\le$ Rank 30)")
            
            # Gaining vs Losing
            g_col, l_col = st.columns(2)
            with g_col:
                st.write("#### Top Gaining Keywords")
                gaining_kw = df_working_queries[df_working_queries['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(20)
                st.dataframe(gaining_kw, use_container_width=True)
                download_csv_btn(gaining_kw, "gaining_keywords.csv")
                
                st.write("#### New Discovered Keywords")
                new_kw = df_working_queries[(df_working_queries['Clicks'] > 0) & (df_working_queries['Clicks_Delta'] == df_working_queries['Clicks'])].head(20)
                st.dataframe(new_kw, use_container_width=True)
                download_csv_btn(new_kw, "discovered_keywords.csv")
                
            with l_col:
                st.write("#### Top Losing Keywords")
                losing_kw = df_working_queries[df_working_queries['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(20)
                st.dataframe(losing_kw, use_container_width=True)
                download_csv_btn(losing_kw, "losing_keywords.csv")
                
                st.write("#### Lost/Dropped Keywords")
                lost_kw = df_working_queries[(df_working_queries['Clicks'] == 0) & (df_working_queries['Clicks_Delta'] < 0)].head(20)
                st.dataframe(lost_kw, use_container_width=True)
                download_csv_btn(lost_kw, "dropped_keywords.csv")
                
            st.markdown("---")
            
            # Position Specific Divisions
            st.write("#### Keywords Ranking Positions 4–10 (Striking Distance Quick Wins)")
            striking_kw = df_working_queries[(df_working_queries['Position'] >= 4.0) & (df_working_queries['Position'] <= 10.0)].sort_values(by='Impressions', ascending=False).head(20)
            st.dataframe(striking_kw, use_container_width=True)
            download_csv_btn(striking_kw, "rank_4_10_quickwins.csv")

            st.write("#### Keywords Ranking Positions 11–20 (Optimization Opportunities)")
            page2_opps = df_working_queries[(df_working_queries['Position'] >= 11.0) & (df_working_queries['Position'] <= 20.0)].sort_values(by='Impressions', ascending=False).head(20)
            st.dataframe(page2_opps, use_container_width=True)
            download_csv_btn(page2_opps, "rank_11_20_opportunities.csv")

        # =========================================================================
        # TAB 2: PAGES
        # =========================================================================
        with tab_pages:
            st.markdown("### 📄 Page-Level Structural Diagnostics ($\le$ Rank 30)")
            
            pg1, pg2 = st.columns(2)
            with pg1:
                st.write("#### Top Gaining Pages")
                gaining_p = df_p[df_p['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(20)
                st.dataframe(gaining_p, use_container_width=True)
                download_csv_btn(gaining_p, "gaining_pages.csv")
                
                st.write("#### Pages with Increased Clicks but Lower CTR (Click Bloat)")
                bloat_pages = df_p[(df_p['Clicks_Delta'] > 0) & (df_p['CTR'] < 2.0)].sort_values(by='Clicks', ascending=False).head(20)
                st.dataframe(bloat_pages, use_container_width=True)
                download_csv_btn(bloat_pages, "clicks_bloated_low_ctr_pages.csv")
                
            with pg2:
                st.write("#### Top Losing Pages")
                losing_p = df_p[df_p['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(20)
                st.dataframe(losing_p, use_container_width=True)
                download_csv_btn(losing_p, "losing_pages.csv")
                
                st.write("#### Pages with Declining Rankings")
                rank_drop_p = df_p[df_p['Position_Delta'] > 0.5].sort_values(by='Position_Delta', ascending=False).head(20)
                st.dataframe(rank_drop_p, use_container_width=True)
                download_csv_btn(rank_drop_p, "declining_rank_pages.csv")

        # =========================================================================
        # TAB 3: CTR & RANKINGS
        # =========================================================================
        with tab_ctr_rank:
            st.markdown("### 📉 CTR Performance & Ranking Drops")
            
            st.write("#### Underachieving CTRs (Actual CTR < Expected Benchmark)")
            ctr_leaks = []
            queries_under_30 = df_working_queries[(df_working_queries['Position'] <= 30.0) & (df_working_queries['Impressions'] >= MIN_IMPRESSIONS)]
            for _, row in queries_under_30.iterrows():
                kw = row['Queries']
                clicks = row['Clicks']
                impr = row['Impressions']
                actual_ctr = row['CTR']
                pos = max(1, min(30, int(round(row['Position']))))
                
                benchmark = CTR_BENCHMARKS.get(pos, 1.0)
                if actual_ctr < (benchmark * 0.7):
                    ctr_leaks.append({
                        "Query": kw,
                        "Position": round(row['Position'], 1),
                        "Actual CTR %": round(actual_ctr, 2),
                        "Expected Benchmark %": round(benchmark, 2),
                        "Estimated Click Deficit": int((impr * (benchmark / 100)) - clicks),
                        "Impressions": int(impr)
                    })
            df_ctr_leaks = pd.DataFrame(ctr_leaks).sort_values(by='Estimated Click Deficit', ascending=False) if ctr_leaks else pd.DataFrame()
            st.dataframe(df_ctr_leaks.head(20), use_container_width=True)
            download_csv_btn(df_ctr_leaks, "all_ctr_efficiency_deficits.csv")
            
            cr1, cr2 = st.columns(2)
            with cr1:
                st.write("#### Biggest Ranking Drops (Rank Change > 0)")
                ranking_drops = df_working_queries[df_working_queries['Position_Delta'] > 0].sort_values(by='Position_Delta', ascending=False).head(25)
                st.dataframe(ranking_drops, use_container_width=True)
                download_csv_btn(ranking_drops, "ranking_drops.csv")
                
            with cr2:
                st.write("#### Keywords That Dropped Off Page 1 (Ranks 1-10 -> 11-30)")
                dropped_page_1 = df_working_queries[(df_working_queries['Position'] > 10.0) & (df_working_queries['Position'] - df_working_queries['Position_Delta'] <= 10.0)].head(25)
                st.dataframe(dropped_page_1, use_container_width=True)
                download_csv_btn(dropped_page_1, "dropped_page_1.csv")

        # =========================================================================
        # TAB 4: COUNTRIES & DEVICES
        # =========================================================================
        with tab_countries_devices:
            st.markdown("### 🌍 Technical Layout and Geographical Analysis")
            
            # Country performance
            if 'Countries' in g_data:
                st.write("#### Geolocation Traffic Changes")
                st.dataframe(g_data['Countries'].head(15), use_container_width=True)
                download_csv_btn(g_data['Countries'], "country_performance.csv")
                
            # Device Performance
            if 'Devices' in g_data:
                st.write("#### Device Visibility & Performance Splitting")
                st.dataframe(g_data['Devices'], use_container_width=True)
                download_csv_btn(g_data['Devices'], "device_performance.csv")

        # =========================================================================
        # TAB 5: RECOVERY OPPORTUNITIES
        # =========================================================================
        with tab_opportunities:
            st.markdown("### 🚀 Recovery Opportunities ($\le$ Rank 30)")
            
            # 1. Cannibalization Map
            st.write("#### 🎯 Active Ranking Clashes (Multiple Pages for the same Keyword)")
            cannibal_list = []
            queries_sorted = df_working_queries[df_working_queries['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False)
            
            for idx, q_row in queries_sorted.iterrows():
                query_txt = q_row['Queries']
                q_pos = q_row['Position']
                
                matching_urls = df_p[
                    (df_p['Position'] <= 30.0) &
                    (df_p['Position'] >= q_pos - MAX_CANNIBAL_GAP) & 
                    (df_p['Position'] <= q_pos + MAX_CANNIBAL_GAP) &
                    (df_p['Impressions'] >= MIN_IMPRESSIONS / 2)
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
                                "Rank Gap Offset": round(abs(primary_pos - cannibal_pos), 1)
                            })
            df_clashes = pd.DataFrame(cannibal_list).drop_duplicates(subset=['Conflicting Query', 'Cannibal Competing URL (FIX)']) if cannibal_list else pd.DataFrame()
            if not df_clashes.empty:
                st.dataframe(df_clashes.head(15), use_container_width=True)
                download_csv_btn(df_clashes, "cannibalization_clashes.csv")
            else:
                st.info("No cannibalization mapped within positions 1 to 30.")
                
            st.markdown("---")
            
            # Actionable Priority Lists
            st.write("### 🚨 Recovery Priorities & Implementation Board")
            
            p_update_col, k_focus_col = st.columns(2)
            with p_update_col:
                st.markdown("""
                <div style="background-color: #fffbeb; border-left: 5px solid #f59e0b; padding: 15px; border-radius: 4px;">
                    <h4 style="margin:0 0 10px 0; color: #b45309;">📑 Pages Needing Immediate Content Refresh</h4>
                    <p style="font-size:0.9rem; color: #78350f;">These pages lost the most rankings but still capture search volume. Update details, refresh headings, and improve helpful content depth.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                top_refresh_pages = df_p[df_p['Position_Delta'] > 1.0].sort_values(by='Impressions', ascending=False).head(10)
                st.dataframe(top_refresh_pages[['Pages', 'Position', 'Position_Delta', 'Impressions']], use_container_width=True)
                
            with k_focus_col:
                st.markdown("""
                <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 15px; border-radius: 4px;">
                    <h4 style="margin:0 0 10px 0; color: #b91c1c;">🔑 Keywords To Focus On First (Top Clicks Dropped)</h4>
                    <p style="font-size:0.9rem; color: #7f1d1d;">High-traffic keywords that require immediate focus to prevent further click decay.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                top_recover_kws = df_working_queries[df_working_queries['Clicks_Delta'] < -5].sort_values(by='Clicks_Delta', ascending=True).head(10)
                st.dataframe(top_recover_kws[['Queries', 'Position', 'Clicks', 'Clicks_Delta']], use_container_width=True)

    else:
        st.error("❌ The uploaded GSC ZIP does not contain 'queries.csv' or 'pages.csv' data structures.")
