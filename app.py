import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# PAGE CONFIGURATION & LAYOUT
# =========================================================================
st.set_page_config(
    page_title="Executive SEO Strategy Hub",
    page_icon="🎯",
    layout="wide"
)

# Premium SaaS UI Styling
st.markdown("""
    <style>
    /* Main App Background & Typography */
    .stApp {
        background-color: #fcfcfd;
    }
    h1, h2, h3 {
        color: #1e293b !important;
        font-family: 'Inter', -apple-system, sans-serif;
    }
    
    /* Elegant Modern Executive Header */
    .executive-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #f8fafc;
        padding: 30px;
        border-radius: 16px;
        margin-bottom: 30px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
    }
    .executive-banner h2 {
        color: #ffffff !important;
        margin-top: 0;
        font-weight: 700;
        letter-spacing: -0.025em;
    }
    .executive-banner p {
        color: #94a3b8;
        font-size: 1.1rem;
        margin-bottom: 0;
    }

    /* Beautiful SaaS Action Cards */
    .seo-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .seo-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
    }
    
    /* Card Left Border Highlights */
    .card-critical { border-left: 5px solid #ef4444; }
    .card-warning { border-left: 5px solid #f59e0b; }
    .card-success { border-left: 5px solid #10b981; }

    /* Custom Badges */
    .badge {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        padding: 4px 10px;
        border-radius: 9999px;
        display: inline-block;
        margin-bottom: 12px;
        letter-spacing: 0.05em;
    }
    .badge-red { background-color: #fee2e2; color: #991b1b; }
    .badge-yellow { background-color: #fef3c7; color: #92400e; }
    .badge-green { background-color: #d1fae5; color: #065f46; }
    
    /* URL Formatting block */
    .url-block {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        padding: 8px 12px;
        border-radius: 6px;
        font-family: monospace;
        font-size: 0.85rem;
        word-break: break-all;
        margin-top: 4px;
        margin-bottom: 12px;
        color: #334155;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎯 SEO Expert Growth & Action Engine")
st.write("Upload your GSC raw ZIP export. The engine automatically unzips, maps, and calculates your top organic optimizations.")

# =========================================================================
# SIDEBAR CONTROLS
# =========================================================================
st.sidebar.header("⚙️ Strategy Tuning")
BRAND_KEYWORD = st.sidebar.text_input("Brand Keyword to Exclude", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Min Impressions", min_value=1, value=100)
MAX_POSITION_GAP = st.sidebar.number_input("Max Keyword Proximity Range", min_value=1, max_value=20, value=10)

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
# CORE APPLICATION EXECUTION
# =========================================================================
uploaded_file = st.file_uploader("Drop your raw GSC ZIP File here:", type=["zip"])

if uploaded_file is None:
    st.info("💡 **Ready to grow?** Upload your GSC zip package, and the app will generate your roadmap instantly.")
else:
    with st.spinner("Analyzing Search Console parameters..."):
        gsc_data = unpack_and_analyze_zip(uploaded_file)
        
    if gsc_data and 'Queries' in gsc_data and 'Pages' in gsc_data:
        df_queries = gsc_data['Queries']
        df_pages = gsc_data['Pages']
        
        # Apply brand filters
        if BRAND_KEYWORD:
            df_queries = df_queries[~df_queries['Queries'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
        
        # =========================================================================
        # REAL-TIME MAPPING & COMPUTATION
        # =========================================================================
        
        # 1. Advanced Structural Cannibalization Mapping
        conflict_results = []
        queries_sorted = df_queries[df_queries['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False)
        
        for idx, q_row in queries_sorted.head(150).iterrows():
            query_txt = q_row['Queries']
            q_pos = q_row['Position']
            q_clicks_delta = q_row['Clicks_Delta']
            
            # Map queries with competing URL sets
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
                        threat = "🔴 High Threat" if q_clicks_delta < 0 else "🟡 Moderate"
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
        # EXECUTIVE BRIEFING RENDERING
        # =========================================================================
        st.markdown(f"""
        <div class="executive-banner">
            <h2>📋 SEO Director's Monday Morning Priorities</h2>
            <p>No spreadsheets, no manual formulas. Here are your exact, mathematically calculated optimizations to run this week.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Layout Division
        col_main, col_sidebar_opps = st.columns([1.1, 0.9])
        
        with col_main:
            st.subheader("🔴 Urgent Priority: Fix Keyword Cannibalization")
            st.write("These competing pages are clashing over the same keywords, hurting your search positions. We mapped the exact URLs so you don't have to look them up:")
            
            if not df_conflicts.empty:
                # Select top 5 critical target scenarios directly
                for idx, row in df_conflicts.head(5).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="seo-card card-critical">
                        <span class="badge badge-red">Cannibalization Conflict #{idx+1}</span>
                        <h4 style="margin: 4px 0; color: #1e293b;">Conflict Query: "{row['Target Keyword']}"</h4>
                        
                        <p style="margin: 10px 0 2px 0; font-size: 0.85rem; font-weight: bold; color: #0f172a;">🛡️ Primary Authority URL (Keep this one):</p>
                        <div class="url-block">{row['Primary Authority URL']}</div>
                        
                        <p style="margin: 0 0 2px 0; font-size: 0.85rem; font-weight: bold; color: #991b1b;">⚠️ Conflicting Cannibal URL (De-optimize this):</p>
                        <div class="url-block">{row['Cannibal Target URL']}</div>
                        
                        <p style="margin: 4px 0; font-size: 0.9rem; color: #334155;">
                            <b>Impact:</b> The primary URL sits at position <b>{row['Primary Position']}</b> while the competitor page is pulling search weight at position <b>{row['Cannibal Position']}</b>.
                        </p>
                        <p style="margin: 8px 0 0 0; font-weight: bold; font-size: 0.85rem; color: #dc2626;">
                            👉 Directive: Edit the <i>Conflicting Cannibal URL</i>. Link from it over to the <i>Primary Authority URL</i> using exact match anchor text: "{row['Target Keyword']}".
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("🎉 Excellent! No duplicate URL keyword overlap discovered in GSC datasets.")
                
            # CTR Opportunities Section
            st.markdown("<br/>", unsafe_allow_html=True)
            st.subheader("⚡ High Impact: Underperforming Page-1 CTR")
            st.write("These keywords rank on Page 1 but are receiving significantly fewer clicks than they should. Revamping their Metadata will instantly boost your traffic.")
            
            if not df_ctr.empty:
                for idx, row in df_ctr.head(3).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="seo-card card-warning">
                        <span class="badge badge-yellow">CTR Boost Target #{idx+1}</span>
                        <h4 style="margin: 4px 0; color: #1e293b;">Query: "{row['Keyword']}"</h4>
                        <p style="margin: 8px 0; font-size: 0.9rem; color: #334155;">
                            Currently ranking at Position <b>{row['Ranking Position']}</b>. Your CTR is only <b>{row['Actual CTR']}%</b> (Industry Benchmark is <b>{row['Benchmark CTR']}%</b>). 
                            You missed out on <b>{row['Click Deficit']} clicks</b> because of this layout gap.
                        </p>
                        <p style="margin: 6px 0 0 0; font-weight: bold; font-size: 0.85rem; color: #d97706;">
                            👉 Directive: Update the Title Tag of the page ranking for this query. Add emotional triggers, years, or lists to make your search result stand out.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
        with col_sidebar_opps:
            st.subheader("🚀 Quick Wins: Page-2 'Striking Distance' Targets")
            st.write("These keywords sit just on Page 2 (Positions 11–20) with high search volume. A quick optimization push will jump them to Page 1 for an easy win.")
            
            if not striking_df.empty:
                for idx, row in striking_df.head(4).reset_index(drop=True).iterrows():
                    st.markdown(f"""
                    <div class="seo-card card-success">
                        <span class="badge badge-green">Striking Distance #{idx+1}</span>
                        <h4 style="margin: 4px 0; color: #1e293b;">Target Keyword: "{row['Queries']}"</h4>
                        <p style="margin: 8px 0; font-size: 0.9rem; color: #334155;">
                            Currently ranking at Position <b>{round(row['Position'], 1)}</b> with <b>{int(row['Impressions'])} total impressions</b>.
                        </p>
                        <p style="margin: 6px 0 0 0; font-weight: bold; font-size: 0.85rem; color: #059669;">
                            👉 Directive: Add this keyword to your page's H2 or H3 headers. Build 1-2 new paragraphs answering common search questions around this term.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No matching striking distance parameters detected.")

        # =========================================================================
        # TABBED VIEW FOR RAW DATA TABLES (FIXED 1-BASED INDEXES)
        # =========================================================================
        st.markdown("---")
        st.subheader("📊 Supplementary Data Tables")
        
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
