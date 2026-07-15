import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# Ensure Excel dependencies are available instantly
try:
    import xlsxwriter
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])
    import xlsxwriter

# =========================================================================
# 1. STYLE ENGINE & ADVANCED CSS STYLING
# =========================================================================
st.set_page_config(
    page_title="Enterprise GSC Forensic Hub",
    page_icon="🛡️",
    layout="wide"
)

# Custom premium stylesheet injection
st.markdown("""
    <style>
    /* Global Background Adjustments */
    .stApp { background-color: #f8fafc; }
    
    /* Global Typography Styling */
    h1, h2, h3, h4, h5 { 
        color: #0f172a !important; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em;
    }
    
    /* Premium Header Area Styling */
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff !important;
        padding: 32px;
        border-radius: 12px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        margin-bottom: 30px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
        border-left: 6px solid #6366f1;
    }
    .hero-banner h1 { color: #ffffff !important; margin: 0 0 10px 0 !important; font-size: 2.2rem !important; }
    .hero-banner p { color: #94a3b8; margin: 0; font-size: 1rem; line-height: 1.5; }
    
    /* Dynamic Performance Grid */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 16px;
        margin-bottom: 25px;
    }
    
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    }
    
    .kpi-val { 
        font-size: 2rem; 
        font-weight: 800; 
        color: #0f172a; 
        font-family: "SF Mono", "Courier New", monospace; 
        line-height: 1.1;
    }
    .kpi-val.positive { color: #10b981; }
    .kpi-val.negative { color: #ef4444; }
    .kpi-val.warning { color: #f59e0b; }
    .kpi-lbl { 
        font-size: 0.75rem; 
        color: #64748b; 
        text-transform: uppercase; 
        letter-spacing: 0.08em; 
        margin-top: 8px;
        font-weight: 600;
    }
    
    /* Warning Cards / Alerts */
    .risk-banner {
        background-color: #fffbeb;
        border-left: 5px solid #f59e0b;
        padding: 16px;
        border-radius: 6px;
        margin-bottom: 20px;
    }
    .risk-banner h4 { color: #b45309 !important; margin: 0 0 6px 0 !important; }
    .risk-banner p { color: #78350f; margin: 0; font-size: 0.9rem; }

    /* Streamlit overrides for premium styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border-radius: 6px;
        color: #475569;
        padding: 8px 16px;
        font-weight: 600;
        transition: background-color 0.2s, color 0.2s;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e2e8f0;
        color: #0f172a;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Custom Download Buttons UI */
    .stDownloadButton button {
        background-color: #0f172a !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 8px 16px !important;
        font-weight: 600 !important;
        transition: background-color 0.2s !important;
    }
    .stDownloadButton button:hover {
        background-color: #1e293b !important;
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================================
# 2. CONFIGURATION & BENCHMARKS
# =========================================================================
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 31):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# 3. SIDEBAR ENGINE CONFIG
# =========================================================================
st.sidebar.markdown("### 🔧 Engine Configurations")
BRAND_TERM = st.sidebar.text_input("Exclude Branded Searches", value="botoxie").lower().strip()
MIN_IMPR_THRESHOLD = st.sidebar.number_input("Minimum Impressions Threshold", min_value=1, value=100)
MAX_CANNIBAL_OFFSET = st.sidebar.slider("Cannibalization Max Position Gap", 1, 15, 8)

# =========================================================================
# 4. DATA CLEANING & PARSING PIPE
# =========================================================================
def parse_gsc_sheet(df, dim_name):
    df.columns = [c.strip() for c in df.columns]
    target_col = next((c for c in df.columns if c.lower() in [
        dim_name.lower(), 'query', 'page', 'device', 'country', 'search appearance', f'top {dim_name.lower()}'
    ]), None)
    
    if not target_col:
        return None
        
    normalized = pd.DataFrame()
    normalized[dim_name] = df[target_col].astype(str).str.strip()
    
    # Scrub UTM tracked parameters immediately
    if dim_name == 'Pages':
        normalized = normalized[~normalized['Pages'].str.lower().str.contains('utm_|_utm|utm=', na=False)]
        
    def extract_stats(keywords, default_val=0.0):
        col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
        diff_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' in c.lower()), None)
        
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(default_val) if col else pd.Series(default_val, index=df.index)
        delta_series = pd.to_numeric(df[diff_col], errors='coerce').fillna(0.0) if diff_col else pd.Series(0.0, index=df.index)
        return val_series, delta_series

    normalized['Clicks'], normalized['Clicks_Delta'] = extract_stats(['click'])
    normalized['Impressions'], normalized['Impressions_Delta'] = extract_stats(['impression'])
    
    ctr_col = next((c for c in df.columns if 'ctr' in c.lower() and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
    if ctr_col:
        normalized['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
        normalized['CTR'] = pd.to_numeric(normalized['CTR'], errors='coerce').fillna(0.0)
    else:
        normalized['CTR'] = (normalized['Clicks'] / normalized['Impressions'] * 100).fillna(0.0)
        
    normalized['Position'], normalized['Position_Delta'] = extract_stats(['position'], default_val=99.0)
    
    # Strictly enforce max position boundary ≤ 30
    normalized = normalized[normalized['Position'] <= 30.0]
    return normalized

def extract_gsc_payload(uploaded_zip):
    results = {}
    try:
        with zipfile.ZipFile(uploaded_zip) as z:
            file_names = z.namelist()
            mappings = {
                'Queries': 'queries.csv',
                'Pages': 'pages.csv',
                'Devices': 'devices.csv',
                'Countries': 'countries.csv'
            }
            for key, pattern in mappings.items():
                matched_file = next((n for n in file_names if pattern in n.lower()), None)
                if matched_file:
                    with z.open(matched_file) as f:
                        raw_df = pd.read_csv(f)
                        clean_df = parse_gsc_sheet(raw_df, key if key in ['Queries', 'Pages'] else 'Name')
                        if clean_df is not None:
                            results[key] = clean_df
            return results
    except Exception:
        return None

def make_csv_download(df, name):
    csv_encoded = df.to_csv(index=False).encode('utf-8')
    st.download_button("💾 Download Sheet", csv_encoded, file_name=name, mime="text/csv")

# =========================================================================
# 5. CORE APP INTERACTIVE WORKSPACE
# =========================================================================
uploaded_file = st.file_uploader("Upload your raw GSC ZIP export package:", type=["zip"])

if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        # Apply brand filtering dynamically
        df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(BRAND_TERM, na=False)].copy() if BRAND_TERM else df_q_raw.copy()
        
        # Calculate dynamic KPI diagnostics
        clicks_curr = df_q['Clicks'].sum()
        clicks_delta = df_q['Clicks_Delta'].sum()
        clicks_prev = max(1, clicks_curr - clicks_delta)
        clicks_change_pct = round((clicks_delta / clicks_prev) * 100, 2)
        
        impr_curr = df_q['Impressions'].sum()
        impr_delta = df_q['Impressions_Delta'].sum()
        impr_prev = max(1, impr_curr - impr_delta)
        impr_change_pct = round((impr_delta / impr_prev) * 100, 2)
        
        avg_pos_shift = round(df_q['Position_Delta'].mean(), 2)
        
        winning_queries_cnt = len(df_q[df_q['Clicks_Delta'] > 0])
        losing_queries_cnt = len(df_q[df_q['Clicks_Delta'] < 0])
        
        # Estimate lost clicks due to CTR limits
        est_lost_clicks = 0
        ctr_gaps_table = []
        for _, row in df_q[(df_q['Position'] <= 30.0) & (df_q['Impressions'] >= MIN_IMPR_THRESHOLD)].iterrows():
            pos = max(1, min(30, int(round(row['Position']))))
            benchmark = CTR_BENCHMARKS.get(pos, 1.0)
            if row['CTR'] < (benchmark * 0.7):
                projected_clicks = (row['Impressions'] * (benchmark / 100)) - row['Clicks']
                if projected_clicks > 0:
                    est_lost_clicks += int(projected_clicks)
                    ctr_gaps_table.append({
                        "Keyword": row['Queries'],
                        "Rank": round(row['Position'], 1),
                        "Actual CTR": f"{round(row['CTR'], 1)}%",
                        "Target CTR": f"{round(benchmark, 1)}%",
                        "Click Gap Loss": int(projected_clicks),
                        "Impressions": int(row['Impressions'])
                    })
        df_ctr_gaps = pd.DataFrame(ctr_gaps_table).sort_values(by="Click Gap Loss", ascending=False) if ctr_gaps_table else pd.DataFrame()
        
        health_score = int(max(10, min(100, 100 - (losing_queries_cnt / max(1, winning_queries_cnt + losing_queries_cnt) * 85))))

        # Render Header Section
        st.markdown(f"""
        <div class="hero-banner">
            <h1>🛡️ Enterprise SEO Forensic Platform</h1>
            <p>Evaluating clean keyword data up to Position 30. Automatic removal of campaign parameters (UTMs) applied. Core organic safety factor computed at <b>{health_score}/100</b>.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # KPI Grid
        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-val {"positive" if clicks_change_pct >= 0 else "negative"}">{clicks_change_pct}%</div>
                <div class="kpi-lbl">Clicks Change</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-val {"positive" if impr_change_pct >= 0 else "negative"}">{impr_change_pct}%</div>
                <div class="kpi-lbl">Impressions Change</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-val warning">{avg_pos_shift}</div>
                <div class="kpi-lbl">Avg Position Shift</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-val negative">{est_lost_clicks:,}</div>
                <div class="kpi-lbl">Lost Click Potential</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Unified Multi-Sheet Excel Compiler
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_q.head(500).to_excel(writer, sheet_name='Clean Queries', index=False)
            df_p.head(500).to_excel(writer, sheet_name='Clean Pages', index=False)
            if not df_ctr_gaps.empty:
                df_ctr_gaps.head(500).to_excel(writer, sheet_name='CTR Click Loss Gaps', index=False)
        xlsx_compiled = output.getvalue()
        
        st.download_button(
            label="📊 Download Complete Excel Audit Package (.xlsx)",
            data=xlsx_compiled,
            file_name="gsc_enterprise_performance_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("<br/>", unsafe_allow_html=True)

        # Main Interface Navigation tabs
        tab_kws, tab_pgs, tab_ctr, tab_can, tab_directives = st.tabs([
            "🔑 Keyword Forensic Reports", 
            "📄 Page Leakages", 
            "📈 CTR Gap Analysis", 
            "🎯 Cannibalization Map", 
            "🤖 Algorithmic Directives"
        ])
        
        # --- TAB 1: KEYWORD FORENSICS ---
        with tab_kws:
            st.markdown("### Keyword Diagnostic Core")
            
            k1, k2 = st.columns(2)
            with k1:
                st.write("#### 📉 Top Gaining Keywords")
                gaining = df_q[df_q['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(25)
                st.dataframe(gaining, use_container_width=True)
                make_csv_download(gaining, "top_gaining_keywords.csv")
                
                st.write("#### 🎯 Striking Distance (Positions 4–10)")
                striking = df_q[(df_q['Position'] >= 4.0) & (df_q['Position'] <= 10.0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(striking, use_container_width=True)
                make_csv_download(striking, "striking_distance_keywords.csv")
                
            with k2:
                st.write("#### 🚨 Top Losing Keywords")
                losing = df_q[df_q['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(25)
                st.dataframe(losing, use_container_width=True)
                make_csv_download(losing, "top_losing_keywords.csv")
                
                st.write("#### 📍 Optimization Opportunities (Positions 11–20)")
                page_two = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 20.0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(page_two, use_container_width=True)
                make_csv_download(page_two, "page_two_opportunities.csv")

        # --- TAB 2: PAGE LEAKAGES ---
        with tab_pgs:
            st.markdown("### Landing Page Forensic Core")
            
            p1, p2 = st.columns(2)
            with p1:
                st.write("#### 📈 Top Gaining Pages")
                pg_gain = df_p[df_p['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(25)
                st.dataframe(pg_gain, use_container_width=True)
                make_csv_download(pg_gain, "top_gaining_pages.csv")
                
                st.write("#### ♻️ Pages Requiring Fresh Content")
                pg_refresh = df_p[(df_p['Position_Delta'] > 0.8) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(pg_refresh, use_container_width=True)
                make_csv_download(pg_refresh, "pages_needing_refresh.csv")
                
            with p2:
                st.write("#### 📉 Top Losing Pages")
                pg_lose = df_p[df_p['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(25)
                st.dataframe(pg_lose, use_container_width=True)
                make_csv_download(pg_lose, "top_losing_pages.csv")
                
                st.write("#### 🔍 Missing Click Potential (Top 10 but low CTR)")
                pg_potential = df_p[(df_p['Position'] <= 10.0) & (df_p['CTR'] < 1.8)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(pg_potential, use_container_width=True)
                make_csv_download(pg_potential, "pages_with_low_ctr_in_top10.csv")

        # --- TAB 3: CTR ANALYSIS ---
        with tab_ctr:
            st.markdown("### Click Efficiency Loss and SERP Click Gaps")
            if not df_ctr_gaps.empty:
                st.dataframe(df_ctr_gaps.head(50), use_container_width=True)
                make_csv_download(df_ctr_gaps, "expected_ctr_deficits.csv")
            else:
                st.info("No significant CTR gaps detected based on position benchmarks.")

        # --- TAB 4: CANNIBALIZATION ---
        with tab_can:
            st.markdown("### Organic Search Conflict Map")
            
            cannibal_list = []
            candidates = df_q[df_q['Impressions'] >= MIN_IMPR_THRESHOLD].sort_values(by='Impressions', ascending=False).head(150)
            
            for _, q_row in candidates.iterrows():
                query_txt = q_row['Queries']
                q_pos = q_row['Position']
                
                # Fetch conflict pages nearby on positions
                matching_urls = df_p[
                    (df_p['Position'] >= q_pos - MAX_CANNIBAL_OFFSET) & 
                    (df_p['Position'] <= q_pos + MAX_CANNIBAL_OFFSET) &
                    (df_p['Impressions'] >= MIN_IMPR_THRESHOLD / 2)
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
                                "Primary Authority Page": primary_url,
                                "Primary Rank": primary_pos,
                                "Competing Page": cannibal_url,
                                "Competing Rank": cannibal_pos,
                                "Overlap Distance": round(abs(primary_pos - cannibal_pos), 1)
                            })
                            
            df_cannibals = pd.DataFrame(cannibal_list).drop_duplicates() if cannibal_list else pd.DataFrame()
            
            if not df_cannibals.empty:
                st.dataframe(df_cannibals.head(50), use_container_width=True)
                make_csv_download(df_cannibals, "search_cannibalization_clashes.csv")
            else:
                st.info("No query cannibalization mapped between positions 1 and 30.")

        # --- TAB 5: ALGORITHMIC DIRECTIVES ---
        with tab_directives:
            st.markdown("### Algorithmic Optimization Roadmap")
            
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("""
                <div class="risk-banner">
                    <h4>📋 Top Priority Refresh Targets</h4>
                    <p>Re-evaluate helpful content structures and intent optimization for these URLs first to regain traffic momentum.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                
                bad_pages = df_p[(df_p['Position_Delta'] > 0.8) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Impressions', ascending=False).head(10)
                if not bad_pages.empty:
                    for i, r in bad_pages.reset_index().iterrows():
                        st.markdown(f"**{i+1}.** `{r['Pages']}` (Rank: **{round(r['Position'],1)}** | Loss: **{r['Clicks_Delta']} clicks**)")
                else:
                    st.success("No critical organic drop issues found across landing pages.")
                    
            with d2:
                st.markdown("""
                <div class="risk-banner">
                    <h4>🔗 Top Internal Linking Targets</h4>
                    <p>These terms rank on page 2. Acquire context-relevant internal links to these pages to boost ranking to page 1.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                
                target_kws = df_q[(df_q['Position'] >= 11) & (df_q['Position'] <= 20)].sort_values(by='Impressions', ascending=False).head(10)
                if not target_kws.empty:
                    for i, r in target_kws.reset_index().iterrows():
                        st.markdown(f"**{i+1}.** `{r['Queries']}` (Current Position: **{round(r['Position'],1)}** | Impressions: **{int(r['Impressions'])}**)")
                else:
                    st.success("All primary keywords have high positions on page 1.")

    else:
        st.error("❌ The uploaded ZIP file does not contain valid 'queries.csv' and 'pages.csv' datasets.")
