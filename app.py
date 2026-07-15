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
# 1. PREMIUM MIDNIGHT DARK THEME ENGINE (FIXED "UPLOADPLOAD" BUG)
# =========================================================================
st.set_page_config(
    page_title="Enterprise GSC Forensic Hub",
    page_icon="🛡️",
    layout="wide"
)

# Custom premium dark stylesheet injection (Safe CSS targets to protect UI widgets)
st.markdown("""
    <style>
    /* Premium Midnight Dark Page Background */
    .stApp { 
        background-color: #0f172a !important; 
    }
    
    /* Safely target only Headings and specified elements to prevent uploader overlap */
    h1, h2, h3, h4, h5, h6, .metric-title, .metric-desc { 
        color: #f8fafc !important; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }
    h1, h2, h3 {
        font-weight: 800 !important;
        letter-spacing: -0.02em;
    }
    
    /* Deep Indigo Gradient Executive Header */
    .hero-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #312e81 100%);
        color: #ffffff !important;
        padding: 30px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
        border-left: 6px solid #60a5fa;
    }
    .hero-banner h1 { 
        color: #ffffff !important; 
        margin: 0 0 8px 0 !important; 
        font-size: 2rem !important; 
    }
    .hero-banner p { 
        color: #93c5fd !important; 
        margin: 0; 
        font-size: 1rem; 
        line-height: 1.5; 
        font-weight: 400;
    }
    
    /* High-Contrast Dark KPI Cards */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 20px;
        margin-bottom: 25px;
    }
    
    .kpi-card {
        background-color: #1e293b !important;
        border: 2px solid #334155 !important;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25);
        text-align: center;
    }
    
    .kpi-val { 
        font-size: 2.2rem; 
        font-weight: 800; 
        color: #f8fafc !important; 
        font-family: "SF Mono", monospace; 
        line-height: 1.1;
    }
    .kpi-val.positive { color: #34d399 !important; }
    .kpi-val.negative { color: #f87171 !important; }
    .kpi-val.warning { color: #fbbf24 !important; }
    
    .kpi-lbl { 
        font-size: 0.8rem; 
        color: #94a3b8 !important; 
        text-transform: uppercase; 
        letter-spacing: 0.08em; 
        margin-top: 10px;
        font-weight: 700;
    }
    
    /* Deep Warning Banners */
    .risk-banner {
        background-color: #78350f !important;
        border: 1px solid #b45309 !important;
        border-left: 6px solid #f59e0b !important;
        padding: 16px;
        border-radius: 8px;
        margin-bottom: 15px;
    }
    .risk-banner h4 { color: #fef3c7 !important; margin: 0 0 6px 0 !important; font-size: 1.05rem !important; }
    .risk-banner p { color: #fde68a !important; margin: 0; font-size: 0.9rem; font-weight: 500; }

    /* Dark Mode Tabs Customization */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #1e293b !important;
        padding: 8px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border-radius: 8px;
        color: #94a3b8 !important;
        padding: 10px 20px;
        font-weight: 700 !important;
        font-size: 0.95rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #334155 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
    }
    
    /* Explicit color for standard Streamlit text elements without breaking input files */
    .stMarkdown p, .stMarkdown span {
        color: #e2e8f0 !important;
    }
    
    /* Keep widget labels crisp and clean */
    label[data-testid="stWidgetLabel"] p {
        color: #cbd5e1 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    
    /* Custom spacing and separators */
    hr {
        border-color: #334155 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Title Area Block
st.markdown("""
<div class="hero-banner">
    <h1>🛡️ Enterprise GSC Forensic Hub</h1>
    <p>A high-performance diagnostic dashboard for raw Search Console exports. Analyzes keywords/pages strictly up to <b>Position 30</b>, scrubs UTM tracking flags automatically, and exposes high-yield organic opportunities.</p>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# 2. INLINE CONFIGURATION BOARD
# =========================================================================
st.markdown("### ⚙️ Engine Control Panel")

cfg_col1, cfg_col2, cfg_col3 = st.columns(3)
with cfg_col1:
    BRAND_TERM = st.text_input("Exclude Branded Searches (Default is Blank):", value="").lower().strip()
with cfg_col2:
    MIN_IMPR_THRESHOLD = st.number_input("Minimum Impressions Threshold:", min_value=1, value=100)
with cfg_col3:
    MAX_CANNIBAL_OFFSET = st.slider("Cannibalization Search Space (Pos. Gap):", 1, 15, 8)

st.markdown("---")

# CTR Benchmark targets configuration
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 31):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# 3. ROBUST GSC SHEET COMPARISON ENGINE
# =========================================================================
def parse_gsc_sheet(df, dim_name):
    df.columns = [c.strip() for c in df.columns]
    
    # Locate the query/page identity column
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
        
    # Robust metric grabber capable of handling standard and comparison GSC schemas
    def extract_stats(keywords, default_val=0.0):
        # Look for the primary current value (avoiding "previous" or "difference" columns)
        col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                    and 'difference' not in c.lower() 
                    and 'previous' not in c.lower() 
                    and 'compare' not in c.lower()), None)
        
        # Look for dynamic change/difference metrics
        diff_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                         and ('difference' in c.lower() or 'delta' in c.lower() or 'change' in c.lower())), None)
        
        # If no explicit "difference" column is present, try calculation with "previous" values if they exist
        prev_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                         and 'previous' in c.lower()), None)
        
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(default_val) if col else pd.Series(default_val, index=df.index)
        
        if diff_col:
            # Clean possible percent or sign formatting in differences
            diff_clean = df[diff_col].astype(str).str.replace('%', '', regex=False).str.replace('+', '', regex=False)
            delta_series = pd.to_numeric(diff_clean, errors='coerce').fillna(0.0)
        elif prev_col and col:
            prev_series = pd.to_numeric(df[prev_col], errors='coerce').fillna(default_val)
            delta_series = val_series - prev_series
        else:
            delta_series = pd.Series(0.0, index=df.index)
            
        return val_series, delta_series

    normalized['Clicks'], normalized['Clicks_Delta'] = extract_stats(['click'])
    normalized['Impressions'], normalized['Impressions_Delta'] = extract_stats(['impression'])
    
    # Extract & clean CTR metrics
    ctr_col = next((c for c in df.columns if 'ctr' in c.lower() and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
    if ctr_col:
        normalized['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
        normalized['CTR'] = pd.to_numeric(normalized['CTR'], errors='coerce').fillna(0.0)
    else:
        normalized['CTR'] = ((normalized['Clicks'] / normalized['Impressions']) * 100).fillna(0.0)
        
    normalized['Position'], normalized['Position_Delta'] = extract_stats(['position'], default_val=99.0)
    
    # Keep strictly within Search Range <= 30
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
    st.download_button("💾 Export CSV Sheet", csv_encoded, key=f"dl_{name}", file_name=name, mime="text/csv")

# =========================================================================
# 4. RUN ANALYTICAL PIPELINE
# =========================================================================
uploaded_file = st.file_uploader("Upload GSC ZIP file below to begin analysis:", type=["zip"])

if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        # Apply brand filtering dynamically
        df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(BRAND_TERM, na=False)].copy() if BRAND_TERM else df_q_raw.copy()
        
        # Calculate comparison trends
        clicks_curr = df_q['Clicks'].sum()
        clicks_delta = df_q['Clicks_Delta'].sum()
        clicks_prev = max(1.0, clicks_curr - clicks_delta)
        clicks_change_pct = round((clicks_delta / clicks_prev) * 100, 2)
        
        impr_curr = df_q['Impressions'].sum()
        impr_delta = df_q['Impressions_Delta'].sum()
        impr_prev = max(1.0, impr_curr - impr_delta)
        impr_change_pct = round((impr_delta / impr_prev) * 100, 2)
        
        avg_pos_shift = round(df_q['Position_Delta'].mean(), 2)
        
        winning_queries_cnt = len(df_q[df_q['Clicks_Delta'] > 0])
        losing_queries_cnt = len(df_q[df_q['Clicks_Delta'] < 0])
        
        # Estimate lost clicks due to CTR deficits
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

        # Render KPI Container
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
                <div class="kpi-val {"positive" if avg_pos_shift < 0 else "negative" if avg_pos_shift > 0 else "warning"}">{avg_pos_shift}</div>
                <div class="kpi-lbl">Avg Position Shift</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-val negative">{est_lost_clicks:,}</div>
                <div class="kpi-lbl">Lost Click Potential</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Multi-Sheet Excel Compiler
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_q.head(500).to_excel(writer, sheet_name='Clean Queries', index=False)
            df_p.head(500).to_excel(writer, sheet_name='Clean Pages', index=False)
            if not df_ctr_gaps.empty:
                df_ctr_gaps.head(500).to_excel(writer, sheet_name='CTR Click Loss Gaps', index=False)
        xlsx_compiled = output.getvalue()
        
        st.download_button(
            label="📊 Download Integrated Multi-Sheet Excel Report (.xlsx)",
            data=xlsx_compiled,
            file_name="gsc_enterprise_performance_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("<br/>", unsafe_allow_html=True)

        # Navigation Hub for Forensic Sheets
        tab_kws, tab_pgs, tab_ctr, tab_can, tab_directives = st.tabs([
            "🔑 Keyword Forensic Reports", 
            "📄 Page Leakages", 
            "📈 CTR Gap Analysis", 
            "🎯 Cannibalization Map", 
            "🤖 Algorithmic Directives"
        ])
        
        # --- TAB 1: KEYWORDS ---
        with tab_kws:
            st.markdown("### Keyword Diagnostic Core")
            
            k1, k2 = st.columns(2)
            with k1:
                st.markdown("#### 📈 Top Gaining Keywords")
                gaining = df_q[df_q['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(25)
                st.dataframe(gaining, use_container_width=True)
                make_csv_download(gaining, "top_gaining_keywords.csv")
                
                st.markdown("#### 🎯 Striking Distance (Positions 4–10)")
                striking = df_q[(df_q['Position'] >= 4.0) & (df_q['Position'] <= 10.0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(striking, use_container_width=True)
                make_csv_download(striking, "striking_distance_keywords.csv")
                
            with k2:
                st.markdown("#### 🚨 Top Losing Keywords")
                losing = df_q[df_q['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(25)
                st.dataframe(losing, use_container_width=True)
                make_csv_download(losing, "top_losing_keywords.csv")
                
                st.markdown("#### 📍 Page Two Optimization Options (Positions 11–20)")
                page_two = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 20.0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(page_two, use_container_width=True)
                make_csv_download(page_two, "page_two_opportunities.csv")

        # --- TAB 2: PAGES ---
        with tab_pgs:
            st.markdown("### Landing Page Forensic Core")
            
            p1, p2 = st.columns(2)
            with p1:
                st.markdown("#### 📈 Top Gaining Pages")
                pg_gain = df_p[df_p['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False).head(25)
                st.dataframe(pg_gain, use_container_width=True)
                make_csv_download(pg_gain, "top_gaining_pages.csv")
                
                st.markdown("#### ♻️ Pages Requiring Fresh Content")
                pg_refresh = df_p[(df_p['Position_Delta'] > 0.8) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Impressions', ascending=False).head(25)
                st.dataframe(pg_refresh, use_container_width=True)
                make_csv_download(pg_refresh, "pages_needing_refresh.csv")
                
            with p2:
                st.markdown("#### 📉 Top Losing Pages")
                pg_lose = df_p[df_p['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(25)
                st.dataframe(pg_lose, use_container_width=True)
                make_csv_download(pg_lose, "top_losing_pages.csv")
                
                st.markdown("#### 🔍 Missing Click Potential (Top 10 with Low CTR)")
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

        # --- TAB 4: CANNIBALIZATION MAP ---
        with tab_can:
            st.markdown("### Organic Search Conflict Map")
            
            cannibal_list = []
            candidates = df_q[df_q['Impressions'] >= MIN_IMPR_THRESHOLD].sort_values(by='Impressions', ascending=False).head(150)
            
            for _, q_row in candidates.iterrows():
                query_txt = q_row['Queries']
                q_pos = q_row['Position']
                
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
                    <h4>📋 Content Refresh Targets</h4>
                    <p>Improve on-page intent optimization and information density for these URLs to capture rising keyword trends.</p>
                </div>
                """, unsafe_allow_html=True)
                
                bad_pages = df_p[(df_p['Position_Delta'] > 0.8) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Impressions', ascending=False).head(10)
                if not bad_pages.empty:
                    for i, r in bad_pages.reset_index().iterrows():
                        st.markdown(f"**{i+1}.** `{r['Pages']}` (Rank: **{round(r['Position'],1)}** | Loss: **{r['Clicks_Delta']} clicks**)")
                else:
                    st.success("No critical page-level drops found.")
                    
            with d2:
                st.markdown("""
                <div class="risk-banner">
                    <h4>🔗 Internal Link Targets</h4>
                    <p>These terms rank on page 2. Send internally pointing links to these URLs to boost rankings up to Page 1.</p>
                </div>
                """, unsafe_allow_html=True)
                
                target_kws = df_q[(df_q['Position'] >= 11) & (df_q['Position'] <= 20)].sort_values(by='Impressions', ascending=False).head(10)
                if not target_kws.empty:
                    for i, r in target_kws.reset_index().iterrows():
                        st.markdown(f"**{i+1}.** `{r['Queries']}` (Current Position: **{round(r['Position'],1)}** | Impressions: **{int(r['Impressions'])}**)")
                else:
                    st.success("All target keywords rank cleanly on page 1.")

    else:
        st.error("❌ The uploaded ZIP file does not contain compatible 'queries.csv' and 'pages.csv' datasets.")
