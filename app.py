import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import re

# Ensure Excel/zip dependencies
try:
    import xlsxwriter
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])
    import xlsxwriter

# =========================================================================
# PREMIUM MIDNIGHT DARK THEME ENGINE
# =========================================================================
st.set_page_config(
    page_title="Algorithmic SEO Detective",
    page_icon="🤖",
    layout="wide"
)

st.markdown("""
    <style>
    /* Premium Midnight Dark Page Background */
    .stApp { 
        background-color: #0f172a !important; 
    }
    
    h1, h2, h3, h4, h5, h6, .directive-header { 
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
        border-left: 6px solid #3b82f6;
    }
    .hero-banner h1 { 
        color: #ffffff !important; 
        margin: 0 0 8px 0 !important; 
        font-size: 2.2rem !important; 
    }
    .hero-banner p { 
        color: #93c5fd !important; 
        margin: 0; 
        font-size: 1.05rem; 
        line-height: 1.5; 
        font-weight: 400;
    }
    
    /* Custom Risk Card & Alert Classes */
    .directive-card {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-left: 5px solid #3b82f6 !important;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    .directive-card.danger {
        border-left-color: #ef4444 !important;
        background-color: #2d1616 !important;
    }
    .directive-card.warning {
        border-left-color: #f59e0b !important;
        background-color: #2d200f !important;
    }
    .directive-card.success {
        border-left-color: #10b981 !important;
        background-color: #0f2d1e !important;
    }
    .directive-card.info {
        border-left-color: #6366f1 !important;
        background-color: #171738 !important;
    }
    
    .directive-title {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
        margin-bottom: 8px !important;
    }
    
    .directive-text {
        font-size: 0.95rem !important;
        line-height: 1.6 !important;
        color: #cbd5e1 !important;
    }
    
    /* Tab Styling Overrides for Contrast */
    button[data-baseweb="tab"] {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: #94a3b8 !important;
    }
    button[aria-selected="true"] {
        color: #38bdf8 !important;
    }

    .url-helper-box {
        background-color: #1e293b;
        border: 1px dashed #475569;
        border-radius: 6px;
        padding: 12px;
        margin-top: 8px;
        font-size: 0.85rem;
        color: #94a3b8;
    }
    
    /* Explicit color for standard Streamlit text elements */
    .stMarkdown p, .stMarkdown span {
        color: #e2e8f0 !important;
    }
    
    label[data-testid="stWidgetLabel"] p {
        color: #cbd5e1 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    
    hr {
        border-color: #334155 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Title Area Block
st.markdown("""
<div class="hero-banner">
    <h1>🕵️ Algorithmic SEO Detective</h1>
    <p>This engine analyzes your Search Console data, matches keywords to actual web landing pages (filtering out images and attachments), and highlights traffic gaps against industry benchmarks.</p>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# CONFIGURATION INPUTS
# =========================================================================
st.markdown("### ⚙️ Forensic Tuning & Parameters")

cfg_col1, cfg_col2, cfg_col3 = st.columns(3)
with cfg_col1:
    BRAND_TERM = st.text_input("Exclude Branded Searches (Default is Blank):", value="").lower().strip()
with cfg_col2:
    MIN_IMPR_THRESHOLD = st.number_input("Minimum Impressions Threshold:", min_value=1, value=100)
with cfg_col3:
    MAX_CANNIBAL_OFFSET = st.slider("Cannibalization Search Space (Pos. Gap):", 1, 15, 6)

st.markdown("---")

# CTR Benchmark targets configuration
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 101):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# DATA CLEANING ENGINE (ROBUST BYPASS FOR GSC HEADER METADATA)
# =========================================================================
def clean_gsc_csv(bytes_data):
    try:
        lines = bytes_data.decode('utf-8').splitlines()
    except UnicodeDecodeError:
        try:
            lines = bytes_data.decode('latin-1').splitlines()
        except Exception:
            return None

    header_idx = 0
    for idx, line in enumerate(lines[:15]):
        lower_line = line.lower()
        if any(term in lower_line for term in ['query', 'queries', 'page', 'pages', 'clicks', 'impressions']):
            header_idx = idx
            break
            
    try:
        df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])))
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None

def parse_gsc_sheet(df, dim_name):
    target_col = None
    if dim_name == 'Queries':
        target_col = next((c for c in df.columns if c.lower() in ['query', 'queries', 'top queries', 'top query', 'search query']), None)
    elif dim_name == 'Pages':
        target_col = next((c for c in df.columns if c.lower() in ['page', 'pages', 'top pages', 'top page', 'landing page', 'url']), None)
        
    if not target_col:
        return None
        
    normalized = pd.DataFrame()
    normalized[dim_name] = df[target_col].astype(str).str.strip()
    
    if dim_name == 'Pages':
        # Remove UTM parameters
        normalized = normalized[~normalized['Pages'].str.lower().str.contains('utm_|_utm|utm=', na=False)]
        # CRITICAL FILTER: Remove static assets like images, pdfs, css, js
        asset_pattern = r'\.(jpg|jpeg|png|gif|webp|svg|pdf|css|js|txt|xml|mp4)$'
        normalized = normalized[~normalized['Pages'].str.lower().str.contains(asset_pattern, na=False)]
        
    def extract_stats(keywords, default_val=0.0):
        col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                    and 'difference' not in c.lower() 
                    and 'previous' not in c.lower() 
                    and 'compare' not in c.lower()), None)
        
        diff_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                         and ('difference' in c.lower() or 'delta' in c.lower() or 'change' in c.lower())), None)
        
        prev_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) 
                         and 'previous' in c.lower()), None)
        
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(default_val) if col else pd.Series(default_val, index=df.index)
        
        if diff_col:
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
    
    ctr_col = next((c for c in df.columns if 'ctr' in c.lower() and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
    if ctr_col:
        normalized['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
        normalized['CTR'] = pd.to_numeric(normalized['CTR'], errors='coerce').fillna(0.0)
    else:
        normalized['CTR'] = ((normalized['Clicks'] / normalized['Impressions']) * 100).fillna(0.0)
        
    normalized['Position'], normalized['Position_Delta'] = extract_stats(['position'], default_val=99.0)
    
    return normalized

def find_best_url_match(query_row, df_pages, max_offset=6.0):
    """
    Advanced Lexical & Rank Matcher: 
    Correlates a query string with your pages list by performing string-token 
    overlap analysis combined with ranking range limitations.
    """
    query_str = str(query_row['Queries']).lower().strip()
    q_pos = query_row['Position']
    
    # Tokenize the query, discarding short stop words
    query_tokens = [re.sub(r'[^a-z0-9]', '', token) for token in query_str.split()]
    query_tokens = [t for t in query_tokens if len(t) > 2 and t not in ['and', 'for', 'the', 'with', 'near']]
    
    # Pre-filter pages with ranking similarity
    candidates = df_pages[
        (df_pages['Position'] >= q_pos - max_offset) & 
        (df_pages['Position'] <= q_pos + max_offset)
    ].copy()
    
    if candidates.empty:
        candidates = df_pages[
            (df_pages['Position'] >= q_pos - 15.0) & 
            (df_pages['Position'] <= q_pos + 15.0)
        ].copy()
        
    if candidates.empty:
        candidates = df_pages.copy()

    best_score = -1
    best_url = None
    
    for _, p_row in candidates.iterrows():
        url_path = str(p_row['Pages']).lower()
        
        # Calculate how many query tokens exist inside the URL string
        match_count = sum(1 for token in query_tokens if token in url_path)
        
        score = match_count
        
        # Tie-breaker penalty for position gap
        position_penalty = abs(p_row['Position'] - q_pos) * 0.1
        score -= position_penalty
        
        if score > best_score:
            best_score = score
            best_url = p_row['Pages']
            
    return best_url if best_url else "Verification Required via GSC"

def extract_gsc_payload(uploaded_zip):
    results = {}
    try:
        with zipfile.ZipFile(uploaded_zip) as z:
            file_names = z.namelist()
            
            for file_name in file_names:
                if '__macosx' in file_name.lower() or not file_name.endswith('.csv'):
                    continue
                    
                with z.open(file_name) as f:
                    raw_df = clean_gsc_csv(f.read())
                    if raw_df is None or raw_df.empty:
                        continue
                    
                    cols_lower = [str(c).lower() for c in raw_df.columns]
                    
                    if any(q_term in cols_lower for q_term in ['query', 'queries', 'top queries', 'top query', 'search query']):
                        clean_df = parse_gsc_sheet(raw_df, 'Queries')
                        if clean_df is not None and not clean_df.empty:
                            results['Queries'] = clean_df
                            
                    elif any(p_term in cols_lower for p_term in ['page', 'pages', 'top pages', 'top page', 'landing page', 'url']):
                        clean_df = parse_gsc_sheet(raw_df, 'Pages')
                        if clean_df is not None and not clean_df.empty:
                            results['Pages'] = clean_df
                            
            return results
    except Exception as e:
        st.error(f"Error reading ZIP structure: {str(e)}")
        return None

# =========================================================================
# FORENSIC PIPELINE EXECUTION
# =========================================================================
uploaded_file = st.file_uploader("Upload GSC ZIP file to begin automated detective diagnostics:", type=["zip"])

if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        # Apply brand filter dynamically
        df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(BRAND_TERM, na=False)].copy() if BRAND_TERM else df_q_raw.copy()
        
        # Calculate comparison trends
        losing_keys = df_q[df_q['Clicks_Delta'] < 0]
        gaining_keys = df_q[df_q['Clicks_Delta'] > 0]
        
        total_lost_clicks = abs(losing_keys['Clicks_Delta'].sum())
        total_gained_clicks = gaining_keys['Clicks_Delta'].sum()
        
        core_hit_score = 0.0
        if total_lost_clicks > 0:
            core_hit_score = round((total_lost_clicks / (total_lost_clicks + total_gained_clicks + 1e-5)) * 100, 1)

        # -----------------------------------------------------------------
        # STRUCTURING TAB-BASED ARCHITECTURE
        # -----------------------------------------------------------------
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🔍 Core Update Diagnostic",
            "📉 Keyword Decay Alerts",
            "🎯 Page 1 CTR Gaps",
            "⚔️ Cannibalization Clashes",
            "🚀 Striking Distance Quick Wins",
            "📋 Execution Blueprint"
        ])

        # === TAB 1: CORE UPDATE HIT DETECTOR ===
        with tab1:
            st.markdown("## Algorithmic Updates Checker")
            if core_hit_score > 65.0:
                st.markdown(f"""
                <div class="directive-card danger">
                    <div class="directive-title">🚨 Systemic Algorithmic Suppression Flagged ({core_hit_score}% Probability)</div>
                    <div class="directive-text">
                        <b>Diagnostic:</b> Over {core_hit_score}% of overall trend movements are strictly negative. Clicks and impressions are dropping simultaneously across uncorrelated keywords. This strongly correlates with a Google Core Algorithm Update or search classifier adjustment rather than a simple indexation glitch. <br/><br/>
                        <b>Detective Action:</b> Audit your site-wide informational value. Avoid surface-level updates. Identify if pages hit hardest feature redundant introductory material, high affiliate/ad ratios, or lack distinct expert author perspectives (E-E-A-T).
                    </div>
                </div>
                """, unsafe_allow_html=True)
            elif core_hit_score > 35.0:
                st.markdown(f"""
                <div class="directive-card warning">
                    <div class="directive-title">⚠️ Moderate Algorithmic Volatility Checked ({core_hit_score}% Probability)</div>
                    <div class="directive-text">
                        <b>Diagnostic:</b> Partial traffic degradation spotted across isolated clusters. This is likely not a site-wide quality penalty, but a sub-topic re-evaluation. Competitors are likely optimizing topical coverage or getting featured in newly introduced AI SERP widgets.<br/><br/>
                        <b>Detective Action:</b> Isolate which specific sub-folders or page templates are decaying. Run content comparison sprints against those that rose in your niche over the last month.
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="directive-card success">
                    <div class="directive-title">✅ No Site-Wide Algorithmic Penalty Detected</div>
                    <div class="directive-text">
                        <b>Diagnostic:</b> Your domain performance shows organic stability or growth. Performance fluctuations are local, standard search mechanics rather than a core automated filter or quality suppression.
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # === TAB 2: KEYWORD DECAY ALERTS ===
        with tab2:
            st.markdown("## Real-time Keyword Decay Alerts")
            st.markdown("These keywords show stable search visibility (ranks and impressions are fine), but **clicks are dying**. This points directly to competitor snippet optimization, seasonal shifts, or search layout changes.")
            
            decay_queries = df_q[
                (df_q['Clicks_Delta'] < 0) & 
                (df_q['Impressions_Delta'] >= 0) & 
                (df_q['Position_Delta'] <= 0.2)
            ].sort_values(by='Clicks_Delta', ascending=True).head(15)

            if not decay_queries.empty:
                for idx, r in decay_queries.reset_index().iterrows():
                    mapped_url = find_best_url_match(r, df_p)
                    st.markdown(f"""
                    *   🔴 **Keyword:** `{r['Queries']}`  
                        *   **Current Rank:** {round(r['Position'], 1)} (Trend: {round(r['Position_Delta'], 2)})  
                        *   **Click Shift:** **{int(r['Clicks_Delta'])} clicks**  
                        *   🎯 **Best Match URL:** `{mapped_url}`
                        *   *Directive:* Overhaul this URL's meta title and description immediately.
                    """)
                st.markdown("""
                <div class="url-helper-box">
                    💡 <b>How to verify the exact URL in GSC:</b> Go to your Google Search Console performance report, click the <b>"Queries"</b> tab, click on your target keyword to filter by it, and then click the <b>"Pages"</b> tab.
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("No active keyword decay flags detected.")

        # === TAB 3: HIGH-VALUE CTR GAPS ===
        with tab3:
            st.markdown("## High-Value CTR Gaps (Page 1)")
            st.markdown("These keywords rank on Page 1 but are receiving **abnormally low CTRs** compared to benchmark metrics. Fixing these will result in immediate traffic injections.")

            ctr_gaps_table = []
            for _, row in df_q[(df_q['Position'] <= 10.0) & (df_q['Impressions'] >= MIN_IMPR_THRESHOLD)].iterrows():
                pos = max(1, min(10, int(round(row['Position']))))
                benchmark = CTR_BENCHMARKS.get(pos, 1.0)
                if row['CTR'] < (benchmark * 0.7):
                    projected_clicks = (row['Impressions'] * (benchmark / 100)) - row['Clicks']
                    if projected_clicks > 5:
                        mapped_url = find_best_url_match(row, df_p)
                        ctr_gaps_table.append({
                            "Keyword": row['Queries'],
                            "Rank": round(row['Position'], 1),
                            "Actual CTR": f"{round(row['CTR'], 1)}%",
                            "Target CTR": f"{round(benchmark, 1)}%",
                            "Click Loss": int(projected_clicks),
                            "URL": mapped_url
                        })
            
            if ctr_gaps_table:
                sorted_gaps = sorted(ctr_gaps_table, key=lambda x: x['Click Loss'], reverse=True)[:15]
                for idx, item in enumerate(sorted_gaps):
                    st.markdown(f"""
                    *   🎯 **Keyword:** `{item['Keyword']}` (Rank: **{item['Rank']}**)  
                        *   **Your CTR:** {item['Actual CTR']} *(Expected Benchmark: {item['Target CTR']})*  
                        *   📉 **Estimated Loss:** **-{item['Click Loss']} Clicks** *(Difference between actual clicks and expected industry baseline clicks at Rank {item['Rank']})*
                        *   🔗 **Target URL:** `{item['URL']}`
                        *   *Directive:* Analyze competitor headers vs your metadata. Adjust the title/meta description on this URL.
                    """)
                st.markdown("""
                <div class="url-helper-box">
                    💡 <b>How to verify the exact URL in GSC:</b> Go to your Google Search Console performance report, click the <b>"Queries"</b> tab, click on your target keyword to filter by it, and then click the <b>"Pages"</b> tab.
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("Your Page 1 CTR profiles are healthy and meeting benchmarks.")

        # === TAB 4: CANNIBALIZATION ===
        with tab4:
            st.markdown("## Search Intent & Cannibalization Clashes")
            st.markdown("These are target queries where multiple URLs are ranking in close proximity, confusing search engines and dividing your organic authority.")

            cannibal_list = []
            candidates = df_q[df_q['Impressions'] >= MIN_IMPR_THRESHOLD].sort_values(by='Impressions', ascending=False).head(200)
            
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
                                "Primary URL": primary_url,
                                "Primary Rank": primary_pos,
                                "Competing URL": cannibal_url,
                                "Competing Rank": cannibal_pos
                            })

            if cannibal_list:
                unique_clashes = {v['Query']: v for v in cannibal_list}.values() # Dedup
                for item in list(unique_clashes)[:20]:
                    st.markdown(f"""
                    *   💥 **Query clash on:** `{item['Query']}`  
                        *   🥇 **Primary Page:** `{item['Primary URL']}` (Rank {item['Primary Rank']})  
                        *   🥈 **Competing Page:** `{item['Competing URL']}` (Rank {item['Competing Rank']})  
                        *   *Directive:* De-optimize the competing page for this keyword. Consolidate them or point a hard internal link with descriptive anchor text from Page 2 to Page 1.
                    """)
            else:
                st.info("No active intent overlaps found across your high-impact keywords.")

        # === TAB 5: STRIKING DISTANCE QUICK WINS ===
        with tab5:
            st.markdown("## Striking Distance Quick Wins (Positions 11–15)")
            st.markdown("These high-impression queries are hovering just off Page 1. They are ripe for an easy ranking boost via internal link building.")

            striking_kws = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 15.0)].sort_values(by='Impressions', ascending=False).head(20)

            if not striking_kws.empty:
                for idx, r in striking_kws.reset_index().iterrows():
                    mapped_url = find_best_url_match(r, df_p)
                    st.markdown(f"""
                    *   🚀 **Keyword:** `{r['Queries']}`  
                        *   **Current Rank:** {round(r['Position'], 1)} | **Impressions:** {int(r['Impressions'])}  
                        *   🔗 **Target Landing Page URL:** `{mapped_url}`  
                        *   *Directive:* Locate 2-3 of your highest-authority articles and add an internal link pointing to this landing page URL using optimized anchor text.
                    """)
                st.markdown("""
                <div class="url-helper-box">
                    💡 <b>How to verify the exact URL in GSC:</b> Go to your Google Search Console performance report, click the <b>"Queries"</b> tab, click on your target keyword to filter by it, and then click the <b>"Pages"</b> tab.
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("No queries currently idling on Page 2 striking distance.")

        # === TAB 6: EXECUTION BLUEPRINT ===
        with tab6:
            st.markdown("## Priority Implementation Blueprint")
            st.markdown("""
            Review the findings mapped across your diagnostic panels and prioritize your execution as follows:
            
            1. **Resolve Cannibalization Clashes (Tab 4):** Clear up intent conflicts first so Google knows exactly which URL to send ranking signals to.
            2. **Optimize CTR Gaps (Tab 3):** Rewrite title tags and description snippets for those Page 1 terms showing high impressions but poor actual click rates. 
            3. **Inject Striking Distance Authority (Tab 5):** Pass internal page authority (link equity) to your Page 2 assets to easily nudge them into Page 1 territory.
            """)

    else:
        st.error("❌ ZIP processing succeeded, but the code could not isolate the core 'Queries' or 'Pages' data frames. Please verify you are uploading an authentic zip download directly from the Google Search Console UI's Export function.")
