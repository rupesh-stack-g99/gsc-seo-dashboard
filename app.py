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
# THEME-AGNOSTIC ADAPTIVE ENGINE + MAXIMUM WIDTH LAYOUT
# =========================================================================
st.set_page_config(
    page_title="GSC Forensic Overview",
    page_icon="🕵️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* COMPACT SIDEBAR & ULTRA-WIDE RESULTS CONTAINER OVERRIDES */
    [data-testid="stSidebar"] {
        min-width: 14rem !important;
        max-width: 14rem !important;
        width: 14rem !important;
    }
    [data-testid="stAppViewBlockContainer"] {
        max-width: 96% !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-top: 2rem !important;
    }
    
    /* Typography & Headers - Dynamically adapt to active theme */
    h1, h2, h3, h4, h5, h6, .directive-header { 
        color: var(--text-color) !important; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }
    h1, h2, h3 {
        font-weight: 800 !important;
        letter-spacing: -0.02em;
    }
    
    /* Hero banner keeps its distinct dark background so white text is readable in both modes */
    .hero-banner {
        background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
        color: #ffffff !important;
        padding: 30px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
        border-left: 6px solid #6366f1;
    }
    .hero-banner h1 { 
        color: #ffffff !important; 
        margin: 0 0 8px 0 !important; 
        font-size: 2.2rem !important; 
    }
    .hero-banner p { 
        color: #c7d2fe !important; 
        margin: 0; 
        font-size: 1.05rem; 
        line-height: 1.5; 
        font-weight: 400;
    }
    
    /* Directive cards adapt to light/dark system settings automatically */
    .directive-card {
        background-color: var(--secondary-background-color) !important;
        border: 1px solid rgba(128, 128, 128, 0.2) !important;
        border-left: 5px solid #3b82f6 !important;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .directive-card.danger {
        border-left-color: #ef4444 !important;
        background-color: rgba(239, 68, 68, 0.1) !important;
    }
    .directive-card.warning {
        border-left-color: #f59e0b !important;
        background-color: rgba(245, 158, 11, 0.1) !important;
    }
    .directive-card.success {
        border-left-color: #10b981 !important;
        background-color: rgba(16, 185, 129, 0.1) !important;
    }
    
    .directive-title {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
        margin-bottom: 8px !important;
        color: var(--text-color) !important;
    }
    
    .directive-text {
        font-size: 0.95rem !important;
        line-height: 1.6 !important;
        color: var(--text-color) !important;
    }
    
    button[data-baseweb="tab"] {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
    }

    /* Requirement Box adapts seamlessly via transparent background tints */
    .upload-requirements-box {
        background-color: rgba(79, 70, 229, 0.08);
        border: 2px dashed #4f46e5;
        border-radius: 8px;
        padding: 20px;
        margin-top: 15px;
        margin-bottom: 20px;
    }

    .url-helper-box {
        background-color: var(--secondary-background-color);
        border: 1px dashed rgba(128, 128, 128, 0.3);
        border-radius: 6px;
        padding: 12px;
        margin-top: 8px;
        font-size: 0.85rem;
        color: var(--text-color);
    }

    .warning-tag {
        background-color: #7c2d12;
        color: #fdba74;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: bold;
        display: inline-block;
        margin-top: 4px;
    }

    .verified-tag {
        background-color: #064e3b;
        color: #6ee7b7;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: bold;
        display: inline-block;
        margin-top: 4px;
    }
    
    .stMarkdown p, .stMarkdown span {
        color: var(--text-color) !important;
    }
    
    label[data-testid="stWidgetLabel"] p {
        color: var(--text-color) !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    
    hr {
        border-color: rgba(128, 128, 128, 0.2) !important;
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================================
# COMPACT SIDEBAR CONFIGURATION
# =========================================================================
st.sidebar.markdown("### ⚙️ Forensic Tuning")

BRAND_INPUT = st.sidebar.text_input(
    "Exclude Branded Keywords:", 
    value="", 
    placeholder="e.g. brand, clinic, dr smith"
).lower().strip()

MIN_IMPR_THRESHOLD = st.sidebar.number_input("Min. Impressions:", min_value=1, value=100)
MAX_CANNIBAL_OFFSET = st.sidebar.slider("Cannibalization Gap:", 1, 15, 6)

st.sidebar.markdown("---")
SHOW_UNVERIFIED = st.sidebar.checkbox("🔍 Include unverified URLs", value=False)

# =========================================================================
# MAIN CONTENT AREA
# =========================================================================
st.markdown("""
<div class="hero-banner">
    <h1>🕵️ GSC Forensic Overview & Diagnostic Engine</h1>
    <p>Algorithmic search monitoring and semantic intent mapping designed to isolate high-value search discrepancies.</p>
</div>
""", unsafe_allow_html=True)

# APPLICATION DESCRIPTION (COLLAPSED BY DEFAULT)
with st.expander("📖 View Forensic Capability & Core Functionality (What this app does & finds)", expanded=False):
    st.markdown("""
    ## ⚙️ Forensic Capability & Core Functionality

    This application serves as an automated search intelligence auditor that processes raw, multi-dimensional Google Search Console data structures. It replaces time-consuming spreadsheet lookups with high-precision algorithmic auditing:

    * **What it does:** It cleanses, parses, and cross-references multi-dimensional performance sets (Queries and Pages) via a high-precision semantic matching algorithm.
    * **What it finds:**
        * **Systemic Core Update Impact:** Calculates the exact balance of search term decay vs. growth to flag potential sitewide algorithmic updates.
        * **Hidden Keyword Decay:** Pinpoints queries losing substantial click-through metrics while keeping stable impression scores.
        * **High-Value Page 1 CTR Gaps:** Detects terms ranking on Page 1 that are performing below natural CTR curves, isolating traffic loss.
        * **Keyword Cannibalization Clashes:** Flags competing internal landing pages ranking for identical search terms.
        * **Striking-Distance Quick Wins:** Finds queries hovering on the cusp of page one (positions 11–15) with high impressions.
    """)

# =========================================================================
# REQUIRED EXPORT INSTRUCTIONS & FILE UPLOADER (NOW AT THE TOP)
# =========================================================================
st.markdown("""
<div class="upload-requirements-box">
    <h3 style="margin-top:0; color: #4f46e5 !important;">⚠️ GSC Export Requirement Checklist</h3>
    <p style="margin-bottom:8px; font-size:0.95rem;">To construct comparative trend diagnostics, you must upload the <b>unzipped raw export zip</b> directly generated by Google Search Console:</p>
    <ul style="margin-top:0; margin-bottom:0; font-size:0.95rem; line-height: 1.6;">
        <li>Go to Google Search Console performance menu.</li>
        <li>Set your Date filter range to: <b>Compare last 3 months to previous period</b>.</li>
        <li>Click the <b>Export</b> button in the top right corner and choose <b>Download ZIP</b>.</li>
        <li>Upload that unaltered ZIP archive below.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload GSC ZIP Archive here:", type=["zip"])

st.markdown("---")

# CTR Reference Curve
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 101):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# DATA CLEANING ENGINE & BACKEND HELPERS
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
        normalized = normalized[~normalized['Pages'].str.lower().str.contains('utm_|_utm|utm=', na=False)]
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

def find_best_url_match_precise(query_row, df_pages, max_offset=6.0):
    """
    High-Precision Semantic Matching:
    Enforces distinct semantic scoring constraints to link queries to target URLs,
    returning both the matched URL and a boolean confidence level flag.
    """
    query_str = str(query_row['Queries']).lower().strip()
    q_pos = query_row['Position']
    
    query_tokens = [re.sub(r'[^a-z0-9]', '', token) for token in query_str.split()]
    core_nouns = [t for t in query_tokens if len(t) > 2 and t not in ['and', 'for', 'the', 'with', 'near', 'in', 'nj', 'newjersey']]
    
    candidates = df_pages[
        (df_pages['Position'] >= q_pos - max_offset) & 
        (df_pages['Position'] <= q_pos + max_offset)
    ].copy()
    
    if candidates.empty:
        candidates = df_pages.copy()

    best_score = -9999
    best_url = None
    
    for _, p_row in candidates.iterrows():
        url_path = str(p_row['Pages']).lower()
        score = 0
        
        token_matches = sum(1 for token in core_nouns if token in url_path)
        score += (token_matches * 5)
        
        for critical_word in ['kybella', 'earlobe', 'piercing', 'mounjaro', 'tirzepatide', 'botox']:
            if critical_word in query_str:
                if critical_word in url_path:
                    score += 20  
                else:
                    score -= 15  
                    
        for geo in ['weehawken', 'hoboken', 'jersey']:
            if geo in query_str and geo in url_path:
                score += 5
                
        score -= (abs(p_row['Position'] - q_pos) * 0.2)
        
        if score > best_score:
            best_score = score
            best_url = p_row['Pages']
            
    is_highly_confident = True
    if best_url:
        any_token_in_url = any(token in best_url.lower() for token in core_nouns)
        if not any_token_in_url and len(core_nouns) > 0:
            is_highly_confident = False
            
    return best_url if best_url else "Manual GSC Check Required", is_highly_confident

# -------------------------------------------------------------------------
# NEW BACKEND HELPER: AUTOMATED CONTENT & METRICS RECOMMENDATIONS ENGINE
# -------------------------------------------------------------------------
def generate_seo_recommendations(page_url, keywords):
    """
    Programmatically generates targeted Meta Titles, Descriptions, 
    Headings, and Copy Blurps utilizing the decaying target keywords.
    """
    # Grab the folder or last slug of the URL to give the UI a clean page title context
    clean_topic = page_url.split('/')[-2] if page_url.endswith('/') else page_url.split('/')[-1]
    clean_topic = clean_topic.replace('-', ' ').replace('_', ' ').title()
    if not clean_topic or clean_topic == "":
        clean_topic = "Core Target Page"
        
    primary_kw = keywords[0].title() if len(keywords) > 0 else "Our Services"
    secondary_kws = ", ".join([k.lower() for k in keywords[1:3]]) if len(keywords) > 1 else ""
    
    # 1. Meta Title (Enforces ideal ~50-60 character standard)
    meta_title = f"{primary_kw} Services - {clean_topic}"
    if len(meta_title) > 60:
        meta_title = meta_title[:57] + "..."
        
    # 2. Meta Description (Targeting precise ~145-160 character boundary limit)
    meta_desc = f"Looking for {primary_kw.lower()}? We offer specialized, high-impact strategies tailored for your needs."
    if secondary_kws:
        meta_desc += f" Proudly handling {secondary_kws}."
    meta_desc += " Get started today!"
    if len(meta_desc) > 160:
        meta_desc = meta_desc[:157] + "..."
        
    # 3. Dynamic Headings
    h1_tag = f"Specialized {primary_kw}"
    h2_tag = f"Industry-Leading Results for {primary_kw}"
    
    # 4. Copy-paste ready Content Blurb
    content_blurb = (
        f"Optimizing your performance starts with target precision. Our dedicated team approaches "
        f"<b>{primary_kw.lower()}</b> with industry-leading practices to deliver outstanding results. "
        f"By focusing on specific solutions for {secondary_kws if secondary_kws else primary_kw.lower()}, "
        f"we ensure your custom strategies are streamlined, efficient, and built for growth."
    )
    
    return {
        "title": meta_title,
        "desc": meta_desc,
        "h1": h1_tag,
        "h2": h2_tag,
        "blurb": content_blurb
    }

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
if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        # MULTI-KEYWORD EXCLUSION SYSTEM (Regex OR builder)
        if BRAND_INPUT:
            exclusions = [x.strip() for x in BRAND_INPUT.split(",") if x.strip()]
            if exclusions:
                regex_pattern = "|".join(exclusions)
                df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(regex_pattern, na=False, regex=True)].copy()
            else:
                df_q = df_q_raw.copy()
        else:
            df_q = df_q_raw.copy()
        
        losing_keys = df_q[df_q['Clicks_Delta'] < 0]
        gaining_keys = df_q[df_q['Clicks_Delta'] > 0]
        total_lost_clicks = abs(losing_keys['Clicks_Delta'].sum())
        total_gained_clicks = gaining_keys['Clicks_Delta'].sum()
        
        core_hit_score = 0.0
        if total_lost_clicks > 0:
            core_hit_score = round((total_lost_clicks / (total_lost_clicks + total_gained_clicks + 1e-5)) * 100, 1)

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
            st.markdown("""
            *This diagnostic analysis measures the systemic stability of the organic profile. By evaluating global ratios of decaying keywords against ascending terms, it assesses whether traffic contractions point toward site-wide algorithmic suppression or minor seasonal turbulence.*
            """)
            if core_hit_score > 65.0:
                st.markdown(f"""<div class="directive-card danger"><div class="directive-title">🚨 Systemic Algorithmic Suppression Flagged ({core_hit_score}% Probability)</div></div>""", unsafe_allow_html=True)
            else:
                st.markdown("""<div class="directive-card success"><div class="directive-title">✅ No Site-Wide Algorithmic Penalty Detected</div></div>""", unsafe_allow_html=True)

        # === TAB 2: KEYWORD DECAY ALERTS ===
        with tab2:
            st.markdown("## Real-time Keyword Decay Alerts")
            st.markdown("""
            *This panel exposes critical, high-exposure query drops. It isolates key search metrics where impressions remain highly stable or growing, yet actual click volume drops significantly—signaling a shift in search features, user search layout, or competitor targeting updates.*
            """)
            decay_queries = df_q[(df_q['Clicks_Delta'] < 0) & (df_q['Impressions_Delta'] >= 0) & (df_q['Position_Delta'] <= 0.2)].sort_values(by='Clicks_Delta', ascending=True).head(15)

            if not decay_queries.empty:
                for idx, r in decay_queries.reset_index().iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    
                    if confident or SHOW_UNVERIFIED:
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match - Verify URL</span>'
                        st.markdown(f"""
                        *   🔴 **Keyword:** `{r['Queries']}`  
                            *   **Current Rank:** {round(r['Position'], 1)}  
                            *   **Click Shift:** **{int(r['Clicks_Delta'])} clicks**  
                            *   🎯 **Best Match URL:** `{mapped_url}` {badge}
                        """, unsafe_allow_html=True)

        # === TAB 3: HIGH-VALUE CTR GAPS ===
        with tab3:
            st.markdown("## High-Value CTR Gaps (Page 1)")
            st.markdown("""
            *This tracking sheet measures positions on page one (positions 1-10) against performance models. Keywords yielding CTR scores under 70% of standard expectations are flagged, highlighting urgent optimization targets.*
            """)
            ctr_gaps_table = []
            for _, row in df_q[(df_q['Position'] <= 10.0) & (df_q['Impressions'] >= MIN_IMPR_THRESHOLD)].iterrows():
                pos = max(1, min(10, int(round(row['Position']))))
                benchmark = CTR_BENCHMARKS.get(pos, 1.0)
                if row['CTR'] < (benchmark * 0.7):
                    projected_clicks = (row['Impressions'] * (benchmark / 100)) - row['Clicks']
                    if projected_clicks > 5:
                        mapped_url, confident = find_best_url_match_precise(row, df_p)
                        ctr_gaps_table.append({
                            "Keyword": row['Queries'],
                            "Rank": round(row['Position'], 1),
                            "Actual CTR": f"{round(row['CTR'], 1)}%",
                            "Target CTR": f"{round(benchmark, 1)}%",
                            "Click Loss": int(projected_clicks),
                            "URL": mapped_url,
                            "Confident": confident
                        })
            
            if ctr_gaps_table:
                sorted_gaps = sorted(ctr_gaps_table, key=lambda x: x['Click Loss'], reverse=True)[:15]
                for idx, item in enumerate(sorted_gaps):
                    confident = item['Confident']
                    
                    if confident or SHOW_UNVERIFIED:
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Verification Recommended via GSC</span>'
                        st.markdown(f"""
                        *   🎯 **Keyword:** `{item['Keyword']}` (Rank: **{item['Rank']}**)  
                            *   **Your CTR:** {item['Actual CTR']} *(Expected Benchmark: {item['Target CTR']})*  
                            *   📉 **Estimated Loss:** **-{item['Click Loss']} Clicks** *(Baseline variant comparison)*
                            *   🔗 **Target URL:** `{item['URL']}` {badge}
                            *   *Directive:* Overhaul metadata optimization rules on this specific landing page.
                        """, unsafe_allow_html=True)
                st.markdown("""
                <div class="url-helper-box">
                    💡 <b>How to get 100% exact mappings:</b> If you notice complex local terms cross-bleeding, go into Google Search Console, filter by that specific query, click the <b>"Pages"</b> tab, and use that specific URL.
                </div>
                """, unsafe_allow_html=True)

        # === TAB 4: CANNIBALIZATION ===
        with tab4:
            st.markdown("## Search Intent & Cannibalization Clashes")
            st.markdown("""
            *This reporting space visualizes cannibalization where several internal landing pages conflict within the same general performance window. This conflict splits rankings, preventing a single page from advancing higher.*
            """)
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
                unique_clashes = {v['Query']: v for v in cannibal_list}.values()
                for item in list(unique_clashes)[:20]:
                    st.markdown(f"""
                    *   💥 **Query clash on:** `{item['Query']}`  
                        *   🥇 **Primary Page:** `{item['Primary URL']}` (Rank {item['Primary Rank']})  
                        *   🥈 **Competing Page:** `{item['Competing URL']}` (Rank {item['Competing Rank']})  
                    """)

        # === TAB 5: STRIKING DISTANCE QUICK WINS ===
        with tab5:
            st.markdown("## Striking Distance Quick Wins (Positions 11–15)")
            st.markdown("""
            *This panel highlights high-opportunity keywords holding stable organic baseline patterns just outside page one (positions 11-15). Minor adjustments to content relevance and internal links can lift these terms onto page one to unlock higher click distributions.*
            """)
            striking_kws = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 15.0)].sort_values(by='Impressions', ascending=False).head(20)

            if not striking_kws.empty:
                for idx, r in striking_kws.reset_index().iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    
                    if confident or SHOW_UNVERIFIED:
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Verify Target Asset</span>'
                        st.markdown(f"""
                        *   🚀 **Keyword:** `{r['Queries']}`  
                            *   **Current Rank:** {round(r['Position'], 1)} | **Impressions:** {int(r['Impressions'])}  
                            *   🔗 **Target Landing Page URL:** `{mapped_url}` {badge}
                        """, unsafe_allow_html=True)

        # =========================================================================
        # === TAB 6: EXECUTION BLUEPRINT (UPGRADED! PLACED SUGGESTIONS HERE) ===
        # =========================================================================
        with tab6:
            st.markdown("## 📋 Priority Implementation & Custom SEO Blueprint")
            st.markdown("""
            *This custom roadmap identifies your top-decaying URLs and **automatically designs target optimization suggestions** (Metas, H1, H2, and content patches) to make site implementation faster and easier.*
            """)
            
            # Map queries to identify priority target URLs automatically
            priority_pages_map = {}
            for _, r in decay_queries.head(12).iterrows():
                mapped_url, confident = find_best_url_match_precise(r, df_p)
                if mapped_url and mapped_url != "Manual GSC Check Required":
                    if mapped_url not in priority_pages_map:
                        priority_pages_map[mapped_url] = {"keywords": [], "loss": 0}
                    if r['Queries'] not in priority_pages_map[mapped_url]["keywords"]:
                        priority_pages_map[mapped_url]["keywords"].append(r['Queries'])
                    priority_pages_map[mapped_url]["loss"] += abs(r['Clicks_Delta'])
            
            # Sort mapped assets by total click loss to isolate the top 3 high-priority targets
            sorted_priority_pages = sorted(priority_pages_map.items(), key=lambda x: x[1]["loss"], reverse=True)[:3]
            
            if sorted_priority_pages:
                for idx, (url, details) in enumerate(sorted_priority_pages):
                    kws = details["keywords"][:3]
                    recs = generate_seo_recommendations(url, kws)
                    
                    st.markdown(f"""
                    <div class="directive-card warning" style="margin-top: 25px;">
                        <span class="warning-tag">🚨 HIGH PRIORITY FOCUS PAGE #{idx+1}</span>
                        <div class="directive-title" style="margin-top: 10px;">URL: <a href="{url}" target="_blank" style="color: #60a5fa;">{url}</a></div>
                        <p style="margin: 0; font-size: 0.9rem;"><b>Cumulative Target Click Loss:</b> -{int(details['loss'])} clicks</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Columns to present on-page and metadata recommendations Side-by-Side
                    col_meta, col_onpage = st.columns(2)
                    
                    with col_meta:
                        st.markdown("#### 🔍 Recommended Meta Configuration")
                        st.text_input(
                            f"Suggested Meta Title (Char Count: {len(recs['title'])}/60)", 
                            value=recs['title'], 
                            key=f"title_{idx}"
                        )
                        st.text_area(
                            f"Suggested Meta Description (Char Count: {len(recs['desc'])}/160)", 
                            value=recs['desc'], 
                            key=f"desc_{idx}", 
                            height=80
                        )
                        
                    with col_onpage:
                        st.markdown("#### ✍️ Recommended Heading Framework")
                        st.text_input("Suggested Target H1 Heading:", value=recs['h1'], key=f"h1_{idx}")
                        st.text_input("Suggested Supporting H2 Heading:", value=recs['h2'], key=f"h2_{idx}")
                    
                    # Ready-to-go content patch block
                    st.markdown("#### 📝 Copy-paste Content Update Block")
                    st.info(recs['blurb'])
                    st.markdown("---")
            else:
                st.info("No high-decay pages were identified to construct recommendations. Your site trends are currently stable!")
    else:
        st.error("❌ Data formatting processing configuration mismatch.")
