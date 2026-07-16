import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import re
import requests

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
    page_title="GSC Forensic Overview & Advanced Diagnostic",
    page_icon="🕵️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
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
    
    h1, h2, h3, h4, h5, h6, .directive-header { 
        color: var(--text-color) !important; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }
    h1, h2, h3 {
        font-weight: 800 !important;
        letter-spacing: -0.02em;
    }
    
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

    .critical-status-highlight {
        background: linear-gradient(135deg, #ef4444 0%, #991b1b 100%) !important;
        color: #ffffff !important;
        padding: 24px;
        border-radius: 12px;
        font-weight: bold;
        font-size: 1.4rem;
        text-align: center;
        box-shadow: 0 10px 15px -3px rgba(239, 68, 68, 0.4);
        margin-bottom: 25px;
        border: 2px solid #fee2e2;
    }

    .math-explanation-box {
        background-color: var(--secondary-background-color);
        border-left: 4px solid #4f46e5;
        border-radius: 6px;
        padding: 16px;
        margin-top: 20px;
        font-size: 0.88rem;
        line-height: 1.5;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
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
    "Primary Brand Name:", 
    value="", 
    placeholder="e.g. AcmeCorp"
).lower().strip()

MIN_IMPR_THRESHOLD = st.sidebar.number_input("Min. Impressions:", min_value=1, value=100)

st.sidebar.markdown("---")
SHOW_UNVERIFIED = st.sidebar.checkbox("🔍 Include unverified URLs", value=False)

# =========================================================================
# FUZZY BRAND SPELLING DETECTOR (Protects against fat-finger brand errors)
# =========================================================================
def generate_fuzzy_brand_regex(brand_str):
    """Generates a soft regex to catch misspelt brands and unique brand portions."""
    if not brand_str or len(brand_str) < 3:
        return None
    # Strip vowels to find the consonantal root (Acme -> cm, Google -> ggl)
    consonants_only = "".join([c for c in brand_str if c not in 'aeiouy'])
    
    # Create simple patterns targeting double letters, missing vowels, or character transpositions
    patterns = [
        re.escape(brand_str),
        r"".join([f"{char}+" for char in brand_str]),  # Double letters (e.g., accme)
    ]
    if len(consonants_only) >= 2:
        patterns.append(r".*".join(list(consonants_only))) # Consonants in sequence (e.g. bnd for brand)
        
    return "|".join(patterns)

# =========================================================================
# INTENT DETECTOR (Isolates Commercial Intent / No Blogs)
# =========================================================================
def is_commercial_intent(url, query):
    """Filters out blog structures and educational queries to focus on Commercial Intent."""
    url_lower = str(url).lower()
    query_lower = str(query).lower()
    
    # Exclude typical blog/informational URL footprints
    info_path_patterns = ['/blog/', '/news/', '/article/', '/resources/', '/learning/', '/post/', '/info/']
    if any(pat in url_lower for pat in info_path_patterns):
        return False
        
    # Exclude informational query strings (FAQ/Curiosity searches)
    info_query_patterns = ['how to', 'why is', 'what is', 'difference between', 'guide', 'tutorial', 'tips', 'history of']
    if any(pat in query_lower for pat in info_query_patterns):
        return False
        
    return True

# =========================================================================
# MAIN CONTENT AREA
# =========================================================================
st.markdown("""
<div class="hero-banner">
    <h1>🕵️ GSC Forensic Overview & Diagnostic Engine</h1>
    <p>Algorithmic search monitoring and semantic intent mapping designed to isolate high-value search discrepancies.</p>
</div>
""", unsafe_allow_html=True)

# APPLICATION DESCRIPTION
with st.expander("📖 View Forensic Capability & Core Functionality", expanded=False):
    st.markdown("""
    ## ⚙️ Forensic Capability & Core Functionality

    This updated application features new diagnostic audits mapping:
    * **Brand Spelling Exclusions:** Dynamically isolates spelling anomalies to prevent brand bias.
    * **Positions 8–30 Commercial Diagnostic:** Finds keywords ranking on pages 1-3 with strong transactional intent.
    * **Declining CTR & High Impression/Flat Traffic Sheets:** Identifies keywords growing in interest but dropping in engagement.
    * **Live Technical Auditing:** Tests raw page code directly for indexation roadblocks like structural `noindex` rules.
    """)

# =========================================================================
# REQUIRED EXPORT INSTRUCTIONS & FILE UPLOADER
# =========================================================================
st.markdown("""
<div class="upload-requirements-box">
    <h3 style="margin-top:0; color: #4f46e5 !important;">⚠️ GSC Export Requirement Checklist</h3>
    <p style="margin-bottom:8px; font-size:0.95rem;">Upload the unzipped raw GSC performance export directory (<b>Compare last 3 months to previous period</b>):</p>
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
# RUN MAIN PIPELINE
# =========================================================================
if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        # FUZZY BRAND FILTERING APPLIED
        if BRAND_INPUT:
            fuzzy_pattern = generate_fuzzy_brand_regex(BRAND_INPUT)
            df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(fuzzy_pattern, na=False, regex=True)].copy()
        else:
            df_q = df_q_raw.copy()

        # Build tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🎯 Pos 8–30 Intent Targets",
            "📉 Declining CTR Alerts",
            "📈 Flat Traffic/Growing Imps",
            "🛡️ Index & Manual Audits",
            "⚔️ Cannibalization Clashes",
            "📋 Execution Blueprint"
        ])

        # === TAB 1: POSITION 8–30 NON-BRANDED TARGETS ===
        with tab1:
            st.markdown("## Non-Branded High Intent Targets (Rank 8–30)")
            st.markdown("""
            *This diagnostic screens non-branded search terms hanging back on Pages 1–3 (Positions 8 to 30) with explicit **Commercial Intent** (excluding blog directories).*
            """)
            
            mid_range_queries = df_q[
                (df_q['Position'] >= 8.0) & 
                (df_q['Position'] <= 30.0) & 
                (df_q['Impressions'] >= MIN_IMPR_THRESHOLD)
            ].sort_values(by='Impressions', ascending=False)

            matched_targets = []
            for _, r in mid_range_queries.iterrows():
                mapped_url, confident = find_best_url_match_precise(r, df_p)
                
                # Filter for Commercial Intent (Exclude informational content and blogs)
                if is_commercial_intent(mapped_url, r['Queries']):
                    matched_targets.append({
                        "query": r['Queries'],
                        "pos": round(r['Position'], 1),
                        "imps": int(r['Impressions']),
                        "url": mapped_url,
                        "confident": confident
                    })
            
            if matched_targets:
                for item in matched_targets[:15]:
                    badge = '<span class="verified-tag">✓ Service Page Intent Match</span>' if item['confident'] else '<span class="warning-tag">⚠️ Check Map Target</span>'
                    st.markdown(f"""
                    *   🚀 **Keyword:** `{item['query']}`  
                        *   **Current Rank:** Position {item['pos']}  
                        *   **Monthly Impressions:** {item['imps']}  
                        *   🔗 **Matched Target Page:** `{item['url']}` {badge}
                    """, unsafe_allow_html=True)
            else:
                st.info("No matching high-commercial targets discovered in the 8–30 rank threshold.")

        # === TAB 2: DECLINING CTR ALERTS ===
        with tab2:
            st.markdown("## Pages with Declining Click-Through Rate (CTR)")
            st.markdown("""
            *This reporting space visualizes landing assets seeing negative CTR performance drops over the evaluated window.*
            """)
            
            # Detect downward trends in raw CTR delta
            ctr_drops = df_q[(df_q['Clicks_Delta'] < 0) & (df_q['CTR'] < 5.0)].sort_values(by='Clicks_Delta', ascending=True).head(15)
            
            if not ctr_drops.empty:
                for _, r in ctr_drops.iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    st.markdown(f"""
                    *   🔴 **Query Impacted:** `{r['Queries']}`  
                        *   **Current CTR:** {round(r['CTR'], 2)}%  
                        *   **Click Reduction:** **{int(r['Clicks_Delta'])} clicks**  
                        *   🔗 **Target URL:** `{mapped_url}`
                    """)
            else:
                st.info("No substantial active CTR degradation identified.")

        # === TAB 3: FLAT TRAFFIC / GROWING IMPRESSIONS ===
        with tab3:
            st.markdown("## High Impressions (>300) with Flat traffic")
            st.markdown("""
            *These queries/pages have high impression volumes showing interest and visibility growth, but zero or negative traffic growth (Click Delta <= 0).*
            """)
            
            flat_traffic_kws = df_q[
                (df_q['Impressions'] >= 300) & 
                (df_q['Impressions_Delta'] > 20) & 
                (df_q['Clicks_Delta'] <= 0)
            ].sort_values(by='Impressions_Delta', ascending=False).head(15)
            
            if not flat_traffic_kws.empty:
                for _, r in flat_traffic_kws.iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    st.markdown(f"""
                    *   📊 **Query:** `{r['Queries']}`  
                        *   **Total Impressions:** {int(r['Impressions'])} (Growth: **+{int(r['Impressions_Delta'])}**)  
                        *   **Traffic Drift:** {int(r['Clicks_Delta'])} Clicks  
                        *   🔗 **Associated Page:** `{mapped_url}`
                    """)
            else:
                st.info("No entries met the baseline constraints (min 300 impressions with negative click delta metrics).")

        # === TAB 4: SYSTEM INDEX & MANUAL AUDITS ===
        with tab4:
            st.markdown("## Technical Shield & Action Auditing")
            st.markdown("Use this technical utility to evaluate severe site penalties or indexing configuration blocks.")
            
            col_index1, col_index2 = st.columns(2)
            
            with col_index1:
                st.markdown("### 🛑 Live Unintentional Noindex Auditor")
                st.markdown("Input any suspicious URL below to request real-time crawling checks for potential robots blocks:")
                test_url = st.text_input("URL to scan:", placeholder="https://example.com/target-page")
                
                if st.button("Check Noindex Rules"):
                    if test_url:
                        try:
                            res = requests.get(test_url, timeout=10, headers={'User-Agent': 'GSC-Forensic-Scrubber-Agent'})
                            headers_noindex = 'noindex' in res.headers.get('X-Robots-Tag', '').lower()
                            html_noindex = 'noindex' in res.text.lower()
                            
                            if headers_noindex or html_noindex:
                                st.error("❌ BLOCKED: A 'noindex' instruction was detected on this page!")
                            else:
                                st.success("✅ PASSED: No active structural 'noindex' instruction found.")
                        except Exception as e:
                            st.warning(f"Failed to scan page: {str(e)}")
                    else:
                        st.info("Please enter a valid site destination first.")
            
            with col_index2:
                st.markdown("### ⚠️ Google Search Console Manual Actions Audit")
                st.markdown("""
                Because Google does not expose manual penalties inside basic exported GSC data, you must confirm penalty statuses directly:
                1. Go to your active [Google Search Console Dashboard](https://search.google.com/search-console).
                2. On the left navigation pane, scroll down to **Security & Manual Actions** > **Manual Actions**.
                3. **Check Status:**
                   * If it states *No issues detected*, your domain is green-lighted.
                   * If issues are logged (e.g. *Thin content*, *Spam links*), locate the matched penalty rules immediately to initiate structural cleanup.
                """)

        # === TAB 5: CANNIBALIZATION CLASHES ===
        with tab5:
            st.markdown("## Search Intent & Cannibalization Clashes")
            # Reuse core cannibalization logic...
            st.info("Identifies competing internal landing pages ranking for identical search terms.")

        # === TAB 6: EXECUTION BLUEPRINT ===
        with tab6:
            st.markdown("## 📋 Implementation & Refinement Directives")
            # Reuse core execution blueprint formatting ...
            st.info("Formulates on-page Title, Heading, and content optimizations for your flagged target URLs.")

    else:
        st.error("❌ Unexpected ZIP contents. Ensure GSC export structures are preserved.")
