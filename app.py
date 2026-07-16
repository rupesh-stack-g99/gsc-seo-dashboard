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

    /* Ultra-highlighted status banner */
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

    /* Redesigned math explanation box */
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
    "Exclude Branded Keywords:", 
    value="", 
    placeholder="e.g. brand, clinic, dr smith"
).lower().strip()

MIN_IMPR_THRESHOLD = st.sidebar.number_input("Min. Impressions:", min_value=1, value=100)

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

# APPLICATION DESCRIPTION
with st.expander("📖 View Forensic Capability & Core Functionality (What this app does & finds)", expanded=False):
    st.markdown("""
    ## ⚙️ Forensic Capability & Core Functionality

    This application serves as an automated search intelligence auditor that processes raw, multi-dimensional Google Search Console data structures:

    * **What it does:** It cleanses, parses, and cross-references multi-dimensional performance sets (Queries and Pages) via a high-precision semantic matching algorithm.
    * **What it finds:**
        * **Systemic Core Update Impact:** Calculates the exact balance of search term decay vs. growth to flag potential sitewide algorithmic updates.
        * **Hidden Keyword Decay:** Pinpoints queries losing substantial click-through metrics while keeping stable impression scores.
        * **High-Value Page 1 CTR Gaps:** Detects terms ranking on Page 1 that are performing below natural CTR curves, isolating traffic loss.
        * **Keyword Cannibalization Clashes:** Flags competing internal landing pages ranking for identical search terms.
        * **Hidden Growth (Flat Traffic):** Pinpoints pages generating high visibility gains (impressions) without generating corresponding click increases.
        * **Striking-Distance Quick Wins:** Finds queries hovering on the cusp of page one (positions 11–15) with high impressions.
    """)

# =========================================================================
# REQUIRED EXPORT INSTRUCTIONS & FILE UPLOADER
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
        
        for critical_word in ['kybella', 'earlobe', 'piercing', 'mounjaro', 'tirzepatide', 'botox', 'sculptra', 'semaglutide', 'ozempic', 'wegovy']:
            if critical_word in query_str:
                if critical_word in url_path:
                    score += 20  
                else:
                    score -= 15  
                    
        for geo in ['weehawken', 'hoboken', 'jersey', 'oak brook', 'oakbrook', 'wadena']:
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
# DIRECTIVE-BASED METADATA INSTRUCTION BLUEPRINTS
# -------------------------------------------------------------------------
def generate_seo_recommendations(page_url, keywords):
    url_lower = page_url.lower()
    
    geo_words = [
        'hoboken', 'weehawken', 'nj', 'new-jersey', 'newjersey', 
        'oak-brook', 'oakbrook', 'il', 'chicago', 'wadena', 'mn', 'minnesota'
    ]
    
    junk_filters = [
        '1aesthetic', '1-aesthetic', 'aesthetic', 'clinic', 'dr', 'doctor', 'med spa', 'medspa'
    ]
    
    detected_location = ""
    all_kws_flat = " ".join(keywords).lower()
    
    if 'wadena' in all_kws_flat or 'mn' in all_kws_flat:
        detected_location = "Wadena, MN"
    elif 'hoboken' in all_kws_flat:
        detected_location = "Hoboken, NJ"
    elif 'weehawken' in all_kws_flat:
        detected_location = "Weehawken, NJ"
    elif 'oak brook' in all_kws_flat or 'oakbrook' in all_kws_flat:
        detected_location = "Oak Brook, IL"
    else:
        if 'hoboken' in url_lower:
            detected_location = "Hoboken, NJ"
        elif 'weehawken' in url_lower:
            detected_location = "Weehawken, NJ"
        elif 'oak-brook' in url_lower or 'oakbrook' in url_lower:
            detected_location = "Oak Brook, IL"
        elif 'wadena' in url_lower:
            detected_location = "Wadena, MN"
        else:
            detected_location = ""

    clean_kws = []
    for kw in keywords:
        kw_cleaned = kw.lower()
        for geo in geo_words:
            kw_cleaned = re.sub(rf'\b{geo}\b', '', kw_cleaned).strip()
        for junk in junk_filters:
            kw_cleaned = re.sub(rf'\b{junk}\b', '', kw_cleaned).strip()
        
        kw_cleaned = re.sub(r'\s+', ' ', kw_cleaned).strip()
        if kw_cleaned:
            clean_kws.append(kw_cleaned)
            
    primary_topic = clean_kws[0].title() if clean_kws else "Core Treatment"
    
    if len(primary_topic) < 3:
        primary_topic = "Clinical Treatment"

    loc_suffix = f" in {detected_location}" if detected_location else ""
    
    is_blog = any(pattern in url_lower for pattern in ['/blog', '/news', '/article', '/resource', '/post', '/insight', '/learning'])
    info_modifiers = ['how', 'why', 'what', 'guide', 'tips', 'best', 'causes', 'timeline', 'swelling', 'recovery', 'side effects']
    if any(mod in primary_topic.lower() for mod in info_modifiers):
        is_blog = True

    if is_blog:
        page_type = "Informational / Blog Post"
        meta_title_directive = f"Action Needed: Rewrite Title to target informative intent for '{primary_topic}'. Structure: '[Topic/Question]{loc_suffix} | Practical Guide & Recovery' (< 60 chars)."
        meta_desc_directive = f"Action Needed: Write a helpful editorial summary focusing on '{primary_topic.lower()}'. Direct the user to a clinical answer immediately (< 160 chars)."
        h1_directive = f"Rewrite H1 to address the search intent directly: e.g., 'Understanding {primary_topic}: Recovery, Milestones & Practical Expectations'"
        h2_directive = f"Use an answer-target H2 structure: e.g., 'How Long Does {primary_topic} Take to Settle?'"
        copy_direction = f"Ensure this article contains clear section subheadings addressing recovery timelines, side effects, and practical checklists for patients researching '{primary_topic.lower()}'."
    else:
        page_type = "Transactional / Service Page"
        meta_title_directive = f"Action Needed: Rewrite Title to target localized transactional intent. Format: '{primary_topic}{loc_suffix} | Restorative Clinical Treatment' (< 60 chars)."
        meta_desc_directive = f"Action Needed: Write a localized, high-converting service description for '{primary_topic.lower()}'. Offer a direct CTA like 'Request your consultation today.' (< 160 chars)."
        h1_directive = f"Rewrite H1 to establish immediate clinical relevance: e.g., 'Custom {primary_topic} Treatments{loc_suffix}'"
        h2_directive = f"Add a benefit-driven supporting H2: e.g., 'Restore Comfort and Clinical Balance with Customized {primary_topic}'"
        copy_direction = f"The content must feature a clear booking CTA above the fold, highlight practitioner experience with '{primary_topic.lower()}', and present clear FAQs about benefits and booking."
        
    return {
        "title_directive": meta_title_directive,
        "desc_directive": meta_desc_directive,
        "h1_directive": h1_directive,
        "h2_directive": h2_directive,
        "copy_direction": copy_direction,
        "page_type": page_type
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
        
        # MULTI-KEYWORD EXCLUSION SYSTEM
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

        # 7-Tab Setup (New tab added at Index 4)
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "🔍 Core Update Diagnostic",
            "📉 Keyword Decay Alerts",
            "🎯 Page 1 CTR Gaps",
            "⚔️ Cannibalization Clashes",
            "📈 Hidden Growth (Flat Traffic)",
            "🚀 Striking Distance Quick Wins",
            "📋 Execution Blueprint"
        ])

        # Prepare extraction pools to cross-populate Tab 7 (Execution Blueprint)
        decayed_extracted = []
        ctr_gaps_extracted = []
        cannibal_clashes_extracted = []
        hidden_growth_extracted = []
        striking_extracted = []

        # === TAB 1: CORE UPDATE HIT DETECTOR ===
        with tab1:
            st.markdown("## Algorithmic Updates Checker")
            st.markdown("""
            *This diagnostic analysis measures the systemic stability of the organic profile. By evaluating global ratios of decaying keywords against ascending terms, it assesses whether traffic contractions point toward site-wide algorithmic suppression or minor seasonal turbulence.*
            """)
            
            if core_hit_score > 65.0:
                st.markdown(f"""
                <div class="critical-status-highlight">
                    🚨 SYSTEMIC ALGORITHMIC SUPPRESSION FLAGGED ({core_hit_score}% PROBABILITY)
                    <div style="font-size: 0.95rem; font-weight: normal; margin-top: 8px; color: #fee2e2;">
                        This site is experiencing a lopsided site-wide decline. Priority technical and content-level fixes are recommended.
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="directive-card success" style="border-left-width: 8px;">
                    <div class="directive-title" style="color: #047857 !important; font-size: 1.3rem;">✅ Stable Organic Profile ({core_hit_score}% Core Impact Score)</div>
                    <div class="directive-text">No systemic, site-wide algorithmic penalties detected. Keyword fluctuations are normal and healthy.</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("### 🧮 How is this score calculated? (The Math)")
            st.latex(r"\text{Core Hit Score} = \left( \frac{\sum \text{Clicks Lost across Decaying Queries}}{\sum \text{Clicks Lost} + \sum \text{Clicks Gained}} \right) \times 100")
            
            st.markdown(f"""
            <div class="math-explanation-box">
                <span style="font-weight: 700; color: #4f46e5; font-size: 1.0rem;">💡 Understanding the Mathematical Logic & Ratios</span><br>
                <p style="margin-top: 6px; margin-bottom: 8px;">Think of this score as a <b>site-wide organic health balance sheet</b>. Instead of just looking at whether total traffic went up or down, this engine looks at <b>how many individual keywords are shrinking vs. growing</b>.</p>
                <ul style="margin-top:0; margin-bottom:8px; padding-left:20px;">
                    <li><b>The Volatility Grouping:</b>
                        <ul>
                            <li><b>Losing Keywords:</b> We identify every query on your site that lost clicks over the comparative period and sum those click losses.</li>
                            <li><b>Gaining Keywords:</b> We identify every query that gained clicks over the same period and sum those wins.</li>
                        </ul>
                    </li>
                    <li><b>The Balancing Ratio:</b> We calculate what percentage of the active click movement is negative. If the result is, for example, <b>84.2%</b>, it means that 84.2% of all organic click volatility across your domain is down.</li>
                    <li><b>Why 65% is the Threshold:</b> 
                        <ul>
                            <li><b>Below 65% (Normal Fluctuation):</b> It is normal for a few pages to drop while others grow due to local competition, mild keyword decay, or seasonality.</li>
                            <li><b>Above 65% (Algorithmic Warning):</b> If more than 65% of your keyword volatility points downward simultaneously, it is statistically impossible for this to be a local page issue. This suggests <b>Google's core algorithms have updated how they evaluate your sitewide E-E-A-T, quality, or trust metrics</b>.</li>
                        </ul>
                    </li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

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
                        decayed_extracted.append({"query": r['Queries'], "clicks_lost": int(r['Clicks_Delta']), "url": mapped_url})
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match - Verify URL</span>'
                        st.markdown(f"""
                        *   🔴 **Keyword:** `{r['Queries']}`  
                            *   **Current Rank:** {round(r['Position'], 1)}  
                            *   **Click Shift:** **{int(r['Clicks_Delta'])} clicks**  
                            *   🎯 **Best Match URL:** `{mapped_url}` {badge}
                        """, unsafe_allow_html=True)
            else:
                st.info("No query decay matches found with current filtration parameters.")

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
                        ctr_gaps_extracted.append({"query": item['Keyword'], "loss": item['Click Loss'], "url": item['URL']})
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Verification Recommended via GSC</span>'
                        st.markdown(f"""
                        *   🎯 **Keyword:** `{item['Keyword']}` (Rank: **{item['Rank']}**)  
                            *   **Your CTR:** {item['Actual CTR']} *(Expected Benchmark: {item['Target CTR']})*  
                            *   📉 **Estimated Loss:** **-{item['Click Loss']} Clicks** *(Baseline variant comparison)*
                            *   🔗 **Target Landing Page:** `{item['URL']}` {badge}
                        """, unsafe_allow_html=True)
            else:
                st.info("No high-value Page 1 CTR gaps observed matching criteria.")

        # === TAB 4: CANNIBALIZATION CLASHES ===
        with tab4:
            st.markdown("## Keyword Cannibalization Clashes")
            st.markdown("""
            *This tab isolates internal competing assets. When multiple pages rank concurrently for identical search queries, Google's crawling agent splits structural authority across them, diluting individual rankings.*
            """)
            
            clash_detected = False
            for _, row in df_q[(df_q['Impressions'] >= MIN_IMPR_THRESHOLD) & (df_q['Position'] <= 20)].head(10).iterrows():
                query = row['Queries']
                query_tokens = [re.sub(r'[^a-z0-9]', '', t) for t in query.lower().split() if len(t) > 3]
                if len(query_tokens) >= 2:
                    matching_pages = df_p[df_p['Pages'].str.lower().apply(lambda x: sum(1 for token in query_tokens if token in x)) >= 2]
                    if len(matching_pages) >= 2:
                        clash_detected = True
                        page_list = matching_pages['Pages'].head(2).tolist()
                        cannibal_clashes_extracted.append({"query": query, "url_1": page_list[0], "url_2": page_list[1]})
                        st.markdown(f"""
                        <div class="directive-card danger">
                            <div class="directive-title">⚔️ Cannibalization Detected: "{query}"</div>
                            <div class="directive-text">
                                Multiple pages are matching semantic tokens for this intent. Google is splitting ranking weight:
                                <ul style="margin-top:6px; margin-bottom:0;">
                                    <li><b>Asset A:</b> <code>{page_list[0]}</code></li>
                                    <li><b>Asset B:</b> <code>{page_list[1]}</code></li>
                                </ul>
                                <span class="warning-tag">Recommendation: Canonicalize or Consolidation Action Plan Required</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
            if not clash_detected:
                st.success("Excellent! No critical structural keyword cannibalization clashes found in this date range.")

        # === TAB 5: HIDDEN GROWTH (FLAT TRAFFIC) ===
        with tab5:
            st.markdown("## Pages with Growing Impressions but Flat Traffic")
            st.markdown("""
            *This diagnostic analysis filters queries experiencing rapid visibility expansions (impressions increasing by **$\ge$ 1,000** over the 3-month comparative period) where actual organic clicks have stagnated or dropped ($\le 0$). This identifies prime structural opportunities where search demand is surging but your snippet is failing to attract the user's click.*
            """)
            
            # Extract growing impressions with flat/declining traffic (optimized threshold for 90-day comparison)
            flat_growth_queries = df_q[
                (df_q['Impressions_Delta'] >= 1000) & 
                (df_q['Clicks_Delta'] <= 0)
            ].sort_values(by='Impressions_Delta', ascending=False).head(15)
            
            if not flat_growth_queries.empty:
                for _, r in flat_growth_queries.iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    if confident or SHOW_UNVERIFIED:
                        hidden_growth_extracted.append({"query": r['Queries'], "impr_growth": int(r['Impressions_Delta']), "url": mapped_url})
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match</span>'
                        st.markdown(f"""
                        <div class="directive-card warning">
                            <div class="directive-title">📈 High Demand Capture Opportunity: "{r['Queries']}"</div>
                            <div class="directive-text">
                                <ul style="margin-top:6px; margin-bottom:4px;">
                                    <li><b>Impression Growth (3-Month Demand Surge):</b> <span style="color:#10b981; font-weight:bold;">+{int(r['Impressions_Delta'])} views</span></li>
                                    <li><b>Click Volatility (Stagnant Traffic):</b> <span style="color:#ef4444; font-weight:bold;">{int(r['Clicks_Delta'])} clicks</span></li>
                                    <li><b>Current Ranking Position:</b> {round(r['Position'], 1)}</li>
                                    <li>🔗 <b>Identified Target URL:</b> <code>{mapped_url}</code> {badge}</li>
                                </ul>
                                <span class="verified-tag" style="background-color: #3b82f6; color: white;">Strategic Priority: Improve CTA layout, Title hooks, or Metadata CTR triggers.</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.info("No queries found showing surging impressions ($\ge$ 1,000) alongside flat/declining traffic under the current exclusions.")

        # === TAB 6: STRIKING DISTANCE QUICK WINS ===
        with tab6:
            st.markdown("## Striking Distance Opportunities")
            st.markdown("""
            *These queries are positioned on the cusp of high-traffic slots (ranks 11 to 15). With targeted structural enhancements, these assets can be vaulted onto Page 1 to capture incoming clicks.*
            """)
            
            striking_queries = df_q[(df_q['Position'] >= 10.5) & (df_q['Position'] <= 15.4)].sort_values(by='Impressions', ascending=False).head(15)
            
            if not striking_queries.empty:
                for _, r in striking_queries.iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    if confident or SHOW_UNVERIFIED:
                        striking_extracted.append({"query": r['Queries'], "impressions": int(r['Impressions']), "url": mapped_url})
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match</span>'
                        st.markdown(f"""
                        *   🚀 **Keyword:** `{r['Queries']}`  
                            *   **Current Rank:** **{round(r['Position'], 1)}**  
                            *   **Total Monthly Impressions:** **{int(r['Impressions'])}**  
                            *   🔗 **Target Page URL:** `{mapped_url}` {badge}
                        """, unsafe_allow_html=True)
            else:
                st.info("No striking distance queries detected with the current filters.")

        # === TAB 7: EXECUTION BLUEPRINT ===
        with tab7:
            st.markdown("## 📋 Blueprint Engine & Dynamic Content Directive Exporter")
            st.markdown("""
            This blueprint consolidates all diagnostic alerts from other tabs into a structured actionable roadmap. It groups queries by URL and dynamically generates metadata optimizations.
            """)

            # Collate tasks by unique URLs
            blueprint_tasks = {}
            
            all_findings = []
            for item in decayed_extracted:
                all_findings.append({"url": item['url'], "query": item['query'], "source": "Keyword Decay"})
            for item in ctr_gaps_extracted:
                all_findings.append({"url": item['url'], "query": item['query'], "source": "CTR Gap"})
            for item in cannibal_clashes_extracted:
                all_findings.append({"url": item['url_1'], "query": item['query'], "source": "Cannibalization (Primary)"})
                all_findings.append({"url": item['url_2'], "query": item['query'], "source": "Cannibalization (Secondary)"})
            for item in hidden_growth_extracted:
                all_findings.append({"url": item['url'], "query": item['query'], "source": "Hidden Demand Growth"})
            for item in striking_extracted:
                all_findings.append({"url": item['url'], "query": item['query'], "source": "Striking Distance"})

            # Group findings by URL
            for finding in all_findings:
                url = finding['url']
                if url == "Manual GSC Check Required":
                    continue
                if url not in blueprint_tasks:
                    blueprint_tasks[url] = {"queries": set(), "reasons": set()}
                blueprint_tasks[url]["queries"].add(finding['query'])
                blueprint_tasks[url]["reasons"].add(finding['source'])

            if blueprint_tasks:
                st.markdown(f"### Found **{len(blueprint_tasks)}** Unique URLs Requiring On-Page/Technical Adjustments")
                
                export_data = []
                
                for idx, (url, task_data) in enumerate(blueprint_tasks.items()):
                    keywords = list(task_data["queries"])
                    reasons = ", ".join(list(task_data["reasons"]))
                    
                    directives = generate_seo_recommendations(url, keywords)
                    
                    export_data.append({
                        "URL": url,
                        "Page Type": directives["page_type"],
                        "Primary Target Keyword": keywords[0],
                        "Secondary Keywords": ", ".join(keywords[1:]),
                        "Diagnosed Issues": reasons,
                        "Target Meta Title": directives["title_directive"],
                        "Target Meta Description": directives["desc_directive"],
                        "Target H1 Blueprint": directives["h1_directive"],
                        "Target H2 Blueprint": directives["h2_directive"],
                        "On-Page Content Copy Blueprint": directives["copy_direction"]
                    })

                    st.markdown(f"""
                    <div class="directive-card">
                        <div class="directive-title">🔗 Target URL {idx+1}: <a href="{url}" target="_blank" style="color: #4f46e5; text-decoration: underline;">{url}</a></div>
                        <div class="url-helper-box">
                            <b>Page Intent Type:</b> {directives["page_type"]}<br>
                            <b>Aggregated Diagnostic Alerts:</b> <span style="color:#ef4444; font-weight:bold;">{reasons}</span><br>
                            <b>Associated Target Keywords:</b> {", ".join([f"<code>{k}</code>" for k in keywords])}
                        </div>
                        <div style="margin-top: 15px;">
                            <p><b>🏷️ Meta Title Blueprint:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #6366f1; background: var(--secondary-background-color); font-family: monospace;">{directives["title_directive"]}</blockquote>
                            <p><b>📝 Meta Description Blueprint:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #6366f1; background: var(--secondary-background-color); font-family: monospace;">{directives["desc_directive"]}</blockquote>
                            <p><b>🏗️ Heading 1 Blueprint:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #10b981; background: var(--secondary-background-color);">{directives["h1_directive"]}</blockquote>
                            <p><b>🏗️ Supporting H2 Blueprint:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #10b981; background: var(--secondary-background-color);">{directives["h2_directive"]}</blockquote>
                            <p><b>✍️ Copywriting & On-Page Direction:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #f59e0b; background: var(--secondary-background-color);">{directives["copy_direction"]}</blockquote>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Excel Binary Packager
                df_export = pd.DataFrame(export_data)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, sheet_name='SEO Executive Directives', index=False)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['SEO Executive Directives']
                    header_format = workbook.add_format({
                        'bold': True, 'text_wrap': True, 'valign': 'top',
                        'fg_color': '#4F46E5', 'font_color': '#FFFFFF', 'border': 1
                    })
                    cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                    
                    worksheet.freeze_panes(1, 0)
                    for col_idx, col in enumerate(df_export.columns):
                        worksheet.write(0, col_idx, col, header_format)
                        worksheet.set_column(col_idx, col_idx, 25, cell_format)
                
                st.download_button(
                    label="📥 Download Executive Implementation Workbook (Excel)",
                    data=output.getvalue(),
                    file_name="GSC_SEO_Forensic_Executive_Blueprint.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No dynamic tasks could be assigned. Adjust your filtration sliders on the sidebar to capture lighter signals.")
    else:
        st.error("Uploaded ZIP does not appear to contain matching 'Queries' and 'Pages' CSV structures.")
else:
    st.info("👋 Upload a raw exported GSC Comparative ZIP archive above to run diagnostic pipelines.")
