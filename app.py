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
    page_title="GSC Forensic Overview (3-Month Comparison)",
    page_icon="🕵️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* COMPACT SIDEBAR & ULTRA-WIDE RESULTS CONTAINER OVERRIDES */
    [data-testid="stSidebar"] {
        min-width: 15rem !important;
        max-width: 15rem !important;
        width: 15rem !important;
    }
    [data-testid="stAppViewBlockContainer"] {
        max-width: 96% !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-top: 2rem !important;
    }
    
    /* Typography & Headers */
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
    <p>Algorithmic search monitoring and 3-month performance comparison mapping to isolate high-value search discrepancies.</p>
</div>
""", unsafe_allow_html=True)

# APPLICATION DESCRIPTION
with st.expander("📖 View Forensic Capability & Core Functionality", expanded=False):
    st.markdown("""
    ## ⚙️ Forensic Capability & Core Functionality

    This application serves as an automated search intelligence auditor comparing **last 3 months to previous period** GSC performance sets:

    * **Core Update Diagnostic:** Measures 3-month sitewide volume movement (gains vs. losses) to flag potential algorithmic impact.
    * **Page 1 CTR Gaps:** Isolates keywords ranking on Page 1 (ranks 1–10) over the last 3 months underperforming expected CTR benchmarks.
    * **Cannibalization Clashes:** Flags competing internal landing pages matching for identical query intent sets and details both ranking positions on SERPs.
    * **Striking Distance Quick Wins:** Finds queries hovering between position 8 and 30 that expanded visibility or impressions over the last 3 months.
    * **Execution Blueprint (Top 5):** Aggregates weighted diagnostic threats to highlight the Top 5 priority landing pages needing execution.
    """)

# =========================================================================
# REQUIRED EXPORT INSTRUCTIONS
# =========================================================================
st.markdown("""
<div class="upload-requirements-box">
    <h3 style="margin-top:0; color: #4f46e5 !important;">⚠️ GSC Export Requirement Checklist</h3>
    <p style="margin-bottom:8px; font-size:0.95rem;">To construct comparative 3-month diagnostics, upload the <b>unzipped raw ZIP archive</b> exported from Google Search Console:</p>
    <ul style="margin-top:0; margin-bottom:0; font-size:0.95rem; line-height: 1.6;">
        <li>Go to Google Search Console performance menu.</li>
        <li>Set Date filter range to: <b>Compare last 3 months to previous period</b>.</li>
        <li>Click <b>Export</b> in the top right corner and select <b>Download ZIP</b>.</li>
        <li>Upload that unaltered ZIP file directly below.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload GSC ZIP Archive (3-Month Comparison layout):", 
    type=["zip"], 
    key="gsc_zip_uploader"
)

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
    
    geo_words = ['hoboken', 'weehawken', 'nj', 'new-jersey', 'newjersey', 'oak-brook', 'oakbrook', 'il', 'chicago', 'wadena', 'mn', 'minnesota']
    junk_filters = ['1aesthetic', '1-aesthetic', 'aesthetic', 'clinic', 'dr', 'doctor', 'med spa', 'medspa']
    
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
            
    primary_topic = clean_kws[0].title() if clean_kws else "Core Service"
    if len(primary_topic) < 3:
        primary_topic = "Clinical Service"

    loc_suffix = f" in {detected_location}" if detected_location else ""
    
    is_blog = any(pattern in url_lower for pattern in ['/blog', '/news', '/article', '/resource', '/post', '/insight', '/learning'])
    info_modifiers = ['how', 'why', 'what', 'guide', 'tips', 'best', 'causes', 'timeline', 'swelling', 'recovery', 'side effects']
    if any(mod in primary_topic.lower() for mod in info_modifiers):
        is_blog = True

    if is_blog:
        page_type = "Informational / Blog Post"
        meta_title_directive = f"Action Needed: Rewrite Title to target informative intent for '{primary_topic}'. Structure: '[Topic/Question]{loc_suffix} | Guide & Insights' (< 60 chars)."
        meta_desc_directive = f"Action Needed: Write a direct editorial answer for '{primary_topic.lower()}'. Focus on clarifying user intent (< 160 chars)."
        h1_directive = f"Rewrite H1: e.g., 'Understanding {primary_topic}: Recovery, Timeline & Expectations'"
        h2_directive = f"Add section H2: e.g., 'What to Expect During Your {primary_topic} Journey'"
        copy_direction = f"Include structured subheadings covering timelines, patient considerations, and key takeaways for '{primary_topic.lower()}'."
    else:
        page_type = "Transactional / Service Page"
        meta_title_directive = f"Action Needed: Rewrite Title to target localized transactional intent. Format: '{primary_topic}{loc_suffix} | Professional Treatments' (< 60 chars)."
        meta_desc_directive = f"Action Needed: Write a high-converting description for '{primary_topic.lower()}'. Add a clear CTA like 'Book your consultation today.' (< 160 chars)."
        h1_directive = f"Rewrite H1: e.g., 'Expert {primary_topic} Services{loc_suffix}'"
        h2_directive = f"Add supporting H2: e.g., 'Why Choose Our Team for {primary_topic}'"
        copy_direction = f"Feature a clear booking CTA above the fold, highlight clinical experience with '{primary_topic.lower()}', and answer common conversion FAQs."
        
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
        
        has_comparison_metrics = 'Clicks_Delta' in df_q.columns and (df_q['Clicks_Delta'].abs().sum() > 0)
        
        # STRICT 5-TAB SELECTION AS REQUESTED
        tab_names = [
            "🔍 Core Update Diagnostic",
            "🎯 Page 1 CTR Gaps",
            "⚔️ Cannibalization Clashes",
            "🚀 Striking Distance Quick Wins",
            "📋 Execution Blueprint (Top 5)"
        ]
        
        tabs = st.tabs(tab_names)

        # Priority extraction pools for Execution Blueprint
        ctr_gaps_extracted = []
        cannibal_clashes_extracted = []
        striking_extracted = []

        # === TAB 1: CORE UPDATE DIAGNOSTIC ===
        with tabs[0]:
            st.markdown("## Algorithmic Updates Checker")
            
            if not has_comparison_metrics:
                st.warning("⚠️ **3-Month Comparison Data Missing:** Please ensure you uploaded a ZIP generated using GSC's 'Compare last 3 months to previous period' date filter.")
            else:
                st.markdown("""
                *Measures total volume movement across all keywords comparing the **last 3 months to the previous 3-month period**. It checks whether domain traffic drops indicate site-wide algorithmic penalties or routine keyword movement.*
                """)
                losing_keys = df_q[df_q['Clicks_Delta'] < 0]
                gaining_keys = df_q[df_q['Clicks_Delta'] > 0]
                total_lost_clicks = abs(losing_keys['Clicks_Delta'].sum())
                total_gained_clicks = gaining_keys['Clicks_Delta'].sum()
                
                core_hit_score = 0.0
                if total_lost_clicks > 0:
                    core_hit_score = round((total_lost_clicks / (total_lost_clicks + total_gained_clicks + 1e-5)) * 100, 1)

                if core_hit_score > 65.0:
                    st.markdown(f"""
                    <div class="critical-status-highlight">
                        🚨 SYSTEMIC ALGORITHMIC SUPPRESSION FLAGGED ({core_hit_score}% PROBABILITY)
                        <div style="font-size: 0.95rem; font-weight: normal; margin-top: 8px; color: #fee2e2;">
                            Significant sitewide click drops observed across the 3-month comparative window. Immediate technical and content updates required.
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="directive-card success" style="border-left-width: 8px;">
                        <div class="directive-title" style="color: #047857 !important; font-size: 1.3rem;">✅ Stable Organic Profile ({core_hit_score}% Core Impact Score)</div>
                        <div class="directive-text">No sitewide algorithmic issues detected over the 3-month comparative period.</div>
                    </div>
                    """, unsafe_allow_html=True)

        # === TAB 2: PAGE 1 CTR GAPS ===
        with tabs[1]:
            st.markdown("## High-Value Page 1 CTR Gaps")
            st.markdown("""
            *Isolates queries ranking on Page 1 (ranks 1–10) over the last 3 months performing under standard CTR benchmarks, revealing missed click potential.*
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
                        badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match</span>'
                        st.markdown(f"""
                        *   🎯 **Keyword:** `{item['Keyword']}` (Rank: **{item['Rank']}**)  
                            *   **Your CTR:** {item['Actual CTR']} *(Benchmark: {item['Target CTR']})*  
                            *   📉 **Estimated Clicks Lost:** **-{item['Click Loss']} Clicks**
                            *   🔗 **Target URL:** `{item['URL']}` {badge}
                        """, unsafe_allow_html=True)
            else:
                st.info("No high-value Page 1 CTR gaps observed matching criteria.")

        # === TAB 3: CANNIBALIZATION CLASHES ===
        with tabs[2]:
            st.markdown("## ⚔️ Keyword Cannibalization Clashes")
            st.markdown("""
            *Detects queries where **2 or more distinct internal URLs** are ranking simultaneously in search results. Check their individual positions below to determine if they are competing on the **same SERP**.*
            """)
            
            clash_detected = False
            candidate_queries = df_q[(df_q['Impressions'] >= MIN_IMPR_THRESHOLD) & (df_q['Position'] <= 50)].sort_values(by='Impressions', ascending=False)
            
            for _, q_row in candidate_queries.iterrows():
                query = q_row['Queries']
                query_tokens = [re.sub(r'[^a-z0-9]', '', t) for t in query.lower().split() if len(t) > 3]
                
                if len(query_tokens) >= 1:
                    matching_pages = df_p[df_p['Pages'].str.lower().apply(
                        lambda x: sum(1 for token in query_tokens if token in x) >= max(1, len(query_tokens) - 1)
                    )].copy()
                    
                    matching_pages = matching_pages.sort_values(by='Impressions', ascending=False).drop_duplicates(subset=['Pages'])
                    
                    if len(matching_pages) >= 2:
                        clash_detected = True
                        top_competing_pages = matching_pages.head(3)
                        
                        page_list = top_competing_pages['Pages'].tolist()
                        cannibal_clashes_extracted.append({"query": query, "url_1": page_list[0], "url_2": page_list[1]})
                        
                        ranks = top_competing_pages['Position'].tolist()
                        both_page_one = all(r <= 10.0 for r in ranks[:2])
                        serp_status = "⚠️ Active Direct SERP Competition (Both on Page 1)" if both_page_one else "⚡ Keyword Splitting / Alternating Ranks"
                        
                        st.markdown(f"""
                        <div class="directive-card warning" style="border-left: 6px solid #f59e0b !important;">
                            <div class="directive-title" style="font-size: 1.15rem; color: #b45309 !important;">
                                ⚔️ Clashing Keyword: <code>{query}</code>
                            </div>
                            <div class="directive-text" style="margin-bottom: 10px;">
                                <b>SERP Collision Status:</b> <span class="warning-tag" style="font-size: 0.85rem;">{serp_status}</span>
                            </div>
                            <table style="width:100%; border-collapse: collapse; margin-top: 8px; font-size: 0.9rem;">
                                <tr style="background-color: rgba(128,128,128,0.1); text-align: left;">
                                    <th style="padding: 6px 10px;">Competing URL Path</th>
                                    <th style="padding: 6px 10px;">Rank / Pos</th>
                                    <th style="padding: 6px 10px;">Clicks</th>
                                    <th style="padding: 6px 10px;">Impressions</th>
                                </tr>
                        """, unsafe_allow_html=True)
                        
                        for _, p_row in top_competing_pages.iterrows():
                            st.markdown(f"""
                                <tr style="border-bottom: 1px solid rgba(128,128,128,0.2);">
                                    <td style="padding: 6px 10px;"><code>{p_row['Pages']}</code></td>
                                    <td style="padding: 6px 10px;"><b>{round(p_row['Position'], 1)}</b></td>
                                    <td style="padding: 6px 10px;">{int(p_row['Clicks'])}</td>
                                    <td style="padding: 6px 10px;">{int(p_row['Impressions'])}</td>
                                </tr>
                            """, unsafe_allow_html=True)
                            
                        st.markdown("""
                            </table>
                            <div style="margin-top:10px; font-size:0.85rem; color:var(--text-color);">
                                💡 <b>Recommendation:</b> Decide which URL has higher conversion intent. Add a <code>rel="canonical"</code> tag pointing to the primary page, adjust internal anchor links, or consolidate thin content into the stronger ranking page.
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
            if not clash_detected:
                st.success("✅ No keyword cannibalization clashes found matching current filters.")

        # === TAB 4: STRIKING DISTANCE QUICK WINS (3-MONTH COMPARISON LOGIC) ===
        with tabs[3]:
            st.markdown("## Striking Distance Quick Wins (3-Month Comparison)")
            
            if not has_comparison_metrics:
                st.warning("⚠️ **3-Month Comparison Data Missing:** Upload a GSC export using 'Compare last 3 months to previous period' filter to view comparison metrics.")
            else:
                st.markdown("""
                *Identifies queries currently positioned between **ranks 8 and 30** that show **positive impression growth over the last 3 months compared to the prior 3-month period**. These terms represent fast, high-impact Page 1 wins.*
                """)
                
                striking_queries = df_q[
                    (df_q['Position'] >= 7.5) & 
                    (df_q['Position'] <= 30.4) & 
                    (df_q['Impressions_Delta'] > 0)
                ].sort_values(by='Impressions_Delta', ascending=False).head(15)
                
                if not striking_queries.empty:
                    for _, r in striking_queries.iterrows():
                        mapped_url, confident = find_best_url_match_precise(r, df_p)
                        if confident or SHOW_UNVERIFIED:
                            striking_extracted.append({"query": r['Queries'], "impressions": int(r['Impressions']), "impr_delta": int(r['Impressions_Delta']), "url": mapped_url})
                            badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Low Token Match</span>'
                            st.markdown(f"""
                            *   🚀 **Keyword:** `{r['Queries']}`  
                                *   **Current Rank:** **{round(r['Position'], 1)}**  
                                *   **3-Month Impression Growth:** <span style="color:#10b981; font-weight:bold;">+{int(r['Impressions_Delta'])} impressions</span> *(Total: {int(r['Impressions'])})*  
                                *   🔗 **Target Page URL:** `{mapped_url}` {badge}
                            """, unsafe_allow_html=True)
                else:
                    st.info("No striking distance queries with positive 3-month impression growth detected.")

        # === TAB 5: EXECUTION BLUEPRINT (TOP 5 PRIORITY) ===
        with tabs[4]:
            st.markdown("## 📋 Execution Blueprint: Top 5 Priority Action Items")
            st.markdown("""
            Automatically aggregates and scores issues across Page 1 CTR gaps, Cannibalization conflicts, and 3-Month Striking Distance gains to deliver your **Top 5 highest priority landing pages**.
            """)

            blueprint_tasks = {}

            for item in ctr_gaps_extracted:
                url = item['url']
                if url == "Manual GSC Check Required": continue
                if url not in blueprint_tasks:
                    blueprint_tasks[url] = {"queries": set(), "reasons": set(), "score": 0.0}
                blueprint_tasks[url]["queries"].add(item['query'])
                blueprint_tasks[url]["reasons"].add("CTR Gap")
                blueprint_tasks[url]["score"] += item['loss']

            for item in cannibal_clashes_extracted:
                for u_key in ['url_1', 'url_2']:
                    url = item[u_key]
                    if url == "Manual GSC Check Required": continue
                    if url not in blueprint_tasks:
                        blueprint_tasks[url] = {"queries": set(), "reasons": set(), "score": 0.0}
                    blueprint_tasks[url]["queries"].add(item['query'])
                    blueprint_tasks[url]["reasons"].add("Cannibalization Clash")
                    blueprint_tasks[url]["score"] += 150.0

            for item in striking_extracted:
                url = item['url']
                if url == "Manual GSC Check Required": continue
                if url not in blueprint_tasks:
                    blueprint_tasks[url] = {"queries": set(), "reasons": set(), "score": 0.0}
                blueprint_tasks[url]["queries"].add(item['query'])
                blueprint_tasks[url]["reasons"].add("3-Month Striking Distance Growth")
                blueprint_tasks[url]["score"] += (item['impr_delta'] * 0.05)

            top_5_tasks = dict(sorted(blueprint_tasks.items(), key=lambda x: x[1]['score'], reverse=True)[:5])

            if top_5_tasks:
                st.markdown(f"### Isolated **{len(top_5_tasks)}** Priority URLs Requiring Action")
                
                export_data = []
                
                for idx, (url, task_data) in enumerate(top_5_tasks.items()):
                    keywords = list(task_data["queries"])
                    reasons = ", ".join(list(task_data["reasons"]))
                    
                    directives = generate_seo_recommendations(url, keywords)
                    
                    export_data.append({
                        "Priority Rank": f"Rank {idx+1}",
                        "Priority Score": round(task_data["score"], 1),
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
                    <div class="directive-card" style="border-left: 6px solid #ef4444 !important;">
                        <div class="directive-title" style="color: #ef4444 !important;">
                            🚨 PRIORITY CRITICAL #{idx+1} [Impact Score: {round(task_data["score"], 1)}]
                        </div>
                        <div class="directive-title">🔗 URL: <a href="{url}" target="_blank" style="color: #4f46e5; text-decoration: underline;">{url}</a></div>
                        <div class="url-helper-box">
                            <b>Page Type:</b> {directives["page_type"]}<br>
                            <b>Diagnosed Alert:</b> <span style="color:#ef4444; font-weight:bold;">{reasons}</span><br>
                            <b>Target Keywords:</b> {", ".join([f"<code>{k}</code>" for k in keywords])}
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
                            <p><b>✍️ On-Page Copy Blueprint:</b></p>
                            <blockquote style="margin: 5px 0; padding: 8px 15px; border-left: 3px solid #f59e0b; background: var(--secondary-background-color);">{directives["copy_direction"]}</blockquote>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                df_export = pd.DataFrame(export_data)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, sheet_name='Top 5 Priority Directives', index=False)
                    workbook  = writer.book
                    worksheet = writer.sheets['Top 5 Priority Directives']
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
                    label="📥 Download Top 5 Priority Implementation Workbook (Excel)",
                    data=output.getvalue(),
                    file_name="GSC_SEO_Top_5_Priority_Blueprint.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No dynamic tasks could be assigned. Adjust your filtration parameters on the sidebar.")
    else:
        st.error("Uploaded ZIP does not appear to contain matching 'Queries' and 'Pages' CSV files.")
else:
    st.info("👋 Upload a raw GSC ZIP archive (3-Month Comparison layout) above to begin your audit.")
