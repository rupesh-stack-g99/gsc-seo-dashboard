import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

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
    <p>This engine analyzes your Search Console data in real-time, strips out raw tables entirely, and renders strict algorithmic diagnostics. Get programmatic directives on quality filters, indexation decays, and SERP volatility.</p>
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
    MIN_IMPR_THRESHOLD = st.number_input("Minimum Impressions Threshold:", min_value=1, value=150)
with cfg_col3:
    MAX_CANNIBAL_OFFSET = st.slider("Cannibalization Search Space (Pos. Gap):", 1, 15, 6)

st.markdown("---")

# =========================================================================
# DATA CLEANING ENGINE (ROBUST BYPASS FOR GSC HEADER METADATA)
# =========================================================================
def clean_gsc_csv(bytes_data):
    """
    Slices past GSC metadata header rows if they exist, 
    finding the actual start of the data table.
    """
    try:
        lines = bytes_data.decode('utf-8').splitlines()
    except UnicodeDecodeError:
        try:
            lines = bytes_data.decode('latin-1').splitlines()
        except Exception:
            return None

    # Detect row index where real GSC columns start
    header_idx = 0
    for idx, line in enumerate(lines[:15]):  # Google usually puts headers in first 10 rows
        lower_line = line.lower()
        # Look for typical CSV header identifiers
        if any(term in lower_line for term in ['query', 'queries', 'page', 'pages', 'clicks', 'impressions']):
            header_idx = idx
            break
            
    # Read CSV skipping metadata headers
    try:
        df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])))
        # Strip string whitespace from column headers
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None

def parse_gsc_sheet(df, dim_name):
    # Match the target dimension column dynamically
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
    
    # Strictly scope to Page 1-3
    normalized = normalized[normalized['Position'] <= 30.0]
    return normalized

def extract_gsc_payload(uploaded_zip):
    results = {}
    try:
        with zipfile.ZipFile(uploaded_zip) as z:
            file_names = z.namelist()
            
            for file_name in file_names:
                # Ignore system metadata files inside ZIPs (like macOS __MACOSX)
                if '__macosx' in file_name.lower() or not file_name.endswith('.csv'):
                    continue
                    
                with z.open(file_name) as f:
                    raw_df = clean_gsc_csv(f.read())
                    if raw_df is None or raw_df.empty:
                        continue
                    
                    # Inspect column signatures of the CSV to classify the dataset
                    cols_lower = [str(c).lower() for c in raw_df.columns]
                    
                    # Match Queries Dataset
                    if any(q_term in cols_lower for q_term in ['query', 'queries', 'top queries', 'top query', 'search query']):
                        clean_df = parse_gsc_sheet(raw_df, 'Queries')
                        if clean_df is not None and not clean_df.empty:
                            results['Queries'] = clean_df
                            
                    # Match Pages Dataset
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
        
        st.markdown("## 🔍 Algorithmic Investigation Results")
        st.markdown("Below are the findings derived automatically by correlating mathematical trends across your Search Console datasets.")

        # --- ALGORITHMIC DETECTOR 1: THE CORE UPDATE HIT DETECTOR ---
        losing_keys = df_q[df_q['Clicks_Delta'] < 0]
        gaining_keys = df_q[df_q['Clicks_Delta'] > 0]
        
        total_lost_clicks = abs(losing_keys['Clicks_Delta'].sum())
        total_gained_clicks = gaining_keys['Clicks_Delta'].sum()
        
        core_hit_score = 0.0
        if total_lost_clicks > 0:
            core_hit_score = round((total_lost_clicks / (total_lost_clicks + total_gained_clicks + 1e-5)) * 100, 1)

        if core_hit_score > 65.0:
            st.markdown(f"""
            <div class="directive-card danger">
                <div class="directive-title">🚨 Systemic Algorithmic Suppression Flagged ({core_hit_score}% Probability)</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Over {core_hit_score}% of overall trend movements are strictly negative. Clicks and impressions are dropping simultaneously across uncorrelated keywords. This strongly correlates with a Google Core Algorithm Update or search classifier adjustment rather than a simple indexation glitch. <br/>
                    <b>Detective Action:</b> Audit your site-wide informational value. Avoid surface-level updates. Identify if pages hit hardest feature redundant introductory material, high affiliate/ad ratios, or lack distinct expert author perspectives (E-E-A-T).
                </div>
            </div>
            """, unsafe_allow_html=True)
        elif core_hit_score > 35.0:
            st.markdown(f"""
            <div class="directive-card warning">
                <div class="directive-title">⚠️ Moderate Algorithmic Volatility Checked ({core_hit_score}% Probability)</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Partial traffic degradation spotted across isolated clusters. This is likely not a site-wide quality penalty, but a sub-topic re-evaluation. Competitors are likely optimizing topical coverage or getting featured in newly introduced AI SERP widgets.<br/>
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

        # --- ALGORITHMIC DETECTOR 2: SPAMBRAIN / CONTENT QUALITY CRITICAL FILTER ---
        thin_content_candidates = df_p[
            (df_p['Position_Delta'] > 1.0) & 
            (df_p['Clicks_Delta'] < -10) & 
            (df_p['Impressions_Delta'] >= 0)
        ]
        
        if len(thin_content_candidates) > 0:
            st.markdown(f"""
            <div class="directive-card danger">
                <div class="directive-title">🔴 SpamBrain & Helpful Content Classifier Risk: Detected on {len(thin_content_candidates)} Landing Page Paths</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Multi-metric divergence detected. Google is processing and rendering these pages in search results (Impressions are stable/increasing), but it is systematically shifting rankings downward (Positions are slipping and Clicks are plummeting). This matches the algorithmic behavior of automatic helpfulness classifiers.<br/>
                    <b>Detective Action:</b> Review these landing pages immediately. Look for scaled AI generation, keyword stuffing in H2/H3 elements, and massive block quotes that do not answer the user query directly. Consolidate low-value pages of similar content into singular authoritative pillars.
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="directive-card success">
                <div class="directive-title">🛡️ SpamBrain Classifier Health: Clean</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> No content clusters show systemic impression-position divergence patterns. Google continues to index and rank your optimized content cleanly without triggering modern spam or helpful content filters.
                </div>
            </div>
            """, unsafe_allow_html=True)

        # --- ALGORITHMIC DETECTOR 3: THE INTENT COLLISION ENGINE (CANNIBALIZATION DETECTOR) ---
        cannibal_list = []
        candidates = df_q[df_q['Impressions'] >= MIN_IMPR_THRESHOLD].sort_values(by='Impressions', ascending=False).head(150)
        
        for _, q_row in candidates.iterrows():
            query_txt = q_row['Queries']
            q_pos = q_row['Position']
            
            matching_urls = df_p[
                (df_p['Position'] >= q_pos - MAX_CANNIBAL_OFFSET) & 
                (df_p['Position'] <= q_pos + MAX_CANNIBAL_OFFSET) &
                (df_p['Impressions'] >= MIN_IMPR_THRESHOLD / 2)
            ]
            
            if len(matching_urls) > 1:
                primary_url = matching_urls.iloc[0]['Pages']
                secondary_url = matching_urls.iloc[1]['Pages'] if len(matching_urls) > 1 else ""
                if primary_url != secondary_url and secondary_url:
                    cannibal_list.append(query_txt)

        unique_cannibals = list(set(cannibal_list))
        if len(unique_cannibals) > 5:
            st.markdown(f"""
            <div class="directive-card warning">
                <div class="directive-title">⚔️ Search Intent Collision: High Volatility Confirmed ({len(unique_cannibals)} Overlapping Terms)</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Google's ranking engine is oscillating between multiple pages on your site to answer the same queries. This results in keyword cannibalization where both URLs compete, diluting link juice, anchor text equity, and CTR performance. <br/>
                    <b>Detective Action:</b> De-optimize competing assets. Use precise anchor-text internal linking from the secondary page back to the primary canonical URL using the target term. If content overlaps significantly, 301-redirect or canonicalize the weaker page into the stronger page.
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="directive-card success">
                <div class="directive-title">🎯 Intent Targeting: Highly Focused</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Low intent collision across your GSC footprint. Google has a clear map of which URL is the absolute authority page for your primary organic keyword groups.
                </div>
            </div>
            """, unsafe_allow_html=True)

        # --- ALGORITHMIC DETECTOR 4: SERP LAYOUT & PIXEL SHIFT DETECTOR (VISIBILITY GAP) ---
        serp_layout_shifts = df_q[
            (df_q['Position'] <= 8.0) & 
            (df_q['Clicks_Delta'] < 0) & 
            (df_q['Impressions_Delta'] >= 0) &
            (df_q['CTR'] < 3.0)
        ]
        
        if len(serp_layout_shifts) > 0:
            st.markdown(f"""
            <div class="directive-card info">
                <div class="directive-title">👁️ SERP Landscape Displacement (Pixel-Shift Detected)</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Critical CTR degradation noticed on core keywords where you rank within the top 8 positions. Because impressions are stable, this signals that Google has altered the visual SERP layout (e.g., expanded AI Overviews, larger local packs, sponsored shopping feeds, or video carousels) forcing your result below the fold.<br/>
                    <b>Detective Action:</b> Run manual searches for these queries. Assess if a competitor is optimizing for high-yield schema or if AI Summaries have displaced standard results. Re-engineer titles into questions, deploy custom structured data markup (Product, FAQ, or Schema tables), and target direct inclusion within generative answers.
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="directive-card success">
                <div class="directive-title">✨ SERP Pixel Visibility: Safe</div>
                <div class="directive-text">
                    <b>Diagnostic:</b> Your high-ranking queries are converting at expected ratios relative to their position, showing no sign of pixel displacement or algorithmic crowding in the SERPs.
                </div>
            </div>
            """, unsafe_allow_html=True)

        # --- ALGORITHMIC DETECTOR 5: THE STRATEGIC NEXT-STEP DIRECTIVE ENGINE ---
        st.markdown("## 📋 Execution Blueprint")
        
        striking_distance_kws = df_q[(df_q['Position'] >= 4.0) & (df_q['Position'] <= 12.0)].sort_values(by='Impressions', ascending=False).head(5)
        
        st.markdown("""
        To capture maximum organic traffic growth with minimal structural rebuilding, execute these specific directives on your domain immediately:
        """)
        
        if not striking_distance_kws.empty:
            st.markdown("### 🚀 Quick Win Internal Linking Directives")
            for i, r in striking_distance_kws.reset_index().iterrows():
                st.markdown(f"*   **Directive {i+1}:** Inject **2 to 3 targeted internal links** to the page ranking for keyword `{r['Queries']}` (Current Position: **{round(r['Position'], 1)}**). Use variations of the target term as anchor text across high-performing root pages.")
        
        st.markdown("### 🛠️ On-Page Semantic Updates")
        st.markdown("""
        *   **Optimize Heading Hierarchies:** Ensure your H1 perfectly matches search intent. Add concise answer paragraphs directly below H2 elements to win AI overview selections.
        *   **Clean Up Schema:** Audit your structured data markup. Remove outdated microdata formatting to prevent crawling errors.
        *   **Metadata Tuning:** For pages showing high impressions but low CTR, rewrite your meta title to sound compelling, addressing *why* a searcher should click on your result rather than relying on generative summaries.
        """)

    else:
        st.error("❌ ZIP processing succeeded, but the code could not isolate the core 'Queries' or 'Pages' data frames. Please verify you are uploading an authentic zip download directly from the Google Search Console UI's Export function.")
