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
        color: #e2e8f0;
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

st.markdown("""
<div class="hero-banner">
    <h1>🕵️ Algorithmic SEO Detective (High-Precision Edition)</h1>
    <p>This upgraded release uses hard semantic scoring filters to map keywords to exact core landing page slugs, identifying mismatched intents with a secondary automated validation flag.</p>
</div>
""", unsafe_allow_html=True)

# =========================================================================
# CONFIGURATION INPUTS
# =========================================================================
st.markdown("### ⚙️ Forensic Tuning & Parameters")

cfg_col1, cfg_col2, cfg_col3 = st.columns(3)
with cfg_col1:
    BRAND_TERM = st.text_input("Exclude Branded Searches:", value="").lower().strip()
with cfg_col2:
    MIN_IMPR_THRESHOLD = st.number_input("Minimum Impressions Threshold:", min_value=1, value=100)
with cfg_col3:
    MAX_CANNIBAL_OFFSET = st.slider("Cannibalization Search Space (Pos. Gap):", 1, 15, 6)

st.markdown("---")

CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 101):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# DATA CLEANING ENGINE
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
    
    # Filter candidates by position first
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
        
        # Rule 1: Clean token overlap scoring
        token_matches = sum(1 for token in core_nouns if token in url_path)
        score += (token_matches * 5)
        
        # Rule 2: Strict anchor keyword bonus (e.g., if query has kybella, url MUST have kybella)
        for critical_word in ['kybella', 'earlobe', 'piercing', 'mounjaro', 'tirzepatide', 'botox']:
            if critical_word in query_str:
                if critical_word in url_path:
                    score += 20  # Heavy structural weight bonus
                else:
                    score -= 15  # Penalty for mismatching core service
                    
        # Rule 3: Geo-location alignment check
        for geo in ['weehawken', 'hoboken', 'jersey']:
            if geo in query_str and geo in url_path:
                score += 5
                
        # Rule 4: Small position gap penalty
        score -= (abs(p_row['Position'] - q_pos) * 0.2)
        
        if score > best_score:
            best_score = score
            best_url = p_row['Pages']
            
    # Confidence analysis validation flag
    is_highly_confident = True
    if best_url:
        # If none of the actual service nouns are inside the URL, mark as unverified anomaly
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
# FORENSIC PIPELINE EXECUTION
# =========================================================================
uploaded_file = st.file_uploader("Upload GSC ZIP file to begin automated diagnostics:", type=["zip"])

if uploaded_file is not None:
    gsc = extract_gsc_payload(uploaded_file)
    
    if gsc and 'Queries' in gsc and 'Pages' in gsc:
        df_q_raw = gsc['Queries'].copy()
        df_p = gsc['Pages'].copy()
        
        df_q = df_q_raw[~df_q_raw['Queries'].str.lower().str.contains(BRAND_TERM, na=False)].copy() if BRAND_TERM else df_q_raw.copy()
        
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
            if core_hit_score > 65.0:
                st.markdown(f"""<div class="directive-card danger"><div class="directive-title">🚨 Systemic Algorithmic Suppression Flagged ({core_hit_score}% Probability)</div></div>""", unsafe_allow_html=True)
            else:
                st.markdown("""<div class="directive-card success"><div class="directive-title">✅ No Site-Wide Algorithmic Penalty Detected</div></div>""", unsafe_allow_html=True)

        # === TAB 2: KEYWORD DECAY ALERTS ===
        with tab2:
            st.markdown("## Real-time Keyword Decay Alerts")
            decay_queries = df_q[(df_q['Clicks_Delta'] < 0) & (df_q['Impressions_Delta'] >= 0) & (df_q['Position_Delta'] <= 0.2)].sort_values(by='Clicks_Delta', ascending=True).head(15)

            if not decay_queries.empty:
                for idx, r in decay_queries.reset_index().iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
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
                    badge = '<span class="verified-tag">✓ Confident Match</span>' if item['Confident'] else '<span class="warning-tag">⚠️ Verification Recommended via GSC</span>'
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
            striking_kws = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 15.0)].sort_values(by='Impressions', ascending=False).head(20)

            if not striking_kws.empty:
                for idx, r in striking_kws.reset_index().iterrows():
                    mapped_url, confident = find_best_url_match_precise(r, df_p)
                    badge = '<span class="verified-tag">✓ Confident Match</span>' if confident else '<span class="warning-tag">⚠️ Verify Target Asset</span>'
                    st.markdown(f"""
                    *   🚀 **Keyword:** `{r['Queries']}`  
                        *   **Current Rank:** {round(r['Position'], 1)} | **Impressions:** {int(r['Impressions'])}  
                        *   🔗 **Target Landing Page URL:** `{mapped_url}` {badge}
                    """, unsafe_allow_html=True)

        # === TAB 6: EXECUTION BLUEPRINT ===
        with tab6:
            st.markdown("## Priority Implementation Blueprint")
            st.markdown("Proceed with standard internal linking optimization protocols on high-confidence matched elements.")
    else:
        st.error("❌ Data formatting processing configuration mismatch.")
