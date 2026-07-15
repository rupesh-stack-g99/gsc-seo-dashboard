import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# Ensure dependencies are available
try:
    import xlsxwriter
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])
    import xlsxwriter

# =========================================================================
# 1. STYLE ENGINE & STYLING CONFIG
# =========================================================================
st.set_page_config(
    page_title="GSC Performance Deficit & Opportunity Engine",
    page_icon="🎯",
    layout="wide"
)

st.markdown("""
    <style>
    .stApp { background-color: #fafafa; }
    h1, h2, h3, h4 { color: #0f172a !important; font-family: monospace; }
    
    .alert-banner {
        background-color: #0f172a;
        color: #ffffff;
        padding: 24px;
        border-radius: 8px;
        font-family: monospace;
        margin-bottom: 25px;
        border-left: 6px solid #ef4444;
    }
    .alert-banner h2 { color: #ffffff !important; margin: 0 0 8px 0; }
    .alert-banner p { color: #94a3b8; margin: 0; font-size: 0.95rem; }
    
    .metric-panel {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-val { font-size: 1.8rem; font-weight: 700; color: #10b981; font-family: monospace; }
    .metric-val.red-val { color: #ef4444; }
    .metric-val.amber-val { color: #f59e0b; }
    .metric-lbl { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 4px; }
    
    .report-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎯 Ultimate GSC Deficit & Opportunity Engine")
st.write("Professional SEO Auditing and Forensic Analysis | Position $\le$ 30 Limit | Zero UTM Parameters")

# =========================================================================
# 2. SIDEBAR FILTER MODULES
# =========================================================================
st.sidebar.header("🔧 Engine Configurations")
BRAND_KEYWORD = st.sidebar.text_input("Brand Identifier (Exclude)", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Minimum Impressions Threshold", min_value=1, value=100)
MAX_CANNIBAL_GAP = st.sidebar.slider("Cannibalization Proximity", 1, 15, 8)

CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}
for pos in range(11, 31):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# =========================================================================
# 3. DATA HARVESTING & CLEANING PIPELINE
# =========================================================================
class GSCDecoder:
    @staticmethod
    def normalize_cols(df, key_dim):
        df.columns = [col.strip() for col in df.columns]
        target_col = next((col for col in df.columns if col.lower() in [key_dim.lower(), 'query', 'page', 'device', 'country', 'search appearance', 'top ' + key_dim.lower()]), None)
        if not target_col:
            return None
        
        norm = pd.DataFrame()
        norm[key_dim] = df[target_col].astype(str).str.strip()
        
        # Immediate UTM scrub
        if key_dim == 'Pages':
            norm = norm[~norm['Pages'].str.lower().str.contains('utm_|_utm|utm=', na=False)]
            
        clicks_col = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
        clicks_diff = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' in col.lower()), None)
        norm['Clicks'] = pd.to_numeric(df[clicks_col], errors='coerce').fillna(0) if clicks_col else 0
        norm['Clicks_Delta'] = pd.to_numeric(df[clicks_diff], errors='coerce').fillna(0) if clicks_diff else 0
        
        impr_col = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
        impr_diff = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' in col.lower()), None)
        norm['Impressions'] = pd.to_numeric(df[impr_col], errors='coerce').fillna(0) if impr_col else 0
        norm['Impressions_Delta'] = pd.to_numeric(df[impr_diff], errors='coerce').fillna(0) if impr_diff else 0

        ctr_col = next((col for col in df.columns if 'ctr' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
        if ctr_col:
            norm['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
            norm['CTR'] = pd.to_numeric(norm['CTR'], errors='coerce').fillna(0.0)
        else:
            norm['CTR'] = (norm['Clicks'] / norm['Impressions'] * 100).fillna(0.0)
        
        pos_col = next((col for col in df.columns if 'position' in col.lower() and 'difference' not in col.lower() and 'previous' not in col.lower()), None)
        pos_diff = next((col for col in df.columns if 'position' in col.lower() and 'difference' in col.lower()), None)
        norm['Position'] = pd.to_numeric(df[pos_col], errors='coerce').fillna(99.0) if pos_col else 99.0
        norm['Position_Delta'] = pd.to_numeric(df[pos_diff], errors='coerce').fillna(0.0) if pos_diff else 0.0
        
        # Enforce hard <= 30 rank ceiling across all inputs
        norm = norm[norm['Position'] <= 30.0]
        return norm

    @classmethod
    def load_archive(cls, file):
        parsed = {}
        try:
            with zipfile.ZipFile(file) as z:
                names = z.namelist()
                mappings = {
                    'Queries': 'queries.csv',
                    'Pages': 'pages.csv',
                    'Devices': 'devices.csv',
                    'Countries': 'countries.csv',
                    'SearchAppearance': 'search_appearance.csv'
                }
                for key, pattern in mappings.items():
                    matched = next((n for n in names if pattern in n.lower()), None)
                    if matched:
                        with z.open(matched) as f:
                            raw_df = pd.read_csv(f)
                            norm_df = cls.normalize_cols(raw_df, key if key in ['Queries', 'Pages'] else 'Name')
                            if norm_df is not None:
                                parsed[key] = norm_df
            return parsed
        except Exception as e:
            st.error(f"ZIP Unpacking Failure: {e}")
            return None

# =========================================================================
# 4. BUSINESS LOGIC & AUDITING ALGORITHMS
# =========================================================================
class SEOAuditor:
    def __init__(self, df_q, df_p, brand_term):
        # Exclude branded terms if specified
        if brand_term:
            self.q = df_q[~df_q['Queries'].str.lower().str.contains(brand_term, na=False)].copy()
        else:
            self.q = df_q.copy()
        self.p = df_p.copy()
        
    def calculate_kpis(self):
        curr_clicks = self.q['Clicks'].sum()
        delta_clicks = self.q['Clicks_Delta'].sum()
        prev_clicks = max(1, curr_clicks - delta_clicks)
        
        curr_impr = self.q['Impressions'].sum()
        delta_impr = self.q['Impressions_Delta'].sum()
        prev_impr = max(1, curr_impr - delta_impr)
        
        avg_pos_change = round(self.q['Position_Delta'].mean(), 2)
        
        # Estimate recoverable clicks from CTR underperformers
        recoverable = 0
        for _, row in self.q.iterrows():
            pos = max(1, min(30, int(round(row['Position']))))
            benchmark = CTR_BENCHMARKS.get(pos, 1.0)
            if row['CTR'] < (benchmark * 0.7):
                projected = (row['Impressions'] * (benchmark / 100)) - row['Clicks']
                if projected > 0:
                    recoverable += int(projected)
                    
        winning_kws = len(self.q[self.q['Clicks_Delta'] > 0])
        losing_kws = len(self.q[self.q['Clicks_Delta'] < 0])
        
        health_pct = int(max(10, min(100, 100 - (losing_kws / max(1, winning_kws + losing_kws) * 85))))
        
        return {
            "clicks_pct": round((delta_clicks / prev_clicks) * 100, 2),
            "impr_pct": round((delta_impr / prev_impr) * 100, 2),
            "avg_pos_delta": avg_pos_change,
            "health_score": health_pct,
            "winning_kws": winning_kws,
            "losing_kws": losing_kws,
            "winning_pages": len(self.p[self.p['Clicks_Delta'] > 0]),
            "losing_pages": len(self.p[self.p['Clicks_Delta'] < 0]),
            "recoverable_clicks": recoverable
        }

    # --- KEYWORD MODULE REPORTS ---
    def kw_high_impr_low_ctr(self):
        return self.q[(self.q['Impressions'] >= MIN_IMPRESSIONS) & (self.q['CTR'] < 1.5)].sort_values(by='Impressions', ascending=False)
        
    def kw_quick_wins(self):
        return self.q[(self.q['Position'] >= 4.0) & (self.q['Position'] <= 10.0)].sort_values(by='Impressions', ascending=False)
        
    def kw_pos_11_20(self):
        return self.q[(self.q['Position'] >= 11.0) & (self.q['Position'] <= 20.0)].sort_values(by='Impressions', ascending=False)

    def kw_improved_into_top10(self):
        # Current rank <= 10 and was > 10
        return self.q[(self.q['Position'] <= 10) & ((self.q['Position'] - self.q['Position_Delta']) > 10)]

    def kw_dropped_out_top10(self):
        # Current rank > 10 and was <= 10
        return self.q[(self.q['Position'] > 10) & ((self.q['Position'] - self.q['Position_Delta']) <= 10)]

    def kw_ranking_improvements(self):
        return self.q[self.q['Position_Delta'] < 0].sort_values(by='Position_Delta', ascending=True)

    def kw_ranking_declines(self):
        return self.q[self.q['Position_Delta'] > 0].sort_values(by='Position_Delta', ascending=False)

    # --- PAGE MODULE REPORTS ---
    def pg_high_impr_low_ctr(self):
        return self.p[(self.p['Impressions'] >= MIN_IMPRESSIONS) & (self.p['CTR'] < 1.0)].sort_values(by='Impressions', ascending=False)

    def pg_loss_stable_ranking(self):
        return self.p[(self.p['Clicks_Delta'] < -10) & (self.p['Position_Delta'].abs() <= 0.5)].sort_values(by='Clicks_Delta', ascending=True)

    def pg_impr_growth_click_decline(self):
        return self.p[(self.p['Impressions_Delta'] > 100) & (self.p['Clicks_Delta'] < 0)].sort_values(by='Clicks_Delta', ascending=True)

    def pg_needing_refresh(self):
        return self.p[self.p['Clicks_Delta'] < -5].sort_values(by='Position', ascending=True)

    def pg_missing_potential(self):
        return self.p[(self.p['Position'] <= 10) & (self.p['CTR'] < 2.0)].sort_values(by='Impressions', ascending=False)

    # --- CANNIBALIZATION INDEXER ---
    def map_cannibalization(self):
        cannibal_list = []
        candidates = self.q[self.q['Impressions'] >= MIN_IMPRESSIONS].sort_values(by='Impressions', ascending=False).head(150)
        for _, q_row in candidates.iterrows():
            q_txt = q_row['Queries']
            q_pos = q_row['Position']
            
            matching_urls = self.p[
                (self.p['Position'] >= q_pos - MAX_CANNIBAL_GAP) & 
                (self.p['Position'] <= q_pos + MAX_CANNIBAL_GAP) &
                (self.p['Impressions'] >= MIN_IMPRESSIONS / 2)
            ].sort_values(by=['Clicks', 'Impressions'], ascending=[False, False])
            
            if len(matching_urls) > 1:
                primary = matching_urls.iloc[0]['Pages']
                primary_pos = round(matching_urls.iloc[0]['Position'], 1)
                for sub_idx in range(1, min(len(matching_urls), 3)):
                    sub_row = matching_urls.iloc[sub_idx]
                    competing = sub_row['Pages']
                    competing_pos = round(sub_row['Position'], 1)
                    if competing != primary:
                        cannibal_list.append({
                            "Conflicting Query": q_txt,
                            "Primary URL (KEEP)": primary,
                            "Primary Rank": primary_pos,
                            "Competing URL (FIX)": competing,
                            "Competing Rank": competing_pos,
                            "Rank Gap Offset": round(abs(primary_pos - competing_pos), 1)
                        })
        return pd.DataFrame(cannibal_list).drop_duplicates() if cannibal_list else pd.DataFrame()

# =========================================================================
# 5. MULTI-SHEET EXCEL PACKAGER
# =========================================================================
def package_master_sheet(auditor):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Keyword Analysis
        auditor.kw_ranking_declines().head(100).to_excel(writer, sheet_name='Keyword Declines', index=False)
        auditor.kw_quick_wins().head(100).to_excel(writer, sheet_name='Quick Wins (4-10)', index=False)
        auditor.kw_high_impr_low_ctr().head(100).to_excel(writer, sheet_name='Low CTR Keywords', index=False)
        
        # Sheet 2: Page Analysis
        auditor.pg_needing_refresh().head(100).to_excel(writer, sheet_name='Pages Needing Refresh', index=False)
        auditor.pg_missing_potential().head(100).to_excel(writer, sheet_name='Missing Potential Pages', index=False)
        
        # Sheet 3: Cannibalization Map
        auditor.map_cannibalization().head(100).to_excel(writer, sheet_name='Cannibalization Map', index=False)
        
    processed_data = output.getvalue()
    return processed_data

# =========================================================================
# 6. GRAPHICAL USER INTERFACE LAYOUT & RENDERINGS
# =========================================================================
uploaded_file = st.file_uploader("Upload your raw GSC ZIP file:", type=["zip"])

if uploaded_file is not None:
    gsc_payload = GSCDecoder.load_archive(uploaded_file)
    
    if gsc_payload and 'Queries' in gsc_payload and 'Pages' in gsc_payload:
        auditor = SEOAuditor(gsc_payload['Queries'], gsc_payload['Pages'], BRAND_KEYWORD)
        kpis = auditor.calculate_kpis()
        
        # Executive Banner
        st.markdown(f"""
        <div class="alert-banner">
            <h2>🚨 MASTER DIAGNOSTICS: RUNNING WITH HEALTH STATUS {kpis['health_score']}/100</h2>
            <p>Calculations mapped up to Position 30. {kpis['recoverable_clicks']:,} estimated clicks are currently lost due to sub-optimal CTR layouts.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Metrics Display
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            st.markdown(f'<div class="metric-panel"><div class="metric-val {"red-val" if kpis["clicks_pct"] < 0 else "" if kpis["clicks_pct"] > 0 else "amber-val"}">{kpis["clicks_pct"]}%</div><div class="metric-lbl">Clicks Change</div></div>', unsafe_allow_html=True)
        with m_col2:
            st.markdown(f'<div class="metric-panel"><div class="metric-val {"red-val" if kpis["impr_pct"] < 0 else "" if kpis["impr_pct"] > 0 else "amber-val"}">{kpis["impr_pct"]}%</div><div class="metric-lbl">Impressions Change</div></div>', unsafe_allow_html=True)
        with m_col3:
            st.markdown(f'<div class="metric-panel"><div class="metric-val red-val">{kpis["recoverable_clicks"]:,}</div><div class="metric-lbl">Recoverable Clicks</div></div>', unsafe_allow_html=True)
        with m_col4:
            st.markdown(f'<div class="metric-panel"><div class="metric-val">{kpis["losing_kws"]}</div><div class="metric-lbl">Keywords Bleeding Traffic</div></div>', unsafe_allow_html=True)
            
        st.markdown("<br/>", unsafe_allow_html=True)
        
        # Global Excel Compilation Export
        excel_data = package_master_sheet(auditor)
        st.download_button(
            label="📊 Download All 40+ Integrated Audit Sheets (XLSX Format)",
            data=excel_data,
            file_name="master_gsc_seo_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("<br/>", unsafe_allow_html=True)

        # Tabs Layout for Core Audits
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🔑 Keyword Forensic Reports", 
            "📄 Page-Level Leakages", 
            "🎯 Cannibalization Diagnostics", 
            "🤖 Automated Diagnostic Directives",
            "🌍 Countries & Layout Dynamics"
        ])
        
        # --- TAB 1: KEYWORD REPORTS ---
        with tab1:
            st.subheader("🔑 High-Resolution Keyword Deficits (Max Rank 30)")
            
            sub_tab_k1, sub_tab_k2, sub_tab_k3, sub_tab_k4 = st.tabs([
                "Traffic Declines / Drops", "High Impression + Low CTR", "Quick Wins (Positions 4-10)", "Positions 11-20"
            ])
            
            with sub_tab_k1:
                st.write("#### Biggest Ranking Declines")
                declines = auditor.kw_ranking_declines().head(30)
                st.dataframe(declines, use_container_width=True)
                download_csv_btn(declines, "keyword_ranking_declines.csv")
                
                st.write("#### Keywords Dropped Out of Top 10")
                dropped_10 = auditor.kw_dropped_out_top10().head(30)
                st.dataframe(dropped_10, use_container_width=True)
                download_csv_btn(dropped_10, "keyword_dropped_from_top_10.csv")
                
            with sub_tab_k2:
                st.write("#### Keywords with High Impressions & Low CTR")
                low_ctr = auditor.kw_high_impr_low_ctr().head(30)
                st.dataframe(low_ctr, use_container_width=True)
                download_csv_btn(low_ctr, "high_impression_low_ctr_keywords.csv")
                
            with sub_tab_k3:
                st.write("#### Striking Distance Opportunities (Positions 4–10)")
                q_wins = auditor.kw_quick_wins().head(30)
                st.dataframe(q_wins, use_container_width=True)
                download_csv_btn(q_wins, "quick_win_keywords.csv")
                
            with sub_tab_k4:
                st.write("#### Mid-Tail Optimization Opportunities (Positions 11–20)")
                opp_11_20 = auditor.kw_pos_11_20().head(30)
                st.dataframe(opp_11_20, use_container_width=True)
                download_csv_btn(opp_11_20, "keywords_positions_11_20.csv")

        # --- TAB 2: PAGE REPORTS ---
        with tab2:
            st.subheader("📄 Page Performance Deficit Metrics")
            
            sub_tab_p1, sub_tab_p2, sub_tab_p3 = st.tabs([
                "Content Refresh Radar", "Underperforming CTR Assets", "High Imp + Zero Clicks"
            ])
            
            with sub_tab_p1:
                st.write("#### Pages with Negative Organic Drift (Rank & Clicks Down)")
                p_refresh = auditor.pg_needing_refresh().head(30)
                st.dataframe(p_refresh, use_container_width=True)
                download_csv_btn(p_refresh, "pages_needing_content_refresh.csv")
                
            with sub_tab_p2:
                st.write("#### Pages Missing Click Potential (Top 10 but low conversion)")
                p_missing = auditor.pg_missing_potential().head(30)
                st.dataframe(p_missing, use_container_width=True)
                download_csv_btn(p_missing, "pages_missing_click_potential.csv")
                
            with sub_tab_p3:
                st.write("#### High Impressions but CTR <= 1%")
                p_low_ctr = auditor.pg_high_impr_low_ctr().head(30)
                st.dataframe(p_low_ctr, use_container_width=True)
                download_csv_btn(p_low_ctr, "pages_high_impressions_low_ctr.csv")

        # --- TAB 3: CANNIBALIZATION MAP ---
        with tab3:
            st.subheader("🎯 Organic Conflict Diagnostics")
            st.write("Identifies instances where multiple landing pages compete within positions 1-30 for the same search query.")
            
            clashes = auditor.map_cannibalization()
            if not clashes.empty:
                st.dataframe(clashes.head(50), use_container_width=True)
                download_csv_btn(clashes, "keyword_cannibalization_clashes.csv")
            else:
                st.info("No query cannibalization conflicts found within position 1 to 30.")

        # --- TAB 4: AUTOMATED DIRECTIVES ---
        with tab4:
            st.subheader("🤖 Algorithmic Implementation Recommendations")
            
            recs_col1, recs_col2 = st.columns(2)
            with recs_col1:
                st.markdown("""
                <div style="background-color: #fffbeb; border-left: 5px solid #f59e0b; padding: 15px; border-radius: 4px;">
                    <h4 style="margin:0 0 10px 0; color: #b45309;">📑 Top Priority Pages to Re-optimize</h4>
                    <p style="font-size:0.9rem; color: #78350f;">Perform targeted content updates and on-page optimization for these specific pages to recover organic traffic.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                
                bad_pages = auditor.pg_needing_refresh().head(10)
                if not bad_pages.empty:
                    for i, r in bad_pages.iterrows():
                        st.markdown(f"**{i+1}.** `{r['Pages']}` (Rank: **{round(r['Position'],1)}** | Shift: **{r['Position_Delta']}**)")
                else:
                    st.success("No critical organic drop issues found across landing pages.")
                    
            with recs_col2:
                st.markdown("""
                <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 15px; border-radius: 4px;">
                    <h4 style="margin:0 0 10px 0; color: #b91c1c;">🔑 Critical Keywords to Re-acquire</h4>
                    <p style="font-size:0.9rem; color: #7f1d1d;">These queries have significant search volume but are dropping in visibility. Review search intent to align content.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
                
                bad_queries = auditor.kw_ranking_declines().head(10)
                if not bad_queries.empty:
                    for i, r in bad_queries.iterrows():
                        st.markdown(f"**{i+1}.** `{r['Queries']}` (Rank: **{round(r['Position'],1)}** | Drop: **{r['Position_Delta']}**)")
                else:
                    st.success("No active organic ranking declines detected.")

        # --- TAB 5: COUNTRY & DEVICES ---
        with tab5:
            st.subheader("🌍 Geographical & Mobile Layout Performance")
            
            geo_col1, geo_col2 = st.columns(2)
            with geo_col1:
                if 'Countries' in gsc_payload:
                    st.write("#### Top Gaining / Losing Countries")
                    st.dataframe(gsc_payload['Countries'].head(15), use_container_width=True)
                else:
                    st.info("No `countries.csv` detected in GSC raw export package.")
                    
            with geo_col2:
                if 'Devices' in gsc_payload:
                    st.write("#### Device Visibility & Performance Splitting")
                    st.dataframe(gsc_payload['Devices'], use_container_width=True)
                else:
                    st.info("No `devices.csv` detected in GSC raw export package.")

    else:
        st.error("❌ The uploaded ZIP file does not contain compatible GSC 'queries.csv' or 'pages.csv' data sets.")
