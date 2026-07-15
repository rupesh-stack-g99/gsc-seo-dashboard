import streamlit as st
import pandas as pd
from config import CSS_STYLING
from utils import extract_gsc_zip
from dashboard import render_dashboard
from keyword_analysis import run_keyword_analysis
from page_analysis import run_page_analysis
from ctr_analysis import run_ctr_analysis
from ranking_analysis import run_ranking_analysis
from opportunities import map_cannibalization
from ai_recommendations import generate_recommendations
from charts import plot_top_losers, plot_ctr_vs_position
from exports import compile_audit_workbook

# Inject core CSS styling
st.markdown(CSS_STYLING, unsafe_allow_html=True)

# Application Title
st.title("🛡️ Enterprise GSC Audit and Analytics Engine")
st.write("Professional Diagnostic Workspace | Max Rank Ceiling: 30 | Strict UTM Exclusions")

# Setup raw ZIP source dropzone
uploaded_zip = st.file_uploader("Drop GSC Export Zip file to run diagnostics:", type=["zip"])

if uploaded_zip is not None:
    # Run data parsing
    gsc_payload = extract_gsc_zip(uploaded_zip)
    
    if gsc_payload and 'Queries' in gsc_payload and 'Pages' in gsc_payload:
        # Extract operational dataframes
        df_q = gsc_payload['Queries']
        df_p = gsc_payload['Pages']
        
        # Render high level performance indices
        render_dashboard(df_q, df_p)
        
        # Render multi-sheet download
        cannibal_map = map_cannibalization(df_q, df_p)
        xlsx_book = compile_audit_workbook(df_q, df_p, cannibal_map)
        
        st.download_button(
            label="💾 Download Complete Multi-Sheet XLSX Audit Package",
            data=xlsx_book,
            file_name="gsc_enterprise_performance_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Create core application layout tabs
        tab_kws, tab_pgs, tab_ctr, tab_ranks, tab_can, tab_directives = st.tabs([
            "🔑 Keywords", "📄 Pages", "📉 CTR Gaps", "↕️ Rankings", "🎯 Cannibalization", "🤖 Directives"
        ])
        
        with tab_kws:
            st.subheader("Query Audit Modules")
            kw_results = run_keyword_analysis(df_q)
            
            c1, c2 = st.columns(2)
            with c1:
                st.write("#### Top Gaining Keywords")
                st.dataframe(kw_results['gaining'].head(20), use_container_width=True)
                st.write("#### New Keywords Discovered")
                st.dataframe(kw_results['new'].head(20), use_container_width=True)
            with c2:
                st.write("#### Top Losing Keywords")
                st.dataframe(kw_results['losing'].head(20), use_container_width=True)
                st.write("#### Lost Keywords")
                st.dataframe(kw_results['lost'].head(20), use_container_width=True)
                
            fig_losers = plot_top_losers(df_q)
            if fig_losers:
                st.plotly_chart(fig_losers, use_container_width=True)

        with tab_pgs:
            st.subheader("Landing Page Audit Modules")
            pg_results = run_page_analysis(df_p)
            
            p1, p2 = st.columns(2)
            with p1:
                st.write("#### Top Gaining Pages")
                st.dataframe(pg_results['gaining'].head(20), use_container_width=True)
                st.write("#### Pages with Click Loss + Stable Rankings")
                st.dataframe(pg_results['click_loss_stable_rank'].head(20), use_container_width=True)
            with p2:
                st.write("#### Top Losing Pages")
                st.dataframe(pg_results['losing'].head(20), use_container_width=True)
                st.write("#### Pages Needing Content Refresh")
                st.dataframe(pg_results['refresh_needed'].head(20), use_container_width=True)

        with tab_ctr:
            st.subheader("Click Efficiency & CTR Deficits")
            ctr_results = run_ctr_analysis(df_q)
            
            st.write("#### Expected CTR Deficits")
            st.dataframe(ctr_results['ctr_gaps'].head(25), use_container_width=True)
            
            fig_dist = plot_ctr_vs_position(df_q)
            if fig_dist:
                st.plotly_chart(fig_dist, use_container_width=True)

        with tab_ranks:
            st.subheader("SERP Volatility & Positioning")
            rank_results = run_ranking_analysis(df_q)
            
            r1, r2 = st.columns(2)
            with r1:
                st.write("#### Biggest Ranking Drop-offs")
                st.dataframe(rank_results['drops'].head(20), use_container_width=True)
                st.write("#### Strike Zone Keywords (Positions 2-3)")
                st.dataframe(rank_results['near_first'].head(20), use_container_width=True)
            with r2:
                st.write("#### Biggest Ranking Improvements")
                st.dataframe(rank_results['improvements'].head(20), use_container_width=True)
                st.write("#### Edge of Page 1 Opportunities (Positions 11-15)")
                st.dataframe(rank_results['near_page_one'].head(20), use_container_width=True)

        with tab_can:
            st.subheader("Keyword Cannibalization Audit Map")
            if not cannibal_map.empty:
                st.dataframe(cannibal_map.head(40), use_container_width=True)
            else:
                st.info("No query cannibalization mapped between positions 1 and 30.")

        with tab_directives:
            st.subheader("On-Board Algorithmic Audit Instructions")
            ai_directives = generate_recommendations(df_q, df_p)
            
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("""
                <div style="background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
                    <h4 style="margin:0 0 10px 0; color: #b91c1c;">📋 Content Refresh Worksheets</h4>
                    <p style="font-size:0.9rem; color: #7f1d1d; margin:0;">These pages have dropping search traction and need immediate copy depth updates.</p>
                </div>
                """, unsafe_allow_html=True)
                st.dataframe(ai_directives['recover_pages'][['Pages', 'Position', 'Clicks_Delta']], use_container_width=True)
                
            with d2:
                st.markdown("""
                <div style="background-color: #fffbeb; border-left: 5px solid #f59e0b; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
                    <h4 style="margin:0 0 10px 0; color: #b45309;">📑 Internal Link Acquisition Targets</h4>
                    <p style="font-size:0.9rem; color: #78350f; margin:0;">These terms are on page 2. Send external internal links directly here to boost them to Page 1.</p>
                </div>
                """, unsafe_allow_html=True)
                st.dataframe(ai_directives['internal_links'][['Queries', 'Position', 'Impressions']], use_container_width=True)

    else:
        st.error("❌ Invalid GSC File Format. Please ensure queries.csv and pages.csv are included.")
