import streamlit as st
import pandas as pd
from config import CTR_BENCHMARKS

def get_dashboard_kpis(df_q, df_p):
    curr_clicks = df_q['Clicks'].sum()
    delta_clicks = df_q['Clicks_Delta'].sum()
    prev_clicks = max(1, curr_clicks - delta_clicks)
    clicks_pct = (delta_clicks / prev_clicks) * 100
    
    curr_impr = df_q['Impressions'].sum()
    delta_impr = df_q['Impressions_Delta'].sum()
    prev_impr = max(1, curr_impr - delta_impr)
    impr_pct = (delta_impr / prev_impr) * 100
    
    avg_ctr_curr = df_q['CTR'].mean()
    avg_ctr_prev = (df_q['Clicks'] - df_q['Clicks_Delta']).sum() / max(1, (df_q['Impressions'] - df_q['Impressions_Delta']).sum()) * 100
    ctr_change = avg_ctr_curr - avg_ctr_prev
    
    avg_pos_change = df_q['Position_Delta'].mean()
    
    winning_kws = len(df_q[df_q['Clicks_Delta'] > 0])
    losing_kws = len(df_q[df_q['Clicks_Delta'] < 0])
    winning_pages = len(df_p[df_p['Clicks_Delta'] > 0])
    losing_pages = len(df_p[df_p['Clicks_Delta'] < 0])
    
    # Est. Recoverable clicks (CTR < Benchmark targets)
    rec_clicks = 0
    for _, row in df_q.iterrows():
        pos = max(1, min(30, int(round(row['Position']))))
        benchmark = CTR_BENCHMARKS.get(pos, 1.0)
        if row['CTR'] < (benchmark * 0.7):
            pot = (row['Impressions'] * (benchmark / 100)) - row['Clicks']
            if pot > 0:
                rec_clicks += int(pot)
                
    health_score = int(max(10, min(100, 100 - (losing_kws / max(1, winning_kws + losing_kws) * 85))))
    
    return {
        "clicks_pct": round(clicks_pct, 2),
        "impr_pct": round(impr_pct, 2),
        "ctr_change": round(ctr_change, 2),
        "avg_pos_change": round(avg_pos_change, 2),
        "winning_kws": winning_kws,
        "losing_kws": losing_kws,
        "winning_pages": winning_pages,
        "losing_pages": losing_pages,
        "health_score": health_score,
        "recoverable_clicks": rec_clicks
    }

def render_dashboard(df_q, df_p):
    kpi = get_dashboard_kpis(df_q, df_p)
    
    st.markdown(f"""
    <div class="alert-banner">
        <h2>📊 EXECUTIVE SEO SUMMARY DASHBOARD</h2>
        <p>Current organic site stability evaluation: <b>{kpi['health_score']}/100</b> Core Health Score.</p>
    </div>
    """, unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-value {"positive" if kpi["clicks_pct"]>=0 else "negative"}">{kpi["clicks_pct"]}%</div><div class="metric-label">Click Growth</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-value">{kpi["winning_kws"]:,}</div><div class="metric-label">Winning Keywords</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-value {"positive" if kpi["impr_pct"]>=0 else "negative"}">{kpi["impr_pct"]}%</div><div class="metric-label">Impression Growth</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-value negative">{kpi["losing_kws"]:,}</div><div class="metric-label">Losing Keywords</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-value warning">{kpi["avg_pos_change"]}</div><div class="metric-label">Avg Position Shift</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-value">{kpi["winning_pages"]:,}</div><div class="metric-label">Winning Pages</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><div class="metric-value negative">{kpi["recoverable_clicks"]:,}</div><div class="metric-label">Recoverable Clicks</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-value negative">{kpi["losing_pages"]:,}</div><div class="metric-label">Losing Pages</div></div>', unsafe_allow_html=True)
