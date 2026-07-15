import pandas as pd

def run_keyword_analysis(df_q, min_impr=100):
    reports = {}
    
    # 1. Gaining vs Losing
    reports['gaining'] = df_q[df_q['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False)
    reports['losing'] = df_q[df_q['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True)
    
    # 2. New vs Lost
    reports['new'] = df_q[(df_q['Clicks'] > 0) & (df_q['Clicks_Delta'] == df_q['Clicks'])]
    reports['lost'] = df_q[(df_q['Clicks'] == 0) & (df_q['Clicks_Delta'] < 0)]
    
    # 3. Impression Scale vs Low CTR
    reports['high_impr_low_ctr'] = df_q[(df_q['Impressions'] >= min_impr) & (df_q['CTR'] < 1.5)].sort_values(by='Impressions', ascending=False)
    
    # 4. Quick Wins Positions 4-10
    reports['quick_wins'] = df_q[(df_q['Position'] >= 4.0) & (df_q['Position'] <= 10.0)].sort_values(by='Impressions', ascending=False)
    
    # 5. Position 11-20 Opportunities
    reports['pos_11_20'] = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 20.0)].sort_values(by='Impressions', ascending=False)
    
    # 6. Improved / Dropped Top 10
    reports['improved_top10'] = df_q[(df_q['Position'] <= 10.0) & ((df_q['Position'] - df_q['Position_Delta']) > 10.0)]
    reports['dropped_top10'] = df_q[(df_q['Position'] > 10.0) & ((df_q['Position'] - df_q['Position_Delta']) <= 10.0)]
    
    return reports
