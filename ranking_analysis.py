import pandas as pd

def run_ranking_analysis(df_q):
    reports = {}
    
    # 1. Rank improvements / Drops
    reports['improvements'] = df_q[df_q['Position_Delta'] < 0].sort_values(by='Position_Delta', ascending=True)
    reports['drops'] = df_q[df_q['Position_Delta'] > 0].sort_values(by='Position_Delta', ascending=False)
    
    # 2. Easy Wins Ranks 4-8
    reports['easy_wins'] = df_q[(df_q['Position'] >= 4.0) & (df_q['Position'] <= 8.0)].sort_values(by='Impressions', ascending=False)
    
    # 3. Positions 2-3 (Strike Zone)
    reports['near_first'] = df_q[(df_q['Position'] >= 2.0) & (df_q['Position'] <= 3.9)].sort_values(by='Impressions', ascending=False)
    
    # 4. Positions 11-15 (Edge of Page 1)
    reports['near_page_one'] = df_q[(df_q['Position'] >= 11.0) & (df_q['Position'] <= 15.0)].sort_values(by='Impressions', ascending=False)
    
    return reports
