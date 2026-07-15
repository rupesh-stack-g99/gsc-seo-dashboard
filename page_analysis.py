import pandas as pd

def run_page_analysis(df_p, min_impr=100):
    reports = {}
    
    # 1. Gainers vs Losers
    reports['gaining'] = df_p[df_p['Clicks_Delta'] > 0].sort_values(by='Clicks_Delta', ascending=False)
    reports['losing'] = df_p[df_p['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True)
    
    # 2. High Impression + Low CTR Pages
    reports['high_impr_low_ctr'] = df_p[(df_p['Impressions'] >= min_impr) & (df_p['CTR'] < 1.0)].sort_values(by='Impressions', ascending=False)
    
    # 3. Pages with Click Loss but Stable Rankings
    reports['click_loss_stable_rank'] = df_p[(df_p['Clicks_Delta'] < -5) & (df_p['Position_Delta'].abs() <= 0.5)].sort_values(by='Clicks_Delta', ascending=True)
    
    # 4. Impression Growth but Click Decline (Mismatched Intent)
    reports['impr_up_click_down'] = df_p[(df_p['Impressions_Delta'] > 50) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Clicks_Delta', ascending=True)
    
    # 5. Content Refresh Candidates
    reports['refresh_needed'] = df_p[(df_p['Position_Delta'] > 1.0) & (df_p['Clicks_Delta'] < 0)].sort_values(by='Impressions', ascending=False)
    
    # 6. Missing Click Potential (High Rank, Bad CTR)
    reports['missing_potential'] = df_p[(df_p['Position'] <= 10.0) & (df_p['CTR'] < 2.0)].sort_values(by='Impressions', ascending=False)
    
    return reports
