import pandas as pd

def generate_recommendations(df_q, df_p):
    recommendations = {}
    
    # 1. Recovery Priorities (Pages with high click drops)
    recommendations['recover_pages'] = df_p[df_p['Clicks_Delta'] < -5].sort_values(by='Clicks_Delta', ascending=True).head(10)
    
    # 2. Key Term Priority Recovery
    recommendations['recover_keywords'] = df_q[df_q['Clicks_Delta'] < -5].sort_values(by='Clicks_Delta', ascending=True).head(10)
    
    # 3. Layout Title Optimization candidates (High Rank + Low CTR)
    recommendations['title_opts'] = df_q[(df_q['Position'] <= 10) & (df_q['CTR'] < 1.5)].sort_values(by='Impressions', ascending=False).head(10)
    
    # 4. Internal Linking Target Pages (High Impression + Low Clicks + Position 11-20)
    recommendations['internal_links'] = df_q[(df_q['Position'] >= 11) & (df_q['Position'] <= 20)].sort_values(by='Impressions', ascending=False).head(10)
    
    return recommendations
