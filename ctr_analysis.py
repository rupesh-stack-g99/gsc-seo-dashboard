import pandas as pd
from config import CTR_BENCHMARKS

def run_ctr_analysis(df_q, min_impr=100):
    reports = {}
    
    # 1. CTR Winners and Losers
    reports['winners'] = df_q[df_q['Clicks_Delta'] > 0].sort_values(by='CTR', ascending=False)
    reports['losers'] = df_q[df_q['Clicks_Delta'] < 0].sort_values(by='CTR', ascending=True)
    
    # 2. Calculation of Estimated Click Gaps
    ctr_gaps = []
    candidates = df_q[(df_q['Position'] <= 30.0) & (df_q['Impressions'] >= min_impr)]
    for _, row in candidates.iterrows():
        kw = row['Queries']
        clicks = row['Clicks']
        impr = row['Impressions']
        actual_ctr = row['CTR']
        pos = max(1, min(30, int(round(row['Position']))))
        
        benchmark = CTR_BENCHMARKS.get(pos, 1.0)
        if actual_ctr < (benchmark * 0.7):
            lost_clicks = int((impr * (benchmark / 100)) - clicks)
            if lost_clicks > 0:
                ctr_gaps.append({
                    "Query": kw,
                    "Rank": round(row['Position'], 1),
                    "Actual CTR": f"{round(actual_ctr, 1)}%",
                    "Expected CTR": f"{round(benchmark, 1)}%",
                    "Estimated Click Deficit": lost_clicks,
                    "Total Impressions": int(impr)
                })
                
    reports['ctr_gaps'] = pd.DataFrame(ctr_gaps).sort_values(by='Estimated Click Deficit', ascending=False) if ctr_gaps else pd.DataFrame()
    return reports
