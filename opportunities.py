import pandas as pd

def map_cannibalization(df_q, df_p, min_impr=100, max_gap=8):
    cannibal_list = []
    candidates = df_q[df_q['Impressions'] >= min_impr].sort_values(by='Impressions', ascending=False).head(150)
    
    for _, q_row in candidates.iterrows():
        query_txt = q_row['Queries']
        q_pos = q_row['Position']
        
        # Pull competing URLs fighting in the nearby position vicinity
        matching_urls = df_p[
            (df_p['Position'] <= 30.0) &
            (df_p['Position'] >= q_pos - max_gap) & 
            (df_p['Position'] <= q_pos + max_gap) &
            (df_p['Impressions'] >= min_impr / 2)
        ].sort_values(by=['Clicks', 'Impressions'], ascending=[False, False])
        
        if len(matching_urls) > 1:
            primary_url = matching_urls.iloc[0]['Pages']
            primary_pos = round(matching_urls.iloc[0]['Position'], 1)
            primary_clicks = int(matching_urls.iloc[0]['Clicks'])
            
            for sub_idx in range(1, min(len(matching_urls), 3)):
                sub_row = matching_urls.iloc[sub_idx]
                cannibal_url = sub_row['Pages']
                cannibal_pos = round(sub_row['Position'], 1)
                cannibal_clicks = int(sub_row['Clicks'])
                
                if cannibal_url != primary_url:
                    cannibal_list.append({
                        "Conflicting Query": query_txt,
                        "Primary URL (KEEP)": primary_url,
                        "Primary Rank": primary_pos,
                        "Primary Clicks": primary_clicks,
                        "Competing URL (FIX)": cannibal_url,
                        "Competing Rank": cannibal_pos,
                        "Competing Clicks": cannibal_clicks,
                        "Rank Gap Offset": round(abs(primary_pos - cannibal_pos), 1)
                    })
                    
    return pd.DataFrame(cannibal_list).drop_duplicates() if cannibal_list else pd.DataFrame()
