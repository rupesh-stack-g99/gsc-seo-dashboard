import pandas as pd
import numpy as np
import zipfile
import re

def clean_and_normalize(df, dim_name):
    df.columns = [col.strip() for col in df.columns]
    
    # Locate index column
    target_col = next((col for col in df.columns if col.lower() in [
        dim_name.lower(), 'query', 'page', 'device', 'country', 'search appearance', f'top {dim_name.lower()}'
    ]), None)
    
    if not target_col:
        return None
        
    norm = pd.DataFrame()
    norm[dim_name] = df[target_col].astype(str).str.strip()
    
    # UTM scrub
    if dim_name == 'Pages':
        norm = norm[~norm['Pages'].str.lower().str.contains(r'utm_|_utm|utm=', na=False, regex=True)]
        
    # Extract numerical matrices safely
    def find_and_parse(keywords, default=0.0):
        col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
        diff_col = next((c for c in df.columns if any(k in c.lower() for k in keywords) and 'difference' in c.lower()), None)
        
        val = pd.to_numeric(df[col], errors='coerce').fillna(default) if col else pd.Series(default, index=df.index)
        val_delta = pd.to_numeric(df[diff_col], errors='coerce').fillna(0.0) if diff_col else pd.Series(0.0, index=df.index)
        return val, val_delta

    norm['Clicks'], norm['Clicks_Delta'] = find_and_parse(['click'])
    norm['Impressions'], norm['Impressions_Delta'] = find_and_parse(['impression'])
    
    # Standardize CTR parsing
    ctr_col = next((c for c in df.columns if 'ctr' in c.lower() and 'difference' not in c.lower() and 'previous' not in c.lower()), None)
    if ctr_col:
        norm['CTR'] = df[ctr_col].astype(str).str.replace('%', '', regex=False)
        norm['CTR'] = pd.to_numeric(norm['CTR'], errors='coerce').fillna(0.0)
    else:
        norm['CTR'] = (norm['Clicks'] / norm['Impressions'] * 100).fillna(0.0)
        
    norm['Position'], norm['Position_Delta'] = find_and_parse(['position'], default=99.0)
    
    # Apply global <= 30 rank boundary
    norm = norm[norm['Position'] <= 30.0]
    return norm

def extract_gsc_zip(uploaded_file):
    dfs = {}
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            names = z.namelist()
            targets = {
                'Queries': 'queries.csv',
                'Pages': 'pages.csv',
                'Devices': 'devices.csv',
                'Countries': 'countries.csv',
                'SearchAppearance': 'search_appearance.csv'
            }
            for key, file_pattern in targets.items():
                matched = next((n for n in names if file_pattern in n.lower()), None)
                if matched:
                    with z.open(matched) as f:
                        df = pd.read_csv(f)
                        norm_df = clean_and_normalize(df, key if key in ['Queries', 'Pages'] else 'Name')
                        if norm_df is not None:
                            dfs[key] = norm_df
        return dfs
    except Exception:
        return None
