import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# PAGE CONFIGURATION & LAYOUT
# =========================================================================
st.set_page_config(
    page_title="SEO Cannibalization & De-Optimization Dashboard",
    page_icon="🎯",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .metric-box {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 8px;
        border-left: 5px solid #ff4b4b;
        margin-bottom: 15px;
    }
    .focus-card {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        margin-bottom: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎯 GSC ZIP & CSV SEO Diagnostic Engine")
st.write("Upload a combined Query-Page map CSV or GSC export to identify keyword conflicts.")

# =========================================================================
# CONFIGURATION SIDEBAR
# =========================================================================
st.sidebar.header("🛠️ Processing Rules & Thresholds")

BRAND_KEYWORD = st.sidebar.text_input("Brand Keyword to Exclude", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Min Query Impressions (Last 3M)", min_value=1, value=100)
MAX_POSITION_GAP = st.sidebar.number_input("Max Position Difference (Proximity Limit)", min_value=1, max_value=20, value=10)
MAX_POSITION_LIMIT = st.sidebar.number_input("Max Allowed Position (Filter Boundary)", min_value=10, max_value=100, value=50)

st.sidebar.markdown("---")
st.sidebar.info("""
**How the Proximity Rule Works:**
Only conflicts where the **Primary Page** and the **Cannibal Page** rank within **10 positions** of each other will be surfaced.
""")

# =========================================================================
# DETECTOR & PARSING ENGINE
# =========================================================================
def standardize_df_columns(df):
    """
    Standardizes variable naming across standard and comparison schemas.
    """
    df.columns = [col.strip() for col in df.columns]
    
    # Fuzzy match headers
    page_col = next((col for col in df.columns if col.lower() in ['page', 'landing page', 'urls', 'pages', 'url', 'landing_page']), None)
    query_col = next((col for col in df.columns if col.lower() in ['query', 'keyword', 'search term', 'queries', 'search_query']), None)
    
    if not page_col or not query_col:
        return None, False

    is_comparison = any('difference' in col.lower() or 'last' in col.lower() or 'compare' in col.lower() for col in df.columns)
    
    standardized_df = pd.DataFrame()
    standardized_df['Query'] = df[query_col].astype(str).str.strip()
    standardized_df['Page'] = df[page_col].astype(str).apply(lambda x: x.split('#')[0].strip())
    
    if is_comparison:
        clicks_col = next((col for col in df.columns if 'clicks' in col.lower() and 'last' in col.lower()), None)
        impr_col = next((col for col in df.columns if 'impressions' in col.lower() and 'last' in col.lower()), None)
        ctr_col = next((col for col in df.columns if 'ctr' in col.lower() and 'last' in col.lower()), None)
        pos_col = next((col for col in df.columns if 'position' in col.lower() and 'last' in col.lower()), None)
        
        standardized_df['Clicks'] = pd.to_numeric(df[clicks_col if clicks_col else 'Clicks'], errors='coerce').fillna(0)
        standardized_df['Impressions'] = pd.to_numeric(df[impr_col if impr_col else 'Impressions'], errors='coerce').fillna(0)
        
        ctr_series = df[ctr_col] if ctr_col else df['CTR']
        if ctr_series.dtype == object:
            ctr_series = ctr_series.str.replace('%', '', regex=False)
        standardized_df['CTR'] = pd.to_numeric(ctr_series, errors='coerce').fillna(0) / 100.0
        
        standardized_df['Position'] = pd.to_numeric(df[pos_col if pos_col else 'Position'], errors='coerce').fillna(99.0)
    else:
        clicks_col = next((col for col in df.columns if 'clicks' in col.lower()), 'Clicks')
        impr_col = next((col for col in df.columns if 'impressions' in col.lower()), 'Impressions')
        ctr_col = next((col for col in df.columns if 'ctr' in col.lower()), 'CTR')
        pos_col = next((col for col in df.columns if 'position' in col.lower()), 'Position')
        
        standardized_df['Clicks'] = pd.to_numeric(df[clicks_col], errors='coerce').fillna(0)
        standardized_df['Impressions'] = pd.to_numeric(df[impr_col], errors='coerce').fillna(0)
        
        ctr_series = df[ctr_col]
        if ctr_series.dtype == object:
            ctr_series = ctr_series.str.replace('%', '', regex=False)
        standardized_df['CTR'] = pd.to_numeric(ctr_series, errors='coerce').fillna(0)
        if standardized_df['CTR'].max() > 1.0:
            standardized_df['CTR'] = standardized_df['CTR'] / 100.0
            
        standardized_df['Position'] = pd.to_numeric(df[pos_col], errors='coerce').fillna(99.0)
        
    return standardized_df, is_comparison


def extract_gsc_data_from_file(uploaded_file):
    """
    Parses standalone CSVs or zip files.
    """
    filename = uploaded_file.name.lower()
    
    # --- Case 1: ZIP File ---
    if filename.endswith(".zip"):
        try:
            with zipfile.ZipFile(uploaded_file) as z:
                file_list = z.namelist()
                
                # Check for combined files
                search_results_file = next((f for f in file_list if "search_results.csv" in f.lower() or "search results.csv" in f.lower()), None)
                
                if search_results_file:
                    with z.open(search_results_file) as f:
                        df = pd.read_csv(f)
                    return standardize_df_columns(df)
                
                # If they only uploaded queries and pages separately
                queries_file = next((f for f in file_list if "queries.csv" in f.lower()), None)
                pages_file = next((f for f in file_list if "pages.csv" in f.lower()), None)
                
                if queries_file and pages_file:
                    st.error("""
                    ### ⚠️ Incompatible ZIP File Structure
                    You uploaded a native GSC ZIP file. 
                    * Google's default `Queries.csv` contains only queries.
                    * Google's default `Pages.csv` contains only pages.
                    
                    They cannot be analyzed because they aren't mapped together.
                    
                    **How to resolve:** Read the step-by-step instructions below the upload box to obtain a combined Query-Page map.
                    """)
                    return None
                
                st.error("❌ Invalid ZIP Archive: No compatible GSC CSVs found inside.")
                return None
                
        except Exception as e:
            st.error(f"Failed to unpack ZIP file: {e}")
            return None
            
    # --- Case 2: Standalone CSV ---
    elif filename.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
            result = standardize_df_columns(df)
            if result is None or result[0] is None:
                st.error("""
                ### ❌ Missing Mapped Columns
                The CSV you uploaded doesn't contain both 'Query' and 'Page' columns.
                
                **Your file's actual columns:** `{}`
                
                To solve this, please see the guide below on how to export a combined Query-Page CSV.
                """.format(list(df.columns)))
                return None
            return result
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
            return None
        
    return None

# =========================================================================
# FILE UPLOAD CONSOLE
# =========================================================================
st.subheader("📂 Import Search Console Datasets")
uploaded_file = st.file_uploader("Upload GSC Export (CSV or ZIP)", type=["csv", "zip"])

# Helper instructions if nothing is uploaded or error occurs
if uploaded_file is None:
    st.info("💡 **Tips for exporting your GSC Data:**")
    st.markdown("""
    *   **The Problem:** Standard GSC exports separate your Queries and Pages into isolated lists.
    *   **The Easy Fix:** Install the free **Search Analytics for Sheets** Google Sheets extension, fetch your GSC data grouped by both **Query** and **Page** at the same time, download as a CSV, and drop it here!
    """)

# =========================================================================
# PIPELINE EXECUTION
# =========================================================================
if uploaded_file is not None:
    extracted_data = extract_gsc_data_from_file(uploaded_file)
    
    if extracted_data is not None:
        clean_df, is_compare = extracted_data
        
        if clean_df is not None:
            st.success(f"⚡ File parsed! Loaded {'GSC Period Comparison' if is_compare else 'Standard GSC Performance'} dataset.")
            
            # --- 1. FILTER BRAND & BOUNDARIES ---
            if BRAND_KEYWORD:
                clean_df = clean_df[~clean_df['Query'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
            clean_df = clean_df[clean_df['Position'] <= MAX_POSITION_LIMIT]
            
            # --- 2. AGGREGATE DUPLICATES (URL roll-up) ---
            rolled_df = clean_df.groupby(['Query', 'Page']).agg({
                'Clicks': 'sum',
                'Impressions': 'sum',
                'CTR': 'mean',
                'Position': 'mean'
            }).reset_index()
            
            # --- 3. EVALUATE CONFLICTS ---
            unique_queries = rolled_df['Query'].unique()
            deopt_results = []
            
            for query in unique_queries:
                pages_data = rolled_df[rolled_df['Query'] == query].copy()
                
                if len(pages_data) > 1:
                    pages_data = pages_data.sort_values(by=['Clicks', 'Impressions'], ascending=[False, False])
                    
                    total_impressions = pages_data['Impressions'].sum()
                    if total_impressions < MIN_IMPRESSIONS:
                        continue
                        
                    primary_row = pages_data.iloc[0]
                    primary_url = primary_row['Page']
                    primary_clicks = int(primary_row['Clicks'])
                    primary_pos = round(primary_row['Position'], 1)
                    
                    for index in range(1, len(pages_data)):
                        cannibal_row = pages_data.iloc[index]
                        cannibal_url = cannibal_row['Page']
                        cannibal_clicks = int(cannibal_row['Clicks'])
                        cannibal_pos = round(cannibal_row['Position'], 1)
                        
                        pos_diff = abs(primary_pos - cannibal_pos)
                        
                        if pos_diff < MAX_POSITION_GAP:
                            deopt_results.append({
                                "Keyword/Query": query,
                                "Primary URL": primary_url,
                                "Primary Clicks": primary_clicks,
                                "Primary Position": primary_pos,
                                "Cannibal URL to De-Optimize": cannibal_url,
                                "Cannibal Clicks": cannibal_clicks,
                                "Cannibal Position": cannibal_pos,
                                "Position Gap": round(pos_diff, 1)
                            })
            
            final_deopt_df = pd.DataFrame(deopt_results)
            
            # =========================================================================
            # REPORTING INTERFACE
            # =========================================================================
            if not final_deopt_df.empty:
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.metric("Total Conflicts Identified", len(final_deopt_df))
                with col_m2:
                    st.metric("Unique Keywords Affected", final_deopt_df['Keyword/Query'].nunique())
                with col_m3:
                    avg_gap = round(final_deopt_df['Position Gap'].mean(), 1)
                    st.metric("Average Rank Gap", f"{avg_gap} Positions")
                
                csv_buffer = io.StringIO()
                final_deopt_df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()
                
                st.download_button(
                    label="💾 Download '[SEO] To De-Optimize' Clean Report (CSV)",
                    data=csv_data,
                    file_name="seo_to_de_optimize_report.csv",
                    mime="text/csv"
                )
                
                st.subheader("📋 Directives: [SEO] To De-Optimize")
                st.dataframe(final_deopt_df, use_container_width=True)
                
                # =========================================================================
                # ACTIONABLE RECOMMENDATIONS
                # =========================================================================
                st.markdown("---")
                st.subheader("🔥 Strategic SEO Action Plan")
                
                focus_df = final_deopt_df.groupby('Primary URL').agg({
                    'Keyword/Query': 'count',
                    'Primary Clicks': 'sum',
                    'Cannibal URL to De-Optimize': 'nunique'
                }).rename(columns={
                    'Keyword/Query': 'Total Cannibalized Keywords',
                    'Primary Clicks': 'Estimated Core Clicks',
                    'Cannibal URL to De-Optimize': 'Aggressor Pages'
                }).sort_values(by='Total Cannibalized Keywords', ascending=False).reset_index().head(5)
                
                col_l, col_r = st.columns([1, 1])
                
                with col_l:
                    st.markdown("### 🎯 Top 5 Pages to Focus Optimization")
                    st.write("These core primary landing pages are facing heavy rank friction from nearby pages. Clearing up their internal signals will unleash major traffic recoveries.")
                    
                    for idx, row in focus_df.iterrows():
                        st.markdown(f"""
                        <div class="focus-card">
                            <strong style="color:#ff4b4b;">Priority #{idx+1}: {row['Primary URL']}</strong><br/>
                            • Competing Queries: <b>{row['Total Cannibalized Keywords']}</b><br/>
                            • Aggressor Landing Pages competing: <b>{row['Aggressor Pages']}</b><br/>
                            • Current Organic Click Pool: <b>{row['Estimated Core Clicks']}</b>
                        </div>
                        """, unsafe_allow_html=True)
                
                with col_r:
                    st.markdown("### 🛡️ Critical De-Optimization Playbook")
                    st.markdown("""
                    For the **Cannibal URLs to De-Optimize** identified on your screen:
                    
                    *   **Title Tags & Header Checks:** Review the Cannibal URL. If it targets the exact core phrase of the primary page in its Title or H1, rephrase it to focus on a long-tail variant or helper search intent.
                    *   **Internal Link Cleanups:** Locate internal text anchor links pointing to the Cannibal URL with your primary keyword. Edit those hyperlinks to point directly to the **Primary URL** instead.
                    *   **Query-intent Mapping:** If the secondary page is informational and the primary is transactional, convert the informational page's direct keyword mentions into helpful text blocks that naturally link back up to your primary commercial page.
                    """)
                    
            else:
                st.warning("✅ Clean SEO Horizon! No cannibalization targets found where competing pages rank within 10 positions of each other.")
