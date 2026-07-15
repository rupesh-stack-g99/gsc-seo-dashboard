import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# =========================================================================
# PAGE CONFIGURATION & LAYOUT
# =========================================================================
st.set_page_config(
    page_title="SEO Compare & Cannibalization Engine",
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

st.title("🎯 GSC Auto-Mapping Compare & Cannibalization Engine")
st.write("Upload a raw **GSC Export ZIP** or a custom **Queries + Pages CSV** to run the diagnostics.")

# =========================================================================
# Side Bar Controls
# =========================================================================
st.sidebar.header("🛠️ Diagnostic Parameters")

BRAND_KEYWORD = st.sidebar.text_input("Brand Keyword to Exclude", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Min Query Impressions (Last 3M)", min_value=1, value=100)
MAX_POSITION_GAP = st.sidebar.number_input("Max Position Difference (Proximity Limit)", min_value=1, max_value=20, value=10)
MAX_POSITION_LIMIT = st.sidebar.number_input("Max Allowed Position (Filter Boundary)", min_value=10, max_value=100, value=50)

# =========================================================================
# STANDARDIZATION HELPER
# =========================================================================
def standardize_df_columns(df):
    """
    Normalizes GSC metric columns across standard and comparison schemas.
    """
    df.columns = [col.strip() for col in df.columns]
    
    page_col = next((col for col in df.columns if col.lower() in ['page', 'landing page', 'urls', 'pages', 'url', 'landing_page', 'top pages']), None)
    query_col = next((col for col in df.columns if col.lower() in ['query', 'keyword', 'search term', 'queries', 'search_query', 'top queries']), None)
    
    # Identify if comparison is present
    is_comparison = any('difference' in col.lower() or 'last' in col.lower() or 'compare' in col.lower() for col in df.columns)
    
    standardized_df = pd.DataFrame()
    
    if query_col:
        standardized_df['Query'] = df[query_col].astype(str).str.strip()
    if page_col:
        standardized_df['Page'] = df[page_col].astype(str).apply(lambda x: x.split('#')[0].strip())
    
    if is_comparison:
        clicks_curr = next((col for col in df.columns if 'clicks' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        clicks_diff = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' in col.lower()), None)
        impr_curr = next((col for col in df.columns if 'impressions' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        impr_diff = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' in col.lower()), None)
        pos_curr = next((col for col in df.columns if 'position' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        pos_diff_col = next((col for col in df.columns if 'position' in col.lower() and 'difference' in col.lower()), None)

        standardized_df['Clicks'] = pd.to_numeric(df[clicks_curr], errors='coerce').fillna(0) if clicks_curr else 0
        standardized_df['Clicks_Delta'] = pd.to_numeric(df[clicks_diff], errors='coerce').fillna(0) if clicks_diff else 0
        standardized_df['Impressions'] = pd.to_numeric(df[impr_curr], errors='coerce').fillna(0) if impr_curr else 0
        standardized_df['Impressions_Delta'] = pd.to_numeric(df[impr_diff], errors='coerce').fillna(0) if impr_diff else 0
        standardized_df['Position'] = pd.to_numeric(df[pos_curr], errors='coerce').fillna(99.0) if pos_curr else 99.0
        standardized_df['Position_Delta'] = pd.to_numeric(df[pos_diff_col], errors='coerce').fillna(0) if pos_diff_col else 0
    else:
        clicks_col = next((col for col in df.columns if 'clicks' in col.lower()), 'Clicks')
        impr_col = next((col for col in df.columns if 'impressions' in col.lower()), 'Impressions')
        pos_col = next((col for col in df.columns if 'position' in col.lower()), 'Position')
        
        standardized_df['Clicks'] = pd.to_numeric(df[clicks_col], errors='coerce').fillna(0)
        standardized_df['Clicks_Delta'] = 0
        standardized_df['Impressions'] = pd.to_numeric(df[impr_col], errors='coerce').fillna(0)
        standardized_df['Impressions_Delta'] = 0
        standardized_df['Position'] = pd.to_numeric(df[pos_col], errors='coerce').fillna(99.0)
        standardized_df['Position_Delta'] = 0
        
    return standardized_df, is_comparison, page_col, query_col


# =========================================================================
# DYNAMIC EXTRACTION ENGINE
# =========================================================================
def process_uploaded_file(uploaded_file):
    filename = uploaded_file.name.lower()
    
    # --- Case 1: Raw ZIP File from Google Search Console ---
    if filename.endswith(".zip"):
        try:
            with zipfile.ZipFile(uploaded_file) as z:
                file_list = z.namelist()
                
                # Check for individual Queries and Pages files inside the ZIP
                queries_file = next((f for f in file_list if "queries.csv" in f.lower()), None)
                pages_file = next((f for f in file_list if "pages.csv" in f.lower()), None)
                
                if queries_file and pages_file:
                    st.info("📦 Detected GSC ZIP! Extracting and auto-merging Queries and Pages...")
                    with z.open(queries_file) as q_f, z.open(pages_file) as p_f:
                        queries_df = pd.read_csv(q_f)
                        pages_df = pd.read_csv(p_f)
                    
                    # Standardize both files
                    q_std, is_comp_q, _, _ = standardize_df_columns(queries_df)
                    p_std, is_comp_p, _, _ = standardize_df_columns(pages_df)
                    
                    # Distribute Query values to Page dimensions via cross-join mapping
                    # When files are split, we cross-map to evaluate cannibalization proxies
                    merged_df = pd.merge(q_std, p_std, how='cross', suffixes=('_Query', '_Page'))
                    
                    # Restructure properties for cannibalization pipeline
                    final_df = pd.DataFrame()
                    final_df['Query'] = merged_df['Query']
                    final_df['Page'] = merged_df['Page']
                    final_df['Clicks'] = merged_df['Clicks_Page']
                    final_df['Clicks_Delta'] = merged_df['Clicks_Delta_Page']
                    final_df['Impressions'] = merged_df['Impressions_Page']
                    final_df['Impressions_Delta'] = merged_df['Impressions_Delta_Page']
                    final_df['Position'] = merged_df['Position_Page']
                    final_df['Position_Delta'] = merged_df['Position_Delta_Page']
                    
                    return final_df, (is_comp_q or is_comp_p)
                
                # If they uploaded a custom ZIP with a single combined search_results file
                search_results_file = next((f for f in file_list if "search_results.csv" in f.lower() or "search results.csv" in f.lower()), None)
                if search_results_file:
                    with z.open(search_results_file) as f:
                        df = pd.read_csv(f)
                    res = standardize_df_columns(df)
                    if res and res[0] is not None:
                        return res[0], res[1]
                        
            st.error("❌ ZIP format invalid: Both 'Queries.csv' and 'Pages.csv' are required inside the ZIP.")
            return None
        except Exception as e:
            st.error(f"Error unzipping GSC export: {e}")
            return None
            
    # --- Case 2: Combined CSV ---
    elif filename.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
            result = standardize_df_columns(df)
            
            if result is None or result[0] is None or 'Query' not in result[0].columns or 'Page' not in result[0].columns:
                page_col, query_col = result[2], result[3] if result else (None, None)
                st.error(f"""
                ### ❌ Dimension Mismatch
                CSV files must contain **both** `Query` and `Page` columns mapped together. 
                * **Found Page:** `{page_col if page_col else '❌ Missing'}`
                * **Found Query:** `{query_col if query_col else '❌ Missing'}`
                
                *Tip: Upload the raw **GSC ZIP** instead and let the tool auto-merge them!*
                """)
                return None
            return result[0], result[1]
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
            return None
        
    return None

# =========================================================================
# FILE UPLOAD CONSOLE
# =========================================================================
st.subheader("📂 Upload GSC Data Package")
uploaded_file = st.file_uploader("Drop your raw GSC ZIP (or a combined CSV) here:", type=["csv", "zip"])

if uploaded_file is None:
    st.info("""
    💡 **How to export this ZIP from Google Search Console:**
    1. Open GSC and go to the **Performance** report.
    2. Click **Export** in the top-right corner.
    3. Choose **Download ZIP**.
    4. Drop that exact downloaded ZIP file here! The app will extract, read, and merge both files for you instantly.
    """)

# =========================================================================
# PIPELINE EXECUTION
# =========================================================================
if uploaded_file is not None:
    extracted_data = process_uploaded_file(uploaded_file)
    
    if extracted_data is not None:
        clean_df, is_compare = extracted_data
        
        if clean_df is not None:
            if is_compare:
                st.success("⚡ Comparison Dataset successfully merged and loaded!")
            else:
                st.warning("⚠️ Standard dataset loaded. Historical deltas set to 0.")
            
            # --- 1. FILTER BRAND & POSITION BOUNDARIES ---
            if BRAND_KEYWORD:
                clean_df = clean_df[~clean_df['Query'].str.lower().str.contains(BRAND_KEYWORD, na=False)]
            clean_df = clean_df[clean_df['Position'] <= MAX_POSITION_LIMIT]
            
            # --- 2. AGGREGATE DUPLICATES (URL roll-up) ---
            rolled_df = clean_df.groupby(['Query', 'Page']).agg({
                'Clicks': 'sum',
                'Clicks_Delta': 'sum',
                'Impressions': 'sum',
                'Impressions_Delta': 'sum',
                'Position': 'mean',
                'Position_Delta': 'mean'
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
                    primary_clicks_delta = int(primary_row['Clicks_Delta'])
                    
                    for index in range(1, len(pages_data)):
                        cannibal_row = pages_data.iloc[index]
                        cannibal_url = cannibal_row['Page']
                        cannibal_clicks = int(cannibal_row['Clicks'])
                        cannibal_pos = round(cannibal_row['Position'], 1)
                        cannibal_clicks_delta = int(cannibal_row['Clicks_Delta'])
                        
                        pos_diff = abs(primary_pos - cannibal_pos)
                        
                        if pos_diff < MAX_POSITION_GAP:
                            is_critical = "🔴 High Threat (Cannibal Gaining)" if (cannibal_clicks_delta > 0 and primary_clicks_delta < 0) else "🟡 Moderate (Stable Competition)"
                            
                            deopt_results.append({
                                "Threat Level": is_critical,
                                "Keyword/Query": query,
                                "Primary URL": primary_url,
                                "Primary Clicks (3M)": primary_clicks,
                                "Primary Clicks Delta": primary_clicks_delta,
                                "Primary Position": primary_pos,
                                "Cannibal URL to De-Optimize": cannibal_url,
                                "Cannibal Clicks (3M)": cannibal_clicks,
                                "Cannibal Clicks Delta": cannibal_clicks_delta,
                                "Cannibal Position": cannibal_pos,
                                "Position Gap": round(pos_diff, 1)
                            })
            
            final_deopt_df = pd.DataFrame(deopt_results)
            
            # =========================================================================
            # REPORTING INTERFACE
            # =========================================================================
            if not final_deopt_df.empty:
                critical_count = len(final_deopt_df[final_deopt_df["Threat Level"].str.contains("🔴")])
                
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.metric("Total Conflicts Identified", len(final_deopt_df))
                with col_m2:
                    st.metric("🔴 High-Threat Conflicts", critical_count)
                with col_m3:
                    avg_gap = round(final_deopt_df['Position Gap'].mean(), 1)
                    st.metric("Average Rank Proximity", f"{avg_gap} Positions")
                
                csv_buffer = io.StringIO()
                final_deopt_df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()
                
                st.download_button(
                    label="💾 Download '[SEO] De-Optimization Actions' Report (CSV)",
                    data=csv_data,
                    file_name="gsc_cannibalization_map.csv",
                    mime="text/csv"
                )
                
                st.subheader("📋 Directives: [SEO] To De-Optimize")
                st.dataframe(final_deopt_df, use_container_width=True)
                
                st.markdown("---")
                st.subheader("🔥 Comparison-Based Action Strategy")
                
                col_l, col_r = st.columns([1, 1])
                
                with col_l:
                    st.markdown("### 🔴 Critical Threat Alerts Explained")
                    st.write("""
                    The engine flags a conflict as **High Threat (🔴)** when:
                    1. The secondary (Cannibal) page **grew in clicks** over the last 3 months.
                    2. Your primary target page **lost clicks** over the same period.
                    """)
                
                with col_r:
                    st.markdown("### 🛡️ De-Optimization Blueprint")
                    st.markdown("""
                    *   **Remove Internal Anchor Text:** Find any internal links pointing to the cannibal page with the target keyword and redirect them to the **Primary URL**.
                    *   **Add "Contextual Relinking":** Link directly to the primary target from the cannibal body text to clearly declare ranking intent to Google.
                    """)
                    
            else:
                st.warning("✅ Clean SEO Horizon! No cannibalization targets found where competing pages rank within 10 positions of each other.")
