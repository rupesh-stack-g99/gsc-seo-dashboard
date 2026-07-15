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

# Custom CSS for polished interface
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

st.title("🎯 GSC 3-Month Comparison & De-Optimization Engine")
st.write("Upload your GSC 3-Month Compare Export. This engine automatically matches your comparison intervals and maps keyword-level conflicts.")

# =========================================================================
# Side Bar Controls
# =========================================================================
st.sidebar.header("🛠️ Diagnostic Parameters")

BRAND_KEYWORD = st.sidebar.text_input("Brand Keyword to Exclude", value="botoxie").lower().strip()
MIN_IMPRESSIONS = st.sidebar.number_input("Min Query Impressions (Last 3M)", min_value=1, value=100)
MAX_POSITION_GAP = st.sidebar.number_input("Max Position Difference (Proximity Limit)", min_value=1, max_value=20, value=10)
MAX_POSITION_LIMIT = st.sidebar.number_input("Max Allowed Position (Filter Boundary)", min_value=10, max_value=100, value=50)

st.sidebar.markdown("---")
st.sidebar.info("""
**How Comparison Logic Works:**
The script auto-detects column names like `Clicks (Last 3 months)` vs `Clicks (Previous 3 months)`. It evaluates the **current** conflict and warns you if the cannibalizing page is actively gaining ground while your primary page is losing ground.
""")

# =========================================================================
# DETECTOR & PARSING ENGINE
# =========================================================================
def standardize_df_columns(df):
    """
    Parses and standardizes GSC Compare datasets and standard datasets alike.
    """
    df.columns = [col.strip() for col in df.columns]
    
    # Locate page and query column structures
    page_col = next((col for col in df.columns if col.lower() in ['page', 'landing page', 'urls', 'pages', 'url', 'landing_page']), None)
    query_col = next((col for col in df.columns if col.lower() in ['query', 'keyword', 'search term', 'queries', 'search_query']), None)
    
    if not page_col or not query_col:
        return None, False

    # Check if this is a Compare file
    is_comparison = any('difference' in col.lower() or 'last' in col.lower() or 'compare' in col.lower() for col in df.columns)
    
    standardized_df = pd.DataFrame()
    standardized_df['Query'] = df[query_col].astype(str).str.strip()
    standardized_df['Page'] = df[page_col].astype(str).apply(lambda x: x.split('#')[0].strip())
    
    if is_comparison:
        # Dynamic extraction of compare intervals (e.g., 'Clicks (Last 3 months)' or 'Difference Clicks')
        clicks_curr = next((col for col in df.columns if 'clicks' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        clicks_diff = next((col for col in df.columns if 'clicks' in col.lower() and 'difference' in col.lower()), None)
        
        impr_curr = next((col for col in df.columns if 'impressions' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        impr_diff = next((col for col in df.columns if 'impressions' in col.lower() and 'difference' in col.lower()), None)
        
        pos_curr = next((col for col in df.columns if 'position' in col.lower() and ('last' in col.lower() or 'recent' in col.lower())), None)
        pos_diff_col = next((col for col in df.columns if 'position' in col.lower() and 'difference' in col.lower()), None)

        # Map to localized names
        standardized_df['Clicks'] = pd.to_numeric(df[clicks_curr], errors='coerce').fillna(0) if clicks_curr else 0
        standardized_df['Clicks_Delta'] = pd.to_numeric(df[clicks_diff], errors='coerce').fillna(0) if clicks_diff else 0
        
        standardized_df['Impressions'] = pd.to_numeric(df[impr_curr], errors='coerce').fillna(0) if impr_curr else 0
        standardized_df['Impressions_Delta'] = pd.to_numeric(df[impr_diff], errors='coerce').fillna(0) if impr_diff else 0
        
        standardized_df['Position'] = pd.to_numeric(df[pos_curr], errors='coerce').fillna(99.0) if pos_curr else 99.0
        standardized_df['Position_Delta'] = pd.to_numeric(df[pos_diff_col], errors='coerce').fillna(0) if pos_diff_col else 0
    else:
        # Fall back to standard columns if no comparison is found
        clicks_col = next((col for col in df.columns if 'clicks' in col.lower()), 'Clicks')
        impr_col = next((col for col in df.columns if 'impressions' in col.lower()), 'Impressions')
        pos_col = next((col for col in df.columns if 'position' in col.lower()), 'Position')
        
        standardized_df['Clicks'] = pd.to_numeric(df[clicks_col], errors='coerce').fillna(0)
        standardized_df['Clicks_Delta'] = 0
        standardized_df['Impressions'] = pd.to_numeric(df[impr_col], errors='coerce').fillna(0)
        standardized_df['Impressions_Delta'] = 0
        standardized_df['Position'] = pd.to_numeric(df[pos_col], errors='coerce').fillna(99.0)
        standardized_df['Position_Delta'] = 0
        
    return standardized_df, is_comparison


def extract_gsc_data_from_file(uploaded_file):
    """
    Parses ZIP and standalone CSV files containing compare metrics.
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
                
                queries_file = next((f for f in file_list if "queries.csv" in f.lower()), None)
                pages_file = next((f for f in file_list if "pages.csv" in f.lower()), None)
                
                if queries_file and pages_file:
                    st.error("""
                    ### ⚠️ Unmapped ZIP File Structure
                    GSC default exports place Queries and Pages into separate files. They do not share a primary key to connect them.
                    
                    **Action needed:**
                    Please export your compared search data using a Google Sheet extension like **Search Analytics for Sheets** mapping BOTH Page & Query as grouping dimensions together, and download that sheet as a `.csv`.
                    """)
                    return None
                
                st.error("❌ ZIP format invalid: No compatible files detected.")
                return None
                
        except Exception as e:
            st.error(f"Error reading ZIP file: {e}")
            return None
            
    # --- Case 2: Standalone CSV ---
    elif filename.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
            result = standardize_df_columns(df)
            if result is None or result[0] is None:
                st.error(f"""
                ### ❌ Columns Mismatched
                Your file has these headers: `{list(df.columns)}`
                
                We need both **Query** and **Page** headers to run the analysis.
                """)
                return None
            return result
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
            return None
        
    return None

# =========================================================================
# FILE UPLOAD CONSOLE
# =========================================================================
st.subheader("📂 Upload GSC 3-Month Compare Export")
uploaded_file = st.file_uploader("Upload GSC Compare Dataset (CSV or ZIP)", type=["csv", "zip"])

if uploaded_file is None:
    st.info("""
    💡 **Quick Setup Guide:** 
    1. Go to Google Search Console -> **Performance** -> **Compare last 3 months to previous period**.
    2. Since GSC separates keywords and pages, use **Search Analytics for Sheets** (free Google Sheets extension).
    3. Group by both **Query** and **Page** at the same time and request your comparison metrics.
    4. Download that Sheet as a `.csv` and drag-and-drop it here.
    """)

# =========================================================================
# PIPELINE EXECUTION
# =========================================================================
if uploaded_file is not None:
    extracted_data = extract_gsc_data_from_file(uploaded_file)
    
    if extracted_data is not None:
        clean_df, is_compare = extracted_data
        
        if clean_df is not None:
            if is_compare:
                st.success("⚡ Comparison Dataset Loaded Successfully!")
            else:
                st.warning("⚠️ This is a standard single-period file. The delta metrics (e.g., changes in traffic) will default to 0.")
            
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
                            # Alert levels
                            # High Priority if cannibal page is gaining traffic (+ clicks) while the primary is losing traffic (- clicks)
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
                # Group stats
                critical_count = len(final_deopt_df[final_deopt_df["Threat Level"].str.contains("🔴")])
                
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.metric("Total Conflicts Identified", len(final_deopt_df))
                with col_m2:
                    st.metric("🔴 High-Threat Conflicts", critical_count)
                with col_m3:
                    avg_gap = round(final_deopt_df['Position Gap'].mean(), 1)
                    st.metric("Average Rank Proximity", f"{avg_gap} Positions")
                
                # Download Report
                csv_buffer = io.StringIO()
                final_deopt_df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()
                
                st.download_button(
                    label="💾 Download '[SEO] De-Optimization Actions' Clean Compare Report (CSV)",
                    data=csv_data,
                    file_name="gsc_3m_compare_seo_targets.csv",
                    mime="text/csv"
                )
                
                st.subheader("📋 Directives: [SEO] To De-Optimize")
                st.dataframe(final_deopt_df, use_container_width=True)
                
                # =========================================================================
                # ACTIONABLE RECOMMENDATIONS
                # =========================================================================
                st.markdown("---")
                st.subheader("🔥 Comparison-Based Action Strategy")
                
                col_l, col_r = st.columns([1, 1])
                
                with col_l:
                    st.markdown("### 🔴 Critical Threat Alerts Explained")
                    st.write("""
                    The engine flags a conflict as **High Threat (🔴)** when:
                    1. The secondary (Cannibal) page **grew in clicks** over the last 3 months.
                    2. Your primary target page **lost clicks** over the same period.
                    
                    **Action Plan for High-Threat URLs:**
                    *   **Redirect or Canonicalize:** If the cannibal page serves no unique intent, implement a `301 redirect` or set a canonical tag pointing to your Primary page.
                    *   **Intent Segregation:** If you want to keep both pages, change the headings (H1, H2s) on the cannibal page so they don't target the same keyword.
                    """)
                
                with col_r:
                    st.markdown("### 🛡️ De-Optimization Blueprint")
                    st.markdown("""
                    For **Moderate (🟡)** target pages:
                    
                    *   **Remove Internal Anchor Text:** Find any internal links linking to your cannibal page with the exact target keyword. Switch their links to point to the **Primary URL** instead.
                    *   **Add "Contextual Relinking":** At the top of the cannibal page, add a clean text block that says, *"Looking for [Primary Keyword]? Check out our complete guide here [Link to Primary URL]"*. This passes strong thematic context to Google.
                    """)
                    
            else:
                st.warning("✅ Clean SEO Horizon! No cannibalization targets found where competing pages rank within 10 positions of each other.")
