import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="Free GSC SEO Automation Dashboard", layout="wide")
st.title("📊 Google Search Console Automation Dashboard")
st.markdown("Upload your GSC exported CSV data to instantly find quick-win SEO opportunities.")

# 2. File Uploader Sidebar
st.sidebar.header("Upload GSC Data")
uploaded_file = st.sidebar.file_uploader(
    "Upload 'Queries.csv' or 'Pages.csv' from GSC export", type=["csv"]
)

# Helper function to aggressively clean numeric columns
def clean_numeric_col(series):
    # Convert to string, remove commas, spaces, percentage signs, and coerce errors to NaN
    cleaned = pd.to_numeric(series.astype(str).str.replace(r'[,\s%]', '', regex=True), errors='coerce')
    # Fill any missing/broken values with 0
    return cleaned.fillna(0)

if uploaded_file is not None:
    try:
        # Load data (Skip blank lines and ignore potential corruption)
        df = pd.read_csv(uploaded_file, encoding='utf-8', skip_blank_lines=True)
        
        # Clean up column names (lowercase and strip spaces)
        df.columns = [col.lower().strip() for col in df.columns]
        
        # Check for the core columns
        required_cols = ['clicks', 'impressions', 'ctr', 'position']
        if not all(col in df.columns for col in required_cols):
            st.error(f"The CSV must contain these columns: {required_cols}. Found columns: {list(df.columns)}")
            st.stop()
        
        # Force convert all numeric columns to float/int to fix the 'str' vs 'int' bug
        df['clicks'] = clean_numeric_col(df['clicks']).astype(int)
        df['impressions'] = clean_numeric_col(df['impressions']).astype(int)
        df['position'] = clean_numeric_col(df['position']).astype(float)
        
        # Handle CTR explicitly (if it was text like '5.2%', clean_numeric_col turned it into 5.2. We want 0.052)
        if df['ctr'].dtype == object or df['ctr'].max() > 1:
            df['ctr'] = clean_numeric_col(df['ctr'])
            # If values are like 52.5 instead of 0.525, divide by 100
            if df['ctr'].max() > 1:
                df['ctr'] = df['ctr'] / 100

        # Drop rows where impressions or clicks are 0 due to GSC summary rows at the bottom
        first_col = df.columns[0]
        df = df[df[first_col].notna() & (df[first_col].astype(str).str.strip() != '')]
        df = df[~df[first_col].astype(str).str.contains('Total|Summary', case=False, na=False)]

        # Main metrics display
        total_clicks = df['clicks'].sum()
        total_impressions = df['impressions'].sum()
        avg_ctr = total_clicks / total_impressions if total_impressions > 0 else 0
        avg_pos = df['position'].mean()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Clicks", f"{total_clicks:,}")
        col2.metric("Total Impressions", f"{total_impressions:,}")
        col3.metric("Average CTR", f"{avg_ctr:.2%}")
        col4.metric("Average Position", f"{avg_pos:.1f}")
        
        st.markdown("---")

        # -------------------------------------------------------------
        # AUTOMATED INSIGHT 1: Striking Distance Keywords (Pos 11-20)
        # -------------------------------------------------------------
        st.subheader("🎯 Striking Distance Opportunities (Ranked 11-20)")
        st.markdown("These terms are sitting on page 2 of Google. A minor on-page tweak or internal link could push them to page 1!")
        
        # Safe numeric filtering
        striking_df = df[(df['position'] >= 11) & (df['position'] <= 20)].copy()
        striking_df = striking_df.sort_values(by='impressions', ascending=False)
        
        if not striking_df.empty:
            fig_striking = px.scatter(
                striking_df.head(30), 
                x='position', 
                y='impressions', 
                size='clicks', 
                hover_name=first_col,
                title="Top 30 Striking Distance Keywords by Search Volume (Impressions)",
                labels={'position': 'Search Position', 'impressions': 'Impressions'}
            )
            st.plotly_chart(fig_striking, use_container_width=True)
            st.dataframe(striking_df[[first_col, 'clicks', 'impressions', 'position']], use_container_width=True)
        else:
            st.info("No keywords found in the 11-20 position range.")

        st.markdown("---")

        # -------------------------------------------------------------
        # AUTOMATED INSIGHT 2: High Impression, Low CTR (Underperformers)
        # -------------------------------------------------------------
        st.subheader("⚠️ High Impressions but Low CTR (Fix Meta Titles/Descriptions)")
        st.markdown("These items get lots of eyeballs but very few clicks. Improving their metadata will heavily spike traffic.")
        
        median_imps = df['impressions'].median()
        underperforming_df = df[(df['impressions'] > median_imps) & (df['ctr'] < avg_ctr)].copy()
        underperforming_df = underperforming_df.sort_values(by='impressions', ascending=False)

        if not underperforming_df.empty:
            fig_underperform = px.bar(
                underperforming_df.head(15),
                x=first_col,
                y='impressions',
                color='ctr',
                title="Top 15 High Impression Items with Below-Average CTR",
                labels={'ctr': 'Click-Through Rate'}
            )
            st.plotly_chart(fig_underperform, use_container_width=True)
            st.dataframe(underperforming_df[[first_col, 'clicks', 'impressions', 'ctr', 'position']], use_container_width=True)
        else:
            st.info("No underperforming high-impression items found.")

    except Exception as e:
        st.error(f"Error parsing file: {e}. Please ensure you are uploading a valid CSV directly exported from GSC.")

else:
    st.info("👋 Please upload a GSC 'Queries' or 'Pages' CSV file in the sidebar to generate automation insights.")
