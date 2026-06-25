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

if uploaded_file is not None:
    try:
        # Load data (GSC exports often contain a few header rows or use different encodings)
        df = pd.read_csv(uploaded_file, encoding='utf-8')
        
        # Standardize column names (lowercasing to prevent case-sensitive bugs)
        df.columns = [col.lower().strip() for col in df.columns]
        
        # Verify required columns exist
        required_cols = ['clicks', 'impressions', 'ctr', 'position']
        if not all(col in df.columns for col in required_cols):
            st.error(f"The CSV must contain these columns: {required_cols}. Please check your GSC export.")
            st.stop()
            
        # Clean data types (Clean strings like '1.5%' to float 0.015 if necessary)
        if df['ctr'].dtype == object:
            df['ctr'] = df['ctr'].str.rstrip('%').astype('float') / 100
        elif df['ctr'].max() > 1:
            df['ctr'] = df['ctr'] / 100 # Adjust if CTR is exported as whole numbers (e.g. 50 instead of 0.5)

        # Main metrics display
        total_clicks = df['clicks'].sum()
        total_imps = df['impressions'].sum()
        avg_ctr = df['clicks'].sum() / df['impressions'].sum() if total_imps > 0 else 0
        avg_pos = df['position'].mean()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Clicks", f"{total_clicks:,}")
        col2.metric("Total Impressions", f"{total_imps:,}")
        col3.metric("Average CTR", f"{avg_ctr:.2%}")
        col4.metric("Average Position", f"{avg_pos:.1f}")
        
        st.markdown("---")

        # -------------------------------------------------------------
        # AUTOMATED INSIGHT 1: Striking Distance Keywords (Pos 11-20)
        # -------------------------------------------------------------
        st.subheader("🎯 Striking Distance Opportunities (Ranked 11-20)")
        st.markdown("These terms are sitting on page 2 of Google. A minor on-page tweak or internal link could push them to page 1!")
        
        # Filter for keywords between position 11 and 20
        striking_df = df[(df['position'] >= 11) & (df['position'] <= 20)].copy()
        striking_df = striking_df.sort_values(by='impressions', ascending=False)
        
        if not striking_df.empty:
            # Interactive Plotly Chart
            fig_striking = px.scatter(
                striking_df.head(30), 
                x='position', 
                y='impressions', 
                size='clicks', 
                hover_name=df.columns[0], # Dynamically uses 'query' or 'top queries'
                title="Top 30 Striking Distance Keywords by Search Volume (Impressions)",
                labels={'position': 'Search Position', 'impressions': 'Impressions (Search Volume)'}
            )
            st.plotly_chart(fig_striking, use_container_width=True)
            
            # Data table display
            st.dataframe(striking_df[[df.columns[0], 'clicks', 'impressions', 'position']], use_container_width=True)
        else:
            st.info("No keywords found in the 11-20 position range.")

        st.markdown("---")

        # -------------------------------------------------------------
        # AUTOMATED INSIGHT 2: High Impression, Low CTR (Underperformers)
        # -------------------------------------------------------------
        st.subheader("⚠️ High Impressions but Low CTR (Fix Meta Titles/Descriptions)")
        st.markdown("These items get lots of eyeballs but very few clicks. Improving their metadata or search intent alignment will heavily spike traffic.")
        
        # Filter: Impressions above median, but CTR below average
        median_imps = df['impressions'].median()
        underperforming_df = df[(df['impressions'] > median_imps) & (df['ctr'] < avg_ctr)].copy()
        underperforming_df = underperforming_df.sort_values(by='impressions', ascending=False)

        if not underperforming_df.empty:
            fig_underperform = px.bar(
                underperforming_df.head(15),
                x=df.columns[0],
                y='impressions',
                color='ctr',
                title="Top 15 High Impression Items with Below-Average CTR",
                labels={'ctr': 'Click-Through Rate'}
            )
            st.plotly_chart(fig_underperform, use_container_width=True)
            st.dataframe(underperforming_df[[df.columns[0], 'clicks', 'impressions', 'ctr', 'position']], use_container_width=True)
        else:
            st.info("No underperforming high-impression items found.")

    except Exception as e:
        st.error(f"Error parsing file: {e}. Please ensure you are uploading a valid CSV directly exported from GSC.")

else:
    st.info("👋 Please upload a GSC 'Queries' or 'Pages' CSV file in the sidebar to generate automation insights.")
