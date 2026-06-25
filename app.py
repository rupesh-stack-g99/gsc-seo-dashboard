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
        # Load data (Skip blank lines)
        df = pd.read_csv(uploaded_file, encoding='utf-8', skip_blank_lines=True)
        
        # Clean up column names (lowercase and strip spaces)
        df.columns = [col.lower().strip() for col in df.columns]
        
        # Check for the core columns
        required_cols = ['clicks', 'impressions', 'ctr', 'position']
        if not all(col in df.columns for col in required_cols):
            st.error(f"The CSV must contain these columns: {required_cols}. Found columns: {list(df.columns)}")
            st.stop()
        
        first_col = df.columns[0]
        
        # Drop the "Total" summary rows that Google adds to the very bottom
        df = df[df[first_col].notna() & (df[first_col].astype(str).str.strip() != '')]
        df = df[~df[first_col].astype(str).str.contains('Total|Summary', case=False, na=False)]

        # CRITICAL FIX: Convert columns and force errors to 'NaN' (null) instead of keeping them as text strings
        for col in ['clicks', 'impressions', 'position']:
            df[col] = df[col].astype(str).str.replace(r'[,\s]', '', regex=True) # strip out commas/spaces
            df[col] = pd.to_numeric(df[col], errors='coerce') # Force text to actual NaN numbers

        # Clean CTR explicitly
        df['ctr'] = df['ctr'].astype(str).str.replace(r'[,\s%]', '', regex=True)
        df['ctr'] = pd.to_numeric(df['ctr'], errors='coerce')
        # If CTR is listed as a percentage (e.g., 5.2 instead of 0.052), normalize it
        if df['ctr'].max() > 1:
            df['ctr'] = df['ctr'] / 100

        # Drop any leftover rows that failed numeric conversion (ensures NO 'str' variables remain in numbers)
        df = df.dropna(subset=['clicks', 'impressions', 'ctr', 'position'])

        # Now it is safe to convert data types natively
        df['clicks'] = df['clicks'].astype(int)
        df['impressions'] = df['impressions'].astype(int)
        df['position'] = df['position'].astype(float)
        df['ctr'] = df['ctr'].astype(float)

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
        
        # This operation will absolutely work now because positions are purely float data type
        striking_df = df[(df['position'] >= 11.0) & (df['position'] <= 20.0)].copy()
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
        st.error(f"🚨 Code execution stalled. Error: {e}")

else:
    st.info("👋 Please upload a GSC 'Queries' or 'Pages' CSV file in the sidebar to generate automation insights.")
