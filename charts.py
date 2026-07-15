import plotly.express as px

def plot_top_losers(df_q):
    losers = df_q[df_q['Clicks_Delta'] < 0].sort_values(by='Clicks_Delta', ascending=True).head(10)
    if losers.empty:
        return None
    fig = px.bar(
        losers, 
        x='Queries', 
        y='Clicks_Delta', 
        title='Top 10 Clicks Declines', 
        color_discrete_sequence=['#ef4444']
    )
    fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    return fig

def plot_ctr_vs_position(df_q):
    sample = df_q[df_q['Impressions'] >= 50].head(150)
    if sample.empty:
        return None
    fig = px.scatter(
        sample, 
        x='Position', 
        y='CTR', 
        size='Impressions', 
        hover_name='Queries', 
        title='CTR vs Rank Distribution (Size = Impressions)',
        color_discrete_sequence=['#3b82f6']
    )
    fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    return fig
