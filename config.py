# Standard CTR expectations for Page 1
CTR_BENCHMARKS = {
    1: 30.0, 2: 15.0, 3: 10.0, 4: 7.0, 5: 5.0,
    6: 4.0,  7: 3.0,  8: 2.5,  9: 2.0,  10: 1.5
}

# Apply mathematical CTR decay for Positions 11 to 30
for pos in range(11, 31):
    CTR_BENCHMARKS[pos] = round(15.0 / pos, 2)

# Page Layout configuration
CSS_STYLING = """
<style>
    .stApp { background-color: #f8fafc; }
    h1, h2, h3, h4, h5 { color: #0f172a !important; font-family: monospace; }
    
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 12px;
    }
    .metric-value {
        font-size: 1.75rem;
        font-weight: 700;
        color: #0f172a;
        font-family: monospace;
    }
    .metric-value.negative { color: #ef4444; }
    .metric-value.positive { color: #10b981; }
    .metric-value.warning { color: #f59e0b; }
    .metric-label {
        font-size: 0.75rem;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-top: 4px;
        font-weight: 600;
    }
    .alert-banner {
        background-color: #0f172a;
        color: #ffffff;
        padding: 20px;
        border-radius: 8px;
        font-family: monospace;
        margin-bottom: 25px;
        border-left: 6px solid #3b82f6;
    }
    .alert-banner h2 { color: #ffffff !important; margin: 0 0 6px 0; }
    .alert-banner p { color: #94a3b8; margin: 0; font-size: 0.9rem; }
</style>
"""
