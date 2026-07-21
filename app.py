# === TAB 3: CANNIBALIZATION CLASHES ===
        with tabs[2]:
            st.markdown("## ⚔️ Keyword Cannibalization Clashes")
            st.markdown("""
            *Detects queries where **2 or more distinct internal URLs** are ranking simultaneously in search results. Check their individual positions below to determine if they are competing on the **same SERP**.*
            """)
            
            clash_detected = False
            
            # Filter candidates with reasonable impressions and ranking within top 50
            candidate_queries = df_q[(df_q['Impressions'] >= MIN_IMPR_THRESHOLD) & (df_q['Position'] <= 50)].sort_values(by='Impressions', ascending=False)
            
            for _, q_row in candidate_queries.iterrows():
                query = q_row['Queries']
                query_tokens = [re.sub(r'[^a-z0-9]', '', t) for t in query.lower().split() if len(t) > 3]
                
                if len(query_tokens) >= 1:
                    # Search pages dataframe for URLs matching the query tokens
                    matching_pages = df_p[df_p['Pages'].str.lower().apply(
                        lambda x: sum(1 for token in query_tokens if token in x) >= max(1, len(query_tokens) - 1)
                    )].copy()
                    
                    # Deduplicate and ensure at least 2 distinct competing URLs exist
                    matching_pages = matching_pages.sort_values(by='Impressions', ascending=False).drop_duplicates(subset=['Pages'])
                    
                    if len(matching_pages) >= 2:
                        clash_detected = True
                        top_competing_pages = matching_pages.head(3)  # Get top competing pages
                        
                        # Store for Execution Blueprint
                        page_list = top_competing_pages['Pages'].tolist()
                        cannibal_clashes_extracted.append({"query": query, "url_1": page_list[0], "url_2": page_list[1]})
                        
                        # Calculate if both are on Page 1 (Ranks 1–10)
                        ranks = top_competing_pages['Position'].tolist()
                        both_page_one = all(r <= 10.0 for r in ranks[:2])
                        serp_status = "⚠️ Active Direct SERP Competition (Both on Page 1)" if both_page_one else "⚡ Keyword Splitting / Alternating Ranks"
                        
                        # Render Detailed Card
                        st.markdown(f"""
                        <div class="directive-card warning" style="border-left: 6px solid #f59e0b !important;">
                            <div class="directive-title" style="font-size: 1.15rem; color: #b45309 !important;">
                                ⚔️ Clashing Keyword: <code>{query}</code>
                            </div>
                            <div class="directive-text" style="margin-bottom: 10px;">
                                <b>SERP Collision Status:</b> <span class="warning-tag" style="font-size: 0.85rem;">{serp_status}</span>
                            </div>
                            <table style="width:100%; border-collapse: collapse; margin-top: 8px; font-size: 0.9rem;">
                                <tr style="background-color: rgba(128,128,128,0.1); text-align: left;">
                                    <th style="padding: 6px 10px;">Competing URL Path</th>
                                    <th style="padding: 6px 10px;">Rank / Pos</th>
                                    <th style="padding: 6px 10px;">Clicks</th>
                                    <th style="padding: 6px 10px;">Impressions</th>
                                </tr>
                        """, unsafe_allow_html=True)
                        
                        for _, p_row in top_competing_pages.iterrows():
                            st.markdown(f"""
                                <tr style="border-bottom: 1px solid rgba(128,128,128,0.2);">
                                    <td style="padding: 6px 10px;"><code>{p_row['Pages']}</code></td>
                                    <td style="padding: 6px 10px;"><b>{round(p_row['Position'], 1)}</b></td>
                                    <td style="padding: 6px 10px;">{int(p_row['Clicks'])}</td>
                                    <td style="padding: 6px 10px;">{int(p_row['Impressions'])}</td>
                                </tr>
                            """, unsafe_allow_html=True)
                            
                        st.markdown("""
                            </table>
                            <div style="margin-top:10px; font-size:0.85rem; color:var(--text-color);">
                                💡 <b>Recommendation:</b> Decide which URL has higher conversion intent. Add a <code>rel="canonical"</code> tag pointing to the primary page, adjust internal anchor links, or consolidate thin content into the stronger ranking page.
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
            if not clash_detected:
                st.success("✅ No keyword cannibalization clashes found matching current filters.")
