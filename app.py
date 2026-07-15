function detectCannibalization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  
  let deoptSheet = ss.getSheetByName("[SEO] To De-Optimize");
  if (!deoptSheet) deoptSheet = ss.insertSheet("[SEO] To De-Optimize");
  
  // =========================================================================
  // CORE PARAMETERS 
  // =========================================================================
  const MIN_IMPRESSIONS = 100;       
  const MAX_POSITION = 50;          
  const BRAND_KEYWORD = "botoxie"; 
  const CTR_DIVERGENCE_LIMIT = 15;  
  const MIN_TOP_POSITION = 3.0;     
  const MAX_POSITION_GAP = 10;      
  // =========================================================================

  let siteUrl = dashboard.getRange("A1").getValue().trim();
  if (!siteUrl) {
    console.warn("Execution halted: Please enter your GSC Property URL in cell A1 of the Dashboard sheet.");
    return;
  }
  
  deoptSheet.clear();
  
  const today = new Date();
  const thirtyDaysAgo = new Date(today.getTime() - (30 * 24 * 60 * 60 * 1000));
  const formatDate = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  
  const payload = {
    startDate: formatDate(thirtyDaysAgo),
    endDate: formatDate(today),
    dimensions: ["query", "page"],
    rowLimit: 25000 
  };
  
  const apiUrl = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}/searchAnalytics/query`;
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  
  try {
    const httpResponse = UrlFetchApp.fetch(apiUrl, options);
    if (httpResponse.getResponseCode() !== 200) {
      console.error(`API Error: ${httpResponse.getContentText()}`);
      return;
    }
    
    const response = JSON.parse(httpResponse.getContentText());
    if (!response.rows || response.rows.length === 0) {
      console.warn("No data found for the last 30 days in GSC.");
      return;
    }
    
    // =========================================================================
    // DYNAMIC PARSING FOR STANDARD VS COMPARISON DATA
    // =========================================================================
    const keywordMap = new Map();
    const excludeTerm = BRAND_KEYWORD.toLowerCase().trim();
    
    // Read the first row to determine if this is a comparison export
    const sampleRow = response.rows[0];
    const isComparisonFile = sampleRow.hasOwnProperty('clicksDifference') || sampleRow.hasOwnProperty('compareClicks'); 

    for (let i = 0; i < response.rows.length; i++) {
      const row = response.rows[i];
      const rawUrl = row.keys[1].trim().split("#")[0]; 
      const query = row.keys[0].toLowerCase();
      
      // Determine columns based on file schema (Comparison vs Standard GSC API mapping)
      // Standardizes both to use the most recent 3-month window performance
      const clicks = isComparisonFile ? (row.compareClicks || row.clicks || 0) : row.clicks;
      const impressions = isComparisonFile ? (row.compareImpressions || row.impressions || 0) : row.impressions;
      const ctr = isComparisonFile ? (row.compareCtr || row.ctr || 0) : row.ctr;
      const position = isComparisonFile ? (row.comparePosition || row.position || 0) : row.position;
      
      if (excludeTerm && query.includes(excludeTerm)) continue;
      if (position > MAX_POSITION) continue;
      
      if (!keywordMap.has(row.keys[0])) {
        keywordMap.set(row.keys[0], {});
      }
      
      const queryGroup = keywordMap.get(row.keys[0]);
      if (!queryGroup[rawUrl]) {
        queryGroup[rawUrl] = { page: rawUrl, clicks: 0, impressions: 0, ctrSum: 0, posSum: 0, count: 0 };
      }
      
      queryGroup[rawUrl].clicks += clicks;
      queryGroup[rawUrl].impressions += impressions;
      queryGroup[rawUrl].ctrSum += ctr;
      queryGroup[rawUrl].posSum += position;
      queryGroup[rawUrl].count += 1;
    }
    
    const deoptRows = [["Keyword/Query", "Primary URL", "Primary Clicks", "Primary Position", "Cannibal URL to De-Optimize", "Cannibal Clicks", "Cannibal Position"]];
    
    for (let [query, pagesObj] of keywordMap.entries()) {
      const pagesArray = Object.values(pagesObj);
      
      if (pagesArray.length > 1) {
        pagesArray.sort((a, b) => (b.clicks - a.clicks) || (b.impressions - a.impressions));
        
        let totalImpressions = 0;
        const finalizedPages = pagesArray.map(p => {
          totalImpressions += p.impressions;
          return {
            page: p.page,
            clicks: p.clicks,
            impressions: p.impressions,
            ctr: parseFloat(((p.ctrSum / p.count) * 100).toFixed(2)),
            position: parseFloat((p.posSum / p.count).toFixed(1))
          };
        });
        
        if (totalImpressions < MIN_IMPRESSIONS) continue;
        
        const bestPage = finalizedPages[0];
        if (bestPage.position < MIN_TOP_POSITION) continue;
        
        for (let j = 1; j < finalizedPages.length; j++) {
          const secondaryPage = finalizedPages[j];
          const ctrGap = Math.abs(bestPage.ctr - secondaryPage.ctr);
          if (ctrGap > CTR_DIVERGENCE_LIMIT) continue;
          
          const positionDifference = Math.abs(bestPage.position - secondaryPage.position);
          
          if (positionDifference < MAX_POSITION_GAP) {
            deoptRows.push([query, bestPage.page, bestPage.clicks, bestPage.position, secondaryPage.page, secondaryPage.clicks, secondaryPage.position]);
          }
        }
      }
    }
    
    if (deoptRows.length > 1) {
      deoptSheet.getRange(1, 1, deoptRows.length, 7).setValues(deoptRows);
      deoptSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#FCE8E6"); 
      deoptSheet.autoResizeColumns(1, 7);
      console.log(`Analysis complete! Identified ${deoptRows.length - 1} cannibalization targets.`);
    } else {
      console.log("No matching cannibalization targets found.");
    }
    
  } catch (error) {
    console.error("Execution failed: " + error.toString());
  }
}
