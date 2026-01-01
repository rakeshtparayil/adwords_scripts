/**
 * Keywords Performance Report - 7, 14, 30 Days Analysis
 * 
 * This script shows keyword performance across three time periods:
 * - Last 7 Days
 * - Last 14 Days
 * - Last 30 Days
 * 
 * Metrics Displayed:
 * - Cost for each period
 * - Conversions for each period
 * - Cost Per Conversion (CPA) for each period
 * 
 * Sorted by highest cost (30-day period)
 * 
 * Features:
 * - Gmail email notification with detailed table
 * - Shows Campaign, Ad Group, Keyword, Match Type
 * - Easy to identify high-cost keywords
 * - Compare performance across time periods
 * 
 * To use:
 * 1. Go to Google Ads > Tools & Settings > Bulk Actions > Scripts
 * 2. Click the + button to create a new script
 * 3. Paste this code
 * 4. Configure EMAIL_RECIPIENTS below
 * 5. Save and run or schedule
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Gmail Email Recipients - REQUIRED
// Replace with your actual Gmail addresses
const EMAIL_RECIPIENTS = [
  'example@gmail.com',        // Replace with your Gmail
  ''            // Add more recipients as needed
];

// Minimum clicks threshold for each period (to filter out low-traffic keywords)
const MIN_CLICKS_7_DAYS = 1;
const MIN_CLICKS_14_DAYS = 2;
const MIN_CLICKS_30_DAYS = 3;

// Exchange rate: AED to USD (set to 1 if you don't need conversion)
const AED_TO_USD = 3.67;

// Campaign filter (optional - leave empty to include all campaigns)
const CAMPAIGN_NAME_CONTAINS = '';  // Example: 'Brand' to only include brand campaigns

// Keyword status filter
const KEYWORD_STATUS = 'ENABLED';  // Options: 'ENABLED', 'PAUSED', 'ALL'

// Number of keywords to show in email (sorted by cost)
const MAX_KEYWORDS_IN_EMAIL = 100;

// ============================================================================
// MAIN FUNCTION
// ============================================================================

function main() {
  Logger.log('='.repeat(70));
  Logger.log('🚀 Starting Keywords Performance Report (7, 14, 30 Days)');
  Logger.log('='.repeat(70));
  
  try {
    // Validate configuration
    if (EMAIL_RECIPIENTS.length === 0) {
      throw new Error('Please configure EMAIL_RECIPIENTS in the script settings');
    }
    
    Logger.log(`🎯 Keyword Status: ${KEYWORD_STATUS}`);
    if (CAMPAIGN_NAME_CONTAINS) {
      Logger.log(`🎯 Campaign Filter: ${CAMPAIGN_NAME_CONTAINS}`);
    }
    
    // Fetch keywords for all periods (current and previous)
    Logger.log('\n' + '='.repeat(70));
    Logger.log('📊 FETCHING DATA FOR ALL PERIODS');
    Logger.log('='.repeat(70));
    
    Logger.log('📊 Fetching Last 7 Days data...');
    const keywords7Days = fetchKeywordsData('LAST_7_DAYS', MIN_CLICKS_7_DAYS);
    Logger.log(`✅ Found ${keywords7Days.length} keywords (Last 7 Days)`);
    
    Logger.log('📊 Fetching Previous 7 Days data (days 8-14)...');
    const keywordsPrev7Days = fetchKeywordsDataCustomRange(8, 14, MIN_CLICKS_7_DAYS);
    Logger.log(`✅ Found ${keywordsPrev7Days.length} keywords (Previous 7 Days)`);
    
    Logger.log('📊 Fetching Last 14 Days data...');
    const keywords14Days = fetchKeywordsData('LAST_14_DAYS', MIN_CLICKS_14_DAYS);
    Logger.log(`✅ Found ${keywords14Days.length} keywords (Last 14 Days)`);
    
    Logger.log('📊 Fetching Previous 14 Days data (days 15-28)...');
    const keywordsPrev14Days = fetchKeywordsDataCustomRange(15, 28, MIN_CLICKS_14_DAYS);
    Logger.log(`✅ Found ${keywordsPrev14Days.length} keywords (Previous 14 Days)`);
    
    Logger.log('📊 Fetching Last 30 Days data...');
    const keywords30Days = fetchKeywordsData('LAST_30_DAYS', MIN_CLICKS_30_DAYS);
    Logger.log(`✅ Found ${keywords30Days.length} keywords (Last 30 Days)`);
    
    Logger.log('📊 Fetching Previous 30 Days data (days 31-60)...');
    const keywordsPrev30Days = fetchKeywordsDataCustomRange(31, 60, MIN_CLICKS_30_DAYS);
    Logger.log(`✅ Found ${keywordsPrev30Days.length} keywords (Previous 30 Days)`);
    
    // Combine all keywords from all periods (including previous periods for comparison)
    Logger.log('\n' + '='.repeat(70));
    Logger.log('🔍 COMBINING KEYWORDS FROM ALL PERIODS');
    Logger.log('='.repeat(70));
    const allKeywords = combineAllKeywords(
      keywords7Days, keywordsPrev7Days,
      keywords14Days, keywordsPrev14Days,
      keywords30Days, keywordsPrev30Days
    );
    Logger.log(`✅ Total unique keywords: ${allKeywords.length}`);
    
    // Calculate summaries
    const summary7Days = calculateSummary(keywords7Days);
    const summary14Days = calculateSummary(keywords14Days);
    const summary30Days = calculateSummary(keywords30Days);
    
    // Send Gmail report
    Logger.log('\n' + '='.repeat(70));
    Logger.log('📧 Sending Gmail report...');
    Logger.log('='.repeat(70));
    sendGmailReport({
      allKeywords: allKeywords,
      summary7Days: summary7Days,
      summary14Days: summary14Days,
      summary30Days: summary30Days
    });
    Logger.log(`✅ Email sent to: ${EMAIL_RECIPIENTS.join(', ')}`);
    
    Logger.log('\n' + '='.repeat(70));
    Logger.log('✅ Report generation completed successfully!');
    Logger.log('='.repeat(70));
    
    // Display summary
    displaySummary(summary7Days, summary14Days, summary30Days, allKeywords.length);
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(error.stack);
    throw error;
  }
}

// ============================================================================
// DATA FETCHING
// ============================================================================

function fetchKeywordsData(dateRange, minClicks) {
  const keywords = [];
  
  // Build the query
  let query = `
    SELECT
      campaign.id,
      campaign.name,
      ad_group.id,
      ad_group.name,
      ad_group_criterion.keyword.text,
      ad_group_criterion.keyword.match_type,
      ad_group_criterion.criterion_id,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.ctr,
      metrics.average_cpc,
      metrics.conversions,
      metrics.conversions_value
    FROM keyword_view
    WHERE segments.date DURING ${dateRange}
  `;
  
  // Add status filter
  if (KEYWORD_STATUS === 'ENABLED') {
    query += `
      AND campaign.status = 'ENABLED'
      AND ad_group.status = 'ENABLED'
      AND ad_group_criterion.status = 'ENABLED'
    `;
  } else if (KEYWORD_STATUS === 'PAUSED') {
    query += `
      AND ad_group_criterion.status = 'PAUSED'
    `;
  }
  
  // Add campaign filter if specified
  if (CAMPAIGN_NAME_CONTAINS) {
    query += `
      AND campaign.name CONTAINS_IGNORE_CASE '${CAMPAIGN_NAME_CONTAINS}'
    `;
  }
  
  // Filter for minimum clicks
  query += `
    AND metrics.clicks >= ${minClicks}
  `;
  
  // Order by cost descending (most expensive first)
  query += `
    ORDER BY metrics.cost_micros DESC
  `;
  
  Logger.log(`🔍 Executing query for ${dateRange}...`);
  const report = AdsApp.report(query);
  const rows = report.rows();
  
  while (rows.hasNext()) {
    const row = rows.next();
    
    const campaignId = row['campaign.id'];
    const campaignName = row['campaign.name'];
    const adGroupId = row['ad_group.id'];
    const adGroupName = row['ad_group.name'];
    const keywordText = row['ad_group_criterion.keyword.text'];
    const matchType = row['ad_group_criterion.keyword.match_type'];
    const keywordId = row['ad_group_criterion.criterion_id'];
    const impressions = parseInt(row['metrics.impressions']) || 0;
    const clicks = parseInt(row['metrics.clicks']) || 0;
    const costMicros = parseFloat(row['metrics.cost_micros']) || 0;
    const ctr = parseFloat(row['metrics.ctr']) || 0;
    const avgCpcMicros = parseFloat(row['metrics.average_cpc']) || 0;
    const conversions = parseFloat(row['metrics.conversions']) || 0;
    const conversionsValue = parseFloat(row['metrics.conversions_value']) || 0;
    
    // Convert cost from micros to currency, then AED to USD
    const costAED = costMicros / 1000000;
    const cost = costAED / AED_TO_USD;
    
    // Convert CPC from micros
    const cpcAED = avgCpcMicros / 1000000;
    const cpc = cpcAED / AED_TO_USD;
    
    // Calculate CPA (Cost Per Acquisition)
    const cpa = conversions > 0 ? (cost / conversions) : 0;
    
    // Format match type for display
    let matchTypeDisplay = matchType;
    if (matchType === 'EXACT') {
      matchTypeDisplay = 'Exact';
    } else if (matchType === 'PHRASE') {
      matchTypeDisplay = 'Phrase';
    } else if (matchType === 'BROAD') {
      matchTypeDisplay = 'Broad';
    }
    
    keywords.push({
      campaignId: campaignId,
      campaignName: campaignName,
      adGroupId: adGroupId,
      adGroupName: adGroupName,
      keywordText: keywordText,
      matchType: matchTypeDisplay,
      keywordId: keywordId,
      impressions: impressions,
      clicks: clicks,
      cost: cost,
      ctr: ctr * 100,  // Convert to percentage
      cpc: cpc,
      conversions: conversions,
      conversionsValue: conversionsValue,
      cpa: cpa,
      // Unique identifier for matching across periods
      uniqueKey: `${campaignId}_${adGroupId}_${keywordId}`
    });
  }
  
  return keywords;
}

function fetchKeywordsDataCustomRange(daysAgo, daysAgoEnd, minClicks) {
  const keywords = [];
  
  // Calculate date range
  const today = new Date();
  const endDate = new Date(today);
  endDate.setDate(today.getDate() - daysAgo);
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - daysAgoEnd);
  
  const formatDate = (date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
  };
  
  const startDateStr = formatDate(startDate);
  const endDateStr = formatDate(endDate);
  
  // Build the query
  let query = `
    SELECT
      campaign.id,
      campaign.name,
      ad_group.id,
      ad_group.name,
      ad_group_criterion.keyword.text,
      ad_group_criterion.keyword.match_type,
      ad_group_criterion.criterion_id,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.ctr,
      metrics.average_cpc,
      metrics.conversions,
      metrics.conversions_value
    FROM keyword_view
    WHERE segments.date BETWEEN '${startDateStr}' AND '${endDateStr}'
  `;
  
  // Add status filter
  if (KEYWORD_STATUS === 'ENABLED') {
    query += `
      AND campaign.status = 'ENABLED'
      AND ad_group.status = 'ENABLED'
      AND ad_group_criterion.status = 'ENABLED'
    `;
  } else if (KEYWORD_STATUS === 'PAUSED') {
    query += `
      AND ad_group_criterion.status = 'PAUSED'
    `;
  }
  
  // Add campaign filter if specified
  if (CAMPAIGN_NAME_CONTAINS) {
    query += `
      AND campaign.name CONTAINS_IGNORE_CASE '${CAMPAIGN_NAME_CONTAINS}'
    `;
  }
  
  // Filter for minimum clicks
  query += `
    AND metrics.clicks >= ${minClicks}
  `;
  
  query += `
    ORDER BY metrics.cost_micros DESC
  `;
  
  Logger.log(`🔍 Executing custom range query (${daysAgoEnd}-${daysAgo} days ago)...`);
  const report = AdsApp.report(query);
  const rows = report.rows();
  
  while (rows.hasNext()) {
    const row = rows.next();
    
    const campaignId = row['campaign.id'];
    const campaignName = row['campaign.name'];
    const adGroupId = row['ad_group.id'];
    const adGroupName = row['ad_group.name'];
    const keywordText = row['ad_group_criterion.keyword.text'];
    const matchType = row['ad_group_criterion.keyword.match_type'];
    const keywordId = row['ad_group_criterion.criterion_id'];
    const impressions = parseInt(row['metrics.impressions']) || 0;
    const clicks = parseInt(row['metrics.clicks']) || 0;
    const costMicros = parseFloat(row['metrics.cost_micros']) || 0;
    const ctr = parseFloat(row['metrics.ctr']) || 0;
    const avgCpcMicros = parseFloat(row['metrics.average_cpc']) || 0;
    const conversions = parseFloat(row['metrics.conversions']) || 0;
    const conversionsValue = parseFloat(row['metrics.conversions_value']) || 0;
    
    const costAED = costMicros / 1000000;
    const cost = costAED / AED_TO_USD;
    const cpcAED = avgCpcMicros / 1000000;
    const cpc = cpcAED / AED_TO_USD;
    const cpa = conversions > 0 ? (cost / conversions) : 0;
    
    let matchTypeDisplay = matchType;
    if (matchType === 'EXACT') {
      matchTypeDisplay = 'Exact';
    } else if (matchType === 'PHRASE') {
      matchTypeDisplay = 'Phrase';
    } else if (matchType === 'BROAD') {
      matchTypeDisplay = 'Broad';
    }
    
    keywords.push({
      campaignId: campaignId,
      campaignName: campaignName,
      adGroupId: adGroupId,
      adGroupName: adGroupName,
      keywordText: keywordText,
      matchType: matchTypeDisplay,
      keywordId: keywordId,
      impressions: impressions,
      clicks: clicks,
      cost: cost,
      ctr: ctr * 100,
      cpc: cpc,
      conversions: conversions,
      conversionsValue: conversionsValue,
      cpa: cpa,
      uniqueKey: `${campaignId}_${adGroupId}_${keywordId}`
    });
  }
  
  return keywords;
}

// ============================================================================
// COMBINE ALL KEYWORDS FROM ALL PERIODS
// ============================================================================

function combineAllKeywords(keywords7Days, keywordsPrev7Days, keywords14Days, keywordsPrev14Days, keywords30Days, keywordsPrev30Days) {
  // Use a Map to track unique keywords and their data across periods
  const keywordMap = new Map();
  
  // Helper function to add keyword to map
  const addToMap = (kw, periodKey) => {
    if (keywordMap.has(kw.uniqueKey)) {
      const existing = keywordMap.get(kw.uniqueKey);
      existing[periodKey] = kw;
    } else {
      const newEntry = {
        campaignName: kw.campaignName,
        adGroupName: kw.adGroupName,
        keywordText: kw.keywordText,
        matchType: kw.matchType,
        uniqueKey: kw.uniqueKey,
        data7Days: null,
        dataPrev7Days: null,
        data14Days: null,
        dataPrev14Days: null,
        data30Days: null,
        dataPrev30Days: null
      };
      newEntry[periodKey] = kw;
      keywordMap.set(kw.uniqueKey, newEntry);
    }
  };
  
  // Process all periods
  keywords7Days.forEach(kw => addToMap(kw, 'data7Days'));
  keywordsPrev7Days.forEach(kw => addToMap(kw, 'dataPrev7Days'));
  keywords14Days.forEach(kw => addToMap(kw, 'data14Days'));
  keywordsPrev14Days.forEach(kw => addToMap(kw, 'dataPrev14Days'));
  keywords30Days.forEach(kw => addToMap(kw, 'data30Days'));
  keywordsPrev30Days.forEach(kw => addToMap(kw, 'dataPrev30Days'));
  
  // Convert map to array and calculate percentage changes for all metrics
  const allKeywords = Array.from(keywordMap.values()).map(kw => {
    // Calculate percentage changes for cost
    kw.costChange7d = calculatePercentageChange(
      kw.data7Days ? kw.data7Days.cost : 0,
      kw.dataPrev7Days ? kw.dataPrev7Days.cost : 0
    );
    kw.costChange14d = calculatePercentageChange(
      kw.data14Days ? kw.data14Days.cost : 0,
      kw.dataPrev14Days ? kw.dataPrev14Days.cost : 0
    );
    kw.costChange30d = calculatePercentageChange(
      kw.data30Days ? kw.data30Days.cost : 0,
      kw.dataPrev30Days ? kw.dataPrev30Days.cost : 0
    );
    
    // Calculate percentage changes for conversions
    kw.convChange7d = calculatePercentageChange(
      kw.data7Days ? kw.data7Days.conversions : 0,
      kw.dataPrev7Days ? kw.dataPrev7Days.conversions : 0
    );
    kw.convChange14d = calculatePercentageChange(
      kw.data14Days ? kw.data14Days.conversions : 0,
      kw.dataPrev14Days ? kw.dataPrev14Days.conversions : 0
    );
    kw.convChange30d = calculatePercentageChange(
      kw.data30Days ? kw.data30Days.conversions : 0,
      kw.dataPrev30Days ? kw.dataPrev30Days.conversions : 0
    );
    
    // Calculate percentage changes for CPA
    kw.cpaChange7d = calculatePercentageChange(
      kw.data7Days ? kw.data7Days.cpa : 0,
      kw.dataPrev7Days ? kw.dataPrev7Days.cpa : 0
    );
    kw.cpaChange14d = calculatePercentageChange(
      kw.data14Days ? kw.data14Days.cpa : 0,
      kw.dataPrev14Days ? kw.dataPrev14Days.cpa : 0
    );
    kw.cpaChange30d = calculatePercentageChange(
      kw.data30Days ? kw.data30Days.cpa : 0,
      kw.dataPrev30Days ? kw.dataPrev30Days.cpa : 0
    );
    
    return kw;
  });
  
  // Sort by 7-day cost (highest first)
  allKeywords.sort((a, b) => {
    const costA = a.data7Days ? a.data7Days.cost : (a.data14Days ? a.data14Days.cost : (a.data30Days ? a.data30Days.cost : 0));
    const costB = b.data7Days ? b.data7Days.cost : (b.data14Days ? b.data14Days.cost : (b.data30Days ? b.data30Days.cost : 0));
    return costB - costA;
  });
  
  return allKeywords;
}

function calculatePercentageChange(current, previous) {
  if (previous === 0) {
    return current > 0 ? 999 : 0;  // 999 means infinite increase
  }
  return ((current - previous) / previous) * 100;
}

// ============================================================================
// SUMMARY CALCULATIONS
// ============================================================================

function calculateSummary(keywordsData) {
  let totalClicks = 0;
  let totalCost = 0;
  let totalImpressions = 0;
  let totalConversions = 0;
  let totalConversionsValue = 0;
  
  for (const keyword of keywordsData) {
    totalClicks += keyword.clicks;
    totalCost += keyword.cost;
    totalImpressions += keyword.impressions;
    totalConversions += keyword.conversions;
    totalConversionsValue += keyword.conversionsValue;
  }
  
  const avgCPC = totalClicks > 0 ? (totalCost / totalClicks) : 0;
  const avgCTR = totalImpressions > 0 ? ((totalClicks / totalImpressions) * 100) : 0;
  const avgCPA = totalConversions > 0 ? (totalCost / totalConversions) : 0;
  
  return {
    totalKeywords: keywordsData.length,
    totalClicks: totalClicks,
    totalCost: totalCost,
    totalImpressions: totalImpressions,
    totalConversions: totalConversions,
    totalConversionsValue: totalConversionsValue,
    avgCPC: avgCPC,
    avgCTR: avgCTR,
    avgCPA: avgCPA
  };
}

// ============================================================================
// GMAIL EMAIL REPORT
// ============================================================================

function sendGmailReport(data) {
  if (EMAIL_RECIPIENTS.length === 0) {
    return;
  }
  
  try {
    const subject = `📊 Keywords Performance Report - ${data.allKeywords.length} Keywords`;
    
    // Limit keywords in email to top N (sorted by cost)
    const keywordsToShow = data.allKeywords.slice(0, MAX_KEYWORDS_IN_EMAIL);
    
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          body { font-family: Arial, Helvetica, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }
          .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
          .header { background-color: #4285f4; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
          .summary-box { background-color: #f8f9fa; border-left: 4px solid #4285f4; padding: 15px; margin: 20px 0; }
          .metric { display: inline-block; margin: 10px 20px 10px 0; text-align: center; }
          .metric-label { font-size: 11px; color: #666; display: block; text-transform: uppercase; }
          .metric-value { font-size: 24px; font-weight: bold; color: #4285f4; }
          table { border-collapse: collapse; width: 100%; margin: 20px 0; font-size: 11px; }
          th { background-color: #4285f4; color: white; padding: 10px 8px; text-align: left; font-weight: bold; }
          td { padding: 8px; border: 1px solid #ddd; }
          tr:nth-child(even) { background-color: #f8f9fa; }
          tr:hover { background-color: #e8f0fe; }
          .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 11px; color: #666; }
          .period-summary { background-color: #e8f0fe; padding: 10px; border-radius: 5px; margin: 10px 0; }
          .cost-cell { font-weight: bold; color: #d93025; }
          .conv-cell { font-weight: bold; color: #0f9d58; }
          .cpa-cell { color: #666; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2 style="margin: 0; color: white;">📊 Keywords Performance Report</h2>
            <p style="margin: 5px 0 0 0; color: white;">Cost, Conversions & CPA Analysis (7, 14, 30 Days)</p>
          </div>
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #4285f4;">📈 Summary by Period</h3>
            
            <div class="period-summary">
              <strong>📅 Last 7 Days</strong><br>
              <div class="metric">
                <span class="metric-label">Keywords</span>
                <span class="metric-value">${data.summary7Days.totalKeywords}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Cost</span>
                <span class="metric-value">$${data.summary7Days.totalCost.toFixed(2)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Conversions</span>
                <span class="metric-value">${data.summary7Days.totalConversions.toFixed(1)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Avg CPA</span>
                <span class="metric-value">$${data.summary7Days.avgCPA.toFixed(2)}</span>
              </div>
            </div>
            
            <div class="period-summary">
              <strong>📅 Last 14 Days</strong><br>
              <div class="metric">
                <span class="metric-label">Keywords</span>
                <span class="metric-value">${data.summary14Days.totalKeywords}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Cost</span>
                <span class="metric-value">$${data.summary14Days.totalCost.toFixed(2)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Conversions</span>
                <span class="metric-value">${data.summary14Days.totalConversions.toFixed(1)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Avg CPA</span>
                <span class="metric-value">$${data.summary14Days.avgCPA.toFixed(2)}</span>
              </div>
            </div>
            
            <div class="period-summary">
              <strong>📅 Last 30 Days</strong><br>
              <div class="metric">
                <span class="metric-label">Keywords</span>
                <span class="metric-value">${data.summary30Days.totalKeywords}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Cost</span>
                <span class="metric-value">$${data.summary30Days.totalCost.toFixed(2)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Conversions</span>
                <span class="metric-value">${data.summary30Days.totalConversions.toFixed(1)}</span>
              </div>
              <div class="metric">
                <span class="metric-label">Avg CPA</span>
                <span class="metric-value">$${data.summary30Days.avgCPA.toFixed(2)}</span>
              </div>
            </div>
          </div>
          
          <h3 style="color: #4285f4;">📋 Keywords Performance Details</h3>
          <p style="font-size: 13px; color: #666;">
            Showing top ${keywordsToShow.length} keywords sorted by highest cost (last 7 days). 
            ${data.allKeywords.length > MAX_KEYWORDS_IN_EMAIL ? `Total found: ${data.allKeywords.length} keywords.` : ''}
          </p>
          <div style="overflow-x: auto;">
            <table>
              <thead>
                <tr>
                  <th>Campaign</th>
                  <th>Ad Group</th>
                  <th>Keyword</th>
                  <th>Match Type</th>
                  <th>Cost (7d)</th>
                  <th>Conv (7d)</th>
                  <th>CPA (7d)</th>
                  <th>Cost (14d)</th>
                  <th>Conv (14d)</th>
                  <th>CPA (14d)</th>
                  <th>Cost (30d)</th>
                  <th>Conv (30d)</th>
                  <th>CPA (30d)</th>
                </tr>
              </thead>
              <tbody>
                ${keywordsToShow.map(kw => {
                  // Helper function to format value with percentage change
                  const formatWithChange = (value, change, isGoodWhenNegative = true, prefix = '', suffix = '') => {
                    if (!value && value !== 0) return '-';
                    
                    let changeStr = '';
                    if (change !== null && change !== undefined && change !== 0) {
                      if (change === 999) {
                        changeStr = ' <span style="color: #0f9d58; font-size: 10px;">(NEW)</span>';
                      } else {
                        // For cost and CPA: negative is good (green), positive is bad (red)
                        // For conversions: positive is good (green), negative is bad (red)
                        const color = isGoodWhenNegative 
                          ? (change > 0 ? '#d93025' : '#0f9d58')
                          : (change > 0 ? '#0f9d58' : '#d93025');
                        const sign = change > 0 ? '+' : '';
                        changeStr = ` <span style="color: ${color}; font-size: 10px;">(${sign}${change.toFixed(0)}%)</span>`;
                      }
                    }
                    
                    return `${prefix}${value}${suffix}${changeStr}`;
                  };
                  
                  // Format Cost with change
                  const cost7d = kw.data7Days 
                    ? formatWithChange(kw.data7Days.cost.toFixed(2), kw.costChange7d, true, '$', '')
                    : '-';
                  const cost14d = kw.data14Days 
                    ? formatWithChange(kw.data14Days.cost.toFixed(2), kw.costChange14d, true, '$', '')
                    : '-';
                  const cost30d = kw.data30Days 
                    ? formatWithChange(kw.data30Days.cost.toFixed(2), kw.costChange30d, true, '$', '')
                    : '-';
                  
                  // Format Conversions with change
                  const conv7d = kw.data7Days 
                    ? formatWithChange(kw.data7Days.conversions.toFixed(1), kw.convChange7d, false, '', '')
                    : '-';
                  const conv14d = kw.data14Days 
                    ? formatWithChange(kw.data14Days.conversions.toFixed(1), kw.convChange14d, false, '', '')
                    : '-';
                  const conv30d = kw.data30Days 
                    ? formatWithChange(kw.data30Days.conversions.toFixed(1), kw.convChange30d, false, '', '')
                    : '-';
                  
                  // Format CPA with change
                  const cpa7d = kw.data7Days && kw.data7Days.conversions > 0
                    ? formatWithChange(kw.data7Days.cpa.toFixed(2), kw.cpaChange7d, true, '$', '')
                    : (kw.data7Days ? 'N/A' : '-');
                  const cpa14d = kw.data14Days && kw.data14Days.conversions > 0
                    ? formatWithChange(kw.data14Days.cpa.toFixed(2), kw.cpaChange14d, true, '$', '')
                    : (kw.data14Days ? 'N/A' : '-');
                  const cpa30d = kw.data30Days && kw.data30Days.conversions > 0
                    ? formatWithChange(kw.data30Days.cpa.toFixed(2), kw.cpaChange30d, true, '$', '')
                    : (kw.data30Days ? 'N/A' : '-');
                  
                  return `
                  <tr>
                    <td>${kw.campaignName}</td>
                    <td>${kw.adGroupName}</td>
                    <td>${kw.keywordText}</td>
                    <td>${kw.matchType}</td>
                    <td class="cost-cell">${cost7d}</td>
                    <td class="conv-cell">${conv7d}</td>
                    <td class="cpa-cell">${cpa7d}</td>
                    <td class="cost-cell">${cost14d}</td>
                    <td class="conv-cell">${conv14d}</td>
                    <td class="cpa-cell">${cpa14d}</td>
                    <td class="cost-cell">${cost30d}</td>
                    <td class="conv-cell">${conv30d}</td>
                    <td class="cpa-cell">${cpa30d}</td>
                  </tr>
                `;
                }).join('')}
              </tbody>
            </table>
          </div>
          ${data.allKeywords.length > MAX_KEYWORDS_IN_EMAIL ? `
          <p style="font-size: 12px; color: #666; font-style: italic; text-align: center;">
            ... and ${data.allKeywords.length - MAX_KEYWORDS_IN_EMAIL} more keywords
          </p>
          ` : ''}
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #4285f4;">💡 How to Use This Report</h3>
            <ul style="margin: 10px 0; padding-left: 20px;">
              <li><strong>Red (Cost):</strong> Identify high-cost keywords across periods</li>
              <li><strong>Green (Conversions):</strong> See which keywords are converting</li>
              <li><strong>CPA:</strong> Compare cost per acquisition across time periods</li>
              <li><strong>N/A CPA:</strong> Keywords with 0 conversions (consider pausing)</li>
              <li><strong>Sorted by Cost:</strong> Highest cost keywords appear first</li>
            </ul>
          </div>
          
          <div class="footer">
            <p><strong>💰 Note:</strong> All costs converted from AED to USD (rate: ${AED_TO_USD})</p>
            <hr style="margin: 15px 0; border: none; border-top: 1px solid #ddd;">
            <p style="margin: 5px 0;">
              <strong>Report Details:</strong><br>
              Generated by: Google Ads Scripts<br>
              Report Type: Keywords Performance (7, 14, 30 Days)<br>
              Generated on: ${new Date().toLocaleString('en-US', { 
                weekday: 'long', 
                year: 'numeric', 
                month: 'long', 
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                timeZoneName: 'short'
              })}<br>
              Recipients: ${EMAIL_RECIPIENTS.length} Gmail account(s)
            </p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    // Send email to Gmail
    MailApp.sendEmail({
      to: EMAIL_RECIPIENTS.join(','),
      subject: subject,
      htmlBody: htmlBody,
      name: 'Google Ads Performance Reports'
    });
    
    Logger.log(`✅ Gmail report sent successfully`);
    
  } catch (error) {
    Logger.log(`❌ Error sending Gmail: ${error.message}`);
    Logger.log(error.stack);
  }
}

// ============================================================================
// DISPLAY SUMMARY
// ============================================================================

function displaySummary(summary7Days, summary14Days, summary30Days, totalUniqueKeywords) {
  Logger.log('\n📊 FINAL SUMMARY:');
  Logger.log('='.repeat(70));
  Logger.log('Last 7 Days:');
  Logger.log(`  Keywords: ${summary7Days.totalKeywords.toLocaleString()}`);
  Logger.log(`  Cost: $${summary7Days.totalCost.toFixed(2)}`);
  Logger.log(`  Conversions: ${summary7Days.totalConversions.toFixed(1)}`);
  Logger.log(`  Avg CPA: $${summary7Days.avgCPA.toFixed(2)}`);
  Logger.log('');
  Logger.log('Last 14 Days:');
  Logger.log(`  Keywords: ${summary14Days.totalKeywords.toLocaleString()}`);
  Logger.log(`  Cost: $${summary14Days.totalCost.toFixed(2)}`);
  Logger.log(`  Conversions: ${summary14Days.totalConversions.toFixed(1)}`);
  Logger.log(`  Avg CPA: $${summary14Days.avgCPA.toFixed(2)}`);
  Logger.log('');
  Logger.log('Last 30 Days:');
  Logger.log(`  Keywords: ${summary30Days.totalKeywords.toLocaleString()}`);
  Logger.log(`  Cost: $${summary30Days.totalCost.toFixed(2)}`);
  Logger.log(`  Conversions: ${summary30Days.totalConversions.toFixed(1)}`);
  Logger.log(`  Avg CPA: $${summary30Days.avgCPA.toFixed(2)}`);
  Logger.log('');
  Logger.log(`📋 Total Unique Keywords: ${totalUniqueKeywords}`);
  Logger.log('='.repeat(70));
}

// ============================================================================
// END OF SCRIPT
// ============================================================================

