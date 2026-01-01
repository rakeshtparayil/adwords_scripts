/**
 * Quality Score Report - Keyword Level Analysis
 * 
 * This script generates a detailed Quality Score report showing:
 * - Quality Score
 * - Expected Click-Through Rate (CTR)
 * - Ad Relevance
 * - Landing Page Experience
 * - Ad Rank
 * - Campaign and Ad Group information
 * 
 * Time Period: Last 7 Days
 * 
 * Features:
 * - Keyword-level quality metrics
 * - Gmail email notification with detailed table
 * - Color-coded quality indicators
 * - Sorted by Quality Score (lowest first for easy optimization)
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
  'test@gmail.com',        // Replace with your Gmail
   'test2@gmail.com'            // Add more recipients as needed
];

// Time period for performance metrics
const DATE_RANGE = 'LAST_7_DAYS';

// Minimum impressions to include (filter out low-traffic keywords)
const MIN_IMPRESSIONS = 10;

// Campaign filter (optional - leave empty to include all campaigns)
const CAMPAIGN_NAME_CONTAINS = '';  // Example: 'Brand' to only include brand campaigns

// Keyword status filter
const KEYWORD_STATUS = 'ENABLED';  // Options: 'ENABLED', 'PAUSED', 'ALL'

// Number of keywords to show in email
const MAX_KEYWORDS_IN_EMAIL = 150;

// Exchange rate: AED to USD (set to 1 if you don't need conversion)
const AED_TO_USD = 3.67;

// ============================================================================
// MAIN FUNCTION
// ============================================================================

function main() {
  Logger.log('='.repeat(70));
  Logger.log('🚀 Starting Quality Score Report (Last 7 Days)');
  Logger.log('='.repeat(70));
  
  try {
    // Validate configuration
    if (EMAIL_RECIPIENTS.length === 0) {
      throw new Error('Please configure EMAIL_RECIPIENTS in the script settings');
    }
    
    Logger.log(`🎯 Date Range: ${DATE_RANGE}`);
    Logger.log(`🎯 Keyword Status: ${KEYWORD_STATUS}`);
    Logger.log(`🎯 Minimum Impressions: ${MIN_IMPRESSIONS}`);
    if (CAMPAIGN_NAME_CONTAINS) {
      Logger.log(`🎯 Campaign Filter: ${CAMPAIGN_NAME_CONTAINS}`);
    }
    
    // Fetch quality score data
    Logger.log('\n' + '='.repeat(70));
    Logger.log('📊 FETCHING QUALITY SCORE DATA');
    Logger.log('='.repeat(70));
    const qualityScoreData = fetchQualityScoreData();
    Logger.log(`✅ Found ${qualityScoreData.length} keywords with quality score data`);
    
    if (qualityScoreData.length === 0) {
      Logger.log('⚠️  No keywords found with the specified criteria');
      return;
    }
    
    // Calculate summary statistics
    const summary = calculateSummary(qualityScoreData);
    
    // Send Gmail report
    Logger.log('\n' + '='.repeat(70));
    Logger.log('📧 Sending Gmail report...');
    Logger.log('='.repeat(70));
    sendGmailReport({
      keywords: qualityScoreData,
      summary: summary
    });
    Logger.log(`✅ Email sent to: ${EMAIL_RECIPIENTS.join(', ')}`);
    
    Logger.log('\n' + '='.repeat(70));
    Logger.log('✅ Report generation completed successfully!');
    Logger.log('='.repeat(70));
    
    // Display summary
    displaySummary(summary, qualityScoreData.length);
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(error.stack);
    throw error;
  }
}

// ============================================================================
// DATA FETCHING
// ============================================================================

function fetchQualityScoreData() {
  const keywords = [];
  
  // Build the query for quality score data
  let query = `
    SELECT
      campaign.id,
      campaign.name,
      ad_group.id,
      ad_group.name,
      ad_group_criterion.keyword.text,
      ad_group_criterion.keyword.match_type,
      ad_group_criterion.criterion_id,
      ad_group_criterion.quality_info.quality_score,
      ad_group_criterion.quality_info.creative_quality_score,
      ad_group_criterion.quality_info.post_click_quality_score,
      ad_group_criterion.quality_info.search_predicted_ctr,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.conversions,
      metrics.ctr,
      metrics.average_cpc,
      metrics.search_absolute_top_impression_share,
      metrics.search_top_impression_share,
      segments.ad_network_type
    FROM keyword_view
    WHERE segments.date DURING ${DATE_RANGE}
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
  
  // Filter for minimum impressions
  query += `
    AND metrics.impressions >= ${MIN_IMPRESSIONS}
    AND segments.ad_network_type = 'SEARCH'
  `;
  
  // Order by quality score (lowest first for optimization priority)
  query += `
    ORDER BY ad_group_criterion.quality_info.quality_score ASC
  `;
  
  Logger.log('🔍 Executing quality score query...');
  const report = AdsApp.report(query);
  const rows = report.rows();
  
  // Track unique keywords (aggregate across ad network types if needed)
  const keywordMap = new Map();
  
  while (rows.hasNext()) {
    const row = rows.next();
    
    const campaignId = row['campaign.id'];
    const campaignName = row['campaign.name'];
    const adGroupId = row['ad_group.id'];
    const adGroupName = row['ad_group.name'];
    const keywordText = row['ad_group_criterion.keyword.text'];
    const matchType = row['ad_group_criterion.keyword.match_type'];
    const keywordId = row['ad_group_criterion.criterion_id'];
    
    // Quality Score metrics
    const qualityScore = parseInt(row['ad_group_criterion.quality_info.quality_score']) || null;
    const adRelevance = row['ad_group_criterion.quality_info.creative_quality_score'] || 'UNSPECIFIED';
    const landingPageExp = row['ad_group_criterion.quality_info.post_click_quality_score'] || 'UNSPECIFIED';
    const expectedCtr = row['ad_group_criterion.quality_info.search_predicted_ctr'] || 'UNSPECIFIED';
    
    // Performance metrics
    const impressions = parseInt(row['metrics.impressions']) || 0;
    const clicks = parseInt(row['metrics.clicks']) || 0;
    const costMicros = parseFloat(row['metrics.cost_micros']) || 0;
    const conversions = parseFloat(row['metrics.conversions']) || 0;
    const ctr = parseFloat(row['metrics.ctr']) || 0;
    const avgCpcMicros = parseFloat(row['metrics.average_cpc']) || 0;
    const absTopImprShare = parseFloat(row['metrics.search_absolute_top_impression_share']) || 0;
    const topImprShare = parseFloat(row['metrics.search_top_impression_share']) || 0;
    
    // Convert cost from micros to currency, then AED to USD
    const costAED = costMicros / 1000000;
    const cost = costAED / AED_TO_USD;
    
    // Convert CPC from micros
    const cpcAED = avgCpcMicros / 1000000;
    const cpc = cpcAED / AED_TO_USD;
    
    // Format match type for display
    let matchTypeDisplay = matchType;
    if (matchType === 'EXACT') {
      matchTypeDisplay = 'Exact';
    } else if (matchType === 'PHRASE') {
      matchTypeDisplay = 'Phrase';
    } else if (matchType === 'BROAD') {
      matchTypeDisplay = 'Broad';
    }
    
    // Create unique key
    const uniqueKey = `${campaignId}_${adGroupId}_${keywordId}`;
    
    // Aggregate data if keyword already exists
    if (keywordMap.has(uniqueKey)) {
      const existing = keywordMap.get(uniqueKey);
      existing.impressions += impressions;
      existing.clicks += clicks;
      existing.cost += cost;
      existing.conversions += conversions;
      // Recalculate averages
      existing.ctr = existing.impressions > 0 ? (existing.clicks / existing.impressions) * 100 : 0;
      existing.cpc = existing.clicks > 0 ? (existing.cost / existing.clicks) : 0;
      // Average the impression share metrics
      existing.absTopImprShare = (existing.absTopImprShare + absTopImprShare) / 2;
      existing.topImprShare = (existing.topImprShare + topImprShare) / 2;
    } else {
      keywordMap.set(uniqueKey, {
        campaignId: campaignId,
        campaignName: campaignName,
        adGroupId: adGroupId,
        adGroupName: adGroupName,
        keywordText: keywordText,
        matchType: matchTypeDisplay,
        keywordId: keywordId,
        qualityScore: qualityScore,
        expectedCtr: formatQualityComponent(expectedCtr),
        adRelevance: formatQualityComponent(adRelevance),
        landingPageExp: formatQualityComponent(landingPageExp),
        impressions: impressions,
        clicks: clicks,
        cost: cost,
        ctr: ctr * 100,
        cpc: cpc,
        conversions: conversions,
        absTopImprShare: absTopImprShare,
        topImprShare: topImprShare
      });
    }
  }
  
  // Convert map to array
  const keywordsArray = Array.from(keywordMap.values());
  
  // Sort by cost (highest first)
  keywordsArray.sort((a, b) => {
    return b.cost - a.cost;
  });
  
  return keywordsArray;
}

function formatQualityComponent(value) {
  // Convert Google Ads quality component values to readable format
  const mapping = {
    'ABOVE_AVERAGE': 'Above Average',
    'AVERAGE': 'Average',
    'BELOW_AVERAGE': 'Below Average',
    'UNSPECIFIED': 'N/A',
    'UNKNOWN': 'N/A'
  };
  return mapping[value] || value;
}

// ============================================================================
// SUMMARY CALCULATIONS
// ============================================================================

function calculateSummary(keywordsData) {
  let totalKeywords = keywordsData.length;
  let totalImpressions = 0;
  let totalClicks = 0;
  let totalCost = 0;
  let totalConversions = 0;
  
  // Quality score distribution
  let qsDistribution = {
    '1-3': 0,
    '4-6': 0,
    '7-10': 0,
    'null': 0
  };
  
  let sumQualityScore = 0;
  let countWithQS = 0;
  
  for (const keyword of keywordsData) {
    totalImpressions += keyword.impressions;
    totalClicks += keyword.clicks;
    totalCost += keyword.cost;
    totalConversions += keyword.conversions;
    
    if (keyword.qualityScore !== null) {
      sumQualityScore += keyword.qualityScore;
      countWithQS++;
      
      if (keyword.qualityScore <= 3) {
        qsDistribution['1-3']++;
      } else if (keyword.qualityScore <= 6) {
        qsDistribution['4-6']++;
      } else {
        qsDistribution['7-10']++;
      }
    } else {
      qsDistribution['null']++;
    }
  }
  
  const avgQualityScore = countWithQS > 0 ? (sumQualityScore / countWithQS) : 0;
  const avgCTR = totalImpressions > 0 ? ((totalClicks / totalImpressions) * 100) : 0;
  const avgCPC = totalClicks > 0 ? (totalCost / totalClicks) : 0;
  
  return {
    totalKeywords: totalKeywords,
    totalImpressions: totalImpressions,
    totalClicks: totalClicks,
    totalCost: totalCost,
    totalConversions: totalConversions,
    avgQualityScore: avgQualityScore,
    avgCTR: avgCTR,
    avgCPC: avgCPC,
    qsDistribution: qsDistribution
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
    const subject = `📊 Quality Score Report - ${data.keywords.length} Keywords (Last 7 Days)`;
    
    // Limit keywords in email
    const keywordsToShow = data.keywords.slice(0, MAX_KEYWORDS_IN_EMAIL);
    
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
          th { background-color: #4285f4; color: white; padding: 10px 6px; text-align: left; font-weight: bold; }
          td { padding: 8px 6px; border: 1px solid #ddd; }
          tr:nth-child(even) { background-color: #f8f9fa; }
          tr:hover { background-color: #e8f0fe; }
          .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 11px; color: #666; }
          .qs-excellent { background-color: #c8e6c9; font-weight: bold; }
          .qs-good { background-color: #fff9c4; }
          .qs-poor { background-color: #ffcdd2; font-weight: bold; }
          .qs-null { background-color: #e0e0e0; }
          .quality-above { color: #0f9d58; font-weight: bold; }
          .quality-average { color: #f9ab00; }
          .quality-below { color: #d93025; font-weight: bold; }
          .quality-na { color: #999; }
          .distribution { display: inline-block; margin: 10px 15px; padding: 10px; background: #e8f0fe; border-radius: 5px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2 style="margin: 0; color: white;">📊 Quality Score Report</h2>
            <p style="margin: 5px 0 0 0; color: white;">Keyword-Level Quality Analysis (Last 7 Days)</p>
          </div>
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #4285f4;">📖 Quality Score Components Explained</h3>
            <table style="width: 100%; margin: 10px 0; font-size: 12px; border: none;">
              <tr>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Quality Score (1-10)</strong><br>
                  <span style="font-size: 11px;">Google's rating of the quality and relevance of your keywords, ads, and landing pages. Higher scores can lead to lower costs and better ad positions.</span>
                </td>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Expected CTR</strong><br>
                  <span style="font-size: 11px;">How likely your ad will be clicked when shown for this keyword, compared to other advertisers. Based on your ad's past performance.</span>
                </td>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Ad Relevance</strong><br>
                  <span style="font-size: 11px;">How closely your ad matches the intent behind a user's search. Measures how well your ad copy relates to the keyword.</span>
                </td>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Landing Page Experience</strong><br>
                  <span style="font-size: 11px;">How relevant and useful your landing page is to people who click your ad. Considers page load speed, mobile-friendliness, and content relevance.</span>
                </td>
              </tr>
              <tr>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Abs. Top Impr. %</strong><br>
                  <span style="font-size: 11px;">Percentage of impressions shown in the absolute top position (first ad above organic results). Higher is better for visibility.</span>
                </td>
                <td style="border: none; padding: 8px; vertical-align: top; width: 25%;">
                  <strong style="color: #4285f4;">Top Impr. %</strong><br>
                  <span style="font-size: 11px;">Percentage of impressions shown anywhere above organic search results. Indicates how often you appear in premium positions.</span>
                </td>
              </tr>
            </table>
            
            <div style="background: #e8f0fe; padding: 10px; border-radius: 5px; margin: 10px 0;">
              <strong>Rating Scale:</strong> 
              <span style="color: #0f9d58; font-weight: bold;">Above Average</span> | 
              <span style="color: #f9ab00;">Average</span> | 
              <span style="color: #d93025; font-weight: bold;">Below Average</span>
            </div>
          </div>
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #4285f4;">📈 Summary Statistics</h3>
            <div class="metric">
              <span class="metric-label">Total Keywords</span>
              <span class="metric-value">${data.summary.totalKeywords}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Avg Quality Score</span>
              <span class="metric-value">${data.summary.avgQualityScore.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Total Cost</span>
              <span class="metric-value">$${data.summary.totalCost.toFixed(2)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Conversions</span>
              <span class="metric-value">${data.summary.totalConversions.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Avg CTR</span>
              <span class="metric-value">${data.summary.avgCTR.toFixed(2)}%</span>
            </div>
            
            <h4 style="margin: 20px 0 10px 0; color: #4285f4;">Quality Score Distribution:</h4>
            <div class="distribution">
              <strong>Poor (1-3):</strong> ${data.summary.qsDistribution['1-3']} keywords
            </div>
            <div class="distribution">
              <strong>Average (4-6):</strong> ${data.summary.qsDistribution['4-6']} keywords
            </div>
            <div class="distribution">
              <strong>Good (7-10):</strong> ${data.summary.qsDistribution['7-10']} keywords
            </div>
            <div class="distribution">
              <strong>No Data:</strong> ${data.summary.qsDistribution['null']} keywords
            </div>
          </div>
          
          <h3 style="color: #4285f4;">📋 Keyword Quality Score Details</h3>
          <p style="font-size: 13px; color: #666;">
            Showing ${keywordsToShow.length} keywords sorted by highest cost (last 7 days).
            ${data.keywords.length > MAX_KEYWORDS_IN_EMAIL ? `Total found: ${data.keywords.length} keywords.` : ''}
          </p>
          <div style="overflow-x: auto;">
            <table>
              <thead>
                <tr>
                  <th>Campaign</th>
                  <th>Ad Group</th>
                  <th>Keyword</th>
                  <th>Match Type</th>
                  <th>Quality Score</th>
                  <th>Expected CTR</th>
                  <th>Ad Relevance</th>
                  <th>Landing Page Exp</th>
                  <th>Abs Top %</th>
                  <th>Top %</th>
                  <th>Impressions</th>
                  <th>Clicks</th>
                  <th>CTR</th>
                  <th>Cost</th>
                  <th>Conv</th>
                </tr>
              </thead>
              <tbody>
                ${keywordsToShow.map(kw => {
                  // Determine QS cell class
                  let qsClass = 'qs-null';
                  if (kw.qualityScore !== null) {
                    if (kw.qualityScore >= 7) qsClass = 'qs-excellent';
                    else if (kw.qualityScore >= 4) qsClass = 'qs-good';
                    else qsClass = 'qs-poor';
                  }
                  
                  // Format quality components with colors
                  const formatQuality = (value) => {
                    if (value === 'Above Average') return `<span class="quality-above">${value}</span>`;
                    if (value === 'Average') return `<span class="quality-average">${value}</span>`;
                    if (value === 'Below Average') return `<span class="quality-below">${value}</span>`;
                    return `<span class="quality-na">${value}</span>`;
                  };
                  
                  return `
                  <tr>
                    <td>${kw.campaignName}</td>
                    <td>${kw.adGroupName}</td>
                    <td>${kw.keywordText}</td>
                    <td>${kw.matchType}</td>
                    <td class="${qsClass}">${kw.qualityScore !== null ? kw.qualityScore : 'N/A'}</td>
                    <td>${formatQuality(kw.expectedCtr)}</td>
                    <td>${formatQuality(kw.adRelevance)}</td>
                    <td>${formatQuality(kw.landingPageExp)}</td>
                    <td>${(kw.absTopImprShare * 100).toFixed(1)}%</td>
                    <td>${(kw.topImprShare * 100).toFixed(1)}%</td>
                    <td>${kw.impressions.toLocaleString()}</td>
                    <td>${kw.clicks.toLocaleString()}</td>
                    <td>${kw.ctr.toFixed(2)}%</td>
                    <td>$${kw.cost.toFixed(2)}</td>
                    <td>${kw.conversions.toFixed(1)}</td>
                  </tr>
                `;
                }).join('')}
              </tbody>
            </table>
          </div>
          ${data.keywords.length > MAX_KEYWORDS_IN_EMAIL ? `
          <p style="font-size: 12px; color: #666; font-style: italic; text-align: center;">
            ... and ${data.keywords.length - MAX_KEYWORDS_IN_EMAIL} more keywords
          </p>
          ` : ''}
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #4285f4;">💡 How to Read This Report</h3>
            <ul style="margin: 10px 0; padding-left: 20px;">
              <li><strong>Quality Score:</strong> 1-10 scale (7-10 = Good, 4-6 = Average, 1-3 = Poor)</li>
              <li><strong>Expected CTR:</strong> How likely your ad will be clicked</li>
              <li><strong>Ad Relevance:</strong> How closely your ad matches the keyword</li>
              <li><strong>Landing Page Experience:</strong> Quality and relevance of your landing page</li>
              <li><strong>Abs Top %:</strong> % of impressions in absolute top position (first ad)</li>
              <li><strong>Top %:</strong> % of impressions anywhere above organic results</li>
              <li><strong>Green:</strong> Above Average | <strong>Yellow:</strong> Average | <strong>Red:</strong> Below Average</li>
            </ul>
            
            <h3 style="margin-top: 20px; color: #4285f4;">🎯 Optimization Priority</h3>
            <ol style="margin: 10px 0; padding-left: 20px;">
              <li><strong>Focus on QS 1-3 (Poor):</strong> These need immediate attention</li>
              <li><strong>Improve "Below Average" components:</strong> Start with easiest wins</li>
              <li><strong>Expected CTR:</strong> Improve ad copy and test new variations</li>
              <li><strong>Ad Relevance:</strong> Ensure keywords match ad copy closely</li>
              <li><strong>Landing Page:</strong> Improve page speed, relevance, and user experience</li>
            </ol>
          </div>
          
          <div class="footer">
            <p><strong>💰 Note:</strong> All costs converted from AED to USD (rate: ${AED_TO_USD})</p>
            <p><strong>📅 Date Range:</strong> ${DATE_RANGE}</p>
            <hr style="margin: 15px 0; border: none; border-top: 1px solid #ddd;">
            <p style="margin: 5px 0;">
              <strong>Report Details:</strong><br>
              Generated by: Google Ads Scripts<br>
              Report Type: Quality Score Analysis<br>
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

function displaySummary(summary, totalKeywords) {
  Logger.log('\n📊 FINAL SUMMARY:');
  Logger.log('='.repeat(70));
  Logger.log(`Total Keywords: ${totalKeywords}`);
  Logger.log(`Average Quality Score: ${summary.avgQualityScore.toFixed(1)}`);
  Logger.log(`Total Cost: $${summary.totalCost.toFixed(2)}`);
  Logger.log(`Total Conversions: ${summary.totalConversions.toFixed(1)}`);
  Logger.log(`Average CTR: ${summary.avgCTR.toFixed(2)}%`);
  Logger.log('');
  Logger.log('Quality Score Distribution:');
  Logger.log(`  Poor (1-3): ${summary.qsDistribution['1-3']} keywords`);
  Logger.log(`  Average (4-6): ${summary.qsDistribution['4-6']} keywords`);
  Logger.log(`  Good (7-10): ${summary.qsDistribution['7-10']} keywords`);
  Logger.log(`  No Data: ${summary.qsDistribution['null']} keywords`);
  Logger.log('='.repeat(70));
}

// ============================================================================
// END OF SCRIPT
// ============================================================================

