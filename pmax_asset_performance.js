/**
 * Performance Max Asset Analysis Script
 * 
 * This script analyzes assets in Performance Max campaigns using a two-step approach:
 * 1. First gathering asset data and field types
 * 2. Then attempting to get performance data from ad_group_asset_view resource
 */

function main() {
  // Configuration
  const SPREADSHEET_URL = "YOUR SPREADSHEET UEL"; // Replace with your Google Sheet URL
  const DATE_RANGE = "LAST_30_DAYS";
  
  // Initialize spreadsheet
  let spreadsheet;
  if (SPREADSHEET_URL) {
    spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  } else {
    spreadsheet = SpreadsheetApp.create("PMax Asset Performance " + new Date().toISOString().split('T')[0]);
    Logger.log("Created new spreadsheet: " + spreadsheet.getUrl());
  }
  
  // Clear existing sheets except the first one
  const sheets = spreadsheet.getSheets();
  for (let i = 1; i < sheets.length; i++) {
    spreadsheet.deleteSheet(sheets[i]);
  }
  
  // Rename the first sheet to Dashboard
  const dashboardSheet = sheets[0];
  dashboardSheet.setName("Dashboard");
  
  // Get all Performance Max campaigns
  const pmaxCampaigns = getPMaxCampaigns();
  if (pmaxCampaigns.length === 0) {
    Logger.log("No Performance Max campaigns found in this account.");
    return;
  }
  
  Logger.log(`Found ${pmaxCampaigns.length} Performance Max campaigns.`);
  
  // Process assets
  processAssetsWithPerformanceData(pmaxCampaigns, spreadsheet, DATE_RANGE);
  
  // Process asset groups
  processAssetGroups(pmaxCampaigns, spreadsheet, DATE_RANGE);
  
  // Create dashboard
  createDashboard(spreadsheet, pmaxCampaigns);
  
  Logger.log("Analysis complete! Check your spreadsheet for results.");
}

/**
 * Gets all Performance Max campaigns in the account
 */
function getPMaxCampaigns() {
  const pmaxCampaigns = [];
  
  // Directly get Performance Max campaigns
  const performanceMaxCampaignIterator = AdsApp.performanceMaxCampaigns()
    .withCondition("campaign.status IN ('ENABLED', 'PAUSED')")
    .get();
  
  Logger.log(`Total Performance Max campaigns found: ${performanceMaxCampaignIterator.totalNumEntities()}`);
  
  // Add each campaign to our array
  while (performanceMaxCampaignIterator.hasNext()) {
    const campaign = performanceMaxCampaignIterator.next();
    pmaxCampaigns.push(campaign);
  }
  
  return pmaxCampaigns;
}

/**
 * Process assets and attempt to get performance data using ad_group_asset_view
 */
function processAssetsWithPerformanceData(pmaxCampaigns, spreadsheet, dateRange) {
  const sheet = spreadsheet.insertSheet("Asset Performance");
  
  // Set up headers
  const headers = [
    "Campaign", "Asset Group", "Field Type", "Performance Label", 
    "Text Content", "Impressions", "Clicks", "CTR", "Conversions", "Cost"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  let row = 2;
  
  // Process each campaign
  for (const campaign of pmaxCampaigns) {
    const campaignName = campaign.getName();
    const campaignId = campaign.getId();
    
    Logger.log(`Processing assets for campaign: ${campaignName}`);
    
    // Step 1: Get basic asset data
    try {
      const assetReport = AdsApp.report(
        "SELECT " +
        "  campaign.name, " +
        "  asset_group.name, " +
        "  asset_group_asset.field_type, " +
        "  asset_group_asset.performance_label, " +
        "  asset.text_asset.text " +
        "FROM asset_group_asset " +
        "WHERE campaign.id = " + campaignId
      );
      
      const assets = [];
      const assetRows = assetReport.rows();
      let assetCount = 0;
      
      while (assetRows.hasNext()) {
        const rowData = assetRows.next();
        assetCount++;
        
        let textContent = "";
        try {
          textContent = rowData['asset.text_asset.text'] || "";
        } catch (e) {
          // Text content might not exist
        }
        
        assets.push({
          campaignName: rowData['campaign.name'],
          assetGroupName: rowData['asset_group.name'],
          fieldType: rowData['asset_group_asset.field_type'] || "",
          performanceLabel: rowData['asset_group_asset.performance_label'] || "N/A",
          textContent: textContent
        });
      }
      
      Logger.log(`Found ${assetCount} assets for campaign ${campaignName}`);
      
      // Step 2: Try to get performance data using ad_group_asset_view (may not work depending on account access)
      try {
        const perfReport = AdsApp.report(
          "SELECT " +
          "  campaign.name, " +
          "  asset_group.name, " +
          "  ad_group_asset_view.field_type, " +
          "  ad_group_asset_view.performance_label, " +
          "  metrics.impressions, " +
          "  metrics.clicks, " +
          "  metrics.ctr, " +
          "  metrics.conversions, " +
          "  metrics.cost_micros " +
          "FROM ad_group_asset_view " +
          "WHERE campaign.id = " + campaignId + " " +
          "AND segments.date DURING " + dateRange
        );
        
        // Create a map for performance data by asset group and field type
        const perfDataMap = new Map();
        const perfRows = perfReport.rows();
        
        while (perfRows.hasNext()) {
          const rowData = perfRows.next();
          
          const key = rowData['campaign.name'] + '|' + 
                      rowData['asset_group.name'] + '|' + 
                      rowData['ad_group_asset_view.field_type'];
          
          perfDataMap.set(key, {
            impressions: parseInt(rowData['metrics.impressions']) || 0,
            clicks: parseInt(rowData['metrics.clicks']) || 0,
            ctr: rowData['metrics.ctr'] || "0%",
            conversions: parseFloat(rowData['metrics.conversions'] || 0).toFixed(2),
            cost: (rowData['metrics.cost_micros'] / 1000000).toFixed(2)
          });
        }
        
        // Combine asset data with performance data
        for (const asset of assets) {
          const key = asset.campaignName + '|' + asset.assetGroupName + '|' + asset.fieldType;
          const perfData = perfDataMap.get(key) || {
            impressions: 0,
            clicks: 0,
            ctr: "0%",
            conversions: 0,
            cost: 0
          };
          
          const rowValues = [
            asset.campaignName,
            asset.assetGroupName,
            asset.fieldType,
            asset.performanceLabel,
            asset.textContent,
            perfData.impressions,
            perfData.clicks,
            perfData.ctr,
            perfData.conversions,
            perfData.cost
          ];
          
          sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
          row++;
        }
      } catch (perfError) {
        Logger.log(`Could not get performance data: ${perfError.message}`);
        
        // If performance data isn't available, just write the asset data
        for (const asset of assets) {
          const rowValues = [
            asset.campaignName,
            asset.assetGroupName,
            asset.fieldType,
            asset.performanceLabel,
            asset.textContent,
            0, // Impressions
            0, // Clicks
            "0%", // CTR
            0, // Conversions
            0  // Cost
          ];
          
          sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
          row++;
        }
      }
      
    } catch (e) {
      Logger.log(`Error processing campaign ${campaignName}: ${e.message}`);
      
      // Add an error row in the sheet
      const errorValues = [
        campaignName,
        "Error",
        "",
        "",
        `Error: ${e.message}`,
        "",
        "",
        "",
        "",
        ""
      ];
      
      sheet.getRange(row, 1, 1, errorValues.length).setValues([errorValues]);
      row++;
    }
  }
  
  // Format the sheet
  sheet.autoResizeColumns(1, headers.length);
  
  // Attempt specialized query for image assets
  try {
    processImageAssets(pmaxCampaigns, sheet, row);
  } catch (imgError) {
    Logger.log(`Error processing image assets: ${imgError.message}`);
  }
}

/**
 * Process image assets specifically
 */
function processImageAssets(pmaxCampaigns, sheet, startRow) {
  let row = startRow;
  
  for (const campaign of pmaxCampaigns) {
    const campaignName = campaign.getName();
    const campaignId = campaign.getId();
    
    try {
      const imageReport = AdsApp.report(
        "SELECT " +
        "  campaign.name, " +
        "  asset_group.name, " +
        "  asset_group_asset.field_type, " +
        "  asset_group_asset.performance_label, " +
        "  asset.image_asset.full_size.url " +
        "FROM asset_group_asset " +
        "WHERE campaign.id = " + campaignId + " " +
        "AND asset.type = 'IMAGE'"
      );
      
      const imageRows = imageReport.rows();
      let imageCount = 0;
      
      while (imageRows.hasNext()) {
        const rowData = imageRows.next();
        imageCount++;
        
        let imageUrl = "";
        try {
          imageUrl = rowData['asset.image_asset.full_size.url'] || "";
        } catch (e) {
          // Image URL might not exist
        }
        
        const rowValues = [
          rowData['campaign.name'],
          rowData['asset_group.name'],
          rowData['asset_group_asset.field_type'] || "IMAGE",
          rowData['asset_group_asset.performance_label'] || "N/A",
          "Image URL: " + imageUrl,
          0, // Impressions
          0, // Clicks
          "0%", // CTR
          0, // Conversions
          0  // Cost
        ];
        
        sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
        row++;
      }
      
      Logger.log(`Found ${imageCount} image assets for campaign ${campaignName}`);
    } catch (e) {
      Logger.log(`Error processing image assets for campaign ${campaignName}: ${e.message}`);
    }
  }
}

/**
 * Process asset groups for each campaign
 */
function processAssetGroups(pmaxCampaigns, spreadsheet, dateRange) {
  const sheet = spreadsheet.insertSheet("Asset Groups");
  
  // Set up headers
  const headers = [
    "Campaign", "Asset Group", "Ad Strength", "Status",
    "Impressions", "Clicks", "CTR", "Conversions", "Cost"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  let row = 2;
  
  // Process each campaign
  for (const campaign of pmaxCampaigns) {
    const campaignName = campaign.getName();
    const campaignId = campaign.getId();
    
    try {
      // Query asset groups with metrics
      const report = AdsApp.report(
        "SELECT " +
        "  campaign.name, " +
        "  asset_group.name, " +
        "  asset_group.ad_strength, " +
        "  asset_group.status, " +
        "  metrics.impressions, " +
        "  metrics.clicks, " +
        "  metrics.ctr, " +
        "  metrics.conversions, " +
        "  metrics.cost_micros " +
        "FROM asset_group " +
        "WHERE campaign.id = " + campaignId + " " +
        "AND segments.date DURING " + dateRange
      );
      
      const rows = report.rows();
      let rowCount = 0;
      
      while (rows.hasNext()) {
        const rowData = rows.next();
        rowCount++;
        
        const costInCurrency = rowData['metrics.cost_micros'] / 1000000;
        
        const rowValues = [
          rowData['campaign.name'],
          rowData['asset_group.name'],
          rowData['asset_group.ad_strength'] || "UNKNOWN",
          rowData['asset_group.status'] || "UNKNOWN",
          parseInt(rowData['metrics.impressions']) || 0,
          parseInt(rowData['metrics.clicks']) || 0,
          rowData['metrics.ctr'] || "0%",
          parseFloat(rowData['metrics.conversions'] || 0).toFixed(2),
          costInCurrency.toFixed(2)
        ];
        
        sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
        row++;
      }
      
      Logger.log(`Found ${rowCount} asset groups for campaign ${campaignName}`);
      
      // If no asset groups were found with metrics, try a simpler query
      if (rowCount === 0) {
        Logger.log(`No asset group data found with metrics for ${campaignName}, trying simplified approach...`);
        
        const altReport = AdsApp.report(
          "SELECT " +
          "  campaign.name, " +
          "  asset_group.name, " +
          "  asset_group.ad_strength, " +
          "  asset_group.status " +
          "FROM asset_group " +
          "WHERE campaign.id = " + campaignId
        );
        
        const altRows = altReport.rows();
        while (altRows.hasNext()) {
          const rowData = altRows.next();
          
          const rowValues = [
            rowData['campaign.name'],
            rowData['asset_group.name'],
            rowData['asset_group.ad_strength'] || "UNKNOWN",
            rowData['asset_group.status'] || "UNKNOWN",
            0, // Impressions
            0, // Clicks
            "0%", // CTR
            0, // Conversions
            0  // Cost
          ];
          
          sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
          row++;
        }
      }
    } catch (e) {
      Logger.log(`Error processing asset group data for campaign ${campaignName}: ${e.message}`);
      // Add an error row in the sheet
      const errorValues = [
        campaignName,
        "Error",
        "Error: " + e.message.split(':')[0],
        "",
        "",
        "",
        "",
        "",
        ""
      ];
      
      sheet.getRange(row, 1, 1, errorValues.length).setValues([errorValues]);
      row++;
    }
  }
  
  // Format the sheet
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Creates a dashboard summary of campaign performance
 */
function createDashboard(spreadsheet, pmaxCampaigns) {
  const dashboardSheet = spreadsheet.getSheetByName("Dashboard");
  dashboardSheet.clear();
  
  // Set title
  dashboardSheet.getRange("A1").setValue("Performance Max Asset Analysis Dashboard");
  dashboardSheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  
  // Campaign summary
  dashboardSheet.getRange("A3").setValue("Campaign Summary");
  dashboardSheet.getRange("A3").setFontWeight("bold");
  
  const campaignHeaders = ["Campaign Name", "Status", "Budget", "Impressions", "Clicks", "CTR", "Conversions", "Cost"];
  dashboardSheet.getRange(4, 1, 1, campaignHeaders.length).setValues([campaignHeaders]);
  dashboardSheet.getRange(4, 1, 1, campaignHeaders.length).setFontWeight("bold");
  
  let row = 5;
  for (const campaign of pmaxCampaigns) {
    const stats = campaign.getStatsFor("LAST_30_DAYS");
    
    const campaignValues = [
      campaign.getName(),
      campaign.isEnabled() ? "Enabled" : "Paused",
      campaign.getBudget().getAmount().toFixed(2),
      stats.getImpressions(),
      stats.getClicks(),
      (stats.getCtr() * 100).toFixed(2) + "%",
      stats.getConversions().toFixed(2),
      stats.getCost().toFixed(2)
    ];
    
    dashboardSheet.getRange(row, 1, 1, campaignValues.length).setValues([campaignValues]);
    row++;
  }
  
  // Asset performance summary
  row += 2;
  dashboardSheet.getRange(row, 1).setValue("Asset Performance Summary");
  dashboardSheet.getRange(row, 1).setFontWeight("bold");
  row++;
  
  const assetSheet = spreadsheet.getSheetByName("Asset Performance");
  if (assetSheet) {
    try {
      const assetData = assetSheet.getDataRange().getValues();
      // Skip header row
      const fieldTypeMap = new Map();
      
      for (let i = 1; i < assetData.length; i++) {
        const fieldType = assetData[i][2];
        const performanceLabel = assetData[i][3];
        
        if (!fieldType) continue;
        
        if (!fieldTypeMap.has(fieldType)) {
          fieldTypeMap.set(fieldType, {
            total: 0,
            best: 0,
            good: 0,
            low: 0,
            poor: 0,
            other: 0
          });
        }
        
        const stats = fieldTypeMap.get(fieldType);
        stats.total++;
        
        if (performanceLabel === "BEST") stats.best++;
        else if (performanceLabel === "GOOD") stats.good++;
        else if (performanceLabel === "LOW") stats.low++;
        else if (performanceLabel === "POOR") stats.poor++;
        else stats.other++;
      }
      
      // Write summary headers
      const summaryHeaders = ["Field Type", "Total Assets", "Best", "Good", "Low", "Poor", "Other"];
      dashboardSheet.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
      dashboardSheet.getRange(row, 1, 1, summaryHeaders.length).setFontWeight("bold");
      row++;
      
      // Write data for each field type
      for (const [fieldType, stats] of fieldTypeMap.entries()) {
        const summaryValues = [
          fieldType,
          stats.total,
          stats.best,
          stats.good,
          stats.low,
          stats.poor,
          stats.other
        ];
        
        dashboardSheet.getRange(row, 1, 1, summaryValues.length).setValues([summaryValues]);
        row++;
      }
    } catch (e) {
      Logger.log(`Error creating asset summary: ${e.message}`);
    }
  }
  
  // Asset groups summary
  row += 2;
  dashboardSheet.getRange(row, 1).setValue("Asset Group Ad Strength Summary");
  dashboardSheet.getRange(row, 1).setFontWeight("bold");
  row++;
  
  const assetGroupSheet = spreadsheet.getSheetByName("Asset Groups");
  if (assetGroupSheet) {
    try {
      const assetGroupData = assetGroupSheet.getDataRange().getValues();
      // Skip header row
      const strengthMap = new Map();
      
      for (let i = 1; i < assetGroupData.length; i++) {
        const adStrength = assetGroupData[i][2] || "UNKNOWN";
        
        if (!strengthMap.has(adStrength)) {
          strengthMap.set(adStrength, {
            count: 0
          });
        }
        
        const stats = strengthMap.get(adStrength);
        stats.count++;
      }
      
      // Write summary headers
      const summaryHeaders = ["Ad Strength", "Count"];
      dashboardSheet.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
      dashboardSheet.getRange(row, 1, 1, summaryHeaders.length).setFontWeight("bold");
      row++;
      
      // Write data for each strength level
      for (const [strength, stats] of strengthMap.entries()) {
        const summaryValues = [
          strength,
          stats.count
        ];
        
        dashboardSheet.getRange(row, 1, 1, summaryValues.length).setValues([summaryValues]);
        row++;
      }
    } catch (e) {
      Logger.log(`Error creating asset group summary: ${e.message}`);
    }
  }
  
  // Format dashboard
  dashboardSheet.autoResizeColumns(1, campaignHeaders.length);
  
  // Add guidance about the report
  row += 2;
  dashboardSheet.getRange(row, 1).setValue("Notes About This Report");
  dashboardSheet.getRange(row, 1).setFontWeight("bold");
  row++;
  
  const notes = [
    "• Performance metrics (impressions, clicks, etc.) may show as zeros if they can't be accessed through the Google Ads API.",
    "• Asset performance labels (BEST, GOOD, LOW, POOR) provide relative performance insights from Google Ads.",
    "• Campaign-level metrics reflect actual performance data even if asset-level metrics are unavailable.",
    "• Asset performance data requires special API access that some accounts may not have.",
    "• Use this report to identify asset types and optimize based on performance labels.",
    "• For full metrics visualization, use Google Ads UI under Performance Max campaigns > Asset Groups."
  ];
  
  for (const note of notes) {
    dashboardSheet.getRange(row, 1).setValue(note);
    row++;
  }
  
  // Add date information
  row += 1;
  dashboardSheet.getRange(row, 1).setValue("Report generated on: " + new Date().toLocaleString());
}

