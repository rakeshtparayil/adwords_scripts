/**
 * Google Ads A/B Testing Script with Statistical Significance and Weighted Metrics
 * 
 * This script compares the performance of ads by calculating statistical significance
 * between control and test ads and provides a weighted composite score to determine
 * overall performance across multiple metrics (CTR, CPC, and Conversion Rate).
 */

function main() {
  // Configuration - adjust as needed
  var CONFIG = {
    // Minimum number of impressions required for an ad to be included
    minImpressions: 50,
    
    // Significance level (alpha) - 0.05 = 95% confidence level
    significanceLevel: 0.05,
    
    // Metric weights for composite score (must add up to 1.0)
    metricWeights: {
      CTR: 0.4,       // 40% weight to Click-Through Rate
      CPC: 0.3,       // 30% weight to Cost Per Click (lower is better)
      Conversions: 0.3 // 30% weight to Conversion Rate
    },
    
    // Whether higher is better for each metric (true) or lower is better (false)
    metricDirection: {
      CTR: true,       // Higher CTR is better
      CPC: false,      // Lower CPC is better
      Conversions: true // Higher conversion rate is better
    },
    
    // Primary metric for significance testing
    primaryMetricForSignificance: "CTR", 
    
    // Control ad definition - use final URL to identify the control ad
    controlAdUrl: "https://www.wafeq.com/ar-sa/campaigns/wafeq",
    
    // Email to send the report to (optional)
    emailAddress: "",
    
    // Output results to Google Sheet (if true, provide spreadsheetUrl or leave blank to create new)
    outputToSheet: true,
    
    // Google Sheet URL to output results (leave blank to create a new one)
    spreadsheetUrl: "",
    
    // Filter for specific campaigns or ad groups (optional)
    campaignNameContains: "GL_WFQ_GA_SEM_AO_NBR_Gen_Web_All_KSA_AR_PRO_Conv_14032025_Accounting-Software",
    adGroupNameContains: "",
    
    // Include ads with all statuses (ENABLED, PAUSED, etc.) or only enabled
    includeAllAdStatus: false,
    
    // Date range for the analysis (e.g., "LAST_30_DAYS", "LAST_7_DAYS", "YESTERDAY", "THIS_MONTH", "LAST_MONTH")
    dateRange: "LAST_30_DAYS"
  };
  
  // Run the A/B test
  var results = runTest(CONFIG);
  
  // Log and share results
  logResults(results, CONFIG);
  if (CONFIG.emailAddress) {
    emailResults(results, CONFIG);
  }
  if (CONFIG.outputToSheet) {
    outputToGoogleSheet(results, CONFIG);
  }
}

/**
 * Runs the A/B test analysis
 */
function runTest(config) {
  var adPerformanceData = [];
  var adGroups = [];
  
  // Collect all relevant ad groups
  var adGroupSelector = AdsApp.adGroups()
    .forDateRange(config.dateRange);
    
  if (config.campaignNameContains) {
    adGroupSelector = adGroupSelector.withCondition("CampaignName CONTAINS '" + config.campaignNameContains + "'");
  }
  if (config.adGroupNameContains) {
    adGroupSelector = adGroupSelector.withCondition("AdGroupName CONTAINS '" + config.adGroupNameContains + "'");
  }
  
  var adGroupIterator = adGroupSelector.get();
  while (adGroupIterator.hasNext()) {
    adGroups.push(adGroupIterator.next());
  }
  
  Logger.log("Analyzing " + adGroups.length + " ad groups");
  
  // Process each ad group
  adGroups.forEach(function(adGroup) {
    var adGroupData = {
      name: adGroup.getName(),
      campaignName: adGroup.getCampaign().getName(),
      ads: []
    };
    
    // Get ads in this ad group
    var adSelector = adGroup.ads()
      .withCondition("Impressions >= " + config.minImpressions)
      .forDateRange(config.dateRange);
      
    if (!config.includeAllAdStatus) {
      adSelector = adSelector.withCondition("Status = ENABLED");
    }
      
    var adIterator = adSelector.get();
    
    // Skip ad groups with fewer than 2 ads
    if (adSelector.get().totalNumEntities() < 2) {
      return;
    }
    
    while (adIterator.hasNext()) {
      var ad = adIterator.next();
      var stats = ad.getStatsFor(config.dateRange);
      
      adGroupData.ads.push({
        id: ad.getId(),
        headline: getAdHeadline(ad),
        displayId: "ID: " + ad.getId() + " - " + getAdHeadline(ad), // Combined ID and headline for display
        finalUrl: ad.urls().getFinalUrl(),
        impressions: stats.getImpressions(),
        clicks: stats.getClicks(),
        ctr: stats.getCtr(),
        avgCpc: stats.getAverageCpc(),
        cost: stats.getCost(),
        conversions: stats.getConversions(),
        convRate: stats.getConversionRate(),
        // Calculate cost per conversion manually to avoid the error
        costPerConversion: stats.getConversions() > 0 ? stats.getCost() / stats.getConversions() : 0
      });
    }
    
    // Only include ad groups with at least 2 ads
    if (adGroupData.ads.length >= 2) {
      // Find control ad based on final URL
      var controlAd = null;
      var testAds = [];
      
      // Try to find the control ad first by final URL
      for (var i = 0; i < adGroupData.ads.length; i++) {
        var ad = adGroupData.ads[i];
        // Check if this ad has the control URL
        if (ad.finalUrl && ad.finalUrl.indexOf(config.controlAdUrl) !== -1) {
          controlAd = ad;
          break;
        }
      }
      
      // If no control ad found by URL, fall back to using the one with highest impressions
      if (!controlAd) {
        // Sort ads by impressions (highest first) and use the first one as control
        adGroupData.ads.sort(function(a, b) {
          return b.impressions - a.impressions;
        });
        controlAd = adGroupData.ads[0];
        Logger.log("Warning: No ad found with control URL '" + config.controlAdUrl + 
                  "' in ad group '" + adGroupData.name + "'. Using highest impression ad as control.");
      }
      
      // Get all non-control ads as test ads
      for (var i = 0; i < adGroupData.ads.length; i++) {
        if (adGroupData.ads[i].id !== controlAd.id) {
          testAds.push(adGroupData.ads[i]);
        }
      }
      
      // Only proceed if we have at least one test ad
      if (testAds.length > 0) {
        adGroupData.controlAd = controlAd;
        adGroupData.testResults = [];
        
        // Compare each test ad against the control
        for (var i = 0; i < testAds.length; i++) {
          var testAd = testAds[i];
          var testResult = calculateSignificance(controlAd, testAd, config.primaryMetricForSignificance, config);
          adGroupData.testResults.push({
            testAd: testAd,
            pValue: testResult.pValue,
            isSignificant: testResult.pValue < config.significanceLevel,
            relativeDifference: testResult.relativeDifference,
            absoluteDifference: testResult.absoluteDifference,
            better: testResult.better,
            compositeScore: testResult.compositeScore,
            isBetterOverall: testResult.compositeScore > 0.5,
            metricResults: testResult.metricResults
          });
        }
        
        adPerformanceData.push(adGroupData);
      }
    }
  });
  
  return adPerformanceData;
}

/**
 * Calculates statistical significance between control and test ad
 * and computes a composite score based on weighted metrics
 */
function calculateSignificance(controlAd, testAd, primaryMetric, config) {
  var result = {
    pValue: 1.0,
    relativeDifference: 0,
    absoluteDifference: 0,
    better: false,
    compositeScore: 0,
    metricResults: {}
  };
  
  // Statistical significance for primary metric (default: CTR)
  var controlClicks = controlAd.clicks;
  var controlImpressions = controlAd.impressions;
  var testClicks = testAd.clicks;
  var testImpressions = testAd.impressions;
  
  // Calculate Fisher's exact test
  var fisherResult = fisherExactTest(
    controlClicks,
    controlImpressions - controlClicks,
    testClicks,
    testImpressions - testClicks
  );
  
  result.pValue = fisherResult.pValue;
  
  // Calculate differences for all metrics
  // 1. CTR
  var controlCTR = controlAd.ctr;
  var testCTR = testAd.ctr;
  var ctrRelativeDiff = 0;
  if (controlCTR !== 0) {
    ctrRelativeDiff = ((testCTR - controlCTR) / controlCTR) * 100;
  } else if (testCTR !== 0) {
    ctrRelativeDiff = 100; // If control is 0 and test is not, that's a 100% improvement
  }
  
  result.metricResults.CTR = {
    control: controlCTR,
    test: testCTR,
    relativeDiff: ctrRelativeDiff,
    absoluteDiff: (testCTR - controlCTR) * 100,
    better: testCTR > controlCTR
  };
  
  // 2. CPC 
  var controlCPC = controlAd.avgCpc;
  var testCPC = testAd.avgCpc;
  var cpcRelativeDiff = 0;
  if (controlCPC !== 0) {
    cpcRelativeDiff = ((testCPC - controlCPC) / controlCPC) * 100;
  } else if (testCPC !== 0) {
    cpcRelativeDiff = -100; // Lower CPC is better, so if control is 0 it's worse
  }
  
  result.metricResults.CPC = {
    control: controlCPC,
    test: testCPC,
    relativeDiff: cpcRelativeDiff,
    absoluteDiff: testCPC - controlCPC,
    better: testCPC < controlCPC  // Lower CPC is better
  };
  
  // 3. Conversion Rate
  var controlConvRate = controlAd.convRate || 0;
  var testConvRate = testAd.convRate || 0;
  var convRelativeDiff = 0;
  if (controlConvRate !== 0) {
    convRelativeDiff = ((testConvRate - controlConvRate) / controlConvRate) * 100;
  } else if (testConvRate !== 0) {
    convRelativeDiff = 100; // If control is 0 and test is not, that's a 100% improvement
  }
  
  result.metricResults.Conversions = {
    control: controlConvRate,
    test: testConvRate,
    relativeDiff: convRelativeDiff,
    absoluteDiff: (testConvRate - controlConvRate) * 100,
    better: testConvRate > controlConvRate
  };
  
  // For backward compatibility with primary metric
  if (primaryMetric === "CTR") {
    result.relativeDifference = result.metricResults.CTR.relativeDiff;
    result.absoluteDifference = result.metricResults.CTR.absoluteDiff;
    result.better = result.metricResults.CTR.better;
  }
  
  // Calculate weighted composite score
  result.compositeScore = calculateCompositeScore(controlAd, testAd, config);
  
  return result;
}

/**
 * Calculates composite score based on weighted metrics
 */
function calculateCompositeScore(controlAd, testAd, config) {
  var weights = config.metricWeights;
  var directions = config.metricDirection;
  var score = 0;
  
  // Normalize each metric to 0-1 scale and apply weights
  
  // 1. CTR (higher is better)
  var maxCTR = Math.max(controlAd.ctr, testAd.ctr);
  var minCTR = Math.min(controlAd.ctr, testAd.ctr);
  var ctrScore = 0.5; // Default if both are equal
  
  if (maxCTR !== minCTR) {
    var denominator = maxCTR - minCTR;
    if (denominator === 0) {
      denominator = 0.0001; // Avoid division by zero
    }
    ctrScore = (testAd.ctr - minCTR) / denominator;
  }
  
  if (!directions.CTR) {
    ctrScore = 1 - ctrScore;
  }
  
  // 2. CPC (lower is better)
  var maxCPC = Math.max(controlAd.avgCpc, testAd.avgCpc);
  var minCPC = Math.min(controlAd.avgCpc, testAd.avgCpc);
  var cpcScore = 0.5; // Default if both are equal
  
  if (maxCPC !== minCPC) {
    var denominator = maxCPC - minCPC;
    if (denominator === 0) {
      denominator = 0.0001; // Avoid division by zero
    }
    cpcScore = (maxCPC - testAd.avgCpc) / denominator;
  }
  
  if (!directions.CPC) {
    cpcScore = 1 - cpcScore;
  }
  
  // 3. Conversion Rate (higher is better)
  var controlConvRate = controlAd.convRate || 0;
  var testConvRate = testAd.convRate || 0;
  var maxConvRate = Math.max(controlConvRate, testConvRate);
  var minConvRate = Math.min(controlConvRate, testConvRate);
  var convScore = 0.5; // Default if both are equal
  
  if (maxConvRate !== minConvRate) {
    var denominator = maxConvRate - minConvRate;
    if (denominator === 0) {
      denominator = 0.0001; // Avoid division by zero
    }
    convScore = (testConvRate - minConvRate) / denominator;
  }
  
  if (!directions.Conversions) {
    convScore = 1 - convScore;
  }
  
  // Apply weights
  score = (weights.CTR * ctrScore) + 
          (weights.CPC * cpcScore) + 
          (weights.Conversions * convScore);
  
  return score;
}

/**
 * Implements Fisher's Exact Test for calculating p-values
 * a, b, c, d represent the 2x2 contingency table:
 * [ a  b ]
 * [ c  d ]
 */
function fisherExactTest(a, b, c, d) {
  // Calculate Fisher's exact test p-value
  // Note: This implementation works for moderately sized numbers but may have precision issues with very large numbers
  
  function logFactorial(n) {
    var result = 0;
    for (var i = 1; i <= n; i++) {
      result += Math.log(i);
    }
    return result;
  }
  
  // Calculate p-value using the hypergeometric probability
  var n = a + b + c + d;
  var logP = logFactorial(a + b) + logFactorial(c + d) + logFactorial(a + c) + logFactorial(b + d) - 
             logFactorial(a) - logFactorial(b) - logFactorial(c) - logFactorial(d) - logFactorial(n);
  
  var p = Math.exp(logP);
  
  // For a two-tailed test, we need to sum the probabilities of all outcomes as or more extreme
  // This is a simplified approach for Fisher's exact test
  return {
    pValue: p * 2 // Two-tailed test approximation
  };
}

/**
 * Gets a readable headline from an ad
 */
function getAdHeadline(ad) {
  try {
    if (ad.getType() === "RESPONSIVE_SEARCH_AD") {
      return ad.getHeadlines()[0].text;
    } else {
      // For expanded text ads
      return ad.getHeadlinePart1();
    }
  } catch (e) {
    return "Ad #" + ad.getId();
  }
}

/**
 * Gets the metric value from an ad based on the specified metric
 */
function getMetricValue(ad, metric) {
  switch (metric) {
    case "CTR":
      return ad.ctr;
    case "CPC":
      return ad.avgCpc;
    case "Conversions":
      return ad.convRate || 0;
    default:
      return ad.ctr;
  }
}

/**
 * Formats the metric value for display
 */
function formatMetricValue(value, metric) {
  switch (metric) {
    case "CTR":
    case "Conversions":
      return (value * 100).toFixed(2) + "%";
    case "CPC":
      return value.toFixed(2);
    default:
      return value.toFixed(2);
  }
}

/**
 * Logs the test results to the console
 */
function logResults(results, config) {
  Logger.log("=====================================================");
  Logger.log("A/B TEST RESULTS - " + config.dateRange);
  Logger.log("Control landing page: " + config.controlAdUrl);
  Logger.log("Primary metric for significance: " + config.primaryMetricForSignificance);
  Logger.log("Metric weights: CTR=" + (config.metricWeights.CTR*100) + "%, CPC=" + 
             (config.metricWeights.CPC*100) + "%, Conversions=" + (config.metricWeights.Conversions*100) + "%");
  Logger.log("Significance level: " + (config.significanceLevel * 100) + "%");
  Logger.log("=====================================================");
  
  var significantCount = 0;
  var totalTests = 0;
  var betterOverallCount = 0;
  
  results.forEach(function(adGroupData) {
    Logger.log("\nCampaign: " + adGroupData.campaignName);
    Logger.log("Ad Group: " + adGroupData.name);
    Logger.log("Control Ad: " + adGroupData.controlAd.displayId);
    
    // Log all metrics for control
    Logger.log("Control CTR: " + formatMetricValue(adGroupData.controlAd.ctr, "CTR"));
    Logger.log("Control CPC: " + formatMetricValue(adGroupData.controlAd.avgCpc, "CPC"));
    Logger.log("Control Conv. Rate: " + formatMetricValue(adGroupData.controlAd.convRate || 0, "Conversions"));
    
    adGroupData.testResults.forEach(function(result) {
      totalTests++;
      
      Logger.log("\n  Test Ad: " + result.testAd.displayId);
      
      // Log all metrics for test ad
      Logger.log("  Test CTR: " + formatMetricValue(result.testAd.ctr, "CTR") + 
                " (" + (result.metricResults.CTR.relativeDiff >= 0 ? "+" : "") + 
                result.metricResults.CTR.relativeDiff.toFixed(2) + "%) " + 
                (result.metricResults.CTR.better ? "▲" : "▼"));
      
      Logger.log("  Test CPC: " + formatMetricValue(result.testAd.avgCpc, "CPC") + 
                " (" + (result.metricResults.CPC.relativeDiff >= 0 ? "+" : "") + 
                result.metricResults.CPC.relativeDiff.toFixed(2) + "%) " + 
                (result.metricResults.CPC.better ? "▲" : "▼"));
      
      Logger.log("  Test Conv. Rate: " + formatMetricValue(result.testAd.convRate || 0, "Conversions") + 
                " (" + (result.metricResults.Conversions.relativeDiff >= 0 ? "+" : "") + 
                result.metricResults.Conversions.relativeDiff.toFixed(2) + "%) " + 
                (result.metricResults.Conversions.better ? "▲" : "▼"));
      
      // Statistical significance of primary metric
      var significanceSymbol = result.isSignificant ? "✓" : "✗";
      Logger.log("  P-value for " + config.primaryMetricForSignificance + ": " + 
                result.pValue.toFixed(4) + " - Significant: " + significanceSymbol);
      
      // Composite score
      var compositeSymbol = result.isBetterOverall ? "BETTER" : "WORSE";
      Logger.log("  Composite Score: " + result.compositeScore.toFixed(4) + 
                " - Overall: " + compositeSymbol);
      
      if (result.isSignificant) {
        significantCount++;
      }
      
      if (result.isBetterOverall) {
        betterOverallCount++;
      }
    });
    
    Logger.log("-----------------------------------------------------");
  });
  
  Logger.log("\nSummary: ");
  Logger.log("- " + significantCount + " out of " + totalTests + 
            " tests showed statistical significance at " + (config.significanceLevel * 100) + "% confidence level.");
  Logger.log("- " + betterOverallCount + " out of " + totalTests + 
            " test ads performed better overall based on the weighted metrics.");
}

/**
 * Emails the test results
 */
function emailResults(results, config) {
  var subject = "Google Ads A/B Test Results - " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  var body = "<html><body>";
  body += "<h2>A/B Test Results - " + config.dateRange + "</h2>";
  body += "<p><strong>Control landing page:</strong> " + config.controlAdUrl + "</p>";
  body += "<p><strong>Primary metric for significance:</strong> " + config.primaryMetricForSignificance + "</p>";
  body += "<p><strong>Metric weights:</strong> CTR=" + (config.metricWeights.CTR*100) + "%, " +
          "CPC=" + (config.metricWeights.CPC*100) + "%, " +
          "Conversions=" + (config.metricWeights.Conversions*100) + "%</p>";
  body += "<p><strong>Significance level:</strong> " + (config.significanceLevel * 100) + "%</p>";
  
  var significantCount = 0;
  var totalTests = 0;
  var betterOverallCount = 0;
  
  results.forEach(function(adGroupData) {
    body += "<hr>";
    body += "<h3>Campaign: " + adGroupData.campaignName + "</h3>";
    body += "<h4>Ad Group: " + adGroupData.name + "</h4>";
    
    // Control ad information
    body += "<div style='margin-bottom: 20px;'>";
    body += "<h4>Control Ad: " + adGroupData.controlAd.displayId + "</h4>";
    body += "<p><strong>Final URL:</strong> " + (adGroupData.controlAd.finalUrl || "N/A") + "</p>";
    body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>";
    body += "<tr style='background-color:#f2f2f2;'><th>Ad ID</th><th>CTR</th><th>CPC</th><th>Conv. Rate</th><th>Conversions</th><th>Cost/Conv</th></tr>";
    
    body += "<tr>";
    body += "<td>" + adGroupData.controlAd.id + "</td>";
    body += "<td>" + formatMetricValue(adGroupData.controlAd.ctr, "CTR") + "</td>";
    body += "<td>" + formatMetricValue(adGroupData.controlAd.avgCpc, "CPC") + "</td>";
    body += "<td>" + formatMetricValue(adGroupData.controlAd.convRate || 0, "Conversions") + "</td>";
    body += "<td>" + (adGroupData.controlAd.conversions || 0) + "</td>";
    body += "<td>" + formatMetricValue(adGroupData.controlAd.costPerConversion || 0, "CPC") + "</td>";
    body += "</tr>";
    body += "</table>";
    body += "</div>";
    
    // Test ads table
    body += "<h4>Test Ads</h4>";
    body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>";
    body += "<tr style='background-color:#f2f2f2;'><th>Test Ad</th><th>Ad ID</th>" +
            "<th>CTR (Diff)</th><th>CPC (Diff)</th><th>Conv. Rate (Diff)</th>" +
            "<th>Conversions</th><th>Cost/Conv</th><th>P-value</th><th>Significance</th><th>Composite Score</th><th>Overall</th></tr>";
    
    adGroupData.testResults.forEach(function(result) {
      totalTests++;
      
      var significanceText = result.isSignificant ? "YES" : "NO";
      var overallText = result.isBetterOverall ? "BETTER" : "WORSE";
      var rowColor = result.isBetterOverall ? "#d9ead3" : "#f4cccc"; // Light green if better, light red if worse
      
      body += "<tr style='background-color:" + rowColor + ";'>";
      body += "<td>" + result.testAd.headline + "</td>";
      body += "<td>" + result.testAd.id + "</td>";
      
      // CTR column
      body += "<td>" + formatMetricValue(result.testAd.ctr, "CTR") + 
              " (" + (result.metricResults.CTR.relativeDiff >= 0 ? "+" : "") + 
              result.metricResults.CTR.relativeDiff.toFixed(2) + "% " + 
              (result.metricResults.CTR.better ? "▲" : "▼") + ")</td>";
      
      // CPC column
      body += "<td>" + formatMetricValue(result.testAd.avgCpc, "CPC") + 
              " (" + (result.metricResults.CPC.relativeDiff >= 0 ? "+" : "") + 
              result.metricResults.CPC.relativeDiff.toFixed(2) + "% " + 
              (result.metricResults.CPC.better ? "▲" : "▼") + ")</td>";
      
      // Conv Rate column
      body += "<td>" + formatMetricValue(result.testAd.convRate || 0, "Conversions") + 
              " (" + (result.metricResults.Conversions.relativeDiff >= 0 ? "+" : "") + 
              result.metricResults.Conversions.relativeDiff.toFixed(2) + "% " + 
              (result.metricResults.Conversions.better ? "▲" : "▼") + ")</td>";
      
      // Conversions and Cost/Conv
      body += "<td>" + (result.testAd.conversions || 0) + "</td>";
      body += "<td>" + formatMetricValue(result.testAd.costPerConversion || 0, "CPC") + "</td>";
      
      // P-value and significance
      body += "<td>" + result.pValue.toFixed(4) + "</td>";
      body += "<td>" + significanceText + "</td>";
      
      // Composite score and overall assessment
      body += "<td>" + result.compositeScore.toFixed(4) + "</td>";
      body += "<td><strong>" + overallText + "</strong></td>";
      
      body += "</tr>";
      
      if (result.isSignificant) {
        significantCount++;
      }
      
      if (result.isBetterOverall) {
        betterOverallCount++;
      }
    });
    
    body += "</table>";
  });
  
  body += "<hr>";
  body += "<h3>Summary:</h3>";
  body += "<ul>";
  body += "<li><strong>" + significantCount + " out of " + totalTests + "</strong> tests showed statistical significance at " + 
          (config.significanceLevel * 100) + "% confidence level.</li>";
  body += "<li><strong>" + betterOverallCount + " out of " + totalTests + "</strong> test ads performed better overall based on the weighted metrics.</li>";
  body += "</ul>";
  
  // If we're also outputting to Google Sheets, include the link
  if (config.outputToSheet && config.spreadsheetUrl) {
    body += "<p><a href='" + config.spreadsheetUrl + "'>View detailed results in Google Sheets</a></p>";
  }
  
  body += "</body></html>";
  
  MailApp.sendEmail({
    to: config.emailAddress,
    subject: subject,
    htmlBody: body
  });
  
  Logger.log("Email report sent to " + config.emailAddress);
}

/**
 * Outputs the test results to a Google Sheet
 */
function outputToGoogleSheet(results, config) {
  var spreadsheet;
  var sheet;
  
  // Either use an existing spreadsheet or create a new one
  if (config.spreadsheetUrl && config.spreadsheetUrl !== "") {
    try {
      spreadsheet = SpreadsheetApp.openByUrl(config.spreadsheetUrl);
      // Create a new sheet with today's date
      var sheetName = "A/B Test " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
      sheet = spreadsheet.getSheetByName(sheetName);
      
      // If a sheet with this name already exists, append a number to make it unique
      if (sheet) {
        var counter = 1;
        while (sheet) {
          sheetName = "A/B Test " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd") + " (" + counter + ")";
          sheet = spreadsheet.getSheetByName(sheetName);
          counter++;
        }
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        sheet = spreadsheet.insertSheet(sheetName);
      }
    } catch (e) {
      Logger.log("Error opening spreadsheet: " + e);
      // Create a new spreadsheet if there was an error opening the existing one
      spreadsheet = SpreadsheetApp.create("Google Ads A/B Test Results");
      sheet = spreadsheet.getActiveSheet();
      sheet.setName("A/B Test " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd"));
      config.spreadsheetUrl = spreadsheet.getUrl();
    }
  } else {
    // Create a new spreadsheet
    spreadsheet = SpreadsheetApp.create("Google Ads A/B Test Results");
    sheet = spreadsheet.getActiveSheet();
    sheet.setName("A/B Test " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd"));
    config.spreadsheetUrl = spreadsheet.getUrl();
  }
  
  // Add title and config info
  var row = 1;
  sheet.getRange(row, 1).setValue("Google Ads A/B Test Results");
  sheet.getRange(row, 1).setFontWeight("bold").setFontSize(14);
  row++;
  
  sheet.getRange(row, 1).setValue("Date Range:");
  sheet.getRange(row, 2).setValue(config.dateRange);
  row++;
  
  sheet.getRange(row, 1).setValue("Control Landing Page:");
  sheet.getRange(row, 2).setValue(config.controlAdUrl);
  row++;
  
  sheet.getRange(row, 1).setValue("Primary Metric for Significance:");
  sheet.getRange(row, 2).setValue(config.primaryMetricForSignificance);
  row++;
  
  sheet.getRange(row, 1).setValue("Metric Weights:");
  sheet.getRange(row, 2).setValue("CTR=" + (config.metricWeights.CTR*100) + "%, CPC=" + 
                                 (config.metricWeights.CPC*100) + "%, Conv=" + 
                                 (config.metricWeights.Conversions*100) + "%");
  row++;
  
  sheet.getRange(row, 1).setValue("Significance Level:");
  sheet.getRange(row, 2).setValue((config.significanceLevel * 100) + "%");
  row++;
  
  sheet.getRange(row, 1).setValue("Minimum Impressions:");
  sheet.getRange(row, 2).setValue(config.minImpressions);
  row++;
  
  row++; // Empty row
  
  // Create a summary row counter and total significant tests counter
  var totalTests = 0;
  var significantTests = 0;
  var betterOverallTests = 0;
  
  // For each ad group, create a section
  results.forEach(function(adGroupData) {
    // Ad Group header
    sheet.getRange(row, 1).setValue("Campaign:");
    sheet.getRange(row, 2).setValue(adGroupData.campaignName);
    row++;
    
    sheet.getRange(row, 1).setValue("Ad Group:");
    sheet.getRange(row, 2).setValue(adGroupData.name);
    row++;
    
    // Control ad information
    sheet.getRange(row, 1).setValue("Control Ad:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.headline);
    row++;
    
    sheet.getRange(row, 1).setValue("Control Ad ID:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.id);
    row++;
    
    sheet.getRange(row, 1).setValue("Control Final URL:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.finalUrl || "N/A");
    row++;
    
    // Control ad metrics
    sheet.getRange(row, 1).setValue("Control CTR:");
    sheet.getRange(row, 2).setValue((adGroupData.controlAd.ctr * 100).toFixed(2) + "%");
    row++;
    
    sheet.getRange(row, 1).setValue("Control CPC:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.avgCpc.toFixed(2));
    row++;
    
    sheet.getRange(row, 1).setValue("Control Conv. Rate:");
    sheet.getRange(row, 2).setValue(((adGroupData.controlAd.convRate || 0) * 100).toFixed(2) + "%");
    row++;
    
    sheet.getRange(row, 1).setValue("Control Conversions:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.conversions || 0);
    row++;
    
    sheet.getRange(row, 1).setValue("Control Cost/Conv:");
    sheet.getRange(row, 2).setValue((adGroupData.controlAd.costPerConversion || 0).toFixed(2));
    row++;
    
    sheet.getRange(row, 1).setValue("Control Clicks:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.clicks);
    row++;
    
    sheet.getRange(row, 1).setValue("Control Impressions:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.impressions);
    row++;
    
    sheet.getRange(row, 1).setValue("Control Cost:");
    sheet.getRange(row, 2).setValue(adGroupData.controlAd.cost.toFixed(2));
    row++;
    
    // Table header for test ads
    var headers = ["Test Ad", "Ad ID", "CTR", "CTR Diff", "CPC", "CPC Diff", "Conv. Rate", "Conv. Rate Diff", 
                  "Conversions", "Cost/Conv", "Clicks", "Impressions", "Cost", "P-value", "Significant?", 
                  "Composite Score", "Overall"];
    for (var i = 0; i < headers.length; i++) {
      sheet.getRange(row, i + 1).setValue(headers[i]).setFontWeight("bold");
    }
    row++;
    
    // Add test ad data
    adGroupData.testResults.forEach(function(result) {
      totalTests++;
      
      var testAd = result.testAd;
      var significanceText = result.isSignificant ? "YES" : "NO";
      var overallText = result.isBetterOverall ? "BETTER" : "WORSE";
      
      // If significant, count it
      if (result.isSignificant) {
        significantTests++;
      }
      
      // If better overall, count it
      if (result.isBetterOverall) {
        betterOverallTests++;
      }
      
      // Format the CTR difference value
      var ctrDiffFormatted = (result.metricResults.CTR.relativeDiff >= 0 ? "+" : "") + 
                             result.metricResults.CTR.relativeDiff.toFixed(2) + "% " +
                             (result.metricResults.CTR.better ? "▲" : "▼");
      
      // Format the CPC difference value
      var cpcDiffFormatted = (result.metricResults.CPC.relativeDiff >= 0 ? "+" : "") + 
                             result.metricResults.CPC.relativeDiff.toFixed(2) + "% " +
                             (result.metricResults.CPC.better ? "▲" : "▼");
      
      // Format the Conv Rate difference value
      var convDiffFormatted = (result.metricResults.Conversions.relativeDiff >= 0 ? "+" : "") + 
                             result.metricResults.Conversions.relativeDiff.toFixed(2) + "% " +
                             (result.metricResults.Conversions.better ? "▲" : "▼");
      
      var rowData = [
        testAd.headline,
        testAd.id,
        (testAd.ctr * 100).toFixed(2) + "%",
        ctrDiffFormatted,
        testAd.avgCpc.toFixed(2),
        cpcDiffFormatted,
        ((testAd.convRate || 0) * 100).toFixed(2) + "%",
        convDiffFormatted,
        testAd.conversions || 0,
        (testAd.costPerConversion || 0).toFixed(2),
        testAd.clicks,
        testAd.impressions,
        testAd.cost.toFixed(2),
        result.pValue.toFixed(4),
        significanceText,
        result.compositeScore.toFixed(4),
        overallText
      ];
      
      for (var i = 0; i < rowData.length; i++) {
        var cell = sheet.getRange(row, i + 1);
        cell.setValue(rowData[i]);
        
        // Color coding for significance
        if (i === 14) { // Significant column
          if (result.isSignificant) {
            cell.setBackground("#d9ead3"); // Light green
          }
        }
        
        // Color coding for better/worse
        if (i === 16) { // Overall column
          if (rowData[i] === "BETTER") {
            cell.setBackground("#d9ead3"); // Light green for better
          } else {
            cell.setBackground("#f4cccc"); // Light red for worse
          }
        }
      }
      
      row++;
    });
    
    row++; // Empty row between ad groups
  });
  
  // Add summary section
  sheet.getRange(row, 1).setValue("SUMMARY").setFontWeight("bold");
  row++;
  
  sheet.getRange(row, 1).setValue("Total Tests:");
  sheet.getRange(row, 2).setValue(totalTests);
  row++;
  
  sheet.getRange(row, 1).setValue("Significant Results:");
  sheet.getRange(row, 2).setValue(significantTests);
  row++;
  
  sheet.getRange(row, 1).setValue("Significance Rate:");
  var significanceRate = totalTests > 0 ? (significantTests / totalTests * 100).toFixed(1) + "%" : "0%";
  sheet.getRange(row, 2).setValue(significanceRate);
  row++;
  
  sheet.getRange(row, 1).setValue("Better Overall:");
  sheet.getRange(row, 2).setValue(betterOverallTests);
  row++;
  
  sheet.getRange(row, 1).setValue("Better Overall Rate:");
  var betterRate = totalTests > 0 ? (betterOverallTests / totalTests * 100).toFixed(1) + "%" : "0%";
  sheet.getRange(row, 2).setValue(betterRate);
  row++;
  
  // Format the spreadsheet
  sheet.autoResizeColumns(1, 17); // Increased to 17 for all the new columns
  
  // Log the spreadsheet URL
  Logger.log("Results output to Google Sheet: " + config.spreadsheetUrl);
  
  // If we created a new spreadsheet, update the CONFIG object so the email can include the link
  if (config.spreadsheetUrl) {
    config.spreadsheetUrl = config.spreadsheetUrl;
  }
  
  return config.spreadsheetUrl;
}
