/**
 * @name Quality Score Analysis 
 * 
 * @overview Export keyword performance data along with quality score  
 * parameters to the give spreadsheet. 
 * 
 * Please send in your comment, feedback, criticism or occasional Thank You 
 * to support@karooya.com
 * 
 * All AdWords scripts published by Karooya are available here 
 * http://www.karooya.com/blog/category/karooya-adwords-script/ 
 *
 * Copyright (c) 2017 - Karooya Technologies Pvt. Ltd.  All Rights Reserved.
 * http://www.karooya.com
 * 
 * The script was last modified on: 29 April, 2019
 */

// Path to the folder in Google Drive where all the reports are to be created
var REPORTS_FOLDER_PATH = 'Quality Score Analysis';

// Specify a date range for the report 
var DATE_RANGE = "LAST_7_DAYS"; // Other allowed values are: LAST_<NUM>_DAYS (Ex: LAST_90_DAYS) or TODAY, YESTERDAY, THIS_WEEK_SUN_TODAY, THIS_WEEK_MON_TODAY, LAST_WEEK, LAST_WEEK, LAST_BUSINESS_WEEK, LAST_WEEK_SUN_SAT, THIS_MONTH, LAST_MONTH 

// Or specify a custom date range. Format is: yyyy-mm-dd
var USE_CUSTOM_DATE_RANGE = false;
var START_DATE = "<Date in yyyy-mm-dd format>"; // Example "2016-02-01"
var END_DATE = "<Date in yyyy-mm-dd format>"; // Example "2016-02-29"

// Set this to true to only look at currently active campaigns. 
// Set to false to include campaigns that had impressions but are currently paused.
var IGNORE_PAUSED_CAMPAIGNS = true;
 
// Set this to true to only look at currently active ad groups.
// Set to false to include ad groups that had impressions but are currently paused.  
var IGNORE_PAUSED_ADGROUPS = true;

var IGNORE_PAUSED_KEYWORDS = true;
var REMOVE_ZERO_IMPRESSIONS_KW = true;

// Number of top keywords (by impressions) to export to spreadsheet
var MAX_KEYWORDS = 1000000;

/*-- More filter for MCC account --*/
//Is your account a MCC account
var IS_MCC_ACCOUNT = true;

var FILTER_ACCOUNTS_BY_LABEL = false;
var ACCOUNT_LABEL_TO_SELECT = "INSERT_LABEL_NAME_HERE";

var FILTER_ACCOUNTS_BY_IDS = false;
var ACCOUNT_IDS_TO_SELECT = ['INSERT_ACCOUNT_ID_HERE', 'INSERT_ACCOUNT_ID_HERE'];
/*---------------------------------*/

//The script is expected to work with following API version
var API_VERSION = {
  apiVersion: 'v201809'
}
//////////////////////////////////////////////////////////////////////////////
function main() {
  var reportsFolder = getFolder(REPORTS_FOLDER_PATH);
  
  if (!IS_MCC_ACCOUNT) {
    processCurrentAccount(reportsFolder);
  } else {
    var childAccounts  = getManagedAccounts();
    while(childAccounts .hasNext()) {
      var childAccount  = childAccounts .next()
      MccApp.select(childAccount);
      processCurrentAccount(reportsFolder);
    }
  }
  trackEventInAnalytics();
  Logger.log("Done!");
  Logger.log("=========================");
  Logger.log("All the reports are available in the Google Drive folder at following URL: ");
  Logger.log(reportsFolder.getUrl());
  Logger.log("=========================");
}

function getManagedAccounts() {
  var accountSelector = MccApp.accounts();
  if (FILTER_ACCOUNTS_BY_IDS) {
    accountSelector = accountSelector.withIds(ACCOUNT_IDS_TO_SELECT);
  }
  if (FILTER_ACCOUNTS_BY_LABEL) {
    accountSelector = accountSelector.withCondition("LabelNames CONTAINS '" + ACCOUNT_LABEL_TO_SELECT + "'")
  }
  return accountSelector.get();  
}

function processCurrentAccount(reportsFolder) {
  var adWordsAccount = AdWordsApp.currentAccount();
  var spreadsheet = getReportSpreadsheet(reportsFolder, adWordsAccount);

  var accountName = adWordsAccount.getName();
  var currencyCode = adWordsAccount.getCurrencyCode();
  Logger.log("Accesing AdWord account: " + accountName);
  Logger.log("Fetching data from AdWords..");
  var keywordReport = getKeywordReport();

  Logger.log("Computing..");
  var keywordArray = compute(keywordReport);
  var summary = computeSummary(keywordArray);

  Logger.log("Exporting results to spreadsheet..");
  var dateString = Utilities.formatDate(new Date(), adWordsAccount.getTimeZone(), "yyyyMMdd");
  // Export keyword level stats to a spreadsheet
  var newSheetName = dateString + "-Details";
  var sheet = spreadsheet.getSheetByName(newSheetName);
  if (sheet != null) {
    clearDataAndCharts(sheet);
  } else {
    sheet = spreadsheet.insertSheet(newSheetName, 0);
  }
  //keywordReport.exportToSheet(sheet);
  exportToSpreadsheet(keywordArray, sheet, accountName);

  Logger.log("Exporting summary & charts to spreadsheet..");
  // Export summary data and charts to another sheet
  var newSummarySheetName = dateString + "-Summary";
  var summarySheet = spreadsheet.getSheetByName(newSummarySheetName);
  if (summarySheet != null) {
    clearDataAndCharts(summarySheet);
  } else {
    summarySheet = spreadsheet.insertSheet(newSummarySheetName, 0);
  }
  exportSummaryStatsToSpreadsheet(summary, summarySheet);
}

function compute(keywordReport) {
  var reportIterator = keywordReport.rows();
  var keywordArray = new Array();
  while (reportIterator.hasNext()) {
    var kw = reportIterator.next();
    keywordArray.push(kw);
  }
  Logger.log("Total keywords found: " + keywordArray.length);
  keywordArray.sort(getComparator("Impressions", true));
  
  // Truncate  the array after MAX_KEYWORDS limit
  keywordArray = keywordArray.slice(0, MAX_KEYWORDS);
  return keywordArray;
}

function getKeywordReport() {
  var dateRange = getDateRange(",");
  
  var whereStatements = "";
  if (IGNORE_PAUSED_CAMPAIGNS) {
    whereStatements += "AND CampaignStatus = ENABLED ";
  } else {
    whereStatements += "AND CampaignStatus IN ['ENABLED','PAUSED'] ";
  }
  
  if (IGNORE_PAUSED_ADGROUPS) {
    whereStatements += "AND AdGroupStatus = ENABLED ";
  } else {
    whereStatements += "AND AdGroupStatus IN ['ENABLED','PAUSED'] ";
  }

  if (IGNORE_PAUSED_KEYWORDS) {
    whereStatements += "AND Status = ENABLED "; 
  } else {
    whereStatements += "AND Status IN ['ENABLED','PAUSED'] ";
  }
  
  if (REMOVE_ZERO_IMPRESSIONS_KW) {
    whereStatements += "AND Impressions > 0 "; 
  }
  
  var query = "SELECT CampaignId, AdGroupId, Id, CampaignName, AdGroupName, Criteria, KeywordMatchType, QualityScore, SearchPredictedCtr, CreativeQualityScore, PostClickQualityScore,  Clicks, Impressions, Ctr, AverageCpc, Cost, Conversions, CostPerConversion, AveragePosition " +
    "FROM  KEYWORDS_PERFORMANCE_REPORT " +
    "WHERE IsNegative = FALSE AND HasQualityScore=true " + whereStatements +
    "DURING " + dateRange;

  return AdWordsApp.report(query, API_VERSION);
}

function getDateRange(seperator) {
  var dateRange = DATE_RANGE;
  if (USE_CUSTOM_DATE_RANGE) {
    dateRange = START_DATE.replace(/-/g, "") + seperator + END_DATE.replace(/-/g, "");
  } else if (dateRange.match(/LAST_(.*)_DAYS/)) {
    var adWordsAccount = AdWordsApp.currentAccount();
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var numDaysBack = parseInt(dateRange.match(/LAST_(.*)_DAYS/)[1]);
    var today = new Date();
    var endDate = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY), adWordsAccount.getTimeZone(), "yyyyMMdd");// Yesterday
    var startDate = Utilities.formatDate(new Date(today.getTime() - (MILLIS_PER_DAY * numDaysBack)), adWordsAccount.getTimeZone(), "yyyyMMdd");
    dateRange = startDate + seperator + endDate;
  }
  return dateRange;
}

function getComparator(sortFieldName, reverse) {
  return function(obj1, obj2) {
    var retVal = 0;
    var val1 = parseInt(obj1[sortFieldName], 10);
    var val2 = parseInt(obj2[sortFieldName], 10);
    if (val1 < val2)
      retVal = -1;
    else if (val1 > val2)
      retVal = 1;
    else 
      retVal = 0;
    
    if (reverse) {
      retVal = -1 * retVal;
    }
    return retVal;
  }
}

function exportToSpreadsheet(keywordArray, sheet, accountName) {
  var rowsArray = new Array();
  for (var i=0; i<keywordArray.length; i++) {
    var kw = keywordArray[i];
    if (kw["SearchPredictedCtr"]!="Not applicable") {
      rowsArray.push(
          [ kw["CampaignName"], kw["AdGroupName"], kw["Criteria"], kw["KeywordMatchType"], 
            kw["Clicks"], kw["Impressions"], kw["Ctr"], kw["AverageCpc"], kw["Conversions"], 
            kw["Cost"], kw["CostPerConversion"], kw["AveragePosition"], kw["QualityScore"], 
            kw["SearchPredictedCtr"].replace(/\s/g, "_").toUpperCase(), 
            kw["CreativeQualityScore"].replace(/\s/g, "_").toUpperCase(), 
            kw["PostClickQualityScore"].replace(/\s/g, "_").toUpperCase()
          ]
      );
    }
  }

  var colTitleColor = "#03cfcc"; // Aqua
  var summaryRowColor = "#D3D3D3"; // Grey
  var headers = ["Campaign Name", "Ad Group Name", "Keyword", "Match Type", "Clicks", "Impressions", "Ctr", "Avg CPC", "Conversions", "Cost", "Cost Per Conversion", "Avg Position", "Quality Score", "Expected CTR", "Ad Relevance", "Landing Page Experience"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(colTitleColor).setFontWeight("BOLD");
  if (rowsArray.length > 0) {
    sheet.getRange(2, 1, rowsArray.length, headers.length).setValues(rowsArray);
    applyColorCoding(sheet, rowsArray, 13, 1);
    applyColorCoding(sheet, rowsArray, 14, 1);
    applyColorCoding(sheet, rowsArray, 15, 1);
    sheet.getRange(1, 5, rowsArray.length).setNumberFormat("#,##0");
    sheet.getRange(1, 6, rowsArray.length).setNumberFormat("#,##0");
    sheet.getRange(1, 8, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 9, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 10, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 11, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 12, rowsArray.length).setNumberFormat("#0.0");
    sheet.getRange(1, 13, rowsArray.length).setNumberFormat("#,##0");
  }
  for(var i=1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);  
  }
  sheet.setFrozenRows(1);
}

function applyColorCoding(sheet, rowsArray, colIdx, rowOffset) {
  for(var i=0; i<rowsArray.length; i++) {
    var bgColor = getBgColorFor(rowsArray[i][colIdx]);
    sheet.getRange(i+rowOffset+1, colIdx+1).setBackground(bgColor);
  }
}

function getBgColorFor(value) {
  if (value == "BELOW_AVERAGE") return "#FACDB1"; // Red
  if (value == "ABOVE_AVERAGE") return "#CDF6BE"; // Green
  if (value == "AVERAGE") return "#EAEAEA";       // Grey
}

/* ******************************************* */
function computeSummary(keywordArray) {
  var summary = getNewEmptySummaryObject();
  for(var i=0; i < keywordArray.length; i++) {
    addStats(summary, keywordArray[i]);
  }
  setAverages(summary);
  return summary;
}

function addStats(summary, kw) {
  var clicks = cleanAndParseInt(kw["Clicks"], 10);
  var impressions = cleanAndParseInt(kw["Impressions"], 10);
  var cost = cleanAndParseFloat(kw["Cost"], 10);
  var conversions = cleanAndParseFloat(kw["Conversions"], 10);
  groupByValue(clicks, impressions, conversions, cost, summary.byQS, parseInt(kw["QualityScore"]));
  groupByValue(clicks, impressions, conversions, cost, summary.byLPE, kw["PostClickQualityScore"]);
  groupByValue(clicks, impressions, conversions, cost, summary.byExpectedCTR, kw["SearchPredictedCtr"]);
  groupByValue(clicks, impressions, conversions, cost, summary.byAdRelevance, kw["CreativeQualityScore"]);
}

function groupByValue(clicks, impressions, conversions, cost, map, key) {
  var aggStats = map[key];
  if (!aggStats) {
    aggStats = getNewEmptyStatsObject();
    map[key] = aggStats;
  }
  aggStats["Clicks"] += clicks;
  aggStats["Impressions"] += impressions;
  aggStats["Conversions"] += conversions;
  aggStats["Cost"] += cost;
};

function setAverages(summary) {
  for(groupByName in summary) {
    var map = summary[groupByName]
    for(key in map) {
      var stats = map[key];
      if (stats["Clicks"] > 0) {
        stats["AverageCpc"]= stats["Cost"] / stats["Clicks"];
      }
      if (stats["Conversions"])
      stats["CostPerConversion"] = stats["Cost"] / stats["Conversions"];
    }
  }
}

function getNewEmptyStatsObject() {
  return {
    "Clicks":0,
    "Impressions":0,
    "Conversions":0,
    "Cost":0,
    "AverageCpc":0,
    "CostPerConversion":0
  };
}

function getNewEmptySummaryObject() {
  return {
    "byQS": new Object(),
    "byLPE": new Object(),
    "byExpectedCTR": new Object(),
    "byAdRelevance": new Object()
  };
};

function exportSummaryStatsToSpreadsheet(summary, sheet) {
  var startRow = 1;
  exportAggStats(sheet, summary.byQS, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], startRow, "Quality Score", "#FFFF80", Charts.ChartType.COLUMN);
  
  var qsParamValues = ["Above average", "Average", "Below average"];
  var chartOptions = {"slices":{"0":{"color":"#00ff00"},"1":{"color":"#3366cc"},"2":{"color":"#dc3912"}}};
  exportAggStats(sheet, summary.byLPE, qsParamValues, startRow+= 13, "Lading Page Experience", "#f9cda8", Charts.ChartType.PIE, chartOptions);
  exportAggStats(sheet, summary.byExpectedCTR, qsParamValues, startRow+=6, "Expected CTR", "#a8d4f9", Charts.ChartType.PIE, chartOptions);
  exportAggStats(sheet, summary.byAdRelevance, qsParamValues, startRow+=6, "Ad Relevance", "#9aff9a", Charts.ChartType.PIE, chartOptions);
}

function exportAggStats(sheet, map, keyArray, startRow, keyColHeader, bgColor, chartType, chartOptions) {
  var rowsArray = new Array();
  for (var i=0; i < keyArray.length; i++) {
    var key = keyArray[i];
    var stats = map[key];
    if (!stats) {
      stats = getNewEmptyStatsObject();
    }
    rowsArray.push(
      [
        key, stats["Clicks"], stats["Impressions"], stats["Cost"], stats["AverageCpc"],  stats["Conversions"],  stats["CostPerConversion"]
      ]
    );
  }
  
  var headerRow = startRow;
  var headers = [keyColHeader, "Clicks", "Impressions", "Cost", "Avg CPC", "Conversions",  "Cost Per Conversion"];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setBackground(bgColor).setFontWeight("BOLD");

  var firstDataRow = headerRow + 1;
  sheet.getRange(firstDataRow, 1, rowsArray.length, headers.length).setValues(rowsArray).setBackground(bgColor);
//  sheet.getRange(firstDataRow, 1, rowsArray.length).setFontWeight("BOLD");
  sheet.getRange(firstDataRow, 2, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 3, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 4, rowsArray.length).setNumberFormat("#,##0.00");
  sheet.getRange(firstDataRow, 5, rowsArray.length).setNumberFormat("#,##0.00");
  sheet.getRange(firstDataRow, 6, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 7, rowsArray.length).setNumberFormat("#,##0.00");

  for(var i=1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);  
  }
  
  drawChart(sheet, chartType, startRow, rowsArray.length, keyColHeader, chartOptions);
}

function drawChart(sheet, chartType, startRow, numOfRows, colName2, chartOptions) {
  // Creates a column chart for values in range
  var colName1 = "Impressions";
  var range1 = sheet.getRange(startRow+1, 1, numOfRows, 1);
  var range2 = sheet.getRange(startRow+1, 3, numOfRows, 1);

  var chartBuilder = sheet.newChart();
  chartBuilder.addRange(range1).addRange(range2).setChartType(chartType);
  var chartTitle = colName1 + " Vs. " + colName2;
  if (chartType === Charts.ChartType.PIE) {
    chartTitle = colName1 + " Chart for " + colName2;
  }
  chartBuilder.setOption('title', chartTitle);
  chartBuilder.setOption('useFirstColumnAsDomain', true);
  chartBuilder.setOption('legend', {position: 'bottom'});
  chartBuilder.setOption("treatLabelsAsText", true);
  chartBuilder.setOption("hAxis", { "useFormatFromData": true, "title": colName2});
  chartBuilder.setOption("vAxis", { "useFormatFromData": true, "title": colName1});
  if (chartOptions) {
    for(chartOptionName in chartOptions) {
      chartBuilder.setOption(chartOptionName, chartOptions[chartOptionName]);
    }
  }

  chartBuilder.setPosition(startRow, 8, 70-(2*startRow), 2);
  sheet.insertChart(chartBuilder.build());
}

/*
 * Gets the report file (spreadsheet) for the given Adwords account in the given folder.
 * Creates a new spreadsheet if doesn't exist.
 */
function getReportSpreadsheet(folder, adWordsAccount) {
  var accountId = adWordsAccount.getCustomerId();
  var accountName = adWordsAccount.getName();
  var spreadsheet = undefined;
  var files = folder.searchFiles(
      'mimeType = "application/vnd.google-apps.spreadsheet" and title contains "'+ accountId + '"');
  if (files.hasNext()) {
    var file = files.next();
    spreadsheet = SpreadsheetApp.open(file);
  }
  
  if (!spreadsheet) {
    var fileName = accountName + " (" + accountId + ")";
    spreadsheet = SpreadsheetApp.create(fileName);
    var file = DriveApp.getFileById(spreadsheet.getId());
    var oldFolder = file.getParents().next();
    folder.addFile(file);
    oldFolder.removeFile(file);
  }
  return spreadsheet;
}

/*
* Gets the folder in Google Drive for the given folderPath.  
* Creates the folder and all the internediate folders if needed.
*/
function getFolder(folderPath) {
  var folder = DriveApp.getRootFolder();
  var folderNamesArray = folderPath.split("/");
  for(var idx=0; idx < folderNamesArray.length; idx++) {
    var newFolderName = folderNamesArray[idx];
    // Skip if new folder name is empty (possiblly due to slash at the end) 
    if (newFolderName.trim() == "") { 
      continue;
    }
    var folderIterator = folder.getFoldersByName(newFolderName);
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      Logger.log("Creating folder '" + newFolderName + "'");
      folder = folder.createFolder(newFolderName);
    }
  }
  return folder;
}

function clearDataAndCharts(sheet) {
  sheet.clear();
  var charts = sheet.getCharts();
  for(var i=0; i<charts.length; i++) {
    sheet.removeChart(charts[i]);
  }
}
/* ******************************************* */
function cleanAndParseFloat(valueStr) {
  valueStr = cleanValueStr(valueStr);
  return parseFloat(valueStr);
}

function cleanAndParseInt(valueStr) {
  valueStr = cleanValueStr(valueStr);
  return parseInt(valueStr);
}

function cleanValueStr(valueStr) {
  if (valueStr.charAt(valueStr.length - 1) == '%') {
    valueStr = valueStr.substring(0, valueStr.length - 1);
  }
  valueStr = valueStr.replace(/,/g,'');
  return valueStr;
}

/* 
* Tracks the execution of the script as an event in Google Analytics.
* Sends the script name and a random UUID (it is just a random number, required by Analytics).
* The event information just tells us that somewhere someone ran this script. 
* Credit for the idea goes to Russel Savage, who posted his version at http://www.freeadwordsscripts.com/2013/11/track-adwords-script-runs-with-google.html.
* and to Martin Roettgerding, as we learnt the trick from his script at http://www.ppc-epiphany.com/2016/03/11/introducing-the-quality-score-tracker-v3-0
*/
function trackEventInAnalytics() {
  // Create the random UUID from 30 random hex numbers gets them into the format xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx (with y being 8, 9, a, or b).
  var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, 
    function(c) {var r = Math.random()*16|0,v=c=='x'?r:r&0x3|0x8;return v.toString(16);});
  var url = "http://www.google-analytics.com/collect?v=1&t=event&tid=UA-46662882-1&cid=" + uuid 
  + "&ec=AdWords%20Scripts&ea=Script%20Execution&el=QS%20Analysis";
  UrlFetchApp.fetch(url);
}