/**
 * AutomatingAdWords.com - Auto-negative keyword adder
 *
 * This script will automatically add negative keywords based on specified criteria
 * Go to automatingadwords.com for installation instructions & advice
 *
 * Version: 1.7.0
 */

//You may also be interested in this Chrome keyword wrapper!
//https://chrome.google.com/webstore/detail/keyword-wrapper/paaonoglkfolneaaopehamdadalgehbb

var SPREADSHEET_URL = "your-settings-url-here"; //template here: https://docs.google.com/spreadsheets/d/1G8mjtESGW3O1Jtd3l9MvqPA93cwUvFdVwkWsHOvT0W4/edit#gid=0
var SS = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

// Types of compaign to iterate over. Processing for each campaign / ad group
// will be attempted for all types. (We assume that a campaign with a certain
// name can only be one of those types, so all except one will immediately warn
// and continue before doing any real processing.) at the moment, there's only
// a difference between "Shopping" and any other type.
var processCampaignTypes = ["Shopping", "Text"];

// Threshold for logs. Only messages with this level and higher will be logged.
var LOG_THRESHOLD = 2;

var LOGLEVEL_TRACE = 0;
var LOGLEVEL_DEBUG = 1;
var LOGLEVEL_INFO = 2;
var LOGLEVEL_WARN = 3;
var LOGLEVEL_ERROR = 4;

function main() {
  var timeStampCol = 7;
  var sheets = SS.getSheets();
  var sheet;

  // Get array of recorded timestamp per sheet, and sheet name per timestamp.
  // This enables us to start processing the sheet that was processed the
  // longest ago. NOTES:
  // - These timestamps are not used for the query to gather latest queries.
  // - If multiple sheets have the same timestamp, only one of those sheets
  //   will be processed; the unprocessed sheet will be done the next time this
  //   script runs.
  var sheetNamesByTime = {};
  var timesBySheet = [];
  for (var sheetNo in sheets) {
    sheet = sheets[sheetNo];
    var timestamp = sheet.getRange(1, timeStampCol).getValue();

    if (timestamp) {
      sheetNamesByTime[timestamp] = sheet.getName();
      timesBySheet.push(timestamp);
    } else {
      // First run. Add slightly different timestamps because of above note.
      var oldenTimes = new Date();
      oldenTimes.setDate(oldenTimes.getDate() - (1000 + timesBySheet.length));
      sheetNamesByTime[oldenTimes] = sheet.getName();
      timesBySheet.push(oldenTimes);
    }
  }

  // Loop through sheets, starting with the earliest time.
  timesBySheet = timesBySheet.sort(function(a, b){return a-b});
  for (var iSheet in timesBySheet) {
    var sheetName = sheetNamesByTime[timesBySheet[iSheet]];
    log("Checking sheet: " + sheetName);
    sheet = SS.getSheetByName(sheetName);

    var SETTINGS = {};
    SETTINGS["CAMPAIGN_NAME"] = sheet.getRange("A2").getValue();
    SETTINGS["MIN_QUERY_CLICKS"] = sheet.getRange("B2").getValue();
    SETTINGS["MAX_QUERY_CONVERSIONS"] = sheet.getRange("C2").getValue();
    SETTINGS["DATE_RANGE"] = sheet.getRange("D2").getValue();
    SETTINGS["NEGATIVE_MATCH_TYPE"] = sheet.getRange("E2").getValue();
    SETTINGS["CAMPAIGN_LEVEL_QUERIES"] = sheet.getRange("F2").getValue() === "Yes";
    log("Settings: " + JSON.stringify(SETTINGS), LOGLEVEL_DEBUG);

    var numMatchesRow = 4;
    var adGroupNameRow = 5;
    var firstAdGroupRow = 6;

    // Loop through the ad groups listed in the sheet. Use the numMatches row
    // for this; we assume this is always populated.
    var currentAdgroupCol = 2;
    while (sheet.getRange(numMatchesRow, currentAdgroupCol).getValue()) {
      var adGroupMin = sheet.getRange(numMatchesRow, currentAdgroupCol).getValue();
      var adGroupName = sheet.getRange(adGroupNameRow, currentAdgroupCol).getValue();
      log("Checking campaign: " + SETTINGS["CAMPAIGN_NAME"] +"; ad group: " + adGroupName);

      // Iterate through types of campaigns. (We assume that a campaign with a
      // certain name can only be one of those types, so all except one will
      // immediately warn and continue before doing any real processing.)
      for (var t in processCampaignTypes) {
        var campaignType = processCampaignTypes[t];

        // Get AdGroup object; check if this campaign / ad group exists. NOTE:
        // even when we get queries by campaign level, negative keywords are
        // still added for each specified ad group (not for a campaign). So we
        // still iterate through all ad groups and check their names.
        if (campaignType === "Shopping") {
          var adGroupIterator = AdWordsApp.shoppingAdGroups()
        } else {
          var adGroupIterator = AdWordsApp.adGroups()
        }
        adGroupIterator = adGroupIterator
          .withCondition("Name = '" + adGroupName + "'")
          .withCondition("CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'")
          .get();

        if (!adGroupIterator.hasNext()) {
          log("Ad group '" + adGroupName + "' in campaign '" + SETTINGS["CAMPAIGN_NAME"] + "' (campaign type " + campaignType + ") not found in the account. Check if the ad group / campaign name is correct in the sheet.", LOGLEVEL_WARN);
          continue;
        }
        var adGroup = adGroupIterator.next();

        // Get the 'positive' keywords from the sheet for this campaign / ad
        // group. NOTE: if the row contains duplicate keywords, these will
        // count as two matches, which is important if adGroupMin > 1!
        var keywords = [];
        var row = firstAdGroupRow;
        while (sheet.getRange(row, currentAdgroupCol).getValue()) {
          keywords.push(sheet.getRange(row, currentAdgroupCol).getValue());
          row++;
        }
        log("Got " + keywords.length + " 'positive keywords' from sheet: " + keywords, LOGLEVEL_DEBUG);

        // Get the search queries from the campaign
        var report = null;
        log("Getting report of search queries...", LOGLEVEL_DEBUG);
        var query = "SELECT Query" +
          " FROM SEARCH_QUERY_PERFORMANCE_REPORT" +
          " WHERE CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'";
        if (SETTINGS["MIN_QUERY_CLICKS"] != "") {
          query += " AND Clicks > " + SETTINGS["MIN_QUERY_CLICKS"]
        }
        if (SETTINGS["MAX_QUERY_CONVERSIONS"] != "") {
          query += " AND Conversions < " + SETTINGS["MAX_QUERY_CONVERSIONS"]
        }
        if (!SETTINGS["CAMPAIGN_LEVEL_QUERIES"]) {
          query += " AND AdGroupName = '" + adGroupName + "'"
        }
        if (SETTINGS["DATE_RANGE"] != "ALL_TIME") {
          query += " DURING " + SETTINGS["DATE_RANGE"]
        }
        log("Query: " + query, LOGLEVEL_TRACE);
        report = AdWordsApp.report(query);

        var rows = report.rows();
        var negs = [];
        // Loop through this campaign's queries; add anything which doesn't
        // contain our positive keywords to the negs array. (These will be added
        // as negatives later.) Note we're doing this every time while iterating
        // through all types, with the same outcome - but that may change.
        while (rows.hasNext()) {
          var nxt = rows.next();
          var q = nxt.Query;
          var matches = 0;
          var count = 0;

          //loop through the positive keywords (from the sheet)
          for (var k in keywords) {
            //if > min (e.g. 2) of the keywords are in the search term, then neg it
            // log("Checking against keyword: " + keywords[k]);
            // log("Checking against term:    " + q);
            count++;
            //if the keyword is in the query, we have a match. match++
            if (q.indexOf(keywords[k]) > -1) {
              //log(nxt.Query + " - " + q.indexOf(keywords[k]) + " - " + keywords[k] );
              matches++;
            }
            // log("matches: " + matches);
            // log("count: " + count + " - " + keywords.length);

            //if we have reached the end of the positive keywords i.e. checked them all
            //and if the number of matches is leSS than the minimum number of matches for the adgroup (specified on the sheet)
            //then add the query to the negatives array
            if (matches < adGroupMin && count == keywords.length) {
              // log(count + " - " + keywords.length);
              //log("adding negative: " + nxt.Query);
              negs.push(q);
              break;
            }
          }
        }
        log("Found a total of " + negs.length + " negative keywords to add.");

        // Now add the new negative keywords to the ad group.
        log("Adding the negative keywords...");
        for (var neg in negs) {
          neg = addMatchType(negs[neg], SETTINGS);
          adGroup.createNegativeKeyword(neg);
        }
      }//end ad types loop

      currentAdgroupCol++;
    }

    // Set timestamp for sorting sheets next time.
    var date = new Date();
    sheet.getRange(1, timeStampCol).setValue(date);
  }//end time array (sheets) loop

  log("Finished.")
}//end main

function addMatchType(word, SETTINGS) {
  if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "broad") {
    word = word.trim();
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "bmm") {
    word = word.split(" ").map(function (x) {
      return "+" + x
    }).join(" ").trim()
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "phrase") {
    word = '"' + word.trim() + '"'
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "exact") {
    word = '[' + word.trim() + ']'
  } else {
    throw("Error: Match type not recognised. Please provide one of Broad, BMM, Exact or Phrase")
  }
  return word;
}

function log(message, level) {
  if (level === undefined) {
    level = LOGLEVEL_INFO;
  }

  if (level >= LOG_THRESHOLD) {
    Logger.log(message)
  }
}