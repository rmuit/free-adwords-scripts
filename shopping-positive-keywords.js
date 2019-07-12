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

// This affects script behavior IF "Campaign Level Queries" in the sheet is
// "Yes". (If "No", this constant is not used.)
// "Campaign Level Queries" means the user queries are read from all ad groups
// in the campaign.
// If this constant is false, all those queries are still checked against a
// list of keywords in the sheet, for each ad group - and negative keywords are
// created for each ad group.
// If this constant is true, only one row of keywords (row B) is checked in the
// sheet; no ad group is used / needs to be defined in the sheet; and the
// negative keywords are created for the campaign as a whole.
var CAMPAIGN_LEVEL_KEYWORDS = false;

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
    var lastRow = sheet.getLastRow();

    var SETTINGS = {};
    SETTINGS["CAMPAIGN_NAME"] = sheet.getRange("A2").getValue();
    SETTINGS["MIN_QUERY_CLICKS"] = sheet.getRange("B2").getValue();
    SETTINGS["MAX_QUERY_CONVERSIONS"] = sheet.getRange("C2").getValue();
    SETTINGS["DATE_RANGE"] = sheet.getRange("D2").getValue();
    SETTINGS["NEGATIVE_MATCH_TYPE"] = sheet.getRange("E2").getValue();
    SETTINGS["CAMPAIGN_LEVEL_QUERIES"] = sheet.getRange("F2").getValue() === "Yes";
    // CAMPAIGN_LEVEL_KEYWORDS constant should not be used further; only setting.
    SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"] = SETTINGS["CAMPAIGN_LEVEL_QUERIES"] && CAMPAIGN_LEVEL_KEYWORDS;
    log("Settings: " + JSON.stringify(SETTINGS), LOGLEVEL_DEBUG);

    var numMatchesRow = 4;
    var adGroupNameRow = 5;
    var firstAdGroupRow = 6;

    // Loop through the ad groups listed in the sheet. Use the numMatches row
    // for this; we assume this is always populated.
    var currentAdgroupCol = 2;
    while (sheet.getRange(numMatchesRow, currentAdgroupCol).getValue()) {
      var minKeywordMatches = sheet.getRange(numMatchesRow, currentAdgroupCol).getValue();
      var adGroupName;
      if (!SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"]) {
        adGroupName = sheet.getRange(adGroupNameRow, currentAdgroupCol).getValue();
      }

      // Iterate through types of campaigns. (We assume that a campaign with a
      // certain name can only be one of those types, so all except one will
      // immediately warn and continue before doing any real processing.)
      for (var t in processCampaignTypes) {
        var campaignType = processCampaignTypes[t];
        var selector;
        var iterator;

        // Get AdGroup or Campaign object, to add negative keywords to. This
        // also checks for existence of this campaign / ad group.
        if (campaignType === "Shopping") {
          selector = SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"] ? AdWordsApp.shoppingCampaigns() : AdWordsApp.shoppingAdGroups();
        } else {
          selector = SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"] ? AdWordsApp.campaigns() : AdWordsApp.adGroups();
        }
        selector = selector.withCondition("CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'");
        if (!SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"]) {
          selector = selector.withCondition("Name = '" + adGroupName + "'");

          iterator = selector.get();
          if (!iterator.hasNext()) {
            log("Ad group '" + adGroupName + "' in campaign '" + SETTINGS["CAMPAIGN_NAME"] + "' (campaign type " + campaignType + ") not found in the account. Check if the ad group / campaign name is correct in the sheet.", LOGLEVEL_WARN);
            continue;
          }
          log("Checking campaign: " + SETTINGS["CAMPAIGN_NAME"] +"; ad group: " + adGroupName);
        } else {
          iterator = selector.get();
          if (!iterator.hasNext()) {
            log("Campaign '" + SETTINGS["CAMPAIGN_NAME"] + "' (campaign type " + campaignType + ") not found in the account. Check if the campaign name is correct in the sheet.", LOGLEVEL_WARN);
            continue;
          }
          log("Checking campaign: " + SETTINGS["CAMPAIGN_NAME"]);
        }
        var adGroupOrCampaign = iterator.next();

        // Get the 'positive' keywords from the sheet for this campaign / ad
        // group. NOTE: if the row contains duplicate keywords, these will
        // count as two matches, which is important if minKeywordMatches > 1!
        var keywords = [];
        var values = sheet.getRange(firstAdGroupRow, currentAdgroupCol, lastRow).getValues();
        var row = 0;
        while (values[row][0]) {
          keywords.push(values[row][0]);
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
          var queryString = nxt.Query;
          var matches = 0;

          // Loop through the positive keywords (from the sheet); if we find
          // enough matches then stop processing.
          for (var k in keywords) {
            if (containsKeyword(queryString, keywords[k])) {
              matches++;
              // Log 'trace' so that if we select DEBUG as threshold, we only
              // get a list of 'to be added' keywords.
              if (matches >= minKeywordMatches) {
                log("Query '" + queryString + "' contains positive keyword '" + keywords[k] + "'; skipping.", LOGLEVEL_TRACE);
                break;
              }
              log("Query '" + queryString + "' contains positive keyword '" + keywords[k] + "', but continuing (" + matches + " < " + minKeywordMatches + " matches).", LOGLEVEL_TRACE);
            }
          }

          // Add as a negative keyword if we did not find enough matches.
          if (matches < minKeywordMatches) {
            // Use 'debug' level because every script will try to re-add the
            // same keywords, so it's not extremely useful.
            log("Query '" + queryString + "' will be added as negative keyword.", LOGLEVEL_DEBUG);
            negs.push(queryString);
          }
        }

        // Now add the new negative keywords to the ad group / campaign.
        if (negs.length) {
          log("Adding a total of " + negs.length + " negative keywords...");
          for (var neg in negs) {
            neg = addMatchType(negs[neg], SETTINGS);
            adGroupOrCampaign.createNegativeKeyword(neg);
          }
        } else {
          log("Found no new negative keywords to add.");
        }
      }//end ad types loop

      // For campaign level queries, check only column B; we don't want to redo
      // the exact same query report (even with e.g. different nr of matches).
      if (SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"]) {
        log("Not checking further rows in this sheet, since we added negative keywords at the campaign level.", LOGLEVEL_DEBUG);
        break;
      }
      currentAdgroupCol++;
    }

    // Set timestamp for sorting sheets next time.
    var date = new Date();
    sheet.getRange(1, timeStampCol).setValue(date);
  }//end time array (sheets) loop

  log("Finished.")
}//end main

/**
 * Checks if a keyword matches a query string. Not just as a substring,
 * but the whole keyword must match (a) whole word(s) in the query. (Keyword
 * can also be multiple words, which then must be contained in the string in
 * the same order.)
 */
function containsKeyword(string, keyword) {
  // If the string is equal to the keyword, it 'contains' the keyword.
  var match = string === keyword;
  if (!match) {
    // Otherwise, for speed, first search if the keyword is a substring. If so,
    // then go on to determine if the string actually contains this keyword.
    var index = string.indexOf(keyword);
    match = index >= 0;
    if (match) {
      // Check: a keyword which is a substring should not trigger a match. For
      // simplicity, we assume the only non-word characters in the query string
      // (for our purpose) are spaces. So check for spaces or string start/end
      // on both sides.
      var positionAfterMatch = index + keyword.length;
      match = ((index === 0 || string[index - 1] === ' ')
        && (positionAfterMatch === string.length || string[positionAfterMatch] === ' '));
    }
  }
  return match;
}

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
