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
// negative keywords are created for the campaign as a whole. (Any negative
// keywords tht may exist for specific ad groups already are disregarded.)
var CAMPAIGN_LEVEL_KEYWORDS = false;

// Threshold for logs. Only messages with this level and higher will be logged.
var LOG_THRESHOLD = 2;

var LOGLEVEL_TRACE = 0;
var LOGLEVEL_DEBUG = 1;
var LOGLEVEL_INFO = 2;
var LOGLEVEL_WARN = 3;
var LOGLEVEL_ERROR = 4;

// Variable containing positive keywords. Defined outside of main() so we don't
// need to pass it (by value) into functions.
var positiveKeywords = [];

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
        positiveKeywords = [];
        var values = sheet.getRange(firstAdGroupRow, currentAdgroupCol, lastRow).getValues();
        var row = 0;
        while (values[row][0]) {
          positiveKeywords.push(values[row][0]);
          row++;
        }
        log("Got " + positiveKeywords.length + " 'positive keywords' from sheet: " + positiveKeywords, LOGLEVEL_DEBUG);

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

        // Get the existing negative keywords for this campaign / ad group, to
        // check whether keywords from the report are already added. (Note
        // keywords targeted to an ad group are a supergroup of queries
        // targeted to its campaign - or they are a separate group in the API;
        // I'm not sure yet. In the Web UI, they are displayed like a supergroup.
        // This unlike the above report for queries; queries in an ad group are
        // a subgroup of queries for its campaign.)
        //
        // *NOTE* - THIS DOES NOT WORK YET, UNLESS SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"] IS TRUE,
        // because it doesn't work for Shopping: we can't get to existing negative keywords for ad groups.
        // I don't know why exactly, but
        // - Methods ShoppingAdGroup.negativeKeywords() / negativeKeywordLists() don't exist
        //   (and they do for AdGroup, Campaign and ShoppingCampaign)
        // - I don't know the fieldname to use for ShoppingCampaign.negativeKeywords().withCondition() to select
        //   on AdGroup. (It's not 'AdGroup' or 'AdGroupName'.)
        // - But ShoppingAdGroup.createNegativeKeyword _does_ exist, so
        //   "negative keywords for a Shopping ad group" is clearly a thing. As
        //   olso proven by this script which didn't even have
        //   "CAMPAIGN_LEVEL_KEYWORDS" previously.
        // Until this mystery is solved and/or we want to start testing this for
        // 'non-Shopping' ad groups, Let's if() it. This will just mean that
        // the previous behavior of trying again and again to add the same
        // negative keyword, still exists. (If someone wants to test this for
        // 'non-Shopping' ad groups, go ahead and change the if().)
        var checkExistingNegKeywords = SETTINGS["CAMPAIGN_LEVEL_KEYWORDS"];
        var negKeywords = {};
        if (checkExistingNegKeywords) {
          // If the 'match type' of a negative keyword that we find in the
          // map is the same as our setting, we assume we can safely remove it.
          // (This is so that we can remove negative keywords automatically,
          // which were added by an earlier run of this script, after which the
          // keyword was added as a positive one in the sheet to correct the
          // situation. Also negative keywords that consist of multiple words
          // and contain the positive keyword, will be removed automatically.)
          // Otherwise we only warn and leave it to the user to deal with.
          var removeMatchType = SETTINGS["NEGATIVE_MATCH_TYPE"].toUpperCase();

          log("Getting existing negative keywords...", LOGLEVEL_DEBUG);
          var negKeywordCount = 0;
          var negativeKeyword;
          iterator = adGroupOrCampaign.negativeKeywords().get();
          while (iterator.hasNext()) {
            negativeKeyword = iterator.next();
            if (checkNegKeywordAgainstPos(negativeKeyword, removeMatchType)) {
              negKeywords[negativeKeyword.getText()] = 1;
              negKeywordCount++;
            }
          }
          var negKeywordNonListCount = negKeywordCount;

          // Also get keywords from any lists. (This might mean we're re-reading
          // the same list if it has been used in another campaign we previously
          // processed; there's no caching for this.)
          iterator = adGroupOrCampaign.negativeKeywordLists().get();
          // var negKeywordListCount = iterator.totalNumEntities;
          while (iterator.hasNext()) {
            var list = iterator.next();
            // If we find a keyword from this list in the map, don't
            // automatically remove it; just warn.
            removeMatchType = "*" + list.getName();

            var iterator2 = list.negativeKeywords().get();
            while (iterator2.hasNext()) {
              negativeKeyword = iterator2.next();
              if (checkNegKeywordAgainstPos(negativeKeyword, removeMatchType)) {
                negKeywords[negativeKeyword.getText()] = 1;
                negKeywordCount++;
              }
            }
          }
          log("Got " + negKeywordCount + " existing negative keywords (" + (negKeywordCount - negKeywordNonListCount) + " of which are from keyword lists).", LOGLEVEL_DEBUG);
          log(JSON.stringify(negKeywords), LOGLEVEL_TRACE);
        }

        var rows = report.rows();
        var negs = [];
        // Loop through this campaign's queries; add anything which doesn't
        // contain our positive keywords to the negs array. (These will be added
        // as negatives later.) Note we're doing this every time while iterating
        // through all types, with the same outcome - but that may change.
        while (rows.hasNext()) {
          var nxt = rows.next();
          var queryString = nxt.Query;
          var queryWithMatchType = addMatchType(queryString, SETTINGS);

          if (matchingNegativeKeyword(queryWithMatchType, negKeywords)) {
            // Log 'trace' so that if we select DEBUG as threshold, we only
            // get a list of 'to be added' keywords.
            log("Query '" + queryString + "' exists as negative keyword; skipping.", LOGLEVEL_TRACE);
            continue;
          }

          // Loop through the positive keywords (from the sheet); if we find
          // enough matches then stop processing.
          var matches = 0;
          for (var k in positiveKeywords) {
            if (containsKeyword(queryString, positiveKeywords[k])) {
              matches++;
              if (matches >= minKeywordMatches) {
                log("Query '" + queryString + "' contains positive keyword '" + positiveKeywords[k] + "'; skipping.", LOGLEVEL_TRACE);
                break;
              }
              log("Query '" + queryString + "' contains positive keyword '" + positiveKeywords[k] + "', but continuing (" + matches + " < " + minKeywordMatches + " matches).", LOGLEVEL_TRACE);
            }
          }

          // Add as a negative keyword if we did not find enough matches.
          if (matches < minKeywordMatches) {
            negs.push(queryWithMatchType);
            // If we're not checking against existing negative keywords, the
            // script will try to re-add the same keywords every time, so this
            // log is not extremely useful.
            log("Query '" + queryString + "' will be added as negative keyword.", checkExistingNegKeywords ? LOGLEVEL_INFO : LOGLEVEL_DEBUG);
          }
        }

        // Now add the new negative keywords to the ad group / campaign.
        if (negs.length) {
          log("Adding a total of " + negs.length + " negative keywords...");
          for (var neg in negs) {
            adGroupOrCampaign.createNegativeKeyword(negs[neg]);
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

/**
 * Checks a negative keyword against the map of positive keywords; if it is
 * found (meaning that the negative keyword contains any positive keyword), we
 * remove it or log a warning.
 *
 * (So if a positive keyword is multiple words and contains the negative keyword
 * which is 'smaller', then the negative keyword isn't removed.)
 *
 * @param negativeKeyword
 *   AdsApp.​NegativeKeyword object
 * @param removeMatchType
 *   Remove the keyword if it's equal to this match type; otherwise warn. Must
 *   be uppercase. Pass "*<LISTNAME>" to indicate that it should never be
 *   removed because it's part of a negative keyword list; the list name will
 *   be used in the warning message.
 *
 * @return
 *   True if the negative keyword still exists, False if it was removed.
 */
function checkNegKeywordAgainstPos(negativeKeyword, removeMatchType) {
  var exists = true;
  var string = stripKeywordModifiers(negativeKeyword.getText());
  // We're using the file-scoped positiveKeywords variable.
  for (var k in positiveKeywords) {
    if (containsKeyword(string, positiveKeywords[k])) {
      // Negative keyword is equal to / contains a positive keyword in the sheet.
      // Check if we should remove it or just warn.
      if (removeMatchType[0] === '*') {
        log("Keyword '" + positiveKeywords[k] + "' found in sheet has a related negative keyword '" + negativeKeyword.getText() + "', which should likely be removed! The negative keyword is part of a list named '" + removeMatchType.substr(1) + "'.", LOGLEVEL_WARN);
      } else if (removeMatchType !== negativeKeyword.getMatchType()) {
        log("Keyword '" + positiveKeywords[k] + "' found in sheet has a related negative keyword '" + negativeKeyword.getText() + "', which should likely be removed!", LOGLEVEL_WARN);
      } else {
        log("Keyword '" + positiveKeywords[k] + "' found in sheet has a related negative keyword '" + negativeKeyword.getText() + "'. Removing the negative keyword.", LOGLEVEL_WARN);
        negativeKeyword.remove();
        exists = false;
      }
      break;
    }
  }

  return exists;
}

/**
 * Returns True if a query string (already modified with match type) is found
 * in a map of keywords.
 */
function matchingNegativeKeyword(keywordWithMatchType, keywordMap) {
  // We could extend this logic, e.g. if queryWithMatchType == '[word]' and
  // 'word' or '+word' was already defined as a negative keyword, we could
  // regard that as a match. But let's keep it simple for now.
  return keywordMap[keywordWithMatchType] !== undefined;
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

/**
 * Strips keyword modifiers from a keyword, with the intention of
 * having a canonical keyword that we can search for individual words.
 */
function stripKeywordModifiers(keyword) {
  if ((keyword[0] === '[' && keyword[keyword.length - 1] === ']') ||
    (keyword[0] === '"' && keyword[keyword.length - 1] === '"')) {
    // Exact keywords or phrase. Modify nothing else.
    keyword = keyword.substring(1, keyword.length - 1);
  } else {
    // Double quotes and plus signs can both be used inside a search term;
    // they're just word modifiers. Since our caller is interested in just
    // the search words, we just remove them, as follows:
    // - Assume '+' is only used at the beginning of a word or is preceded by
    //   a space.
    // - Assume '"' is only used at the beginning/end of the full string or has
    //   a space before or after it, and there are always pairs in the string.
    // This means we can just remove all of them because they are never literal
    // parts of a keyword. If this is not the case... this code needs adjusting.
    if (keyword.indexOf('"') >= 0) {
      keyword = keyword.split('"').join(' ');
    }
    if (keyword.indexOf('+') >= 0) {
      keyword = keyword.split('+').join(' ');
    }
  }

  return keyword;
}

function log(message, level) {
  if (level === undefined) {
    level = LOGLEVEL_INFO;
  }

  if (level >= LOG_THRESHOLD) {
    Logger.log(message)
  }
}
