var SHEET_NAME = "Marketing Submissions";

function doPost(e) {
  try {
    var sheet = getOrCreateSheet(SHEET_NAME);
    var p = e.parameter;

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(getHeaders());
      sheet.getRange(1, 1, 1, getHeaders().length)
        .setFontWeight("bold")
        .setBackground("#022B6D")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    var row = [
      new Date(),
      p.fullName           || "",
      p.businessName       || "",
      p.whatYouDo          || "",
      p.primaryMarket      || "",
      p.subMarkets         || "",
      p.adTypes            || "",
      p.idealClients       || "",
      p.audienceNotes      || "",
      p.listingAddress     || "",
      p.listingBeds        || "",
      p.listingBaths       || "",
      p.listingSqft        || "",
      p.listingDescription || "",
      p.listingExtras      || "",
      p.differentiator     || "",
      p.contrarian         || "",
      p.wordsHate          || "",
      p.phrasesUse         || "",
      p.voicePreference    || "",
      p.googleAdsId        || "",
      p.metaAdAccountId    || ""
    ];

    sheet.appendRow(row);
    sheet.autoResizeColumns(1, getHeaders().length);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput("Marketing intake endpoint is live.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getHeaders() {
  return [
    "Timestamp",
    "Full Name",
    "Business Name",
    "What You Do",
    "Primary Market",
    "Sub-Markets",
    "Ad Types",
    "Ideal Clients",
    "Audience Notes",
    "Listing Address",
    "Beds",
    "Baths",
    "Sq Ft",
    "Listing Description",
    "Listing Extras",
    "Differentiator",
    "Contrarian View",
    "Words to Avoid",
    "Phrases to Use",
    "Voice Preference",
    "Google Ads ID",
    "Meta Ad Account ID"
  ];
}

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
