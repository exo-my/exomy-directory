// ═══════════════════════════════════════════════
//  ExoMy Directory — Google Apps Script
//  Extensions → Apps Script → paste this
//  Then: Run → syncToGitHub (first time manual)
//  Then: Triggers → On form submit → syncToGitHub
// ═══════════════════════════════════════════════

var GITHUB_TOKEN  = "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";
var GITHUB_OWNER  = "exo-my";
var GITHUB_REPO   = "exomy-directory";
var GITHUB_FILE   = "data.json";
var GITHUB_BRANCH = "main";

// ── Verified column map (counted from your actual sheet row 1) ──
var COL = {
  // col 0  = Timestamp
  // col 1  = Email Address (submitter)
  // col 2  = SECTION 1 header (empty)
  name:           3,   // Business / Shop / Breeder Name
  owner:          4,   // Owner / Representative Name
  bizType:        5,   // Business Type (Physical/Online/Both)
  types:          6,   // What best describes your business?
  otherType:      7,   // If Other, please specify
  yearEst:        8,   // Year Established
  description:    9,   // Business Description
  // col 10 = SECTION 2 header / species question
  species:        10,  // Which species or categories do you specialize in?
  speciesOther:   11,  // If other, Please specify
  // col 12 = SECTION 3 header (empty)
  state:          12,  // SECTION 3: LOCATION — this col holds state value in data rows
  city:           13,  // City/District
  address:        14,  // Business Address
  coverage:       15,  // Service Coverage
  shipping:       16,  // Do you offer shipping or delivery service?
  // col 17 = SECTION 4 header / Facebook
  facebook:       17,  // Facebook Page URL
  instagram:      18,  // Instagram URL
  tiktok:         19,  // TikTok URL
  threads:        20,  // Threads URL
  website:        21,  // Website URL
  primaryLink:    22,  // Primary Business Link
  // col 23 = SECTION 5 header / Phone
  phone:          23,  // Phone Number / WhatsApp
  email:          24,  // Email Address (public contact)
  // col 25 = SECTION 6 header / logo question
  // col 26 = Upload Business Logo
  // col 27 = SECTION 7 header / compliance question
  permits:        28,  // Do you hold any relevant permits / Perhilitan
  verifiedWanted: 29,  // Would you like to be considered for ExoMy Verified Status?
  // col 30 = SECTION 8 header / declaration
  status:         31   // STATUS (approved/pending — for future use)
};

// ── Helpers ──
function clean(val) {
  if (val === null || val === undefined) return "";
  var s = String(val).trim();
  return (s === "nan" || s === "NaN") ? "" : s;
}

function splitMulti(val) {
  var s = clean(val);
  if (!s) return [];
  return s.split(";").map(function(x) { return x.trim(); }).filter(function(x) { return x !== ""; });
}

function cleanYear(val) {
  var s = clean(val);
  if (!s) return "";
  // Remove decimal that Sheets adds to numbers (e.g. "2023.0" → "2023")
  return s.split(".")[0];
}

// ── Main sync function ──
function syncToGitHub() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rows    = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < rows.length; i++) {
    var r    = rows[i];
    var name = clean(r[COL.name]);

    // Skip blank rows, test entries, and section header rows
    if (!name || name === "test" || name === "ExoMy") continue;

    // Optional: only include rows marked approved (remove the // to enable)
    // var rowStatus = clean(r[COL.status]).toLowerCase();
    // if (rowStatus !== "approved") continue;

    members.push({
      name:           name,
      owner:          clean(r[COL.owner]),
      bizType:        clean(r[COL.bizType]),
      types:          splitMulti(r[COL.types]),
      otherType:      clean(r[COL.otherType]),
      yearEst:        cleanYear(r[COL.yearEst]),
      description:    clean(r[COL.description]),
      species:        splitMulti(r[COL.species]),
      speciesOther:   clean(r[COL.speciesOther]),
      state:          clean(r[COL.state]),
      city:           clean(r[COL.city]),
      address:        clean(r[COL.address]),
      coverage:       splitMulti(r[COL.coverage]),
      shipping:       clean(r[COL.shipping]),
      facebook:       clean(r[COL.facebook]),
      instagram:      clean(r[COL.instagram]),
      tiktok:         clean(r[COL.tiktok]),
      threads:        clean(r[COL.threads]),
      website:        clean(r[COL.website]),
      primaryLink:    clean(r[COL.primaryLink]),
      phone:          clean(r[COL.phone]),
      permits:        clean(r[COL.permits]),
      verifiedWanted: clean(r[COL.verifiedWanted]),
      status:         clean(r[COL.status])
      // NOTE: email intentionally excluded from public JSON (spam protection)
    });
  }

  var payload = JSON.stringify({
    members:     members,
    total:       members.length,
    lastUpdated: new Date().toISOString()
  }, null, 2);

  Logger.log("Syncing " + members.length + " members to GitHub...");
  pushToGitHub(payload);
}

// ── Push to GitHub via REST API ──
function pushToGitHub(content) {
  var apiUrl = "https://api.github.com/repos/"
    + GITHUB_OWNER + "/" + GITHUB_REPO + "/contents/" + GITHUB_FILE;

  var headers = {
    "Authorization": "token " + GITHUB_TOKEN,
    "Accept":        "application/vnd.github.v3+json",
    "Content-Type":  "application/json"
  };

  // Step 1: Get current file SHA (required to update an existing file)
  var sha = "";
  try {
    var getResp = UrlFetchApp.fetch(apiUrl, {
      method:             "get",
      headers:            headers,
      muteHttpExceptions: true
    });
    var getCode = getResp.getResponseCode();
    if (getCode === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
      Logger.log("Found existing data.json, SHA: " + sha);
    } else if (getCode === 404) {
      Logger.log("data.json not found — will create new file.");
    } else {
      Logger.log("Unexpected GET response " + getCode + ": " + getResp.getContentText());
    }
  } catch (e) {
    Logger.log("GET error: " + e);
  }

  // Step 2: PUT updated content
  var body = {
    message: "Auto-sync: ExoMy directory updated " + new Date().toLocaleString("ms-MY"),
    content:  Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch:   GITHUB_BRANCH
  };
  if (sha) body.sha = sha; // required when updating; omit when creating

  try {
    var resp = UrlFetchApp.fetch(apiUrl, {
      method:             "put",
      headers:            headers,
      payload:            JSON.stringify(body),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();

    if (code === 200) {
      Logger.log("SUCCESS: data.json updated on GitHub.");
    } else if (code === 201) {
      Logger.log("SUCCESS: data.json created on GitHub for the first time.");
    } else {
      // Log the full error body so you can diagnose
      Logger.log("ERROR " + code + ": " + resp.getContentText());
    }

  } catch (e) {
    Logger.log("PUT error: " + e);
  }
}
