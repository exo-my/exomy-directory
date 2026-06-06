// ═══════════════════════════════════════════════
//  ExoMy Directory — Google Apps Script
//  Paste this in Extensions → Apps Script
// ═══════════════════════════════════════════════

var GITHUB_TOKEN  = "ghp_YOUR_TOKEN_HERE";
var GITHUB_OWNER  = "YOUR_GITHUB_USERNAME";
var GITHUB_REPO   = "exomy-directory";
var GITHUB_FILE   = "data.json";
var GITHUB_BRANCH = "main";

// ── Column indexes (0-based) matching your sheet ──
var COL = {
  name:          3,
  owner:         4,
  bizType:       5,
  types:         6,
  otherType:     7,
  yearEst:       8,
  description:   9,
  species:       10,
  speciesOther:  11,
  state:         12,
  city:          13,
  address:       14,
  coverage:      15,
  shipping:      16,
  facebook:      17,
  instagram:     18,
  tiktok:        19,
  threads:       20,
  website:       21,
  primaryLink:   22,
  phone:         23,
  email:         24,
  permits:       28,
  verifiedWanted:29
};

function clean(val) {
  if (val === null || val === undefined) return "";
  var s = String(val).trim();
  if (s === "nan" || s === "NaN") return "";
  return s;
}

function splitMulti(val) {
  var s = clean(val);
  if (!s) return [];
  return s.split(";").map(function(x){ return x.trim(); }).filter(function(x){ return x !== ""; });
}

function syncToGitHub() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rows  = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < rows.length; i++) {
    var r    = rows[i];
    var name = clean(r[COL.name]);
    if (!name || name === "test" || name === "ExoMy") continue;

    var yearRaw = clean(r[COL.yearEst]);
    if (yearRaw.indexOf(".") !== -1) {
      yearRaw = yearRaw.split(".")[0];
    }

    members.push({
      name:           name,
      owner:          clean(r[COL.owner]),
      bizType:        clean(r[COL.bizType]),
      types:          splitMulti(r[COL.types]),
      otherType:      clean(r[COL.otherType]),
      yearEst:        yearRaw,
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
      email:          clean(r[COL.email]),
      permits:        clean(r[COL.permits]),
      verifiedWanted: clean(r[COL.verifiedWanted]),
      updatedAt:      new Date().toISOString()
    });
  }

  var payload = JSON.stringify({
    members:     members,
    total:       members.length,
    lastUpdated: new Date().toISOString()
  }, null, 2);

  pushToGitHub(payload);
}

function pushToGitHub(content) {
  var apiUrl = "https://api.github.com/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/contents/" + GITHUB_FILE;

  // Get current SHA of the file (needed to update it)
  var sha = "";
  try {
    var getResp = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: {
        "Authorization": "token " + GITHUB_TOKEN,
        "Accept": "application/vnd.github.v3+json"
      },
      muteHttpExceptions: true
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }
  } catch(e) {
    Logger.log("File not found, will create: " + e);
  }

  var body = {
    message: "Auto-sync: ExoMy directory updated " + new Date().toLocaleString(),
    content:  Utilities.base64Encode(content),
    branch:   GITHUB_BRANCH
  };
  if (sha) body.sha = sha;

  var resp = UrlFetchApp.fetch(apiUrl, {
    method:  "put",
    headers: {
      "Authorization": "token " + GITHUB_TOKEN,
      "Content-Type":  "application/json",
      "Accept":        "application/vnd.github.v3+json"
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code === 200 || code === 201) {
    Logger.log("SUCCESS: data.json updated on GitHub. Members: " + JSON.parse(resp.getContentText()).content.name);
  } else {
    Logger.log("ERROR " + code + ": " + resp.getContentText());
  }
}
