// ============================================================
//  ESG Car Wash — Loyalty Program  |  Code.gs  v3.0
//  Sheets: "Users" + "WashLog"
// ============================================================

var USERS_SHEET   = "Users";
var WASHLOG_SHEET = "WashLog";
var ADMIN_SHEET   = "Admin";      // Announcement text lives here
var LEGACY_SHEET  = "Database";

var STREAK_MAX = 9;               // Washes needed to earn a reward choice

// ── ERP Integration ──────────────────────────────────────────
// Deployed URL of the ESG Car Wash Manager ERP web app
var ERP_API_URL = "https://script.google.com/macros/s/AKfycbymraoQGpi9yGe8awT6zRJU004_ptbQ2Mhwpj3j2OkOvDBiy87p84D99o0u_6j69b2I/exec";

// Users columns (0-based) — Balance + Notes added at the end
var U = { USERID:0, FULLNAME:1, PHONE:2, TOTAL:3, CYCLE:4, STATUS:5, SINCE:6, BALANCE:7, NOTES:8 };
var USERS_HDR = ["UserID","FullName","PhoneNumber","TotalLifetimeWashes","CurrentCycle","MemberStatus","MemberSince","Balance","Notes"];

// WashLog columns (0-based)
var W = { TIMESTAMP:0, USERID:1, WASHTYPE:2, REWARD:3 };
var WASHLOG_HDR = ["Timestamp","UserID","WashType","RewardRedeemed"];

// Default manager PIN (change with setManagerPin from GAS editor)
var DEFAULT_PIN = "1234";

// ─── JSON API Entry Points ───────────────────────────────────
// Deploy as Web App → Execute as: Me  → Who has access: Anyone
// Both doGet (reads) and doPost (writes) return JSON.
// GAS automatically adds Access-Control-Allow-Origin: * for
// "Anyone" deployments, so GitHub Pages fetch() calls work fine.

function getAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch(e) { return ""; }
}

// Helper — wrap any object as a JSON response
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry points — both route through handleRequest_ ─────────
function doGet(e) {
  return handleRequest_((e && e.parameter) ? e.parameter : {});
}

function doPost(e) {
  var data = {};
  try {
    // Frontend sends Content-Type: text/plain to avoid CORS preflight;
    // GAS still receives the raw JSON in e.postData.contents
    data = JSON.parse(e.postData.contents);
  } catch(err) {
    return jsonOut_({ success:false, message:"Invalid JSON: " + err.message });
  }
  return handleRequest_(data);
}

function handleRequest_(d) {
  var action = String(d.action || "");
  try {
    switch (action) {
      case "ping":            return jsonOut_({ success:true, message:"ESG API v3 OK" });
      case "setup":           return jsonOut_(setup());
      case "fetchUser":       return jsonOut_(checkUser(String(d.phone || "")));
      case "registerUser":    return jsonOut_(registerUser({ fullName:d.fullName, phone:d.phone }));
      case "addWash":         return jsonOut_(logWash_(d.userID, "Standard Wash"));
      case "logWashByManager":return jsonOut_(logWashByManager(d.userID, d.washType));
      case "getUserByID":     return jsonOut_(getUserByID(String(d.userID || d.id || "")));
      case "getDailyStats":   return jsonOut_(getDailyStats());
      case "verifyPin":       return jsonOut_(verifyManagerPin(String(d.pin || "")));
      case "setPin":          return jsonOut_(setManagerPin(d.currentPin, d.newPin));
      case "getAppUrl":       return jsonOut_({ success:true, url:getAppUrl() });
      case "getAppConfig":    return jsonOut_(getAppConfig());
      case "claimReward":     return jsonOut_(claimReward(d.userId, d.rewardChoice));
      default:                return jsonOut_({ success:false, message:"Unknown action: " + action });
    }
  } catch(err) {
    return jsonOut_({ success:false, message:err.message });
  }
}

// ─── Setup ───────────────────────────────────────────────────
function setup() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var created = [];

    if (!ss.getSheetByName(USERS_SHEET)) {
      var us = ss.insertSheet(USERS_SHEET);
      us.appendRow(USERS_HDR);
      styleHeader_(us, USERS_HDR.length, "#38bdf8");
      us.setColumnWidths(1, USERS_HDR.length, 165);
      created.push("Users");
    }

    if (!ss.getSheetByName(WASHLOG_SHEET)) {
      var wl = ss.insertSheet(WASHLOG_SHEET);
      wl.appendRow(WASHLOG_HDR);
      styleHeader_(wl, WASHLOG_HDR.length, "#10b981");
      wl.setColumnWidth(1, 200);
      wl.setColumnWidths(2, WASHLOG_HDR.length - 1, 160);
      created.push("WashLog");
    }

    // Admin sheet (announcement text in A2)
    if (!ss.getSheetByName(ADMIN_SHEET)) {
      var adm = ss.insertSheet(ADMIN_SHEET);
      adm.getRange("A1").setValue("Announcement");
      adm.getRange("A2").setValue("Welcome to ESG Car Wash! Visit us today.");
      styleHeader_(adm, 1, "#f59e0b");
      created.push("Admin");
    }

    // Seed default PIN if not set
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty("MANAGER_PIN")) {
      props.setProperty("MANAGER_PIN", DEFAULT_PIN);
    }

    return {
      success: true,
      message: created.length ? "Created: " + created.join(", ") : "All sheets exist."
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─── migrateDatabase ─────────────────────────────────────────
function migrateDatabase() {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var old = ss.getSheetByName(LEGACY_SHEET);
    if (!old || old.getLastRow() < 2) return { success: true, message: "Nothing to migrate." };
    setup();
    var users = ss.getSheetByName(USERS_SHEET);
    var data  = old.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var row = data[i], uid = row[1], since = "";
      try { since = Utilities.formatDate(new Date(row[0]), tz, "MMMM yyyy"); } catch(e){}
      if (!findUserRow_(users, uid, "id")) {
        users.appendRow([uid, row[2], row[3], row[4]||0, row[5]||0, row[7]||"Active", since]);
        count++;
      }
    }
    return { success: true, message: "Migrated " + count + " users." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─── PIN Management ──────────────────────────────────────────
function verifyManagerPin(pin) {
  try {
    var props  = PropertiesService.getScriptProperties();
    var stored = props.getProperty("MANAGER_PIN") || DEFAULT_PIN;
    return { success: true, valid: String(pin) === String(stored) };
  } catch (e) {
    return { success: false, valid: false, message: e.message };
  }
}

function setManagerPin(currentPin, newPin) {
  try {
    var check = verifyManagerPin(currentPin);
    if (!check.valid) return { success: false, message: "Current PIN is incorrect." };
    if (!newPin || String(newPin).length < 4) return { success: false, message: "New PIN must be at least 4 digits." };
    PropertiesService.getScriptProperties().setProperty("MANAGER_PIN", String(newPin));
    return { success: true, message: "PIN updated." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─── getUserStats ────────────────────────────────────────────
function getUserStats(userID) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(WASHLOG_SHEET);

    if (!logSheet || logSheet.getLastRow() < 2) {
      return { lifetimeTotal:0, monthlyStreak:0, monthlyRewardRedeemed:false, history:[] };
    }

    var data     = logSheet.getDataRange().getValues();
    var now      = new Date();
    var curMonth = now.getMonth();
    var curYear  = now.getFullYear();
    var tz       = Session.getScriptTimeZone();

    var userRows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][W.USERID]) === String(userID)) userRows.push(data[i]);
    }

    userRows.sort(function(a,b){ return new Date(b[W.TIMESTAMP]) - new Date(a[W.TIMESTAMP]); });

    var lifetimeTotal         = userRows.length;
    var monthlyStreak         = 0;
    var monthlyRewardRedeemed = false;

    for (var j = 0; j < userRows.length; j++) {
      var ts = new Date(userRows[j][W.TIMESTAMP]);
      if (ts.getMonth() === curMonth && ts.getFullYear() === curYear) {
        monthlyStreak++;
        if (String(userRows[j][W.WASHTYPE]) === "Interior Cleaning (Monthly Reward)") {
          monthlyRewardRedeemed = true;
        }
      }
    }

    var history = userRows.slice(0, 10).map(function(row) {
      var fmt   = "";
      var tsStr = "";
      try {
        var tsDate = new Date(row[W.TIMESTAMP]);
        fmt   = Utilities.formatDate(tsDate, tz, "MMM dd, h:mm a");
        tsStr = tsDate.toISOString(); // always a plain string — safe to serialize
      } catch(e){}
      return {
        timestamp: tsStr,            // string, never a Date object
        formatted: fmt,
        washType:  String(row[W.WASHTYPE] || "Standard Wash"),
        isReward:  row[W.REWARD] === "Yes"
      };
    });

    return { lifetimeTotal:lifetimeTotal, monthlyStreak:monthlyStreak, monthlyRewardRedeemed:monthlyRewardRedeemed, history:history };
  } catch (e) {
    return { lifetimeTotal:0, monthlyStreak:0, monthlyRewardRedeemed:false, history:[], error:e.message };
  }
}

// ─── getUserByID ─────────────────────────────────────────────
function getUserByID(userID) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var users = ss.getSheetByName(USERS_SHEET);

    if (!users || users.getLastRow() < 2) return { success:true, found:false, message:"No users in database." };

    var data = users.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][U.USERID]) === String(userID)) {
        var user  = rowToUser_(data[i]);
        var stats = getUserStats(userID);
        return { success:true, found:true, user:user, stats:stats };
      }
    }
    return { success:true, found:false, message:"Member ID not found: " + userID };
  } catch (e) {
    return { success:false, found:false, message:e.message };
  }
}

// ─── checkUser ───────────────────────────────────────────────
function checkUser(phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(USERS_SHEET)) setup();
    var users      = ss.getSheetByName(USERS_SHEET);
    var cleanPhone = phone.replace(/\D/g, "");
    if (users.getLastRow() < 2) return { success:true, found:false };
    var data = users.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][U.PHONE]).replace(/\D/g,"") === cleanPhone) {
        var user  = rowToUser_(data[i]);
        var stats = getUserStats(user.userID);
        return { success:true, found:true, user:user, stats:stats };
      }
    }
    return { success:true, found:false };
  } catch (e) {
    return { success:false, found:false, message:e.message };
  }
}

// ─── registerUser ────────────────────────────────────────────
function registerUser(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(USERS_SHEET)) setup();
    var users = ss.getSheetByName(USERS_SHEET);
    var check = checkUser(payload.phone);
    if (check.found) return { success:false, code:"PHONE_EXISTS", message:"This phone number is already registered." };
    var userID = generateUserID_(users);
    var since  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy");
    users.appendRow([userID, payload.fullName, payload.phone, 0, 0, "Active", since]);
    return {
      success:true, message:"Welcome to ESG Car Wash!",
      user:{ userID:userID, fullName:payload.fullName, phoneNumber:payload.phone, totalWashes:0, currentCycle:0, memberStatus:"Active", memberSince:since, streakMax:STREAK_MAX },
      stats:{ lifetimeTotal:0, monthlyStreak:0, monthlyRewardRedeemed:false, history:[] }
    };
  } catch (e) {
    return { success:false, message:e.message };
  }
}

// ─── incrementWash (client-facing / admin test) ──────────────
function incrementWash(uid) {
  return logWash_(uid, "Standard Wash");
}

// ─── logWashByManager ────────────────────────────────────────
//  washType: "Standard Wash" | "Interior Cleaning (Monthly Reward)"
function logWashByManager(userID, washType) {
  try {
    var result;
    if (washType === "Interior Cleaning (Monthly Reward)") {
      result = redeemMonthlyReward_(userID);
    } else {
      result = logWash_(userID, washType);
    }
    // Mirror every successful manager scan into the ERP washing list
    if (result.success) {
      try {
        syncToERP_(userID, washType, result.user);
        result.erpSynced = true;
      } catch(erpErr) {
        Logger.log("ERP sync error: " + erpErr.message);
        result.erpSynced = false;
        result.erpError  = erpErr.message;
      }
    }
    return result;
  } catch (e) {
    return { success:false, message:e.message };
  }
}

// ─── syncToERP_ ──────────────────────────────────────────────
// Posts the wash entry to the ERP web app's doPost endpoint so
// the manager terminal sees it in the washing list automatically.
function syncToERP_(userID, washType, userData) {
  // Map loyalty wash types to ERP Georgian labels
  var erpWashType = "სტანდარტი";
  if (washType && washType.indexOf("Interior") !== -1) erpWashType = "შიგნიდან";
  else if (washType && washType.indexOf("Free") !== -1) erpWashType = "სხვა";

  var payload = {
    action:       "addLoyaltyEntry",
    userID:       userID,
    customerName: (userData && userData.fullName)    ? userData.fullName    : userID,
    phone:        (userData && userData.phoneNumber) ? userData.phoneNumber : "",
    erpWashType:  erpWashType
  };

  var response = UrlFetchApp.fetch(ERP_API_URL, {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    followRedirects:    true,
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (!result.success) throw new Error(result.message || "ERP rejected the entry");
}

// ─── getDailyStats ───────────────────────────────────────────
function getDailyStats() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(WASHLOG_SHEET);
    var users    = ss.getSheetByName(USERS_SHEET);

    if (!logSheet || logSheet.getLastRow() < 2) return { todayCount:0, recentScans:[] };

    var data    = logSheet.getDataRange().getValues();
    var tz      = Session.getScriptTimeZone();
    var todayDs = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    // Build user ID→name lookup
    var nameMap = {};
    if (users && users.getLastRow() > 1) {
      var ud = users.getDataRange().getValues();
      for (var k = 1; k < ud.length; k++) nameMap[String(ud[k][U.USERID])] = String(ud[k][U.FULLNAME]);
    }

    var todayRows = [];
    for (var i = 1; i < data.length; i++) {
      try {
        var ds = Utilities.formatDate(new Date(data[i][W.TIMESTAMP]), tz, "yyyy-MM-dd");
        if (ds === todayDs) todayRows.push(data[i]);
      } catch(e) {}
    }

    todayRows.sort(function(a,b){ return new Date(b[W.TIMESTAMP]) - new Date(a[W.TIMESTAMP]); });

    var recentScans = todayRows.slice(0, 10).map(function(row) {
      var fmt = "";
      try { fmt = Utilities.formatDate(new Date(row[W.TIMESTAMP]), tz, "HH:mm"); } catch(e){}
      return {
        time:      fmt,
        userID:    String(row[W.USERID]),
        userName:  nameMap[String(row[W.USERID])] || "Unknown",
        washType:  String(row[W.WASHTYPE]),
        isReward:  row[W.REWARD] === "Yes"
      };
    });

    return { todayCount: todayRows.length, recentScans: recentScans };
  } catch (e) {
    return { todayCount:0, recentScans:[], error:e.message };
  }
}

// ─── Private: core wash logic ─────────────────────────────────
function logWash_(uid, requestedType) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var users = ss.getSheetByName(USERS_SHEET);
  var wlog  = ss.getSheetByName(WASHLOG_SHEET);
  if (!users) return { success:false, message:"Users sheet missing. Run setup()." };
  if (!wlog)  return { success:false, message:"WashLog sheet missing. Run setup()." };
  if (users.getLastRow() < 2) return { success:false, message:"No users found." };

  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][U.USERID]) !== String(uid)) continue;

    var total        = parseInt(data[i][U.TOTAL]) || 0;
    var cycle        = parseInt(data[i][U.CYCLE]) || 0;
    var status       = String(data[i][U.STATUS]);
    var rewardEarned = false;
    var message      = "";
    var washType     = requestedType || "Standard Wash";
    var isRedemption = (status === "Reward Available");

    total++;

    if (isRedemption) {
      washType     = "Free Wash (Reward)";
      cycle        = 0;
      status       = "Active";
      message      = "Free wash redeemed! New cycle started.";
    } else {
      cycle++;
      if (cycle >= STREAK_MAX) {
        cycle        = STREAK_MAX;
        status       = "Reward Available";
        rewardEarned = true;
        message      = "Streak complete! Choose your reward.";
      } else {
        message = "Wash recorded. " + (STREAK_MAX - cycle) + " more to reward.";
      }
    }

    wlog.appendRow([new Date(), uid, washType, isRedemption ? "Yes" : "No"]);
    var rowNum = i + 1;
    users.getRange(rowNum, U.TOTAL  + 1).setValue(total);
    users.getRange(rowNum, U.CYCLE  + 1).setValue(cycle);
    users.getRange(rowNum, U.STATUS + 1).setValue(status);

    var updRow = data[i].slice();
    updRow[U.TOTAL] = total; updRow[U.CYCLE] = cycle; updRow[U.STATUS] = status;

    return { success:true, message:message, rewardEarned:rewardEarned, user:rowToUser_(updRow), stats:getUserStats(uid) };
  }
  return { success:false, message:"User ID not found: " + uid };
}

function redeemMonthlyReward_(uid) {
  var stats = getUserStats(uid);
  if (!stats || stats.monthlyStreak < 4)
    return { success:false, code:"STREAK_LOW", message:"Monthly streak requirement not met (" + (stats ? stats.monthlyStreak : 0) + "/4)." };
  if (stats.monthlyRewardRedeemed)
    return { success:false, message:"Monthly Interior Cleaning already redeemed this month." };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wlog  = ss.getSheetByName(WASHLOG_SHEET);
  var users = ss.getSheetByName(USERS_SHEET);
  if (!wlog || !users) return { success:false, message:"Sheets not found." };

  // Log the redemption — doesn't affect the 9-wash cycle
  wlog.appendRow([new Date(), uid, "Interior Cleaning (Monthly Reward)", "Yes"]);

  // Still increment lifetime total
  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][U.USERID]) === String(uid)) {
      var newTotal = (parseInt(data[i][U.TOTAL]) || 0) + 1;
      users.getRange(i + 1, U.TOTAL + 1).setValue(newTotal);
      var updRow = data[i].slice();
      updRow[U.TOTAL] = newTotal;
      return { success:true, message:"Interior Cleaning reward redeemed!", isRedemption:true, user:rowToUser_(updRow), stats:getUserStats(uid) };
    }
  }
  return { success:false, message:"User not found." };
}

// ─── Helpers ──────────────────────────────────────────────────
function rowToUser_(row) {
  return {
    userID:      String(row[U.USERID]   || ""),
    fullName:    String(row[U.FULLNAME] || ""),
    phoneNumber: String(row[U.PHONE]    || ""),
    totalWashes: parseInt(row[U.TOTAL]) || 0,
    currentCycle:parseInt(row[U.CYCLE]) || 0,
    memberStatus:String(row[U.STATUS]   || "Active"),
    memberSince: String(row[U.SINCE]    || ""),
    balance:     parseInt(row[U.BALANCE]) || 0,
    notes:       String(row[U.NOTES]    || ""),
    streakMax:   STREAK_MAX
  };
}

// ─── getAppConfig ─────────────────────────────────────────────
// Returns announcement text from Admin sheet cell A2.
function getAppConfig() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var adm  = ss.getSheetByName(ADMIN_SHEET);
    var text = "";
    if (adm && adm.getLastRow() >= 2) {
      text = String(adm.getRange("A2").getValue() || "");
    }
    return { success:true, announcementText:text, streakMax:STREAK_MAX };
  } catch(e) {
    return { success:false, announcementText:"", streakMax:STREAK_MAX, error:e.message };
  }
}

// ─── claimReward ─────────────────────────────────────────────
// rewardChoice: "NANO_WAX" | "HALF_OFF_DETAIL" | "FREE_WASH"
function claimReward(userId, rewardChoice) {
  try {
    var VALID = ["NANO_WAX","HALF_OFF_DETAIL","FREE_WASH"];
    if (!userId)                       return { success:false, code:"MISSING_ID",     message:"Missing userId." };
    if (VALID.indexOf(rewardChoice)<0) return { success:false, code:"INVALID_REWARD", message:"Invalid reward choice." };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var users = ss.getSheetByName(USERS_SHEET);
    var wlog  = ss.getSheetByName(WASHLOG_SHEET);
    if (!users) return { success:false, message:"Users sheet not found." };

    var data = users.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][U.USERID]) !== String(userId)) continue;

      var balance = (parseInt(data[i][U.BALANCE]) || 0) + 1;
      var rowNum  = i + 1;

      users.getRange(rowNum, U.CYCLE   + 1).setValue(0);
      users.getRange(rowNum, U.STATUS  + 1).setValue("Active");
      users.getRange(rowNum, U.BALANCE + 1).setValue(balance);
      users.getRange(rowNum, U.NOTES   + 1).setValue(rewardChoice);

      if (wlog) wlog.appendRow([new Date(), userId, "Reward: " + rewardChoice, "Yes"]);

      var updRow = data[i].slice();
      updRow[U.CYCLE]   = 0;
      updRow[U.STATUS]  = "Active";
      updRow[U.BALANCE] = balance;
      updRow[U.NOTES]   = rewardChoice;

      return {
        success:      true,
        message:      "Reward claimed: " + rewardChoice,
        rewardChoice: rewardChoice,
        user:         rowToUser_(updRow),
        stats:        getUserStats(userId)
      };
    }
    return { success:false, code:"USER_NOT_FOUND", message:"User not found." };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

function generateUserID_(usersSheet) {
  var existing = [];
  if (usersSheet && usersSheet.getLastRow() > 1) {
    var ids = usersSheet.getRange(2, U.USERID+1, usersSheet.getLastRow()-1, 1).getValues();
    existing = ids.map(function(r){ return r[0]; });
  }
  var id;
  do { id = "ESG-" + Math.floor(1000 + Math.random() * 9000); }
  while (existing.indexOf(id) !== -1);
  return id;
}

function findUserRow_(sheet, value, by) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getDataRange().getValues();
  var col  = (by === "phone") ? U.PHONE : U.USERID;
  for (var i = 1; i < data.length; i++) {
    var cell = String(data[i][col]).replace(by==="phone"?/\D/g:"", "");
    if (cell === String(value).replace(by==="phone"?/\D/g:"","")) return { row:i+1, data:data[i] };
  }
  return null;
}

function styleHeader_(sheet, numCols, accentColor) {
  var hdr = sheet.getRange(1, 1, 1, numCols);
  hdr.setFontWeight("bold").setBackground("#0f172a").setFontColor(accentColor).setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
}
