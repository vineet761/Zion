// ============================================================
// Household Renewals Tracker — Vineet & Linda, 68 Hillfort Drive
// Design: append-only Log (one row per renewal event) + a computed
// Dashboard (always shows the latest renewal per item, auto-sorted).
// Adding a new year's renewal = append ONE row to Log. That's it.
// ============================================================

var LOG_SHEET = "Log";
var DASH_SHEET = "Dashboard";
var OLD_SHEET_NAMES = ["Renewals", "History"]; // leftovers from earlier designs — removed on build

// Log columns (append-only — every renewal/payment ever logged)
var L_CATEGORY = 1, L_ITEM = 2, L_OWNER = 3, L_VEHICLE = 4, L_REF = 5,
    L_DATE = 6, L_AMOUNT = 7, L_NOTES = 8;
var LOG_HEADERS = ["Category", "Item / Vendor", "Owner", "Vehicle Reg / Property",
  "Reference / Policy / Account No.", "Renewal Date", "Amount (EUR)", "Notes"];

// Dashboard columns (computed — one row per item, showing its LATEST renewal)
var D_CATEGORY = 1, D_ITEM = 2, D_OWNER = 3, D_VEHICLE = 4, D_REF = 5,
    D_DATE = 6, D_AMOUNT = 7, D_DAYS = 8, D_STATUS = 9, D_NOTES = 10;
var DASH_HEADERS = ["Category", "Item / Vendor", "Owner", "Vehicle Reg / Property",
  "Reference / Policy / Account No.", "Current Renewal Date", "Current Amount (EUR)",
  "Days Until Expiry", "Status", "Notes"];

var CATEGORY_LIST = ["Motor Tax", "Motor Insurance", "Vehicle NCT", "Home Insurance",
  "Electricity", "TV Licence", "Property Tax", "Other"];
var OWNER_LIST = ["Vineet", "Linda", "Joint", "Household"];

function d(y, m, day) { return new Date(y, m - 1, day); }

// ============================================================
// MENU
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("\ud83c\udfe0 Household Renewals")
    .addItem("\ud83d\udd28 Build Workbook (Setup)", "buildRenewalsWorkbook")
    .addItem("\u2795 Log a Renewal", "logRenewal")
    .addItem("\ud83d\udd04 Refresh Dashboard", "refreshDashboard")
    .addSeparator()
    .addItem("\u2753 How To Use", "showHowTo")
    .addToUi();
}

// ============================================================
// BUILD WORKBOOK — creates Log + Dashboard, removes leftover sheets
// from earlier designs, and seeds the Log with known history
// ============================================================
function buildRenewalsWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var existingLog = ss.getSheetByName(LOG_SHEET);
  if (existingLog && existingLog.getLastRow() > 1) {
    var resp = ui.alert("\u26a0\ufe0f Rebuild Workbook",
      "This will clear and re-seed the Log and Dashboard sheets. Continue?",
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  // Create Log/Dashboard FIRST — deleting old sheets before these exist can
  // leave the spreadsheet with zero sheets momentarily, which Sheets rejects.
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  var dash = ss.getSheetByName(DASH_SHEET) || ss.insertSheet(DASH_SHEET);

  OLD_SHEET_NAMES.forEach(function (name) {
    if (name === LOG_SHEET || name === DASH_SHEET) return;
    var s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  log.clear();
  dash.clear();
  dash.clearConditionalFormatRules();

  setupLogSheet(log);
  setupDashboardSheet(dash);
  seedLog(log);

  refreshDashboard();
  ss.setActiveSheet(dash);
  ui.alert("\u2705 Household Renewals workbook built. Log has every known renewal; " +
    "Dashboard shows the current status per item. See 'How To Use' for next steps.");
}

function setupLogSheet(sheet) {
  sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
    .setFontWeight("bold").setBackground("#0d47a1").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.getRange(2, L_CATEGORY, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CATEGORY_LIST, true).build());
  sheet.getRange(2, L_OWNER, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(OWNER_LIST, true).build());
  sheet.getRange(2, L_DATE, 500, 1).setNumberFormat("dd/mm/yyyy");
  sheet.getRange(2, L_AMOUNT, 500, 1).setNumberFormat("\u20ac#,##0.00");
}

function setupDashboardSheet(sheet) {
  sheet.getRange(1, 1, 1, DASH_HEADERS.length).setValues([DASH_HEADERS])
    .setFontWeight("bold").setBackground("#1a7f3c").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
}

// ============================================================
// SEED LOG — every known renewal/payment event, on hand as of Sep 2026.
// One row per event. TBD = confirm with source document later.
// ============================================================
function seedLog(sheet) {
  var rows = [
    ["TV Licence", "TV Licence", "Vineet", "68 Hillfort Drive", "Ref 03135908440126085",
      d(2026, 8, 31), 160, "Renews annually; period start not on notice"],
    ["Home Insurance", "Zurich Home Insurance", "Joint", "68 Hillfort Drive",
      "Policy 01 PPZ 5077866", d(2026, 10, 14), 227.83, "New Business transaction"],
    ["Electricity", "Electric Ireland", "Household", "68 Hillfort Drive, Naas",
      "Account 951529117", d(2026, 10, 16), "TBD", "Contract renewal date only; amount to confirm"],
    ["Motor Insurance", "Insurer TBD", "Linda", "141D48673",
      "Policy AVDN04280PCR10083854 / Cert 543250001111315", d(2026, 10, 22), 1000,
      "Insurer name not shown on certificate; premium is an estimate"],
    ["Property Tax", "LPT - Revenue.ie", "Joint", "68 Hillfort Drive (PAN 713070)",
      "Property ID 4198631JH", d(2024, 12, 31), 445, "Paid in full"],
    ["Property Tax", "LPT - Revenue.ie", "Joint", "68 Hillfort Drive (PAN 713070)",
      "Property ID 4198631JH", d(2025, 12, 31), 445, "Paid in full"],
    ["Property Tax", "LPT - Revenue.ie", "Joint", "68 Hillfort Drive (PAN 713070)",
      "Property ID 4198631JH", d(2026, 12, 31), 470, "Paid in full"],
    ["Property Tax", "LPT - Revenue.ie", "Joint", "68 Hillfort Drive (PAN 713070)",
      "Property ID 4198631JH", d(2027, 1, 10), "TBD", "2027 due date/amount is an estimate — confirm on Revenue.ie"],
    ["Motor Insurance", "Aviva Insurance Ireland DAC", "Vineet", "181D26379",
      "Policy 932511493", d(2027, 2, 25), "TBD", "Premium not shown on insurance disc/cover letter"],
    ["Motor Tax", "Motor Tax - Revenue.ie", "Linda", "141D48673",
      "Ref 89242639", d(2027, 4, 30), 180, "Car is Linda's; tax registered to Vineet"],
    ["Motor Tax", "Motor Tax - Revenue.ie", "Vineet", "181D26379",
      "Ref 89892484", d(2027, 5, 31), 270, ""],
    ["Vehicle NCT", "NCT - National Car Testing Service", "Linda", "141D48673",
      "Booking ID 84179637", d(2026, 10, 11), "TBD", "Email notification received 11/09/2026; amount TBD"]
  ];
  sheet.getRange(2, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

// ============================================================
// REFRESH DASHBOARD — for each (Item, Vehicle) group in Log, keep only
// the row with the latest Renewal Date, compute Days/Status, sort, write.
// Run this any time after editing Log directly (or it runs automatically
// after 'Log a Renewal').
// ============================================================
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET);
  var dash = ss.getSheetByName(DASH_SHEET);
  if (!log || !dash) return;

  var lastRow = log.getLastRow();
  var latest = {}; // key: item||vehicle -> latest log row (array)
  if (lastRow > 1) {
    var data = log.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
    data.forEach(function (row) {
      var item = row[L_ITEM - 1];
      if (!item) return;
      var key = item + "||" + row[L_VEHICLE - 1];
      var date = row[L_DATE - 1];
      var existing = latest[key];
      var existingDate = existing ? existing[L_DATE - 1] : null;
      var isNewer = date instanceof Date && (!(existingDate instanceof Date) || date > existingDate);
      if (!existing || isNewer) latest[key] = row;
    });
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var dashRows = Object.keys(latest).map(function (key) {
    var row = latest[key];
    var date = row[L_DATE - 1];
    var days = date instanceof Date ? Math.round((date - today) / 86400000) : "";
    var status = days === "" ? "" :
      (days < 0 ? "EXPIRED" : days <= 30 ? "DUE SOON" : days <= 90 ? "UPCOMING" : "OK");
    return [
      row[L_CATEGORY - 1], row[L_ITEM - 1], row[L_OWNER - 1], row[L_VEHICLE - 1], row[L_REF - 1],
      date, row[L_AMOUNT - 1], days, status, row[L_NOTES - 1]
    ];
  });

  dashRows.sort(function (a, b) {
    var da = a[D_DATE - 1], db = b[D_DATE - 1];
    var va = da instanceof Date ? da.getTime() : Infinity;
    var vb = db instanceof Date ? db.getTime() : Infinity;
    return va - vb;
  });

  var lastDashRow = dash.getLastRow();
  if (lastDashRow > 1) dash.getRange(2, 1, lastDashRow - 1, DASH_HEADERS.length).clearContent();

  if (dashRows.length > 0) {
    dash.getRange(2, 1, dashRows.length, DASH_HEADERS.length).setValues(dashRows);
    dash.getRange(2, D_DATE, dashRows.length, 1).setNumberFormat("dd/mm/yyyy");
    dash.getRange(2, D_AMOUNT, dashRows.length, 1).setNumberFormat("\u20ac#,##0.00");
  }

  applyConditionalFormatting(dash, dashRows.length);
  dash.autoResizeColumns(1, DASH_HEADERS.length);
}

function applyConditionalFormatting(sheet, rowCount) {
  sheet.clearConditionalFormatRules();
  if (rowCount === 0) return;
  var range = sheet.getRange(2, 1, rowCount, DASH_HEADERS.length);
  var dueSoon = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2<>"",$H2<=30)')
    .setBackground("#f4c7c3")
    .setRanges([range]).build();
  var upcoming = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2<>"",$H2>30,$H2<=90)')
    .setBackground("#fce8b2")
    .setRanges([range]).build();
  sheet.setConditionalFormatRules([dueSoon, upcoming]);
}

// ============================================================
// LOG A RENEWAL — appends one row to Log. If the item already exists
// (matched by Item + Vehicle), pre-fills Category/Owner/Reference from
// its most recent entry so you only need to type the new date + amount.
// ============================================================
function logRenewal() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET);
  if (!log) { ui.alert("\u274c Run 'Build Workbook (Setup)' first."); return; }

  var itemResp = ui.prompt("Item / Vendor", "Type an existing name to reuse its details, or a new name to add a new item", ui.ButtonSet.OK_CANCEL);
  if (itemResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
  var item = itemResp.getResponseText().trim();

  var vehicleResp = ui.prompt("Vehicle Reg / Property", "(used together with Item to find prior entries)", ui.ButtonSet.OK_CANCEL);
  if (vehicleResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
  var vehicle = vehicleResp.getResponseText().trim();

  var lastRow = log.getLastRow();
  var prior = null;
  if (lastRow > 1) {
    var data = log.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
    for (var r = data.length - 1; r >= 0; r--) {
      if (data[r][L_ITEM - 1] === item && data[r][L_VEHICLE - 1] === vehicle) { prior = data[r]; break; }
    }
  }

  var category, owner, ref;
  if (prior) {
    category = promptWithDefault_(ui, "Category", prior[L_CATEGORY - 1]);
    if (category === null) return;
    owner = promptWithDefault_(ui, "Owner", prior[L_OWNER - 1]);
    if (owner === null) return;
    ref = promptWithDefault_(ui, "Reference / Policy / Account No.", prior[L_REF - 1]);
    if (ref === null) return;
  } else {
    var catResp = ui.prompt("Category", CATEGORY_LIST.join(" / "), ui.ButtonSet.OK_CANCEL);
    if (catResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
    category = catResp.getResponseText().trim();

    var ownerResp = ui.prompt("Owner", OWNER_LIST.join(" / "), ui.ButtonSet.OK_CANCEL);
    if (ownerResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
    owner = ownerResp.getResponseText().trim();

    var refResp = ui.prompt("Reference / Policy / Account No.", "", ui.ButtonSet.OK_CANCEL);
    if (refResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
    ref = refResp.getResponseText().trim();
  }

  var dateResp = ui.prompt("Renewal Date (dd/mm/yyyy)", "", ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
  var renewalDate = parseUkDate_(dateResp.getResponseText().trim());

  var amountResp = ui.prompt("Amount (EUR, blank if unknown)", "", ui.ButtonSet.OK_CANCEL);
  if (amountResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
  var amountText = amountResp.getResponseText().trim();
  var amount = amountText === "" ? "TBD" : Number(amountText);

  var notesResp = ui.prompt("Notes (optional)", "", ui.ButtonSet.OK_CANCEL);
  if (notesResp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return; }
  var notes = notesResp.getResponseText().trim();

  log.appendRow([category, item, owner, vehicle, ref, renewalDate, amount, notes]);
  refreshDashboard();
  ui.alert("\u2705 Logged '" + item + "' renewal. Dashboard updated.");
}

function promptWithDefault_(ui, label, defaultValue) {
  var resp = ui.prompt(label, "Press OK to keep: " + defaultValue, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) { ui.alert("Cancelled."); return null; }
  var text = resp.getResponseText().trim();
  return text === "" ? defaultValue : text;
}

function parseUkDate_(text) {
  if (!text) return "";
  var parts = text.split("/");
  if (parts.length !== 3) return text;
  return d(Number(parts[2]), Number(parts[1]), Number(parts[0]));
}

// ============================================================
// HOW TO USE
// ============================================================
function showHowTo() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("\ud83c\udfe0 Household Renewals — How To Use",
    "Two sheets:\n\n" +
    "\u2022 Log — the source of truth. One row per renewal/payment, ever. Adding a new year's " +
    "renewal is just appending ONE new row here — no columns to find, nothing to remember.\n\n" +
    "\u2022 Dashboard — automatically computed. Shows one row per item with its LATEST renewal " +
    "(by date), Days Until Expiry, and Status, sorted soonest-first. Colour flags: red = due " +
    "within 30 days, amber = due within 90 days.\n\n" +
    "To log a renewal, you have two options:\n" +
    "1) Household Renewals menu > 'Log a Renewal' — if the item already exists, it pre-fills " +
    "Category/Owner/Reference from its last entry, so you usually only type the date + amount.\n" +
    "2) Or just add a row directly to the Log yourself — then run 'Refresh Dashboard' (or it " +
    "happens automatically next time you use 'Log a Renewal').\n\n" +
    "\u26a0\ufe0f The Dashboard is a snapshot, not a live formula — if you edit the Log by hand, " +
    "click 'Refresh Dashboard' afterwards to see the update reflected.\n\n" +
    "Because Log is append-only, history is automatically preserved — nothing is ever overwritten.\n\n" +
    "Email reminders are not set up yet — planned for a later phase.",
    ui.ButtonSet.OK);
}



