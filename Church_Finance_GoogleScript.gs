// ============================================================
// Kildare Christian Fellowship — Financial Workbook
// Sheets: Weekly Offerings | Member Contributions (2026-2050) | Expenses | Summary
// ============================================================

// ---- Global layout constants ----
var TITHE_START_YEAR = 2026;
var TITHE_END_YEAR   = 2050;
var NUM_MONTHS       = 301; // Dec-2025 + Jan-2026 to Dec-2050

// Member Contributions column layout — Email column added at col B
var EMAIL_COL        = 2;                                 // col B = email
var TITHE_COL        = 3;                                 // col C = monthly tithe
var MONTH_COL_START  = 4;                                 // col D = Jan-2026
var MONTH_COL_END    = MONTH_COL_START + NUM_MONTHS - 1; // col KR
var TOTAL_RECV_COL   = MONTH_COL_END + 1;                // col KS
var NOTES_COL        = MONTH_COL_END + 2;                // col KT
var MEMBER_ID_COL    = NOTES_COL + 1;                    // col KU, hidden stable identifier

// ============================================================
// MENU
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("\u26ea Church Finance")
    .addItem("\ud83d\udcca Close Month Balance",              "closeMonthBalance")
    .addItem("\ud83d\udd04 Refresh Weekly Balance",            "refreshWeeklyCollectionsBalance")
    .addItem("\u2795 Add New Member",                 "addNewMember")
    .addItem("\ud83d\udce7 Send Selected Member Statement", "sendAnnualStatements")
    .addItem("\ud83d\udce8 Send Annual Statements to All Members", "sendAllAnnualStatements")
    .addItem("\ud83d\udd17 Show Member Balance Links",     "showMemberLinks")
    .addItem("\u2702\ufe0f Shorten Member Links",              "shortenMemberLinks")
    .addItem("\u2699\ufe0f Set Web App URL",                  "setWebAppUrl")
    .addItem("\ud83d\udd11 Generate Member Tokens",          "generateMemberTokens")
    .addItem("\ud83d\udd04 Rotate Member Links (Revokes Old)", "rotateMemberLinks")
    .addItem("\ud83d\udcc5 Start New Year",                   "startNewYear")
    .addItem("\ud83c\udff7\ufe0f Log New Asset",                  "addNewAsset")
    .addSeparator()
    .addItem("\ud83d\udd04 Safe Rebuild (Preserves Data)", "safeRebuildWorkbook")
    .addItem("\ud83e\uddf9 Prepare This Copy for Sharing", "prepareCopyForSharing")
    .addSeparator()
    .addSubMenu(ui.createMenu("\u26a0\ufe0f Danger Zone")
      .addItem("\ud83d\udd28 Rebuild Workbook (DELETES ALL DATA)", "buildChurchFinanceWorkbook"))
    .addSeparator()
    .addItem("\u2753 How To Use", "showHowTo")
    .addToUi();
}

// ============================================================
// CLOSE MONTH BALANCE — inserts BALANCE row and calculates formulas
// ============================================================
function closeMonthBalance() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Weekly Offerings");
  if (!sheet) { ui.alert("\u274c Weekly Offerings sheet not found."); return; }

  // Find actual last data row by scanning backwards for a non-empty Date cell
  var lastRow = sheet.getLastRow();
  while (lastRow >= 4 && !sheet.getRange(lastRow, 3).getValue()) { lastRow--; }
  if (lastRow < 4) { ui.alert("\u274c No data found in Weekly Offerings."); return; }

  if (String(sheet.getRange(lastRow, 3).getValue()).trim().toUpperCase() === "BALANCE") {
    ui.alert("\u26a0\ufe0f This month is already balanced.\n\nEnter Sunday offerings first, then close.");
    return;
  }

  // Find start of this block and the previous BALANCE row
  var dataStart  = 4;
  var prevBalRow = null;
  for (var r = lastRow - 1; r >= 4; r--) {
    if (String(sheet.getRange(r, 3).getValue()).trim().toUpperCase() === "BALANCE") {
      dataStart  = r + 1;
      prevBalRow = r;
      break;
    }
  }

  var lastYear  = sheet.getRange(lastRow, 1).getValue() || "";
  var lastMonth = sheet.getRange(lastRow, 2).getValue() || "";
  for (var row = dataStart; row <= lastRow; row++) {
    var rowYear = sheet.getRange(row, 1).getValue();
    var rowMonth = sheet.getRange(row, 2).getValue();
    if (String(rowYear) !== String(lastYear) || String(rowMonth) !== String(lastMonth)) {
      ui.alert("\u274c Cannot close this month because rows since the previous balance contain more than one month or year. Correct the entries first.");
      return;
    }
  }

  var resp = ui.prompt(
    "\ud83d\udcca Close Month Balance",
    "Closing: " + lastMonth + " " + lastYear + "\n\nComment (or leave blank):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var comment = resp.getResponseText().trim() || "Balance as of month end";

  var balRow  = lastRow + 1;
  var prevRef = prevBalRow ? "J" + prevBalRow : "B1";

  sheet.getRange(balRow, 1).setValue(lastYear);
  sheet.getRange(balRow, 2).setValue(lastMonth);
  sheet.getRange(balRow, 3).setValue("BALANCE");
  sheet.getRange(balRow, 8).setValue(comment);
  sheet.getRange(balRow, 9).setFormula(
    "=SUM(D" + dataStart + ":D" + lastRow + ")-SUM(E" + dataStart + ":E" + lastRow + ")"
  );
  sheet.getRange(balRow, 10).setFormula("=" + prevRef + "+I" + balRow);

  logMonthlyRentToExpenses(ss, dataStart, lastRow, lastMonth, lastYear, sheet);
  protectClosedMonthRange(sheet, dataStart, balRow);

  SpreadsheetApp.flush();
  var runningBal = sheet.getRange(balRow, 10).getValue();
  ui.alert(
    "\u2705 " + lastMonth + " " + lastYear + " balanced!\n\n" +
    "Running Balance: \u20ac" + Number(runningBal).toFixed(2)
  );
}

// Auto-logs this month's Weekly Offerings col-E total into Expenses for record-keeping.
// Category is tagged "<Month>-<Year>-Monthly Rent" so the Summary formula can exclude it
// from Total Expenses — the amount is already netted into the Weekly Collections Balance.
function logMonthlyRentToExpenses(ss, dataStart, lastRow, lastMonth, lastYear, weeklySheet) {
  var expenseTotal = 0;
  for (var r = dataStart; r <= lastRow; r++) {
    expenseTotal += Number(weeklySheet.getRange(r, 5).getValue()) || 0;
  }
  if (expenseTotal <= 0) return;

  var expSheet = ss.getSheetByName("Expenses");
  if (!expSheet) return;

  var expLastRow = expSheet.getLastRow();
  while (expLastRow >= 3 && !expSheet.getRange(expLastRow, 1).getValue() && !expSheet.getRange(expLastRow, 3).getValue()) { expLastRow--; }
  var newRow = expLastRow + 1;

  var monthAbbrev = {
    January: "Jan", February: "Feb", March: "Mar", April: "Apr",
    May: "May", June: "Jun", July: "Jul", August: "Aug",
    September: "Sep", October: "Oct", November: "Nov", December: "Dec"
  };
  var expDate  = "01-" + (monthAbbrev[String(lastMonth)] || String(lastMonth)) + "-" + lastYear;
  var category = String(lastMonth) + "-" + String(lastYear) + "-Monthly Rent";

  expSheet.getRange(newRow, 1, 1, 6).setValues([[
    expDate,
    category,
    "Auto-logged weekly total from Weekly Offerings (col E) for " + lastMonth + " " + lastYear,
    expenseTotal,
    "Church Fund",
    "Already included in Weekly Collections Balance \u2014 excluded from Total Expenses in Summary"
  ]]);
}

// ============================================================
// SAFE REBUILD — reads all live data, rebuilds structure, restores data
// ============================================================
function safeRebuildWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var resp = ui.alert(
    "\ud83d\udd04 Safe Rebuild",
    "This will:\n" +
    "  \u2705 Preserve ALL existing Weekly Offerings rows\n" +
    "  \u2705 Preserve ALL member tithe payments\n" +
    "  \u2705 Preserve ALL expenses\n" +
    "  \u2705 Preserve ALL assets\n" +
    "  \u2705 Add Email column (col B) to Member Contributions\n" +
    "  \u2705 Remove Net Tithes row from Summary\n" +
    "  \u2705 Update rent instructions\n\n" +
    "Proceed?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var woCapture  = captureWeeklyOfferings(ss);
  var mtCapture  = captureMemberTithes(ss);
  var expCapture = captureExpenses(ss);
  var assetsCapture = captureAssets(ss);
  var backups = {};

  try {
    createSafeRebuildBackups(ss, backups);
    buildWeeklyOfferings(ss, woCapture);
    SpreadsheetApp.flush();
    if (woCapture.rows.length > 0 && ss.getSheetByName("Weekly Offerings").getLastRow() < woCapture.rows.length + 3) {
      throw new Error("Weekly Offerings could not be verified after rebuild.");
    }
    buildMemberTithes(ss, mtCapture.rows, true, mtCapture.hadEmailCol, mtCapture);
    SpreadsheetApp.flush();
    buildExpenses(ss, expCapture.rows);
    SpreadsheetApp.flush();
    buildAssets(ss, assetsCapture.rows);
    SpreadsheetApp.flush();
    buildSummary(ss);
    SpreadsheetApp.flush();
  } catch (err) {
    restoreSafeRebuildBackups(ss, backups);
    ui.alert("\u274c Safe rebuild failed. Your original sheets were restored.\n\nError: " + err.message);
    return;
  }
  deleteSafeRebuildBackups(ss, backups);

  ss.setActiveSheet(ss.getSheetByName("Weekly Offerings"));
  ui.alert(
    "\u2705 Safe rebuild complete!\n\n" +
    "All data has been preserved.\n" +
    "Email column added to Member Contributions \u2014 enter member emails in column B."
  );
}

function createSafeRebuildBackups(ss, backups) {
  var names = ["Weekly Offerings", "Member Contributions", "Expenses", "Assets", "Summary"];
  var suffix = "__KCF_BACKUP__" + new Date().getTime();
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (!sheet) continue;
    var backupName = names[i] + suffix;
    sheet.setName(backupName);
    sheet.hideSheet();
    backups[names[i]] = backupName;
  }
}

function restoreSafeRebuildBackups(ss, backups) {
  for (var name in backups) {
    var replacement = ss.getSheetByName(name);
    if (replacement) ss.deleteSheet(replacement);
    var backup = ss.getSheetByName(backups[name]);
    if (backup) {
      backup.setName(name);
      backup.showSheet();
    }
  }
}

function deleteSafeRebuildBackups(ss, backups) {
  for (var name in backups) {
    var backup = ss.getSheetByName(backups[name]);
    if (backup) ss.deleteSheet(backup);
  }
}

function captureWeeklyOfferings(ss) {
  var sheet = ss.getSheetByName("Weekly Offerings");
  if (!sheet) return { openingBalance: 0, rows: [] };
  var openingBalance = sheet.getRange(1, 2).getValue() || 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return { openingBalance: openingBalance, rows: [] };
  var rows = sheet.getRange(4, 1, lastRow - 3, 10).getValues();
  return { openingBalance: openingBalance, rows: rows };
}

function captureMemberTithes(ss) {
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) return { rows: [], hadEmailCol: false };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var col2Header = String(headers[1] || "").toLowerCase();
  var hadEmailCol = col2Header.indexOf("email") !== -1;
  var totalCol = 0;
  var notesCol = 0;
  var memberIdCol = 0;
  for (var col = 0; col < headers.length; col++) {
    var header = String(headers[col]).toLowerCase();
    if (header.indexOf("total received") !== -1) totalCol = col + 1;
    if (header === "notes") notesCol = col + 1;
    if (header === "member id") memberIdCol = col + 1;
  }
  var monthEndCol = totalCol ? totalCol - 1 : 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { rows: [], hadEmailCol: hadEmailCol, monthEndCol: monthEndCol, notesCol: notesCol, memberIdCol: memberIdCol };
  var readCols = Math.max(monthEndCol, notesCol, memberIdCol);
  var rows = sheet.getRange(3, 1, lastRow - 2, readCols).getValues();
  return { rows: rows, hadEmailCol: hadEmailCol, monthEndCol: monthEndCol, notesCol: notesCol, memberIdCol: memberIdCol };
}

function captureExpenses(ss) {
  var sheet = ss.getSheetByName("Expenses");
  if (!sheet) return { rows: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { rows: [] };
  var rows = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  return { rows: rows };
}

function captureAssets(ss) {
  var sheet = ss.getSheetByName("Assets");
  if (!sheet) return { rows: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { rows: [] };
  var rows = sheet.getRange(3, 1, lastRow - 2, 9).getValues();
  return { rows: rows };
}



// ============================================================
// MASTER BUILD — destructive (adds extra confirmation warning)
// ============================================================
function buildChurchFinanceWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    "\u26a0\ufe0f DESTRUCTIVE REBUILD",
    "This will DELETE ALL existing data and restore only the initial seed data.\n\n" +
    "Use '\ud83d\udd04 Safe Rebuild (Preserves Data)' to keep your data.\n\nAre you absolutely sure?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var typedConfirm = ui.prompt(
    "\u26a0\ufe0f Final Confirmation Required",
    "This is irreversible. Type DELETE (all caps) to permanently erase all financial data, or Cancel to back out.",
    ui.ButtonSet.OK_CANCEL
  );
  if (typedConfirm.getSelectedButton() !== ui.Button.OK || typedConfirm.getResponseText().trim() !== "DELETE") {
    ui.alert("\u274c Rebuild cancelled. No data was changed.");
    return;
  }

  var tempSheet = ss.insertSheet("__temp__" + new Date().getTime());

  try {
    buildWeeklyOfferings(ss, null);
    SpreadsheetApp.flush();
    buildMemberTithes(ss, null, false, false);
    SpreadsheetApp.flush();
    buildExpenses(ss, null);
    SpreadsheetApp.flush();
    buildAssets(ss, null);
    SpreadsheetApp.flush();
    buildSummary(ss);
  } catch (e) {
    ui.alert(
      "\u274c Build failed.\n\nError: " + e.message +
      "\n\nPlease copy this message and report it."
    );
    try { ss.deleteSheet(tempSheet); } catch (_) {}
    return;
  }

  ss.deleteSheet(tempSheet);
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  ss.setActiveSheet(ss.getSheetByName("Weekly Offerings"));
  ui.alert(
    "\u2705 Workbook rebuilt with initial seed data only.\n\n" +
    "Sheets:\n" +
    "  \u2022 Weekly Offerings       (Dec-2025 \u2192 Apr-2026 loaded)\n" +
    "  \u2022 Member Contributions  (Jan-2026 \u2192 Dec-2030, 9 members loaded)\n" +
    "  \u2022 Expenses               (Pastor contributions loaded)\n" +
    "  \u2022 Assets                 (known equipment loaded)\n" +
    "  \u2022 Summary                (live overall balance)"
  );
}

// Use only in a copy of the workbook before sharing it as a template.
function prepareCopyForSharing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "\ud83e\uddf9 Prepare This Copy for Sharing",
    "Use this only in a COPY of the workbook. It permanently removes all financial records, members, emails, member links, and the saved web-app URL from this file.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var weekly = ss.getSheetByName("Weekly Offerings");
  var members = ss.getSheetByName("Member Contributions");
  var expenses = ss.getSheetByName("Expenses");
  var assets = ss.getSheetByName("Assets");
  if (weekly && weekly.getLastRow() >= 4) weekly.getRange(4, 1, weekly.getLastRow() - 3, 10).clearContent();
  if (members && members.getLastRow() >= 3) members.getRange(3, 1, members.getLastRow() - 2, MEMBER_ID_COL).clearContent();
  if (expenses && expenses.getLastRow() >= 3) expenses.getRange(3, 1, expenses.getLastRow() - 2, 6).clearContent();
  if (assets && assets.getLastRow() >= 3) assets.getRange(3, 1, assets.getLastRow() - 2, 9).clearContent();

  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  for (var key in allProps) {
    if (key === "WEB_APP_URL" || key.indexOf("TOKEN_") === 0 || key.indexOf("MEMBER_") === 0 || key.indexOf("SHORTURL_") === 0) {
      props.deleteProperty(key);
    }
  }

  SpreadsheetApp.flush();
  ui.alert(
    "\u2705 This copy is clean and ready to share.\n\nBefore sharing, replace the church name and short name in both script files. The new church must deploy its own web app after adding members."
  );
}

// ============================================================
// onEdit — auto-calculates Month Net + Running Balance
// when "BALANCE" is typed in col C of Weekly Offerings
// ============================================================
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Weekly Offerings") return;

  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  if (editedCol !== 3 || editedRow < 4) return;

  var newValue = String(e.value || "").trim().toUpperCase();
  if (newValue !== "BALANCE") return;

  var dataStart = 4;
  for (var r = editedRow - 1; r >= 4; r--) {
    if (String(sheet.getRange(r, 3).getValue()).trim().toUpperCase() === "BALANCE") {
      dataStart = r + 1;
      break;
    }
  }
  var dataEnd = editedRow - 1;

  if (dataEnd < dataStart) {
    sheet.getRange(editedRow, 9).setValue(0);
    sheet.getRange(editedRow, 10).setFormula("=" + getPreviousBalance(sheet, editedRow));
    return;
  }

  sheet.getRange(editedRow, 9).setFormula(
    "=SUM(D" + dataStart + ":D" + dataEnd + ")" +
    "-SUM(E" + dataStart + ":E" + dataEnd + ")"
  );
  sheet.getRange(editedRow, 10).setFormula(
    "=" + getPreviousBalance(sheet, editedRow) + "+I" + editedRow
  );
}

function getPreviousBalance(sheet, currentRow) {
  for (var r = currentRow - 1; r >= 4; r--) {
    if (String(sheet.getRange(r, 3).getValue()).trim().toUpperCase() === "BALANCE") {
      return "J" + r;
    }
  }
  return "B1";
}

// ============================================================
// SHEET 1: Weekly Offerings
// capture = null → fresh seed build; capture = {openingBalance, rows} → safe rebuild
// ============================================================
function buildWeeklyOfferings(ss, capture) {
  var sheet = ss.getSheetByName("Weekly Offerings");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Weekly Offerings");

  var openingBalance = capture ? (capture.openingBalance || 0) : 0;

  // Row 1: Opening Balance
  sheet.getRange(1, 1).setValue("Opening Balance (\u20ac)");
  sheet.getRange(1, 2).setValue(openingBalance);
  sheet.getRange(1, 1, 1, 10).setBackground("#fff2cc").setFontWeight("bold");
  sheet.getRange(1, 1).setNote(
    "Closing balance carried over from previous year.\n" +
    "Leave as 0 if this is the first year."
  );

  // Row 2: Headers
  var headers = [
    "Year", "Month", "Date",
    "Offering (\u20ac)", "Expenses (\u20ac)", "Expense Type",
    "Proof", "Comments",
    "Month Net (\u20ac)", "Running Balance (\u20ac)"
  ];
  sheet.getRange(2, 1, 1, 10).setValues([headers]);
  sheet.getRange(2, 1, 1, 10)
       .setBackground("#1e3a5f").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  // Row 3: Instructions (rent stays in col E; Close Month Balance auto-logs it to Expenses)
  var instr = sheet.getRange(3, 1, 1, 10);
  instr.merge();
  instr.setValue(
    "\u2193 Enter each Sunday below (Year, Month, Date, Offering, Expenses). " +
    "At month-end: use \u26ea Church Finance \u2192 \ud83d\udcca Close Month Balance to calculate totals automatically. " +
    "\ud83c\udfe0 Rent stays in column E as before \u2014 Close Month Balance auto-logs the month's total to the Expenses sheet for your records."
  );
  instr.setFontStyle("italic").setFontColor("#666666")
       .setBackground("#f3f3f3").setWrap(true);
  sheet.setRowHeight(3, 28);

  // Conditional format: BALANCE rows → yellow + bold
  var balanceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=UPPER($C4)="BALANCE"')
    .setBackground("#fff2cc").setBold(true)
    .setRanges([sheet.getRange("A4:J")]).build();
  sheet.setConditionalFormatRules([balanceRule]);

  // Column widths
  sheet.setColumnWidth(1,  55);   // Year
  sheet.setColumnWidth(2,  90);   // Month
  sheet.setColumnWidth(3, 110);   // Date
  sheet.setColumnWidth(4, 110);   // Offering
  sheet.setColumnWidth(5, 110);   // Expenses
  sheet.setColumnWidth(6, 130);   // Expense Type
  sheet.setColumnWidth(7,  70);   // Proof
  sheet.setColumnWidth(8, 260);   // Comments
  sheet.setColumnWidth(9, 120);   // Month Net
  sheet.setColumnWidth(10, 155);  // Running Balance

  sheet.getRange("J4:J").setBackground("#e8f5e9").setFontWeight("bold");
  sheet.setFrozenRows(2);

  if (capture && capture.rows && capture.rows.length > 0) {
    restoreWeeklyOfferingsData(sheet, capture.rows);
  } else if (capture) {
    // Safe rebuild of an empty sheet or a new year starts empty
  } else {
    populateWeeklyHistory(sheet);
  }
}

// Restore captured Weekly Offerings rows and re-apply BALANCE formulas
// Marks a closed month's rows (Sunday entries + its BALANCE row) as warning-only protected —
// edits are still possible but require an explicit override, so they can no longer be silent.
function protectClosedMonthRange(sheet, startRow, endRow) {
  try {
    var protection = sheet.getRange(startRow, 1, endRow - startRow + 1, 10)
      .protect()
      .setDescription("Closed month \u2014 verify before editing");
    protection.setWarningOnly(true);
  } catch (e) {
    // Protections aren't critical to the close itself — don't block the month close if this fails
  }
}

function restoreWeeklyOfferingsData(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var writeRows = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i].slice(0, 10);
    while (r.length < 10) r.push("");
    r[8] = "";  // clear Month Net — will re-apply formula
    r[9] = "";  // clear Running Balance — will re-apply formula
    writeRows.push(r);
  }
  sheet.getRange(4, 1, writeRows.length, 10).setValues(writeRows);

  // Re-apply BALANCE row formulas
  var prevBalRow = null;
  for (var i = 0; i < rows.length; i++) {
    var sheetRow = i + 4;
    var dateVal  = String(rows[i][2] || "").trim().toUpperCase();
    if (dateVal !== "BALANCE") continue;

    var dataStart = prevBalRow ? prevBalRow + 1 : 4;
    var dataEnd   = sheetRow - 1;

    if (dataEnd >= dataStart) {
      sheet.getRange(sheetRow, 9).setFormula(
        "=SUM(D" + dataStart + ":D" + dataEnd + ")" +
        "-SUM(E" + dataStart + ":E" + dataEnd + ")"
      );
    } else {
      sheet.getRange(sheetRow, 9).setValue(0);
    }
    var prevRef = prevBalRow ? "J" + prevBalRow : "B1";
    sheet.getRange(sheetRow, 10).setFormula("=" + prevRef + "+I" + sheetRow);
    protectClosedMonthRange(sheet, dataStart, sheetRow);
    prevBalRow = sheetRow;
  }
}

// ============================================================
// Pre-populate Weekly Offerings — Dec-2025 through Aug-2026 (all months closed)
// ============================================================
function populateWeeklyHistory(sheet) {
  var blocks = [
    {
      rows: [
        [2025, "December", "14-12-2025",  124,    54,  "", "",    ""],
        [2025, "December", "21-12-2025",  133,    54,  "", "",    ""],
        [2025, "December", "28-12-2025",  120,    54,  "", "\u2713", ""],
        [2026, "January",  "04-01-2026",  140,    54,  "", "",    ""],
        [2026, "January",  "11-01-2026",   86,    54,  "", "",    ""],
        [2026, "January",  "18-01-2026",  123,    54,  "", "",    ""],
        [2026, "January",  "24-01-2026",  "NA",  108,  "", "",    "270 Paid (through Duleek account)"]
      ],
      balYear: 2026, balMonth: "January", balComment: "(Jan rent included in expenses)"
    },
    {
      rows: [
        [2026, "February", "01-02-2026",  171,    54,  "", "", ""],
        [2026, "February", "08-02-2026",  121,    54,  "", "", ""],
        [2026, "February", "15-02-2026",  273.4,  54,  "", "", ""],
        [2026, "February", "22-02-2026",  129,    54,  "", "", "Feb Rent - 216 paid (through Duleek account)"]
      ],
      balYear: 2026, balMonth: "February", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "March", "01-03-2026",  "NA", "NA", "", "", "Common worship"],
        [2026, "March", "08-03-2026",   117,   54,  "", "", ""],
        [2026, "March", "15-03-2026",   130,   54,  "", "", "Amt excluding 1 torn 5 euro note"],
        [2026, "March", "22-03-2026",   138,   54,  "", "", ""],
        [2026, "March", "29-03-2026",    78.2, 54,  "", "", "216 March Rent (through DCF account)"]
      ],
      balYear: 2026, balMonth: "March", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "April", "05-04-2026",  112,    54,  "", "", ""],
        [2026, "April", "12-04-2026",  173.5,  54,  "", "", ""],
        [2026, "April", "19-04-2026",  151,    54,  "", "", ""],
        [2026, "April", "26-04-2026",   88,    54,  "", "", "216 April Rent (through DCF account)"]
      ],
      balYear: 2026, balMonth: "April", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "May", "03-05-2026",  107,  54,  "", "", ""],
        [2026, "May", "10-05-2026",  103,  54,  "", "", ""],
        [2026, "May", "17-05-2026",   75,  54,  "", "", ""],
        [2026, "May", "24-05-2026",  110,  54,  "", "", ""],
        [2026, "May", "31-05-2026",  105,  54,  "", "", "270 May Rent (through DCF account)"]
      ],
      balYear: 2026, balMonth: "May", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "June", "07-06-2026",  116,    54,  "", "", ""],
        [2026, "June", "14-06-2026",  "NA",   90,  "", "", "Common worship - 5 hrs * 18 = 90"],
        [2026, "June", "21-06-2026",  130,    54,  "", "", ""],
        [2026, "June", "28-06-2026",   77.5,  54,  "", "", "252 June Rent (through DCF account)"]
      ],
      balYear: 2026, balMonth: "June", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "July", "05-07-2026",  112.9,   54,  "", "", ""],
        [2026, "July", "12-07-2026",  122.15,  54,  "", "", ""],
        [2026, "July", "19-07-2026",  175.5,   81,  "", "", "Pastor Biju meeting - 4.5 hrs"],
        [2026, "July", "26-07-2026",  123.5,   54,  "", "", "243 July Rent (through DCF account)"]
      ],
      balYear: 2026, balMonth: "July", balComment: "Balance as of month end"
    },
    {
      rows: [
        [2026, "August", "02-08-2026",  134.2,   54,  "", "", ""],
        [2026, "August", "09-08-2026",  137.05,  72,  "", "", ""],
        [2026, "August", "16-08-2026",  107,     54,  "", "", ""],
        [2026, "August", "23-08-2026",  124.6,   54,  "", "", ""],
        [2026, "August", "30-08-2026",   20,     54,  "", "", ""]
      ],
      balYear: 2026, balMonth: "August", balComment: "Balance as of month end"
    }
  ];

  var currentRow = 4;
  var prevBalRow = null;

  for (var b = 0; b < blocks.length; b++) {
    var block     = blocks[b];
    var dataStart = currentRow;

    sheet.getRange(currentRow, 1, block.rows.length, 8).setValues(block.rows);
    currentRow += block.rows.length;

    var dataEnd   = currentRow - 1;
    var balRowIdx = currentRow;

    sheet.getRange(balRowIdx, 1, 1, 8).setValues([[
      block.balYear, block.balMonth, "BALANCE", "", "", "", "", block.balComment
    ]]);
    sheet.getRange(balRowIdx, 9).setFormula(
      "=SUM(D" + dataStart + ":D" + dataEnd + ")" +
      "-SUM(E" + dataStart + ":E" + dataEnd + ")"
    );
    var prevRef = prevBalRow ? "J" + prevBalRow : "B1";
    sheet.getRange(balRowIdx, 10).setFormula("=" + prevRef + "+I" + balRowIdx);
    protectClosedMonthRange(sheet, dataStart, balRowIdx);

    prevBalRow  = balRowIdx;
    currentRow++;   // past balance row
    currentRow++;   // blank spacer row
  }
  // All months up to August 2026 are closed — no open current month block
}

// ============================================================
// SHEET 2: Member Contributions  (Jan-2026 → Dec-2050)
// capturedRowsOrList = null           → use initial seed list
// isCapturedData = true               → capturedRowsOrList is raw sheet rows
// hadEmailCol = whether captured data already had email col (only if isCapturedData=true)
// ============================================================
function buildMemberTithes(ss, capturedRowsOrList, isCapturedData, hadEmailCol, captureInfo) {
  capturedRowsOrList = capturedRowsOrList || buildInitialMemberList();

  var sheet = ss.getSheetByName("Member Contributions");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Member Contributions");

  // Build month headers from Dec-2025 through the configured end year
  var monthNames = ["Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"];
  var monthHeaders = ["Dec-2025"]; // month key 1 = Dec-2025
  for (var y = 2026; y <= TITHE_END_YEAR; y++) {
    for (var m = 0; m < 12; m++) {
      monthHeaders.push(monthNames[m] + "-" + y);
    }
  }

  var headers = ["Member Name", "Email", "Monthly Tithe (\u20ac)"]
    .concat(monthHeaders)
    .concat(["Total Received (\u20ac)", "Notes", "Member ID"]);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#1e3a5f").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  // Row 2: Instructions
  sheet.getRange(2, 1).setValue(
    "Enter \u20ac amounts paid each month (e.g. 100). " +
    "Enter 0 if not paid. Leave future months BLANK. " +
    "Green = paid \u2022 Red = unpaid (0). " +
    "Column B = member email (used to send receipts). " +
    "To record tithe fund expenses, use the \u2018Expenses\u2019 sheet."
  );
  sheet.getRange(2, 1, 1, headers.length)
       .setFontStyle("italic").setFontColor("#666666")
       .setBackground("#f3f3f3").setWrap(true);
  sheet.setRowHeight(2, 32);

  // Conditional formats on month columns (C:BJ)
  var monthRange = sheet.getRange(3, MONTH_COL_START, sheet.getMaxRows() - 2, NUM_MONTHS);
  var unpaidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground("#f4cccc").setFontColor("#cc0000")
    .setRanges([monthRange]).build();
  var paidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#d9ead3")
    .setRanges([monthRange]).build();
  sheet.setConditionalFormatRules([unpaidRule, paidRule]);

  // Column widths
  sheet.setColumnWidth(1, 150);  // Member Name
  sheet.setColumnWidth(2, 200);  // Email
  sheet.setColumnWidth(3, 120);  // Monthly Tithe
  for (var c = MONTH_COL_START; c <= MONTH_COL_END; c++) {
    sheet.setColumnWidth(c, 72); // month cols
  }
  sheet.setColumnWidth(TOTAL_RECV_COL, 140);
  sheet.setColumnWidth(NOTES_COL,      220);
  sheet.hideColumns(MEMBER_ID_COL);

  // Highlight Total Received column
  var tkLetter = columnToLetter(TOTAL_RECV_COL);
  sheet.getRange(tkLetter + "3:" + tkLetter)
       .setBackground("#e8f5e9").setFontWeight("bold");

  // Populate: captured data (safe rebuild) or seed list (fresh build)
  if (isCapturedData && capturedRowsOrList && capturedRowsOrList.length > 0) {
    restoreMemberTithesData(sheet, capturedRowsOrList, hadEmailCol, captureInfo);
  } else {
    var memberList = isCapturedData ? [] : capturedRowsOrList;
    for (var i = 0; i < memberList.length; i++) {
      var mRow   = i + 3;
      var member = memberList[i];
      sheet.getRange(mRow, 1).setValue(member.name);
      sheet.getRange(mRow, 2).setValue(member.email || "");
      sheet.getRange(mRow, 3).setValue(member.tithe);
      sheet.getRange(mRow, MEMBER_ID_COL).setValue(Utilities.getUuid());
      if (member.payments) {
        for (var key in member.payments) {
          var col = MONTH_COL_START + parseInt(key) - 1;
          if (col >= MONTH_COL_START && col <= MONTH_COL_END) {
            sheet.getRange(mRow, col).setValue(member.payments[key]);
          }
        }
      }
      applyMemberFormulas(sheet, mRow);
    }
  }

  sheet.setFrozenRows(2);
}

// Restore captured member rows: old layout (no email col) → new layout (with email col)
function restoreMemberTithesData(sheet, rows, hadEmailCol, captureInfo) {
  var writeRows = [];
  var sourceMonthStart = hadEmailCol ? MONTH_COL_START : MONTH_COL_START - 1;
  var sourceMonthEnd = captureInfo && captureInfo.monthEndCol ? captureInfo.monthEndCol : MONTH_COL_END;
  var sourceNotesCol = captureInfo && captureInfo.notesCol ? captureInfo.notesCol : NOTES_COL;
  var sourceMemberIdCol = captureInfo && captureInfo.memberIdCol ? captureInfo.memberIdCol : 0;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue; // skip blank rows

    var newRow = [];
    while (newRow.length < MEMBER_ID_COL) newRow.push("");
    newRow[0] = row[0];
    newRow[1] = hadEmailCol ? row[1] : "";
    newRow[2] = hadEmailCol ? row[2] : row[1];

    var monthCount = Math.min(sourceMonthEnd - sourceMonthStart + 1, NUM_MONTHS);
    for (var monthOffset = 0; monthOffset < monthCount; monthOffset++) {
      newRow[MONTH_COL_START - 1 + monthOffset] = row[sourceMonthStart - 1 + monthOffset];
    }
    newRow[NOTES_COL - 1] = row[sourceNotesCol - 1] || "";
    newRow[MEMBER_ID_COL - 1] = sourceMemberIdCol ? row[sourceMemberIdCol - 1] : "";
    if (!newRow[MEMBER_ID_COL - 1]) newRow[MEMBER_ID_COL - 1] = Utilities.getUuid();
    writeRows.push(newRow);
  }

  if (writeRows.length === 0) return;
  sheet.getRange(3, 1, writeRows.length, MEMBER_ID_COL).setValues(writeRows);
  for (var r = 0; r < writeRows.length; r++) {
    if (writeRows[r][0]) applyMemberFormulas(sheet, r + 3);
  }
}

// ============================================================
// Initial member list — historical payments Jan-Apr 2026
// keys in payments: 1=Jan-2026, 2=Feb-2026, 3=Mar-2026, 4=Apr-2026
// omit key = blank (not yet due); 0 = confirmed not paid (red)
// ============================================================
// Payment keys: 1=Dec-2025, 2=Jan-2026, 3=Feb-2026, 4=Mar-2026, 5=Apr-2026,
//               6=May-2026, 7=Jun-2026, 8=Jul-2026  (Sep/Oct not yet in old sheet)
function buildInitialMemberList() {
  return [
    { name: "Ajay",         email: "", tithe: 100, payments: { 1:100, 2:100, 3:100, 4:100, 5:100, 6:100, 7:100, 8:100 } },
    { name: "Anish",        email: "", tithe: 100, payments: { 1:0,   2:0,   3:0,   4:0,   5:100, 6:100, 7:100, 8:100 } },
    { name: "Dani",         email: "", tithe: 100, payments: { 1:100, 2:100, 3:100, 4:100, 5:100, 6:100, 7:100        } },
    { name: "Eldho",        email: "", tithe: 100, payments: { 1:100, 2:100, 3:100, 4:100, 5:100, 6:100, 7:200, 8:100, 9:100 } },
    { name: "Lisamma Aunty",email: "", tithe: 100, payments: { 1:100, 2:100, 3:100, 4:100, 5:100, 6:100, 7:100, 8:100, 9:100 } },
    { name: "Neema",        email: "", tithe: 100, payments: { 1:0,   2:0,   3:100, 4:100, 5:100, 6:100, 7:100, 8:100 } },
    { name: "Sajan",        email: "", tithe: 100, payments: { 1:100, 2:100, 3:100, 4:100, 5:100, 6:100, 7:100        } },
    { name: "Vineet",       email: "", tithe: 100, payments: { 1:100, 2:100, 3:500, 4:100, 5:100, 6:100, 7:100, 8:100 } },
    { name: "Reena Sister", email: "", tithe: 100, payments: { 1:0,   2:0,   3:0,   4:0,   5:0,   6:0,   7:100, 8:100 } },
    { name: "John",         email: "", tithe: 100, payments: {                                                         } }
  ];
}

// Writes Total Received SUMIF formula for a member row across all month columns
function applyMemberFormulas(sheet, row) {
  var startLetter  = columnToLetter(MONTH_COL_START); // D
  var endLetter    = columnToLetter(MONTH_COL_END);   // KR
  sheet.getRange(row, TOTAL_RECV_COL).setFormula(
    "=SUMIF(" + startLetter + row + ":" + endLetter + row + ",\">0\")"
  );
}

// ============================================================
// SHEET 3: Expenses — one row per common expense
// capturedRows = null → seed with initial pastor data
// capturedRows = array → restore from safe rebuild capture
// ============================================================
function buildExpenses(ss, capturedRows) {
  var sheet = ss.getSheetByName("Expenses");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Expenses");

  var headers = ["Date", "Category", "Description",
                 "Amount (\u20ac)", "Paid By", "Notes / Proof"];
  sheet.getRange(1, 1, 1, 6).setValues([headers]);
  sheet.getRange(1, 1, 1, 6)
       .setBackground("#1e3a5f").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  // Row 2: Instructions (updated to mention rent from Sept 2026)
  sheet.getRange(2, 1, 1, 6).merge()
       .setValue(
         "Record each common expense here \u2014 one row per expense. " +
         "Category examples: Pastor, Utilities, Event, Supplies, Other. " +
         "\ud83c\udfe0 Monthly rent is auto-logged here (category \u201c<Month>-<Year>-Monthly Rent\u201d) when you Close Month Balance \u2014 it's excluded from the total below since it's already reflected in the Weekly Collections Balance on the Summary sheet. " +
         "The total here is automatically deducted from the Overall Balance in the Summary."
       )
       .setFontStyle("italic").setFontColor("#666666")
       .setBackground("#f3f3f3").setWrap(true);
  sheet.setRowHeight(2, 64);

  // Pre-populate known pastor contributions OR restore captured data
  var initExpenses = capturedRows
    ? capturedRows.filter(function(r) { return r[0] || r[2] || r[3]; })
    : [
        ["30-Dec-2025", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (December)",                        250,     "Church Fund", ""],
        ["25-Jan-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (January)",                         250,     "Church Fund", ""],
        ["01-Feb-2026", "Church inauguration",       "Pastor Joseph - COG",                                                         100,     "Church Fund", ""],
        ["01-Feb-2026", "Church inauguration",       "Pastor Nick - COG",                                                           150,     "Church Fund", ""],
        ["01-Feb-2026", "Church inauguration",       "Food Arrangements",                                                           680,     "Church Fund", ""],
        ["22-Feb-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (Feb)",                             250,     "Church Fund", ""],
        ["29-Mar-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (March)",                           250,     "Church Fund", ""],
        ["19-Apr-2026", "Visiting Pastor",           "Pastor Joseph - COG",                                                         100,     "Church Fund", ""],
        ["02-May-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (April)",                           250,     "Church Fund", ""],
        ["16-May-2026", "June 3 Prayer",             "KCF contribution for Pr Suresh Babu's meeting",                              500,     "Church Fund", ""],
        ["24-May-2026", "Church Podium",             "Ordered from amazon.ie",                                                      160,     "Church Fund", ""],
        ["31-May-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (May)",                             400,     "Church Fund", ""],
        ["14-Jun-2026", "Common Worship",            "KCF - Common worship - Food Arrangements",                                   675,     "Church Fund", ""],
        ["05-Jul-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (June)",                            400,     "Church Fund", ""],
        ["19-Jul-2026", "Visiting Pastor",           "Pastor Biju CX",                                                             150,     "Church Fund", ""],
        ["20-Jul-2026", "Biju CX meeting",           "Food Arrangements",                                                           380,     "Church Fund", ""],
        ["02-Aug-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (July)",                            400,     "Church Fund", ""],
        ["02-Aug-2026", "Visiting Pastor",           "Guest Pastor - Daniel Pastor (UK Street Ministry)",                          100,     "Church Fund", ""],
        ["16-Aug-2026", "Lyrics Screen",             "100 inch Projector Screen",                                                    79,     "Church Fund", ""],
        ["19-Aug-2026", "Speakers",                  "2 RCF ART 912 + 1 Subzero Stage Monitor + 10 Year Warranty",               1857.27,  "Church Fund", ""],
        ["19-Aug-2026", "Speaker Covers",            "RCF ART Speaker Cover + Stage Monitor Wires",                               163.50,  "Church Fund", ""],
        ["22-Aug-2026", "Church Outing",             "Clonfert Farm - KFC Food Arrangements",                                      182,     "Church Fund", ""],
        ["30-Aug-2026", "Pastor",                    "Shoyin Pastor \u2014 Monthly Contribution (August)",                          400,     "Church Fund", ""]
      ];

  if (initExpenses.length > 0) {
    sheet.getRange(3, 1, initExpenses.length, 6).setValues(initExpenses);
  }

  // Conditional format: Amount > 0 \u2192 orange highlight
  var expRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#fce5cd").setBold(true)
    .setRanges([sheet.getRange("D3:D")]).build();
  sheet.setConditionalFormatRules([expRule]);

  // Style the Amount column
  sheet.getRange("D3:D")
       .setBackground("#fff2cc")
       .setNumberFormat("\u20ac#,##0.00");

  // Column widths
  sheet.setColumnWidth(1, 120);  // Date
  sheet.setColumnWidth(2, 140);  // Category
  sheet.setColumnWidth(3, 290);  // Description
  sheet.setColumnWidth(4, 130);  // Amount
  sheet.setColumnWidth(5, 140);  // Paid By
  sheet.setColumnWidth(6, 290);  // Notes / Proof

  sheet.setFrozenRows(1);
}

// ============================================================
// SHEET 5: Assets — fixed asset / equipment register (insurance, inventory, audit trail)
// capturedRows = null → seed with known existing equipment purchases
// capturedRows = array → restore from safe rebuild capture
// ============================================================
function buildAssets(ss, capturedRows) {
  var sheet = ss.getSheetByName("Assets");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Assets");

  var headers = ["Date Acquired", "Asset Name", "Description",
                 "Cost (\u20ac)", "Funding Source", "Location",
                 "Status", "Replacement Value (\u20ac)", "Notes / Serial No."];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#1e3a5f").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange(2, 1, 1, headers.length).merge()
       .setValue(
         "One row per physical asset the church owns (equipment, instruments, furniture). " +
         "Funding Source: Unrestricted (general funds) / Restricted (designated gift) / Donated In-Kind (never touched the church bank account). " +
         "Status: In Use / Damaged / Disposed / Donated Away. " +
         "Use \u26ea Church Finance \u2192 \ud83c\udff7\ufe0f Log New Asset to add a row, or add manually here."
       )
       .setFontStyle("italic").setFontColor("#666666")
       .setBackground("#f3f3f3").setWrap(true);
  sheet.setRowHeight(2, 52);

  var initAssets = capturedRows
    ? capturedRows.filter(function(r) { return r[0] || r[1] || r[3]; })
    : [
        ["24-May-2026", "Church Podium",    "Ordered from amazon.ie",                                    160,     "Unrestricted", "Main Hall", "In Use", 160,     ""],
        ["16-Aug-2026", "Projector Screen", "100 inch Projector Screen",                                  79,      "Unrestricted", "Main Hall", "In Use", 79,      "See Expenses 16-Aug-2026"],
        ["19-Aug-2026", "Speakers",         "2 RCF ART 912 + 1 Subzero Stage Monitor + 10 Year Warranty", 1857.27, "Unrestricted", "Main Hall", "In Use", 1857.27, "See Expenses 19-Aug-2026"],
        ["19-Aug-2026", "Speaker Covers",   "RCF ART Speaker Cover + Stage Monitor Wires",                163.50,  "Unrestricted", "Main Hall", "In Use", 163.50,  "See Expenses 19-Aug-2026"]
      ];

  if (initAssets.length > 0) {
    sheet.getRange(3, 1, initAssets.length, headers.length).setValues(initAssets);
  }

  // Conditional format: flag Status — green for In Use, red for Disposed/Damaged
  var inUseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("In Use")
    .setBackground("#d9ead3")
    .setRanges([sheet.getRange("G3:G")]).build();
  var disposedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Disposed")
    .setBackground("#f4cccc")
    .setRanges([sheet.getRange("G3:G")]).build();
  var damagedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Damaged")
    .setBackground("#fce5cd")
    .setRanges([sheet.getRange("G3:G")]).build();
  sheet.setConditionalFormatRules([inUseRule, disposedRule, damagedRule]);

  sheet.getRange("D3:D").setNumberFormat("\u20ac#,##0.00");
  sheet.getRange("H3:H").setNumberFormat("\u20ac#,##0.00");

  sheet.setColumnWidth(1, 110);  // Date Acquired
  sheet.setColumnWidth(2, 150);  // Asset Name
  sheet.setColumnWidth(3, 260);  // Description
  sheet.setColumnWidth(4, 100);  // Cost
  sheet.setColumnWidth(5, 140);  // Funding Source
  sheet.setColumnWidth(6, 120);  // Location
  sheet.setColumnWidth(7, 100);  // Status
  sheet.setColumnWidth(8, 150);  // Replacement Value
  sheet.setColumnWidth(9, 220);  // Notes / Serial No.

  sheet.setFrozenRows(2);
}
// Net Tithes row removed — not needed
// From Sept 2026: rent moves to Expenses sheet
// ============================================================
function buildSummary(ss) {
  var sheet = ss.getSheetByName("Summary");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Summary");

  var totalRecvLtr = columnToLetter(TOTAL_RECV_COL); // KS

  // ---- Helper: one labelled summary row ----
  function sRow(row, label, formula, bg) {
    sheet.getRange(row, 1).setValue(label)
         .setBackground(bg).setFontWeight("bold")
         .setFontSize(11).setVerticalAlignment("middle");
    sheet.getRange(row, 2).setFormula(formula)
         .setBackground(bg).setFontWeight("bold")
         .setFontSize(11).setHorizontalAlignment("right")
         .setVerticalAlignment("middle").setNumberFormat("\u20ac#,##0.00");
    sheet.getRange(row, 3).setBackground(bg);
    sheet.setRowHeight(row, 38);
  }

  // ---- Helper: dark-header divider row ----
  function headerRow(row, label, bg) {
    sheet.getRange(row, 1, 1, 3).merge()
         .setValue(label)
         .setBackground(bg).setFontColor("#ffffff")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheet.setRowHeight(row, 30);
  }

  // ---- Helper: subtotal row (coloured label + coloured value) ----
  function subtotalRow(row, label, formula, bg) {
    sheet.getRange(row, 1).setValue(label)
         .setBackground(bg).setFontColor("#ffffff")
         .setFontWeight("bold").setFontSize(12).setVerticalAlignment("middle");
    sheet.getRange(row, 2).setFormula(formula)
         .setBackground(bg).setFontColor("#ffffff")
         .setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right")
         .setVerticalAlignment("middle").setNumberFormat("\u20ac#,##0.00");
    sheet.getRange(row, 3).setBackground(bg);
    sheet.setRowHeight(row, 40);
  }

  // ---- Row 1: Title ----
  sheet.getRange(1, 1, 1, 3).merge()
       .setValue("Kildare Christian Fellowship \u2014 Financial Summary")
       .setBackground("#0b2545").setFontColor("#ffffff")
       .setFontWeight("bold").setFontSize(14)
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 44);

  // Row 2: spacer
  sheet.getRange(2, 1, 1, 3).setBackground("#f0f0f0");
  sheet.setRowHeight(2, 6);

  // ── INCOME ──────────────────────────────────────────────
  headerRow(3, "  \u2b07  INCOME", "#1e5f1e");

  // Row 4: Weekly Collections — use the most recent BALANCE row's Running Balance
  sRow(4,
    "Weekly Collections Balance",
    getWeeklyBalanceFormula(),
    "#e8f5e9"
  );

  // Row 5: Total member contributions received
  sRow(5,
    "Total Member Contributions Received",
    "=IFERROR(SUM('Member Contributions'!" + totalRecvLtr + "3:" + totalRecvLtr + "),0)",
    "#d9ead3"
  );

  // Row 6: Total Pooled Income subtotal
  subtotalRow(6, "\ud83d\udcb0  Total Pooled Income", "=B4+B5", "#276227");

  // Row 7: spacer
  sheet.getRange(7, 1, 1, 3).setBackground("#f0f0f0");
  sheet.setRowHeight(7, 6);

  // ── EXPENSES ────────────────────────────────────────────
  headerRow(8, "  \u2b06  EXPENSES", "#7a1e1e");

  // Row 9: Expenses sheet (pastor, events, supplies) — excludes auto-logged
  // "<Month>-<Year>-Monthly Rent" rows, which are already netted into the Weekly Collections Balance
  sRow(9,
    "Expenses Sheet  (Pastor, Events, Supplies \u2014 excludes auto-logged Monthly Rent)",
    "=IFERROR(SUMIF('Expenses'!B3:B,\"<>*Monthly Rent\",'Expenses'!D3:D),0)",
    "#f4cccc"
  );

  // Row 10: Total Expenses subtotal
  subtotalRow(10, "\ud83d\udcb8  Total Expenses", "=B9", "#7a1e1e");

  // Row 11: divider
  sheet.getRange(11, 1, 1, 3).setBackground("#cccccc");
  sheet.setRowHeight(11, 4);

  // Row 12: Overall Balance header
  sheet.getRange(12, 1, 1, 3).merge()
       .setValue("\u2b07  OVERALL BALANCE  (Pooled Income \u2212 Total Expenses)")
       .setBackground("#1e3a5f").setFontColor("#ffffff")
       .setFontWeight("bold").setFontSize(12)
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(12, 36);

  // Row 13: Overall Balance — derived from the subtotal rows above so it can never drift out of sync
  sheet.getRange(13, 1, 1, 3).merge()
       .setFormula("=B6-B10")
       .setBackground("#d9ead3").setFontWeight("bold").setFontSize(22)
       .setHorizontalAlignment("center").setVerticalAlignment("middle")
       .setNumberFormat("\u20ac#,##0.00");
  sheet.setRowHeight(13, 56);

  // Row 14: Cash in Hand (Treasurer) — manually updated, informational only, not part of any total.
  // Covers money members paid directly (e.g. Revolut) that's already counted above but not yet
  // physically transferred into the church bank account.
  sheet.getRange(14, 1).setValue("\ud83d\udcb3  Cash in Hand (Treasurer) \u2014 Pending Transfer to Bank")
       .setBackground("#cfe2f3").setFontWeight("bold")
       .setFontSize(11).setVerticalAlignment("middle");
  sheet.getRange(14, 2).setValue(0)
       .setBackground("#cfe2f3").setFontWeight("bold")
       .setFontSize(11).setHorizontalAlignment("right")
       .setVerticalAlignment("middle").setNumberFormat("\u20ac#,##0.00");
  sheet.getRange(14, 3).setBackground("#cfe2f3");
  sheet.getRange(14, 1).setNote(
    "Money members paid directly to the treasurer (e.g. Revolut) \u2014 already counted in Total Member " +
    "Contributions Received above, but not yet physically transferred into the church bank account.\n" +
    "Update this manually as it changes. Not included in any total above \u2014 informational only, for bank reconciliation."
  );
  sheet.setRowHeight(14, 38);

  // Row 15: spacer
  sheet.getRange(15, 1, 1, 3).setBackground("#f0f0f0");
  sheet.setRowHeight(15, 8);

  // Row 16: tip note (updated)
  sheet.getRange(16, 1, 1, 3).merge()
       .setValue(
         "\ud83d\udca1 Updates live automatically.\n" +
         "Weekly Offerings: enter Sunday amounts in col D, rent stays in col E. Type BALANCE at month-end for totals.\n" +
         "Rent is auto-logged to the Expenses sheet at month-close (excluded from totals \u2014 already included in Weekly Collections Balance above).\n" +
         "Pastor contributions, events, other \u2192 log in the Expenses sheet.\n" +
         "Send statements: \u26ea Church Finance \u2192 \ud83d\udce7 Send Selected Member Statement (or \ud83d\udce8 Send Annual Statements to All Members)."
       )
       .setBackground("#fff2cc").setFontStyle("italic")
       .setFontSize(10).setWrap(true).setVerticalAlignment("top");
  sheet.setRowHeight(16, 80);

  // Column widths
  sheet.setColumnWidth(1, 340);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 100);

  sheet.setFrozenRows(1);
}

function getWeeklyBalanceFormula() {
  return "=IFERROR(INDEX('Weekly Offerings'!J:J,MAX(FILTER(ROW('Weekly Offerings'!C4:C),'Weekly Offerings'!C4:C=\"BALANCE\"))),0)";
}

function refreshWeeklyCollectionsBalance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var weeklySheet = ss.getSheetByName("Weekly Offerings");
  var summarySheet = ss.getSheetByName("Summary");
  if (!weeklySheet || !summarySheet) {
    ui.alert("\u274c Weekly Offerings or Summary sheet not found.");
    return;
  }

  var lastRow = weeklySheet.getLastRow();
  var foundBalance = false;
  for (var row = lastRow; row >= 4; row--) {
    if (String(weeklySheet.getRange(row, 3).getValue()).trim().toUpperCase() === "BALANCE") {
      foundBalance = true;
      break;
    }
  }
  if (!foundBalance) {
    ui.alert("\u26a0\ufe0f No BALANCE row was found in Weekly Offerings. Restore the weekly data from Version History, then close the latest month.");
    return;
  }

  summarySheet.getRange(4, 2).setFormula(getWeeklyBalanceFormula());
  SpreadsheetApp.flush();
  ui.alert("\u2705 Weekly Collections Balance refreshed: \u20ac" + Number(summarySheet.getRange(4, 2).getValue()).toFixed(2));
}

// ============================================================
// ADD NEW MEMBER
// ============================================================
function addNewMember() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      "\u26a0\ufe0f Member Contributions sheet not found.\n" +
      "Please run \u2018Rebuild Workbook\u2019 first."
    );
    return;
  }

  var ui = SpreadsheetApp.getUi();

  var nameResp = ui.prompt("\u2795 Add New Member", "Enter member name:", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  var name = nameResp.getResponseText().trim();
  if (!name) { ui.alert("\u274c Name cannot be empty."); return; }

  var emailResp = ui.prompt("\u2795 Add New Member", "Enter email address (leave blank to add later):", ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  var email = emailResp.getResponseText().trim();

  var titheResp = ui.prompt("\u2795 Add New Member", "Enter monthly tithe amount (\u20ac):", ui.ButtonSet.OK_CANCEL);
  if (titheResp.getSelectedButton() !== ui.Button.OK) return;
  var tithe = parseFloat(titheResp.getResponseText().trim());
  if (isNaN(tithe) || tithe <= 0) {
    ui.alert("\u274c Invalid amount. Must be a number greater than 0.");
    return;
  }

  // Find last row that has a name in col A (avoids stray content pushing row down)
  var lastRow = sheet.getLastRow();
  var colA = lastRow >= 3
    ? sheet.getRange(3, 1, lastRow - 2, 1).getValues()
    : [];
  var lastNameRow = 2;
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] && String(colA[i][0]).trim() !== "") {
      lastNameRow = i + 3;
      break;
    }
  }
  var newRow = lastNameRow + 1;

  sheet.getRange(newRow, 1).setValue(name);
  sheet.getRange(newRow, 2).setValue(email);
  sheet.getRange(newRow, 3).setValue(tithe);
  var memberId = Utilities.getUuid();
  sheet.getRange(newRow, MEMBER_ID_COL).setValue(memberId);
  applyMemberFormulas(sheet, newRow);

  // Auto-generate token and show link if web app URL is already configured
  var props     = PropertiesService.getScriptProperties();
  var webAppUrl = props.getProperty("WEB_APP_URL");
  var tokenMsg  = "";
  if (webAppUrl) {
    var token = props.getProperty("TOKEN_" + memberId);
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, "");
      props.setProperty("TOKEN_" + memberId, token);
      props.setProperty("MEMBER_" + token, memberId);
    }
    var longUrl    = webAppUrl + "?token=" + token;
    var displayUrl = longUrl;
    var cached     = props.getProperty("SHORTURL_" + memberId);
    if (cached) {
      displayUrl = cached;
    } else {
      // alias = firstname + 5 token chars so it's recognisable but not guessable
      var firstName = name.split(/\s+/)[0].toLowerCase().replace(/[^a-z0-9]/g, "");
      var alias = firstName + "-" + token.substring(0, 5);
      try {
        var r2 = UrlFetchApp.fetch(
          "https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl) + "&alias=" + encodeURIComponent(alias),
          { muteHttpExceptions: true }
        );
        var s2 = r2.getContentText().trim();
        if (s2.indexOf("tinyurl.com") === -1) {
          // Alias already taken — fall back to random slug
          r2 = UrlFetchApp.fetch("https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl), { muteHttpExceptions: true });
          s2 = r2.getContentText().trim();
        }
        if (s2.indexOf("tinyurl.com") !== -1) {
          props.setProperty("SHORTURL_" + memberId, s2);
          displayUrl = s2;
        }
      } catch(e) {}
    }
    tokenMsg = "\n\n\ud83d\udd17 Personal link (share with member):\n" + displayUrl;
  } else {
    tokenMsg = "\n\n\u26a0\ufe0f No web app URL set yet.\nRun \u2699\ufe0f Set Web App URL, then \ud83d\udd17 Show Member Balance Links.";
  }

  ui.alert(
    "\u2705 " + name + " added!\n\n" +
    (email ? "Email: " + email : "\u26a0\ufe0f No email \u2014 enter it in column B of Member Contributions.") +
    "\nMonthly tithe: \u20ac" + tithe +
    tokenMsg +
    "\n\n\u2022 Enter past month payments in Member Contributions.\n" +
    "\u2022 Enter 0 for months confirmed not paid.\n" +
    "\u2022 Leave future months blank."
  );
}

// ============================================================
// LOG NEW ASSET
// ============================================================
function addNewAsset() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Assets");
  if (!sheet) {
    ui.alert("\u26a0\ufe0f Assets sheet not found.\nPlease run \u2018Safe Rebuild\u2019 first.");
    return;
  }

  var nameResp = ui.prompt("\ud83c\udff7\ufe0f Log New Asset", "Asset name (e.g. Speakers, Projector, Keyboard):", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  var name = nameResp.getResponseText().trim();
  if (!name) { ui.alert("\u274c Name cannot be empty."); return; }

  var descResp = ui.prompt("\ud83c\udff7\ufe0f Log New Asset", "Description (model, details \u2014 leave blank if none):", ui.ButtonSet.OK_CANCEL);
  if (descResp.getSelectedButton() !== ui.Button.OK) return;
  var description = descResp.getResponseText().trim();

  var costResp = ui.prompt("\ud83c\udff7\ufe0f Log New Asset", "Cost (\u20ac) \u2014 enter 0 if donated in-kind:", ui.ButtonSet.OK_CANCEL);
  if (costResp.getSelectedButton() !== ui.Button.OK) return;
  var cost = parseFloat(costResp.getResponseText().trim());
  if (isNaN(cost) || cost < 0) { ui.alert("\u274c Invalid amount."); return; }

  var fundResp = ui.prompt("\ud83c\udff7\ufe0f Log New Asset", "Funding source \u2014 type Unrestricted, Restricted, or Donated In-Kind:", ui.ButtonSet.OK_CANCEL);
  if (fundResp.getSelectedButton() !== ui.Button.OK) return;
  var fundingSource = fundResp.getResponseText().trim() || "Unrestricted";

  var locResp = ui.prompt("\ud83c\udff7\ufe0f Log New Asset", "Location (e.g. Main Hall \u2014 leave blank if unknown):", ui.ButtonSet.OK_CANCEL);
  if (locResp.getSelectedButton() !== ui.Button.OK) return;
  var location = locResp.getResponseText().trim();

  var lastRow = sheet.getLastRow();
  var newRow  = lastRow < 3 ? 3 : lastRow + 1;
  sheet.getRange(newRow, 1, 1, 9).setValues([[
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy"),
    name, description, cost, fundingSource, location, "In Use", cost, ""
  ]]);

  ui.alert(
    "\u2705 " + name + " added to Assets.\n\n" +
    "Cost: \u20ac" + cost + "\nFunding: " + fundingSource +
    "\n\nUpdate Status or Replacement Value directly in the Assets sheet as needed."
  );
}

// ============================================================
// WEB APP — member self-service balance & history page
// Deploy as: Execute as Me | Access Anyone
// Each member's link: .../exec?token=secure-token
// ============================================================
function doGet(e) {
  var token = String(e.parameter.token || "").trim();
  if (!token) return HtmlService.createHtmlOutput(errorPage("Invalid or missing access link."));

  var props        = PropertiesService.getScriptProperties();
  var memberId = props.getProperty("MEMBER_" + token);
  if (!memberId) return HtmlService.createHtmlOutput(errorPage("Access link is invalid or has expired."));

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) return HtmlService.createHtmlOutput(errorPage("Member Contributions sheet not found."));

  // Find member row
  var lastRow = sheet.getLastRow();
  var memberRow = null;
  for (var r = 3; r <= lastRow; r++) {
    var rowMemberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    if (rowMemberId === memberId) {
      memberRow = sheet.getRange(r, 1, 1, MEMBER_ID_COL).getValues()[0];
      break;
    }
  }
  if (!memberRow) return HtmlService.createHtmlOutput(errorPage("Member record not found. Contact your church administrator."));

  var memberName   = String(memberRow[0]);
  var memberNameHtml = escapeHtml(memberName);
  // Build month rows — labels computed directly to avoid Sheets date auto-conversion
  var mNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var monthRows = "";
  var DISPLAY_START_COL_IDX  = 1;  // Jan-2026
  var timeZone = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var currentYear = Number(Utilities.formatDate(now, timeZone, "yyyy"));
  var currentMonth = Number(Utilities.formatDate(now, timeZone, "M"));
  var CURRENT_MONTH_COL_IDX = Math.min(
    1 + (currentYear - TITHE_START_YEAR) * 12 + currentMonth - 1,
    NUM_MONTHS - 1
  );
  // Also show any later months that already have a recorded value (e.g. paid in advance)
  var lastRecordedColIdx = CURRENT_MONTH_COL_IDX;
  for (var lc = MONTH_COL_START; lc <= MONTH_COL_END; lc++) {
    var lIdx = lc - MONTH_COL_START;
    var lRaw = memberRow[lc - 1];
    if (lRaw !== "" && lRaw !== null && lRaw !== undefined && lIdx > lastRecordedColIdx) {
      lastRecordedColIdx = lIdx;
    }
  }
  var DISPLAY_THROUGH_COL_IDX = Math.min(lastRecordedColIdx, NUM_MONTHS - 1);
  var displayedTotal = 0;
  for (var c = MONTH_COL_START; c <= MONTH_COL_END; c++) {
    var colIdx = c - MONTH_COL_START;
    if (colIdx < DISPLAY_START_COL_IDX)  continue; // skip Dec-2025
    if (colIdx > DISPLAY_THROUGH_COL_IDX) break;   // stop after the current month
    var monthLabel = colIdx === 0 ? "Dec-2025" : mNames[(colIdx - 1) % 12] + "-" + (2026 + Math.floor((colIdx - 1) / 12));
    var raw = memberRow[c - 1];
    var isBlank = raw === "" || raw === null || raw === undefined;
    var val = isBlank ? 0 : Number(raw);

    var amountText, rowStyle, amountStyle;
    if (isBlank) {
      amountText  = "Not recorded";
      rowStyle    = "";
      amountStyle = "color:#888;";
    } else if (val > 0) {
      amountText  = "\u20ac" + raw;
      rowStyle    = "background:#f0fff0;";
      amountStyle = "color:#1e5f1e;font-weight:bold;";
      displayedTotal += val;
    } else {
      amountText  = "0";
      rowStyle    = "";
      amountStyle = "color:#aaa;";
    }
    monthRows += "<tr style='" + rowStyle + "border-bottom:1px solid #eee;'>" +
      "<td style='padding:10px 14px;color:#444;'>" + monthLabel + "</td>" +
      "<td style='padding:10px 14px;text-align:right;" + amountStyle + "'>" + amountText + "</td>" +
      "</tr>";
  }

  var html =
    "<!DOCTYPE html><html><head><meta charset='utf-8'>" +
    "<meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>KCF \u2014 " + memberNameHtml + "</title></head>" +
    "<body style='font-family:Arial,sans-serif;max-width:480px;margin:20px auto;padding:0 12px;'>" +

    "<div style='background:#0b2545;color:white;padding:20px;border-radius:6px 6px 0 0;text-align:center;'>" +
    "<div style='font-size:28px;'>\u271d</div>" +
    "<h2 style='margin:4px 0;'>Kildare Christian Fellowship</h2>" +
    "<p style='margin:0;opacity:0.8;font-size:13px;'>Member Contribution Statement</p>" +
    "</div>" +

    "<div style='background:#1e5f1e;color:white;padding:14px 20px;'>" +
    "<div style='font-size:18px;font-weight:bold;'>" + memberNameHtml + "</div>" +
    "</div>" +

    "<table style='width:100%;border-collapse:collapse;'>" +
    "<tr style='background:#1e3a5f;color:white;'>" +
    "<th style='padding:10px 14px;text-align:left;font-size:13px;font-weight:bold;'>Month</th>" +
    "<th style='padding:10px 14px;text-align:right;font-size:13px;font-weight:bold;'>Amount</th>" +
    "</tr>" +
    monthRows +
    "</table>" +

    "<div style='display:flex;justify-content:space-between;align-items:center;padding:14px 16px;background:#e8f5e9;border-top:2px solid #1e5f1e;'>" +
    "<span style='font-weight:bold;font-size:15px;'>Total Paid (Displayed Months)</span>" +
    "<span style='color:#1e5f1e;font-size:20px;font-weight:bold;'>\u20ac" + displayedTotal.toFixed(2) + "</span>" +
    "</div>" +

    "<div style='padding:14px;background:#fffbe6;border-top:3px solid #f0c000;text-align:center;font-size:13px;color:#666;'>" +
    "Thank you for your faithful giving. God bless you! \ud83d\ude4f" +
    "</div>" +
    "<div style='text-align:center;font-size:11px;color:#aaa;padding:10px;'>" +
    "Kildare Christian Fellowship \u2014 Last updated: " +
    Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd MMM yyyy HH:mm") +
    "</div>" +
    "</body></html>";

  return HtmlService.createHtmlOutput(html).setTitle("KCF \u2014 " + memberName);
}

function errorPage(msg) {
  return "<html><body style='font-family:Arial;padding:30px;text-align:center;'>" +
    "<h2>\u274c " + msg + "</h2>" +
    "<p style='color:#888;'>Contact your church administrator.</p></body></html>";
}

// ============================================================
// START NEW YEAR — archives current Weekly Offerings, opens fresh one
// ============================================================
function startNewYear() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var cur = ss.getSheetByName("Weekly Offerings");
  if (!cur) { ui.alert("\u274c 'Weekly Offerings' sheet not found."); return; }

  // Find closing balance = last BALANCE row's Running Balance (col J)
  var closingBalance = 0;
  var lastRow = cur.getLastRow();
  var lastEntryType = String(cur.getRange(lastRow, 3).getValue()).trim().toUpperCase();
  var lastEntryMonth = String(cur.getRange(lastRow, 2).getValue()).trim().toLowerCase();
  if (lastEntryType !== "BALANCE" || lastEntryMonth !== "december") {
    ui.alert("\u274c Close December with \ud83d\udcca Close Month Balance before starting a new year.");
    return;
  }
  for (var r = lastRow; r >= 4; r--) {
    if (String(cur.getRange(r, 3).getValue()).trim().toUpperCase() === "BALANCE") {
      closingBalance = Number(cur.getRange(r, 10).getValue()) || 0;
      break;
    }
  }

  var latestYear = null;
  for (var entryRow = lastRow; entryRow >= 4; entryRow--) {
    var entryYear = Number(cur.getRange(entryRow, 1).getValue());
    if (entryYear) { latestYear = entryYear; break; }
  }
  if (!latestYear) { ui.alert("\u274c No dated entries found in Weekly Offerings."); return; }
  var expectedYear = latestYear + 1;
  if (expectedYear > TITHE_END_YEAR) {
    ui.alert("\u274c Member Contributions currently supports through " + TITHE_END_YEAR + ". Extend it before starting " + expectedYear + ".");
    return;
  }

  var resp = ui.prompt(
    "\ud83d\udcc5 Start New Year",
    "Closing balance to carry forward: \u20ac" + closingBalance.toFixed(2) +
    "\n\nEnter the new year: " + expectedYear,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var newYear = resp.getResponseText().trim();
  if (newYear !== String(expectedYear)) { ui.alert("\u274c The next year must be " + expectedYear + "."); return; }

  var prevYear    = String(latestYear);
  var archiveName = "Weekly Offerings " + prevYear;
  if (ss.getSheetByName(archiveName)) {
    ui.alert("\u274c Sheet '" + archiveName + "' already exists.");
    return;
  }

  // Archive old sheet, build fresh one with carry-forward balance
  cur.setName(archiveName);
  buildWeeklyOfferings(ss, { openingBalance: closingBalance, blankForNewYear: true });
  ss.setActiveSheet(ss.getSheetByName("Weekly Offerings"));

  ui.alert(
    "\u2705 " + newYear + " started!\n\n" +
    "Opening balance: \u20ac" + closingBalance.toFixed(2) + " (carried from " + prevYear + ").\n" +
    "Old sheet archived as '" + archiveName + "'.\n\n" +
    "Member pages will show the new year automatically."
  );
}

// One-time setup: paste the correct Web App URL from Deploy \u2192 Manage deployments
function setWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    "\u2699\ufe0f Set Web App URL",
    "Paste the Web App URL from:\nDeploy \u2192 Manage deployments \u2192 KCF Member Portal \u2192 Copy URL",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var url = resp.getResponseText().trim().split("?")[0];
  if (!url || url.indexOf("script.google.com") === -1) {
    ui.alert("\u274c Invalid URL.");
    return;
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty("WEB_APP_URL", url);
  // Clear cached short links — they pointed to the old deployment URL
  var allProps = props.getProperties();
  var cleared  = 0;
  for (var key in allProps) {
    if (key.indexOf("SHORTURL_") === 0) { props.deleteProperty(key); cleared++; }
  }
  ui.alert(
    "\u2705 URL saved!" +
    (cleared > 0 ? "\n\u26a0\ufe0f " + cleared + " old short link(s) cleared.\nRun \u2702\ufe0f Shorten Member Links to regenerate." : "\n\nRun \u2702\ufe0f Shorten Member Links to create short links for all members.")
  );
}

// Generates a unique unguessable token per member; run once then reshare links
function generateMemberTokens() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) { ui.alert("Member Contributions sheet not found."); return; }

  var props    = PropertiesService.getScriptProperties();
  var lastRow  = sheet.getLastRow();
  var created  = 0;

  for (var r = 3; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    var memberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    if (!memberId) {
      ui.alert("\u26a0\ufe0f Member IDs are not set up yet. Run \ud83d\udd04 Safe Rebuild (Preserves Data) once, then try again.");
      return;
    }
    if (props.getProperty("TOKEN_" + memberId)) continue;
    var token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty("TOKEN_" + memberId, token);
    props.setProperty("MEMBER_" + token, memberId);
    created++;
  }

  ui.alert(
    "\u2705 " + created + " new token(s) generated.\n\n" +
    "Use \u26ea Church Finance \u2192 \ud83d\udd17 Show Member Balance Links to get the secure URLs."
  );
}

// Replaces every member token and removes old token/short-link mappings.
function rotateMemberLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) { ui.alert("Member Contributions sheet not found."); return; }

  var confirm = ui.alert(
    "\ud83d\udd04 Rotate Member Links",
    "This immediately revokes every existing member link. New secure links will be generated.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var lastRow = sheet.getLastRow();
  var memberIds = [];
  for (var r = 3; r <= lastRow; r++) {
    var memberName = String(sheet.getRange(r, 1).getValue()).trim();
    if (!memberName) continue;
    var memberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    if (!memberId) {
      ui.alert("\u26a0\ufe0f Member IDs are not set up yet. No links were changed. Run \ud83d\udd04 Safe Rebuild (Preserves Data) once, then try again.");
      return;
    }
    if (memberIds.indexOf(memberId) !== -1) {
      ui.alert("\u26a0\ufe0f Duplicate Member ID found. No links were changed. Run \ud83d\udd04 Safe Rebuild (Preserves Data) once, then try again.");
      return;
    }
    memberIds.push(memberId);
  }

  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  for (var key in allProps) {
    if (key.indexOf("TOKEN_") === 0 || key.indexOf("MEMBER_") === 0 || key.indexOf("SHORTURL_") === 0) {
      props.deleteProperty(key);
    }
  }

  var rotated = 0;
  for (var r = 3; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    var memberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    if (!memberId) {
      ui.alert("\u26a0\ufe0f Member IDs are not set up yet. Run \ud83d\udd04 Safe Rebuild (Preserves Data) once, then try again.");
      return;
    }
    var token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty("TOKEN_" + memberId, token);
    props.setProperty("MEMBER_" + token, memberId);
    rotated++;
  }

  ui.alert(
    "\u2705 " + rotated + " member link(s) rotated.\n\n" +
    "All old links are revoked. Run \u2702\ufe0f Shorten Member Links, then \ud83d\udd17 Show Member Balance Links to share the replacements."
  );
}

// Show each member's personal link
function showMemberLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) { ui.alert("Member Contributions sheet not found."); return; }

  var scriptUrl = PropertiesService.getScriptProperties().getProperty("WEB_APP_URL");
  if (!scriptUrl) {
    ui.alert("\u26a0\ufe0f Web App URL not set.\n\nRun: \u26ea Church Finance \u2192 \u2699\ufe0f Set Web App URL");
    return;
  }

  var props   = PropertiesService.getScriptProperties();
  var lastRow = sheet.getLastRow();
  var msg     = "Share these links with each member:\n\n";
  for (var r = 3; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    var memberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    var token = props.getProperty("TOKEN_" + memberId);
    if (!token) {
      msg += name + ":\n[No token \u2014 run \ud83d\udd11 Generate Member Tokens first]\n\n";
    } else {
      var shortUrl = props.getProperty("SHORTURL_" + memberId);
      var fullUrl  = scriptUrl + "?token=" + token;
      msg += name + ":\n  Full:  " + fullUrl + "\n";
      if (shortUrl) msg += "  Short: " + shortUrl + "\n";
      msg += "\n";
    }
  }
  ui.alert(msg);
}

// Generates TinyURL short links for all members and caches them in ScriptProperties
function shortenMemberLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) { ui.alert("Member Contributions sheet not found."); return; }

  var props     = PropertiesService.getScriptProperties();
  var webAppUrl = props.getProperty("WEB_APP_URL");
  if (!webAppUrl) {
    ui.alert("\u26a0\ufe0f Web App URL not set.\nRun \u2699\ufe0f Set Web App URL first.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var created = 0, skipped = 0, failed = 0;

  for (var r = 3; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    var memberId = String(sheet.getRange(r, MEMBER_ID_COL).getValue()).trim();
    if (!memberId) {
      ui.alert("\u26a0\ufe0f Member IDs are not set up yet. Run \ud83d\udd04 Safe Rebuild (Preserves Data) once, then try again.");
      return;
    }
    var token = props.getProperty("TOKEN_" + memberId);
    if (!token) continue;
    if (props.getProperty("SHORTURL_" + memberId)) { skipped++; continue; }

    var longUrl = webAppUrl + "?token=" + token;
    // alias = firstname + 5 token chars so it's recognisable but not guessable
    var firstName = name.split(/\s+/)[0].toLowerCase().replace(/[^a-z0-9]/g, "");
    var alias     = firstName + "-" + token.substring(0, 5);
    try {
      var resp = UrlFetchApp.fetch(
        "https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl) + "&alias=" + encodeURIComponent(alias),
        { muteHttpExceptions: true }
      );
      var shortUrl = resp.getContentText().trim();
      if (shortUrl.indexOf("tinyurl.com") === -1) {
        // Alias already taken — fall back to random slug
        resp     = UrlFetchApp.fetch("https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl), { muteHttpExceptions: true });
        shortUrl = resp.getContentText().trim();
      }
      if (shortUrl.indexOf("tinyurl.com") !== -1) {
        props.setProperty("SHORTURL_" + memberId, shortUrl);
        created++;
      } else {
        failed++;
      }
    } catch(e) {
      failed++;
    }
    Utilities.sleep(300); // avoid TinyURL rate limiting
  }

  ui.alert(
    "\u2705 Done!\n\n" +
    created + " short link(s) created." +
    (skipped > 0 ? "\n" + skipped + " already had short links (unchanged)." : "") +
    (failed > 0 ? "\n\u26a0\ufe0f " + failed + " failed \u2014 full URL used as fallback." : "") +
    "\n\nUse \ud83d\udd17 Show Member Balance Links to see all links."
  );
}

// ============================================================
// EMAIL RECEIPTS
// ============================================================

function sendReceipt(memberName, email, amount, monthYear) {
  var tz     = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var ref    = "KCF-" + Utilities.formatDate(new Date(), tz, "yyyyMMdd") +
               "-" + Math.floor(Math.random() * 900 + 100);
  var dateStr = Utilities.formatDate(new Date(), tz, "dd MMMM yyyy");

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #ddd;border-radius:4px;overflow:hidden;'>" +
    "<div style='background:#0b2545;color:white;padding:20px;text-align:center;'>" +
    "<h2 style='margin:0;font-size:20px;'>\u271d Kildare Christian Fellowship</h2>" +
    "<p style='margin:6px 0 0;font-size:13px;opacity:0.85;'>Official Payment Receipt</p>" +
    "</div>" +
    "<div style='padding:28px;'>" +
    "<table style='width:100%;border-collapse:collapse;font-size:15px;'>" +
    "<tr><td style='padding:10px 0;border-bottom:1px solid #eee;color:#555;width:42%;'>Member</td>" +
    "<td style='padding:10px 0;border-bottom:1px solid #eee;font-weight:bold;'>" + escapeHtml(memberName) + "</td></tr>" +
    "<tr><td style='padding:10px 0;border-bottom:1px solid #eee;color:#555;'>Amount Received</td>" +
    "<td style='padding:10px 0;border-bottom:1px solid #eee;font-weight:bold;font-size:20px;color:#1e5f1e;'>" +
    "\u20ac" + amount + "</td></tr>" +
    "<tr><td style='padding:10px 0;border-bottom:1px solid #eee;color:#555;'>Period</td>" +
    "<td style='padding:10px 0;border-bottom:1px solid #eee;'>" + monthYear + "</td></tr>" +
    "<tr><td style='padding:10px 0;border-bottom:1px solid #eee;color:#555;'>Date Issued</td>" +
    "<td style='padding:10px 0;border-bottom:1px solid #eee;'>" + dateStr + "</td></tr>" +
    "<tr><td style='padding:10px 0;color:#555;'>Reference</td>" +
    "<td style='padding:10px 0;font-family:monospace;font-size:13px;color:#888;'>" + ref + "</td></tr>" +
    "</table>" +
    "<div style='margin-top:24px;padding:14px;background:#f0f8f0;border-left:4px solid #1e5f1e;border-radius:2px;text-align:center;'>" +
    "<p style='margin:0;color:#1e5f1e;font-style:italic;font-size:14px;'>Thank you for your faithful giving. God bless you! \ud83d\ude4f</p>" +
    "</div>" +
    "</div>" +
    "<div style='background:#f5f5f5;padding:10px;text-align:center;font-size:11px;color:#999;'>" +
    "Kildare Christian Fellowship \u2014 This is an official receipt" +
    "</div>" +
    "</div>";

  MailApp.sendEmail({
    to:       email,
    subject:  "KCF Receipt \u2014 " + monthYear + " \u2014 " + memberName,
    htmlBody: html
  });
}

// Sends a full-year contribution statement only to the selected Member Contributions row.
function sendAnnualStatements() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  var selectedRow = ss.getActiveRange().getRow();
  if (sheet.getName() !== "Member Contributions" || selectedRow < 3) {
    ui.alert("\u26a0\ufe0f Select any cell in the member's row on the Member Contributions sheet, then try again.");
    return;
  }

  var data = sheet.getRange(selectedRow, 1, 1, MEMBER_ID_COL).getValues()[0];
  var memberName = String(data[0]).trim();
  var email = String(data[EMAIL_COL - 1]).trim();
  if (!memberName) { ui.alert("\u274c The selected row has no member name."); return; }
  if (!email || email.indexOf("@") === -1) {
    ui.alert("\u274c " + memberName + " does not have a valid email address in column B.");
    return;
  }

  var yearResp = ui.prompt(
    "\ud83d\udce7 Send Annual Statements",
    "Enter the year to send statements for (e.g. 2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (yearResp.getSelectedButton() !== ui.Button.OK) return;
  var yearInput = yearResp.getResponseText().trim();
  if (!/^\d{4}$/.test(yearInput)) { ui.alert("\u274c Enter a 4-digit year."); return; }
  var year = parseInt(yearInput);
  if (year < TITHE_START_YEAR || year > TITHE_END_YEAR) {
    ui.alert("\u274c Statements are available for " + TITHE_START_YEAR + " through " + TITHE_END_YEAR + ".");
    return;
  }

  var props     = PropertiesService.getScriptProperties();
  var scriptUrl = props.getProperty("WEB_APP_URL") || "";
  var mNames    = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  // Array index of January in the selected row. Column D is Dec-2025.
  var janIdx = MONTH_COL_START + (year - TITHE_START_YEAR) * 12;
  var monthRows = "";
  var yearTotal = 0;
  for (var m = 0; m < 12; m++) {
    var val = Number(data[janIdx + m]) || 0;
    var label = mNames[m] + "-" + year;
    var amtTd = val > 0
      ? "<td style='padding:8px 14px;text-align:right;font-weight:bold;color:#1e5f1e;'>\u20ac" + val + "</td>"
      : "<td style='padding:8px 14px;text-align:right;color:#aaa;'>0</td>";
    monthRows += "<tr style='background:" + (val > 0 ? "#f0fff0" : "") + ";border-bottom:1px solid #eee;'>" +
      "<td style='padding:8px 14px;color:#444;'>" + label + "</td>" + amtTd + "</tr>";
    if (val > 0) yearTotal += val;
  }

  var memberId = String(data[MEMBER_ID_COL - 1]).trim();
  var token = props.getProperty("TOKEN_" + memberId);
  var linkBtn = (scriptUrl && token)
    ? "<p style='text-align:center;margin:16px 0 4px;'><a href='" + scriptUrl + "?token=" + token +
      "' style='background:#1e3a5f;color:white;padding:10px 22px;border-radius:4px;text-decoration:none;font-size:13px;'>View Live Statement</a></p>"
    : "";

  var confirm = ui.alert(
    "Send Selected Member Statement",
    "Send the " + year + " annual statement to:\n\n" + memberName + "\n" + email,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    sendAnnualStatementEmail(memberName, email, year, yearTotal, monthRows, linkBtn);
    ui.alert("\u2705 Annual statement sent to " + memberName + " (" + email + ").");
  } catch (err) {
    ui.alert("\u274c Statement could not be sent.\n\nError: " + err.message);
  }
}

// Sends a selected year's statement to every member with a valid email address.
function sendAllAnnualStatements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Member Contributions");
  if (!sheet) { ui.alert("\u274c Member Contributions sheet not found."); return; }

  var yearResp = ui.prompt(
    "Send Annual Statements to All Members",
    "Enter the year to send statements for (e.g. 2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (yearResp.getSelectedButton() !== ui.Button.OK) return;
  var yearInput = yearResp.getResponseText().trim();
  if (!/^\d{4}$/.test(yearInput)) { ui.alert("\u274c Enter a 4-digit year."); return; }
  var year = parseInt(yearInput, 10);
  if (year < TITHE_START_YEAR || year > TITHE_END_YEAR) {
    ui.alert("\u274c Statements are available for " + TITHE_START_YEAR + " through " + TITHE_END_YEAR + ".");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) { ui.alert("\u274c No member data found."); return; }
  var members = sheet.getRange(3, 1, lastRow - 2, MEMBER_ID_COL).getValues();
  var recipients = members.filter(function(row) {
    return String(row[0]).trim() && String(row[EMAIL_COL - 1]).trim().indexOf("@") !== -1;
  });
  if (recipients.length === 0) { ui.alert("\u274c No members with valid email addresses were found."); return; }

  var confirm = ui.alert(
    "Confirm Bulk Annual Statements",
    "Send the " + year + " annual statement to " + recipients.length + " member(s)?\n\nEach recipient will receive only their own statement.",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var props = PropertiesService.getScriptProperties();
  var scriptUrl = props.getProperty("WEB_APP_URL") || "";
  var mNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var janIdx = MONTH_COL_START + (year - TITHE_START_YEAR) * 12;
  var sent = 0;
  var failed = 0;

  for (var i = 0; i < recipients.length; i++) {
    var data = recipients[i];
    var memberName = String(data[0]).trim();
    var email = String(data[EMAIL_COL - 1]).trim();
    var monthRows = "";
    var yearTotal = 0;
    for (var m = 0; m < 12; m++) {
      var val = Number(data[janIdx + m]) || 0;
      var label = mNames[m] + "-" + year;
      var amountCell = val > 0
        ? "<td style='padding:8px 14px;text-align:right;font-weight:bold;color:#1e5f1e;'>\u20ac" + val + "</td>"
        : "<td style='padding:8px 14px;text-align:right;color:#aaa;'>0</td>";
      monthRows += "<tr style='background:" + (val > 0 ? "#f0fff0" : "") + ";border-bottom:1px solid #eee;'>" +
        "<td style='padding:8px 14px;color:#444;'>" + label + "</td>" + amountCell + "</tr>";
      if (val > 0) yearTotal += val;
    }

    var memberId = String(data[MEMBER_ID_COL - 1]).trim();
    var token = props.getProperty("TOKEN_" + memberId);
    var linkBtn = (scriptUrl && token)
      ? "<p style='text-align:center;margin:16px 0 4px;'><a href='" + scriptUrl + "?token=" + token +
        "' style='background:#1e3a5f;color:white;padding:10px 22px;border-radius:4px;text-decoration:none;font-size:13px;'>View Live Statement</a></p>"
      : "";
    try {
      sendAnnualStatementEmail(memberName, email, year, yearTotal, monthRows, linkBtn);
      sent++;
      Utilities.sleep(200);
    } catch (err) {
      failed++;
    }
  }

  ui.alert(
    "\u2705 Bulk annual statements complete.\n\nSent: " + sent +
    (failed ? "\nFailed: " + failed : "")
  );
}

function sendAnnualStatementEmail(memberName, email, year, totalPaid, monthRows, linkBtn) {
  var tz      = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var dateStr = Utilities.formatDate(new Date(), tz, "dd MMMM yyyy");
  var ref     = "KCF-" + year + "-" + memberName.replace(/\s+/g, "").toUpperCase();

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #ddd;border-radius:4px;overflow:hidden;'>" +
    "<div style='background:#0b2545;color:white;padding:20px;text-align:center;'>" +
    "<h2 style='margin:0;font-size:20px;'>\u271d Kildare Christian Fellowship</h2>" +
    "<p style='margin:6px 0 0;font-size:13px;opacity:0.85;'>Annual Contribution Statement \u2014 " + year + "</p>" +
    "</div>" +
    "<div style='background:#1e5f1e;color:white;padding:14px 20px;'>" +
    "<div style='font-size:18px;font-weight:bold;'>" + escapeHtml(memberName) + "</div>" +
    "<div style='font-size:12px;opacity:0.8;'>Ref: " + ref + "</div>" +
    "</div>" +
    "<table style='width:100%;border-collapse:collapse;'>" +
    "<tr style='background:#1e3a5f;color:white;'>" +
    "<th style='padding:10px 14px;text-align:left;font-size:13px;'>Month</th>" +
    "<th style='padding:10px 14px;text-align:right;font-size:13px;'>Amount</th>" +
    "</tr>" +
    monthRows +
    "</table>" +
    "<table style='width:100%;border-collapse:collapse;background:#e8f5e9;border-top:3px solid #1e5f1e;'>" +
    "<tr>" +
    "<td style='padding:16px 14px;font-weight:bold;font-size:15px;color:#333;'>Total Paid " + year + "</td>" +
    "<td style='padding:16px 14px;text-align:right;color:#1e5f1e;font-size:22px;font-weight:bold;'>\u20ac" + totalPaid.toFixed(2) + "</td>" +
    "</tr>" +
    "</table>" +
    linkBtn +
    "<div style='padding:14px;background:#fffbe6;border-top:3px solid #f0c000;text-align:center;font-size:13px;color:#666;'>" +
    "Thank you for your faithful giving. God bless you! \ud83d\ude4f" +
    "</div>" +
    "<div style='background:#f5f5f5;padding:10px;text-align:center;font-size:11px;color:#999;'>" +
    "Kildare Christian Fellowship \u2014 Issued: " + dateStr +
    "</div>" +
    "</div>";

  MailApp.sendEmail({
    to:       email,
    subject:  "KCF Annual Statement \u2014 " + year + " \u2014 " + memberName,
    htmlBody: html
  });
}

// ============================================================
// REFRESH SUMMARY
// ============================================================
function refreshSummary() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ Summary refreshed.");
}

// ============================================================
// Helper: column number to letter  e.g. 64 → "BL"
// ============================================================
function columnToLetter(col) {
  var letter = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
