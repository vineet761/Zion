// ============================================================
// KCF How To Use — paste this file into Apps Script as a new .gs file
// Then add to onOpen() menu:  .addItem("❓ How To Use", "showHowTo")
// ============================================================
function showHowTo() {
  var html = HtmlService.createHtmlOutput(getHowToHtml())
    .setTitle("KCF — How To Use")
    .setWidth(430);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHowToHtml() {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;color:#333;margin:0;padding:12px;}' +
    'h1{background:#0b2545;color:white;margin:-12px -12px 16px;padding:14px 16px;font-size:15px;}' +
    'h2{color:#1e3a5f;font-size:13px;margin:18px 0 6px;border-bottom:2px solid #1e3a5f;padding-bottom:4px;}' +
    '.step{display:flex;gap:8px;margin:5px 0;align-items:flex-start;}' +
    '.num{background:#1e3a5f;color:white;border-radius:50%;min-width:20px;height:20px;' +
      'display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:bold;flex-shrink:0;margin-top:1px;}' +
    '.tip{background:#fffbe6;border-left:3px solid #f0c000;padding:7px 10px;margin:7px 0;font-size:12px;}' +
    '.warn{background:#fce5cd;border-left:3px solid #e65c00;padding:7px 10px;margin:7px 0;font-size:12px;}' +
    'code{background:#f0f0f0;padding:1px 4px;border-radius:2px;font-size:12px;}' +
    'ul{margin:4px 0;padding-left:18px;} li{margin:3px 0;}' +
    '.tag{display:inline-block;background:#1e3a5f;color:white;padding:2px 6px;border-radius:3px;font-size:11px;font-weight:bold;}' +
    '.footer{margin-top:20px;padding:10px;background:#f0f0f0;font-size:11px;color:#999;text-align:center;}' +
    '</style></head><body>' +

    '<h1>\u271d KCF Financial Workbook \u2014 How To Use</h1>' +

    '<h2>\ud83d\udccb The 5 Sheets</h2>' +
    '<ul>' +
    '<li><span class="tag">Weekly Offerings</span> \u2014 Sunday collections &amp; month balances</li>' +
    '<li><span class="tag">Member Contributions</span> \u2014 Individual monthly payments</li>' +
    '<li><span class="tag">Expenses</span> \u2014 Pastor, events, supplies (rent auto-logged here for records)</li>' +
    '<li><span class="tag">Assets</span> \u2014 Equipment / property register (insurance &amp; inventory)</li>' +
    '<li><span class="tag">Summary</span> \u2014 Live overall balance (read-only)</li>' +
    '</ul>' +

    '<h2>\ud83d\udcc5 Every Sunday</h2>' +
    '<div class="step"><div class="num">1</div><div>Open <b>Weekly Offerings</b> sheet</div></div>' +
    '<div class="step"><div class="num">2</div><div>Add a row: <b>Year \u00b7 Month \u00b7 Date \u00b7 Offering (\u20ac) \u00b7 Expenses (\u20ac)</b></div></div>' +
    '<div class="step"><div class="num">3</div><div>Leave Month Net &amp; Running Balance blank \u2014 calculated automatically at month-end</div></div>' +

    '<h2>\ud83d\udcca Close Month (end of each month)</h2>' +
    '<div class="step"><div class="num">1</div><div>Ensure all Sundays for the month are entered</div></div>' +
    '<div class="step"><div class="num">2</div><div>Click <b>\u26ea Church Finance \u2192 \ud83d\udcca Close Month Balance</b></div></div>' +
    '<div class="step"><div class="num">3</div><div>Confirm the month, add a comment (optional), click OK</div></div>' +
    '<div class="tip">\ud83d\udca1 A BALANCE row is inserted with Month Net and Running Balance calculated automatically.</div>' +

    '<h2>\u2795 Adding a New Member</h2>' +
    '<div class="step"><div class="num">1</div><div>Click <b>\u26ea Church Finance \u2192 \u2795 Add New Member</b></div></div>' +
    '<div class="step"><div class="num">2</div><div>Enter their <b>name</b> when prompted</div></div>' +
    '<div class="step"><div class="num">3</div><div>Enter their <b>email address</b> (needed for annual statements)</div></div>' +
    '<div class="step"><div class="num">4</div><div>Enter their <b>monthly tithe (\u20ac)</b> amount</div></div>' +
    '<div class="step"><div class="num">5</div><div>Their secure personal link is shown immediately \u2014 copy and share it</div></div>' +
    '<div class="tip">\ud83d\udca1 Full names are supported. Two people with the same name receive separate private links.</div>' +
    '<div class="tip">\ud83d\udca1 Requires <b>\u2699\ufe0f Set Web App URL</b> to have been run at least once. If not, the link is skipped but you can generate it later via \ud83d\udd17 Show Member Balance Links.</div>' +

    '<h2>\ud83d\udcb0 Tithe Payments (Member Contributions sheet)</h2>' +
    '<ul>' +
    '<li>Enter the <b>\u20ac amount</b> in the correct month column</li>' +
    '<li>Enter <code>0</code> if they did <b>not</b> pay that month</li>' +
    '<li>Leave future months <b>blank</b></li>' +
    '</ul>' +
    '<ul>' +
    '<li>\ud83d\udfe2 Green = paid</li>' +
    '<li>\ud83d\udd34 Red cell = entered as 0 (confirmed not paid)</li>' +
    '<li>\u2b1c Blank = not yet due</li>' +
    '</ul>' +
    '<div class="tip">\ud83d\udca1 If a member pays you directly (e.g. Revolut) instead of through the bank, still mark it paid here as usual. Then update <b>\ud83d\udcb3 Cash in Hand (Treasurer)</b> on the Summary sheet with the running total you\'re holding but haven\'t transferred to the church account yet \u2014 it\'s informational only and isn\'t part of any total, but keeps that money from being forgotten.</div>' +

    '<h2>\ud83d\udce7 Send a Member Statement</h2>' +
    '<div class="step"><div class="num">1</div><div>In <b>Member Contributions</b>, select any cell in the member\'s row. Confirm their email is in <b>column B</b>.</div></div>' +
    '<div class="step"><div class="num">2</div><div>Click <b>\u26ea Church Finance \u2192 \ud83d\udce7 Send Selected Member Statement</b></div></div>' +
    '<div class="step"><div class="num">3</div><div>Enter the year (e.g. 2026) and confirm</div></div>' +
    '<div class="tip">\ud83d\udca1 Only the selected member receives an email. This works safely even when two members have the same name.</div>' +
    '<div class="tip">\ud83d\udce8 To send to everyone with a valid email, use <b>\u26ea Church Finance \u2192 \ud83d\udce8 Send Annual Statements to All Members</b>. It shows the recipient count and requires a separate confirmation.</div>' +

    '<h2>\ud83d\udd17 Member Web Pages (one-time setup)</h2>' +
    '<div class="step"><div class="num">1</div><div>In Apps Script editor: <b>Deploy \u2192 New deployment \u2192 Web App</b><br>Execute as: <b>Me</b> | Access: <b>Anyone</b></div></div>' +
    '<div class="step"><div class="num">2</div><div>Copy the Web App URL</div></div>' +
    '<div class="step"><div class="num">3</div><div>Click <b>\u2699\ufe0f Set Web App URL</b> and paste the URL</div></div>' +
    '<div class="step"><div class="num">4</div><div>Click <b>\ud83d\udd11 Generate Member Tokens</b> (safe to run again at any time)</div></div>' +
    '<div class="step"><div class="num">5</div><div>Click <b>\ud83d\udd17 Show Member Balance Links</b> \u2014 share each person\'s link</div></div>' +
    '<div class="warn">\u26a0\ufe0f Use <b>\ud83d\udd04 Rotate Member Links (Revokes Old)</b> only when a link is compromised or needs replacing. It invalidates every existing member link.</div>' +
    '<div class="warn">\u26a0\ufe0f After upgrading this script, run <b>\ud83d\udd04 Safe Rebuild (Preserves Data)</b> once, then generate and re-share member links. Existing payments are preserved.</div>' +
    '<div class="warn">\u26a0\ufe0f If you create a <b>NEW deployment</b>, repeat steps 2\u20133 and re-share all links \u2014 the URL changes.</div>' +

    '<h2>\ud83d\udcb8 Logging Expenses</h2>' +
    '<ul>' +
    '<li><b>Pastor contributions</b> \u2192 Expenses sheet</li>' +
    '<li><b>Events / food</b> \u2192 Expenses sheet</li>' +
    '<li><b>Rent</b> \u2192 keep entering weekly amounts in Weekly Offerings col E as usual; <b>Close Month Balance</b> auto-logs the month\u2019s total to the Expenses sheet (category \u201c&lt;Month&gt;-&lt;Year&gt;-Monthly Rent\u201d), excluded from the Expenses total since it\u2019s already reflected in the Weekly Collections Balance</li>' +
    '</ul>' +
    '<div class="tip">\ud83d\udca1 The exclusion is just a text match: the Summary formula skips any Expenses row whose Category ends in \u201cMonthly Rent\u201d. Don\u2019t rename or repurpose that category text for a real (non-auto-logged) expense \u2014 it would silently drop out of Total Expenses.</div>' +

    '<h2>\ud83c\udff7\ufe0f Logging Assets / Equipment</h2>' +
    '<div class="step"><div class="num">1</div><div>Click <b>\u26ea Church Finance \u2192 \ud83c\udff7\ufe0f Log New Asset</b></div></div>' +
    '<div class="step"><div class="num">2</div><div>Enter <b>name</b>, <b>description</b>, <b>cost</b> (enter 0 if donated in-kind), <b>funding source</b>, and <b>location</b></div></div>' +
    '<div class="tip">\ud83d\udca1 <b>Funding Source</b>: Unrestricted (general funds) \u00b7 Restricted (a designated gift, e.g. someone sponsors a specific instrument) \u00b7 Donated In-Kind (an item given directly \u2014 never passed through the church bank account, so it never appears in Weekly Offerings or Expenses).</div>' +
    '<div class="tip">\ud83d\udca1 Update <b>Status</b> (In Use / Damaged / Disposed / Donated Away) directly in the Assets sheet as equipment changes over time.</div>' +

    '<h2>\ud83d\uddc3\ufe0f Start New Year</h2>' +
    '<div class="step"><div class="num">1</div><div>Close the final month with <b>\ud83d\udcca Close Month Balance</b></div></div>' +
    '<div class="step"><div class="num">2</div><div>Click <b>\ud83d\udcc5 Start New Year</b> and enter the suggested next year</div></div>' +
    '<div class="step"><div class="num">3</div><div>Current sheet is archived as <b>Weekly Offerings 2026</b>; new blank sheet opens with carry-forward balance</div></div>' +
    '<div class="tip">\ud83d\udca1 Member pages advance to the current calendar year automatically. Member Contributions currently supports through 2050.</div>' +

    '<h2>\ud83c\udfe0 Share a Template</h2>' +
    '<div class="step"><div class="num">1</div><div>Use <b>File \u2192 Make a copy</b> first. Never share the live financial workbook.</div></div>' +
    '<div class="step"><div class="num">2</div><div>In the copied file, click <b>\u26ea Church Finance \u2192 \ud83e\uddf9 Prepare This Copy for Sharing</b>. It removes all records, members, emails, links, and saved web-app settings.</div></div>' +
    '<div class="step"><div class="num">3</div><div>Replace <b>Kildare Christian Fellowship</b> and <b>KCF</b> in both script files with the new church name and short name.</div></div>' +
    '<div class="step"><div class="num">4</div><div>Share only the cleaned copy. The new church deploys its own web app, sets its own URL, and generates new links after adding its members.</div></div>' +
    '<div class="warn">\u26a0\ufe0f A copy starts with separate spreadsheet history. Do not give anyone access to the original workbook, where the live data and previous versions remain.</div>' +

    '<div class="footer">Kildare Christian Fellowship \u2014 Financial Workbook</div>' +
    '</body></html>';
}
