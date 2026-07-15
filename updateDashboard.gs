/**
 * updateDashboard.gs — ARG TV Dashboard update (v2 data model)
 * Sources: End_of_the_Day_Report_7_14_26.xlsx (daily)
 *          Collector_Goal_and_Department_Goal_7-15-2026.xlsx (MTD as of 7/15)
 * Generated + audited: 2026-07-15
 *
 * RUN: open the dashboard Sheet > Extensions > Apps Script > replace all
 * code with this file > Save > Run updateDashboard.
 *
 * v2 DATA MODEL (matches dashboard v2.0):
 *   State      — key/value (see DATA.state)
 *   Collectors — Name | MTD | Pct | Move   (goal $ NEVER written; % only)
 *   Daily      — Name | Accts | Notated | Collected  (created if missing)
 *
 * PRIVACY: per the goals file header ("do not show in dashboard"), goal
 * dollar amounts exist only in the source reports — not in this Sheet,
 * not on screen. Only % of goal is published.
 *
 * NOTE: Move is 0 for all collectors this run — the ranking basis changed
 * from placeholder dollars to true % of goal, so movement vs the old
 * ranking would be misleading. Arrows resume next update.
 */

var DATA = {
  state: {
    monthLabel:     "July 2026 — thru 7/15",
    deptCollected:  785495,
    deptPct:        26.18,
    dailyDate:      "Tuesday, July 14",
    dailyCollected: 36344.96,
    dailySub:       "647 accounts worked \u00b7 774 notated",
    flyerUrl:       "flyer.png",
    banner1Label:   "10-Day Blitz winner",
    banner1Text:    "Ginelle R. — 81.1% of monthly goal by day 10",
    banner2Label:   "Top daily collector — 7/14",
    banner2Text:    "Charles M. — $11,494 collected",
    banner3Label:   "Floor total — 7/14",
    banner3Text:    "647 accounts worked \u00b7 $36,345 collected"
  },
  collectors: [
    // [Name, MTD, Pct]  (as of 7/15; pct per goals file; Move forced 0 this run)
    ["Ginelle R.",  87479.06, 87.48],
    ["Ray L.",      28307.22, 47.18],
    ["Curt A.",     23706.04, 39.51],
    ["Charles M.", 117386.12, 35.57],
    ["Robert D.",   90549.18, 29.21],
    ["Ekrem H.",   281901.70, 27.91],
    ["Gabe T.",     38777.10, 21.54],
    ["Jamal G.",    49384.63, 21.47],
    ["Anthony P.",  27362.80, 15.20],
    ["Ramone W.",   15556.70, 11.97],
    ["Jefrey P.",   25084.43, 11.40]
  ],
  daily: [
    // [Name, AcctsWorked, Notated, CollectedToday]  (EOD 7/14)
    ["Charles M.", 104,  57, 11494.22],
    ["Ekrem H.",   130, 145,  7550.00],
    ["Jamal G.",    70,  73,  6909.25],
    ["Anthony P.",  40,  61,  2218.50],
    ["Gabe T.",     45,  51,  2200.00],
    ["Curt A.",     32,  58,  1127.99],
    ["Ginelle R.",  44,  70,  1100.00],
    ["Jefrey P.",   40,  50,  1100.00],
    ["Ramone W.",   20,  62,  1100.00],
    ["Ray L.",      50,  90,   925.00],
    ["Robert D.",   72,  57,   620.00]
  ]
};

function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stateTab = ss.getSheetByName("State");
  var collTab = ss.getSheetByName("Collectors");
  if (!stateTab || !collTab) throw new Error("Missing 'State' or 'Collectors' tab.");
  var dailyTab = ss.getSheetByName("Daily") || ss.insertSheet("Daily");

  // ---- State (cols A:B; also clears the old template legend in col D) ----
  var keys = Object.keys(DATA.state);
  var stateRows = keys.map(function (k) { return [k, DATA.state[k]]; });
  stateTab.getRange(2, 1, Math.max(stateTab.getLastRow(), stateRows.length + 1), 4).clearContent();
  stateTab.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
  stateTab.getRange(2, 1, stateRows.length, 2).setValues(stateRows);

  // ---- Collectors: Name | MTD | Pct | Move (Move = 0 this run; legend col F cleared) ----
  var collRows = DATA.collectors.map(function (r) { return [r[0], r[1], r[2], 0]; });
  collTab.getRange(2, 1, Math.max(collTab.getLastRow(), collRows.length + 1), 6).clearContent();
  collTab.getRange(1, 1, 1, 4).setValues([["Name", "MTD", "Pct", "Move"]]);
  collTab.getRange(2, 1, collRows.length, 4).setValues(collRows);

  // ---- Daily: Name | Accts | Notated | Collected ----
  var dailyRows = DATA.daily;
  var clearD = Math.max(dailyTab.getLastRow(), dailyRows.length + 1);
  dailyTab.getRange(2, 1, clearD, 4).clearContent();
  dailyTab.getRange(1, 1, 1, 4).setValues([["Name", "Accts", "Notated", "Collected"]]);
  dailyTab.getRange(2, 1, dailyRows.length, 4).setValues(dailyRows);

  Logger.log("Dashboard updated: dept $" + DATA.state.deptCollected.toLocaleString()
    + " (" + DATA.state.deptPct + "%) | daily $" + DATA.state.dailyCollected
    + " | " + collRows.length + " collectors | leader: " + DATA.collectors[0][0]);
}
