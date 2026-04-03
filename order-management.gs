/**
 * ============================================================
 *  ORDER MANAGEMENT AUTOMATION SCRIPT
 *  Developed for: Packly E-Commerce Order Processing
 *  Version: 7.0.0  |  Final — Row-by-Row Direct Mapping
 *
 *  HELPER TAB — Input columns:
 *    A: Mother Items (quantity)   B: Mother INV    C: Order Date
 *    D: Customer Name             E: Phone
 *    F: Sub Items (quantity)      G: Shop Name     H: Sub INV
 *
 *  MAIN TAB — Output columns:
 *    A: Mother INV     B: Order Date    C: Customer Name
 *    D: Phone          E: Shop Name     F: Sub INV
 *    H2: Past Last Data    (grey)
 *    H4: Present Last Data (green)
 *
 *  MATCHING LOGIC (v7 — THE FIX):
 *    Each Helper row that has BOTH a Mother INV (col B) AND
 *    a Sub INV (col H) produces exactly ONE output row.
 *    A and F columns are quantities for display only —
 *    they are NOT used for matching.
 *
 *  DUPLICATE PREVENTION:
 *    Tracker H4 stores the last processed Sub INV.
 *    On next run, all rows up to and including that Sub INV
 *    are skipped. Main tab is cleared and only NEW rows written.
 *
 *  RESET:
 *    "Reset Tracker" clears Main tab data + H2 + H4 (fresh start).
 * ============================================================
 */


// ============================================================
//  CONFIGURATION
// ============================================================
var CFG = {

  TAB: {
    MAIN  : "Main",
    HELPER: "Helper"
  },

  // Helper tab column indices (A=1, B=2 ...)
  H: {
    MOTHER_ITEMS : 1,   // A — quantity (display only, not used for matching)
    MOTHER_INV   : 2,   // B ← matching key
    ORDER_DATE   : 3,   // C
    CUSTOMER_NAME: 4,   // D
    PHONE        : 5,   // E
    SUB_ITEMS    : 6,   // F — quantity (display only, not used for matching)
    SHOP_NAME    : 7,   // G
    SUB_INV      : 8    // H ← matching key
  },

  // Main tab column indices
  M: {
    MOTHER_INV   : 1,   // A
    ORDER_DATE   : 2,   // B
    CUSTOMER_NAME: 3,   // C
    PHONE        : 4,   // D
    SHOP_NAME    : 5,   // E
    SUB_INV      : 6,   // F
    TRACKER_COL  : 8    // H
  },

  TRACKER: {
    PAST_ROW   : 2,     // H2
    PRESENT_ROW: 4      // H4
  },

  OUTPUT_COL_COUNT: 6,
  HEADER_ROW      : 1
};


// ============================================================
//  ON OPEN
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Order Management")
    .addItem("▶  Process Orders", "processOrders")
    .addItem("🗑  Clear Main Tab", "clearMainTab")
    .addItem("🔄  Reset Tracker",  "resetTracker")
    .addSeparator()
    .addItem("❓  Help",           "showHelp")
    .addToUi();
}


// ============================================================
//  UTILITY — safe string trim
// ============================================================
function _str(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim();
}


// ============================================================
//  UTILITY — format date → "dd-MMM-yyyy"
//  Uses spreadsheet timezone to avoid off-by-one date errors
// ============================================================
function _formatDate(val) {
  if (!val || _str(val) === "") return "";
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return _str(val);
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(d, tz, "dd-MMM-yyyy");
  } catch (e) {
    return _str(val);
  }
}


// ============================================================
//  UTILITY — build tracker label from a Main output row
// ============================================================
function _buildLabel(row) {
  if (!row || row.length < CFG.OUTPUT_COL_COUNT) return "N/A";
  return [
    _str(row[CFG.M.MOTHER_INV    - 1]),
    _str(row[CFG.M.ORDER_DATE    - 1]),
    _str(row[CFG.M.CUSTOMER_NAME - 1]),
    _str(row[CFG.M.PHONE         - 1]),
    _str(row[CFG.M.SHOP_NAME     - 1]),
    _str(row[CFG.M.SUB_INV       - 1])
  ].join("  |  ");
}


// ============================================================
//  UTILITY — write and style a tracker cell
// ============================================================
function _writeTracker(sheet, row, col, text, isPast) {
  var cell = sheet.getRange(row, col);
  cell.setValue(text);
  cell.setFontFamily("Arial");
  cell.setFontSize(9);
  cell.setWrap(false);
  if (isPast) {
    cell.setBackground("#f3f3f3");
    cell.setFontColor("#666666");
    cell.setFontWeight("normal");
  } else {
    cell.setBackground("#b6d7a8");
    cell.setFontColor("#1c4a03");
    cell.setFontWeight("bold");
  }
}


// ============================================================
//  UTILITY — clear Main tab data rows (keeps header row)
// ============================================================
function _clearMainDataRows(mainSheet) {
  var lastRow = mainSheet.getLastRow();
  if (lastRow > CFG.HEADER_ROW) {
    mainSheet
      .getRange(CFG.HEADER_ROW + 1, 1,
                lastRow - CFG.HEADER_ROW,
                CFG.OUTPUT_COL_COUNT)
      .clearContent();
  }
}


// ============================================================
//  CORE — parse Helper and build output rows
//
//  MATCHING RULE:
//    For every Helper data row (skip header):
//      - If BOTH Mother INV (col B) AND Sub INV (col H) exist
//        → produce one output row  [MotherINV, Date, Name,
//                                   Phone, ShopName, SubINV]
//      - Rows with only Mother INV or only Sub INV are skipped
//        (they are informational/quantity rows)
//
//  Returns: { rows: [...], skippedRows: number, error: null|string }
// ============================================================
function _parseAndBuild(helperData) {
  var outputRows  = [];
  var skippedRows = 0;

  for (var i = CFG.HEADER_ROW; i < helperData.length; i++) {
    var row    = helperData[i];
    var mInv   = _str(row[CFG.H.MOTHER_INV - 1]);
    var sInv   = _str(row[CFG.H.SUB_INV    - 1]);

    // Both must be present — otherwise skip (quantity-only row)
    if (mInv === "" || sInv === "") {
      skippedRows++;
      continue;
    }

    var shopName = _str(row[CFG.H.SHOP_NAME    - 1]);
    var date     = _formatDate(row[CFG.H.ORDER_DATE   - 1]);
    var name     = _str(row[CFG.H.CUSTOMER_NAME - 1]);
    var phone    = _str(row[CFG.H.PHONE         - 1]);

    outputRows.push([mInv, date, name, phone, shopName, sInv]);
  }

  if (outputRows.length === 0) {
    return {
      rows       : [],
      skippedRows: skippedRows,
      error      : "Helper tab এ কোনো valid row পাওয়া যায়নি।\n" +
                   "প্রতিটি row এ Mother INV (col B) এবং Sub INV (col H) " +
                   "দুটোই থাকতে হবে।"
    };
  }

  return { rows: outputRows, skippedRows: skippedRows, error: null };
}


// ============================================================
//  FILTER — return only rows after the last processed Sub INV
//
//  The tracker H4 stores the last Sub INV that was written
//  to Main tab. We find its LAST occurrence in allRows and
//  return everything after it as "new rows".
// ============================================================
function _getNewRows(allRows, lastSubINV) {
  if (!lastSubINV || lastSubINV === "") {
    return { newRows: allRows, skippedCount: 0 };
  }

  // Walk backwards — find last occurrence of lastSubINV in col F (index 5)
  var cutIndex = -1;
  for (var i = allRows.length - 1; i >= 0; i--) {
    if (_str(allRows[i][CFG.M.SUB_INV - 1]) === lastSubINV) {
      cutIndex = i;
      break;
    }
  }

  if (cutIndex >= 0) {
    return {
      newRows     : allRows.slice(cutIndex + 1),
      skippedCount: cutIndex + 1
    };
  }

  // lastSubINV not found in new Helper data → treat all as new
  return { newRows: allRows, skippedCount: 0 };
}


// ============================================================
//  MAIN — orchestrates the full process
// ============================================================
function processOrders() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Validate tabs ─────────────────────────────────────
  var helperSheet = ss.getSheetByName(CFG.TAB.HELPER);
  var mainSheet   = ss.getSheetByName(CFG.TAB.MAIN);

  if (!helperSheet) {
    ui.alert("❌ Tab পাওয়া যায়নি",
             '"' + CFG.TAB.HELPER + '" নামে কোনো Tab নেই।\n' +
             'Input tab এর নাম "' + CFG.TAB.HELPER + '" রাখুন।',
             ui.ButtonSet.OK);
    return;
  }
  if (!mainSheet) {
    ui.alert("❌ Tab পাওয়া যায়নি",
             '"' + CFG.TAB.MAIN + '" নামে কোনো Tab নেই।\n' +
             'Output tab এর নাম "' + CFG.TAB.MAIN + '" রাখুন।',
             ui.ButtonSet.OK);
    return;
  }

  // ── 2. Read Helper (single batch read) ───────────────────
  var helperData = helperSheet.getDataRange().getValues();

  if (helperData.length <= CFG.HEADER_ROW) {
    ui.alert("⚠️ কোনো Data নেই",
             "Helper tab এ কোনো data row নেই।\nআগে data paste করুন।",
             ui.ButtonSet.OK);
    return;
  }

  // ── 3. Parse and build all output rows ───────────────────
  var built = _parseAndBuild(helperData);

  if (built.error) {
    ui.alert("❌ Data Error", built.error, ui.ButtonSet.OK);
    return;
  }

  var allRows = built.rows;

  // ── 4. Read last tracked Sub INV from H4 ─────────────────
  var lastSubINV = "";
  var h4Val = _str(
    mainSheet.getRange(CFG.TRACKER.PRESENT_ROW, CFG.M.TRACKER_COL).getValue()
  );

  if (h4Val !== "") {
    // H4 format: "Present Last Data :    INV | date | name | phone | shop | SubINV"
    // Sub INV is the LAST segment after the final "|"
    var parts = h4Val.split("|");
    lastSubINV = _str(parts[parts.length - 1]);
  }

  // ── 5. Filter to new rows only ───────────────────────────
  var filtered     = _getNewRows(allRows, lastSubINV);
  var newRows      = filtered.newRows;
  var skippedCount = filtered.skippedCount;

  if (newRows.length === 0) {
    ui.alert(
      "ℹ️ নতুন কিছু নেই",
      "Helper tab এর সব rows আগেই process হয়ে গেছে।\n\n" +
      "Last tracked Sub INV: " + lastSubINV + "\n\n" +
      "নতুন data Helper এ যোগ করুন, তারপর আবার চালান।",
      ui.ButtonSet.OK
    );
    return;
  }

  // ── 6. Save snapshot of current last Main row (for H2) ───
  var pastRowData = null;
  var mainLastRow = mainSheet.getLastRow();
  if (mainLastRow > CFG.HEADER_ROW) {
    pastRowData = mainSheet
      .getRange(mainLastRow, 1, 1, CFG.OUTPUT_COL_COUNT)
      .getValues()[0];
  }

  // ── 7. Confirm dialog ────────────────────────────────────
  var confirmMsg =
    "প্রসেস করতে প্রস্তুত:\n\n" +
    "   Helper এ মোট valid rows   :  " + allRows.length   + "\n" +
    "   আগেই processed (skip)    :  " + skippedCount     + "\n" +
    "   নতুন rows লেখা হবে       :  " + newRows.length   + "\n\n" +
    "📌 Main tab সম্পূর্ণ clear হয়ে\n" +
    "   শুধু নতুন " + newRows.length + " টি row বসবে।\n\n" +
    "Last tracked Sub INV: " +
    (lastSubINV !== "" ? lastSubINV : "কোনোটি নেই — প্রথম run") +
    "\n\nচালিয়ে যেতে চান?";

  var answer = ui.alert("✅ নিশ্চিত করুন", confirmMsg, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  // ── 8. Clear Main tab data rows ──────────────────────────
  _clearMainDataRows(mainSheet);

  // ── 9. Write new rows from row 2 ─────────────────────────
  mainSheet
    .getRange(CFG.HEADER_ROW + 1, 1, newRows.length, CFG.OUTPUT_COL_COUNT)
    .setValues(newRows);

  // ── 10. Snapshot new last row for H4 ─────────────────────
  var newLastRow  = mainSheet.getLastRow();
  var presentData = mainSheet
    .getRange(newLastRow, 1, 1, CFG.OUTPUT_COL_COUNT)
    .getValues()[0];

  // ── 11. Update tracker H2 & H4 ───────────────────────────
  var tCol     = CFG.M.TRACKER_COL;
  var pastText = "Past Last Data :        " +
                 (pastRowData ? _buildLabel(pastRowData) : "N/A  — প্রথম run");
  var preText  = "Present Last Data :    " + _buildLabel(presentData);

  _writeTracker(mainSheet, CFG.TRACKER.PAST_ROW,    tCol, pastText, true);
  _writeTracker(mainSheet, CFG.TRACKER.PRESENT_ROW, tCol, preText,  false);

  // ── 12. Done ─────────────────────────────────────────────
  ui.alert(
    "✅ সফল",
    "প্রসেসিং সম্পন্ন!\n\n" +
    "   নতুন rows লেখা হয়েছে      :  " + newRows.length                 + "\n" +
    "   Main tab এ মোট rows       :  " + (newLastRow - CFG.HEADER_ROW)  + "\n\n" +
    "Tracker আপডেট →  H2 (past)  ও  H4 (present)।",
    ui.ButtonSet.OK
  );
}


// ============================================================
//  CLEAR MAIN TAB — data rows only, keeps header & tracker
// ============================================================
function clearMainTab() {
  var ui = SpreadsheetApp.getUi();

  var answer = ui.alert(
    "🗑  নিশ্চিত করুন",
    "Main tab এর সব data row মুছে যাবে।\n" +
    "Header ও Tracker (H2, H4) অক্ষত থাকবে।\n\n" +
    "নিশ্চিত?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var mainSheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(CFG.TAB.MAIN);
  if (!mainSheet) {
    ui.alert("❌ Error", '"' + CFG.TAB.MAIN + '" tab পাওয়া যায়নি।',
             ui.ButtonSet.OK);
    return;
  }

  _clearMainDataRows(mainSheet);
  ui.alert("✅ সম্পন্ন",
           "Main tab data মুছে গেছে। Tracker অক্ষত।",
           ui.ButtonSet.OK);
}


// ============================================================
//  RESET TRACKER — full reset: Main data + H2 + H4
// ============================================================
function resetTracker() {
  var ui = SpreadsheetApp.getUi();

  var answer = ui.alert(
    "🔄  সম্পূর্ণ Reset — নিশ্চিত করুন",
    "এই action:\n" +
    "  • Main tab এর সব data row মুছবে\n" +
    "  • Tracker H2 ও H4 clear করবে\n\n" +
    "পরবর্তী Process Orders Helper এর সব data কে\n" +
    "নতুন হিসেবে গণ্য করবে।\n\n" +
    "নিশ্চিত?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var mainSheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(CFG.TAB.MAIN);
  if (!mainSheet) {
    ui.alert("❌ Error", '"' + CFG.TAB.MAIN + '" tab পাওয়া যায়নি।',
             ui.ButtonSet.OK);
    return;
  }

  _clearMainDataRows(mainSheet);

  var tCol = CFG.M.TRACKER_COL;
  [CFG.TRACKER.PAST_ROW, CFG.TRACKER.PRESENT_ROW].forEach(function(row) {
    mainSheet.getRange(row, tCol)
      .clearContent()
      .setBackground(null)
      .setFontColor(null)
      .setFontWeight("normal");
  });

  ui.alert("✅ সম্পন্ন",
           "Main tab ও Tracker সম্পূর্ণ reset হয়েছে।\n" +
           "পরবর্তী process সব fresh শুরু করবে।",
           ui.ButtonSet.OK);
}


// ============================================================
//  HELP
// ============================================================
function showHelp() {
  SpreadsheetApp.getUi().alert(
    "📖  How To Use  —  v7.0.0\n\n" +

    "HELPER TAB:\n" +
    "   A: Mother Items (qty)    B: Mother INV  ← matching\n" +
    "   C: Order Date            D: Customer Name\n" +
    "   E: Phone\n" +
    "   F: Sub Items (qty)       G: Shop Name\n" +
    "   H: Sub INV               ← matching\n\n" +

    "MATCHING RULE:\n" +
    "   প্রতিটি Helper row এ Mother INV (B) এবং\n" +
    "   Sub INV (H) দুটোই থাকলে → একটি output row তৈরি হয়।\n" +
    "   A ও F column এর quantity matching এ ব্যবহার হয় না।\n\n" +

    "MAIN TAB OUTPUT:\n" +
    "   A: Mother INV    B: Order Date    C: Customer Name\n" +
    "   D: Phone         E: Shop Name     F: Sub INV\n" +
    "   H2: Past Last Data  (grey)\n" +
    "   H4: Present Last Data  (green)\n\n" +

    "DAILY WORKFLOW:\n" +
    "   1. Helper এ আগের data রেখে নতুন rows যোগ করুন\n" +
    "   2. Menu → ▶ Process Orders\n" +
    "   3. Summary দেখে YES দিন\n" +
    "   4. Main tab clear হয়ে শুধু নতুন rows বসবে\n" +
    "   5. Tracker H2 ও H4 আপডেট হবে\n\n" +

    "RESET:\n" +
    "   🗑 Clear Main Tab  — শুধু data rows মুছে\n" +
    "   🔄 Reset Tracker   — Main data + H2 + H4 সব clear\n\n" +

    "NOTE:\n" +
    "   Tracker H4 থেকে last Sub INV পড়ে duplicate prevent করে।\n" +
    "   Reset করলে পরের process সব data নতুন ধরে।",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
