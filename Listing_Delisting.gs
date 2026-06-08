// ╔══════════════════════════════════════════════════════════════════╗
// ║  FILIGREE INVENTORY MANAGER — PRODUCTION                        ║
// ║                                                                  ║
// ║  Only products in 📦 Posted Products are monitored.             ║
// ║  📦 Posted Products is YOUR sheet — never touched by setup.     ║
// ║                                                                  ║
// ║  Run  ▶ runSetup  once. Then use 🏠 Inventory menu.            ║
// ╚══════════════════════════════════════════════════════════════════╝


// ================================================================
//  CONFIG
// ================================================================

var WEBHOOK_URL = "https://hook.us2.make.com/qi229ejifqv7joaqg1n63zziw630jc3b";
var UPLOAD_WEBHOOK_URL = "https://hook.us2.make.com/n8nwu5ljer8l3jmsxkyvdvqiskbir2ae";
var WEBHOOK_TIMEOUT = 30;  // seconds

var SN = {
  DASHBOARD:      "📊 Dashboard",
  DELIST_QUEUE:   "⬇️ Delist Queue",
  DELIST_HISTORY: "📕 Delist History",
  RELIST_QUEUE:   "⬆️ Relist Queue",
  RELIST_HISTORY: "📗 Relist History",
  POSTED:         "📦 Posted Products"
};

var SUPPLIERS = {
  "Kalalou Raw": {
    HEADER_ROW: 1, DATA_START_ROW: 2,
    INVENTORY_COL: "Available", SKU_COL: "Name",
    NAME_COL: "Display Name", UPC_COL: "UPC",
    SUPPLIER: "Kalalou"
  },
  "Melrose Raw": {
    HEADER_ROW: 1, DATA_START_ROW: 2,
    INVENTORY_COL: "New Avail", SKU_COL: "Vendor SKU",
    NAME_COL: "Product Description", UPC_COL: "UPC Code",
    SUPPLIER: "Melrose"
  }
};

var DELIST_Q_HEADERS   = ["Supplier", "SKU", "Product Name", "UPC", "Product Title", "Found On"];
var DELIST_H_HEADERS   = ["Supplier", "SKU", "Product Name", "UPC", "Product Title", "Sent On"];
var RELIST_Q_HEADERS   = ["Supplier", "SKU", "Product Name", "UPC", "Product Title", "Current Stock", "Found On"];
var RELIST_H_HEADERS   = ["Supplier", "SKU", "Product Name", "UPC", "Product Title", "Sent On"];


// ================================================================
//  SETUP
// ================================================================

function runSetup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var errors = [];

    // 1. Verify supplier sheets and their columns
    var supplierNames = Object.keys(SUPPLIERS);
    for (var i = 0; i < supplierNames.length; i++) {
      var sName = supplierNames[i];
      var cfg   = SUPPLIERS[sName];
      var sSheet = ss.getSheetByName(sName);
      if (!sSheet) { errors.push("Sheet '" + sName + "' not found."); continue; }
      if (sSheet.getLastRow() < 1 || sSheet.getLastColumn() < 1) {
        errors.push("Sheet '" + sName + "' is empty."); continue;
      }
      var h = _safeHeaders(sSheet, cfg.HEADER_ROW);
      if (h.indexOf(cfg.INVENTORY_COL) === -1) errors.push(sName + ": column '" + cfg.INVENTORY_COL + "' not found.\nFirst 10 headers: " + h.slice(0, 10).join(", "));
      if (h.indexOf(cfg.SKU_COL) === -1)       errors.push(sName + ": column '" + cfg.SKU_COL + "' not found.\nFirst 10 headers: " + h.slice(0, 10).join(", "));
    }

    // 2. Verify Posted Products exists with correct columns
    var posted = ss.getSheetByName(SN.POSTED);
    if (!posted) {
      errors.push("Sheet '" + SN.POSTED + "' not found.\nCreate it with columns: SKU, Product Title, Product Description, Posted On");
    } else if (posted.getLastRow() < 1 || posted.getLastColumn() < 1) {
      errors.push("Sheet '" + SN.POSTED + "' has no headers.\nAdd columns: SKU, Product Title, Product Description, Posted On");
    } else {
      var ph = _safeHeaders(posted, 1);
      if (ph.indexOf("SKU") === -1) errors.push("'" + SN.POSTED + "' is missing 'SKU' column.");
      if (ph.indexOf("Product Title") === -1) errors.push("'" + SN.POSTED + "' is missing 'Product Title' column.");
    }

    if (errors.length > 0) {
      ui.alert("Setup Failed\n\n" + errors.join("\n\n"));
      return;
    }

    // 3. Create operational sheets (never touches Posted Products)
    _ensureSheet(SN.DELIST_HISTORY, DELIST_H_HEADERS, "#B71C1C", true);
    _ensureSheet(SN.RELIST_HISTORY, RELIST_H_HEADERS, "#1B5E20", true);
    _ensureSheet(SN.DELIST_QUEUE,   DELIST_Q_HEADERS,  "#C62828", false);
    _ensureSheet(SN.RELIST_QUEUE,   RELIST_Q_HEADERS,  "#2E7D32", false);
    _buildDashboard();

    SpreadsheetApp.flush();

    // 4. Verify webhook
    var webhookOk = _post({action: "test", timestamp: new Date().toISOString()});

    // 5. Refresh and navigate
    _safeRefreshDash();
    var dash = ss.getSheetByName(SN.DASHBOARD);
    if (dash) ss.setActiveSheet(dash);

    var postedCount = Math.max(0, (posted.getLastRow() || 1) - 1);

    ui.alert(
      webhookOk
      ? "Setup Complete\n\n" +
        "Products being monitored: " + postedCount + "\n\n" +
        "• Operational sheets created\n" +
        "• Webhook connected\n\n" +
        "Refresh your browser to see the\n🏠 Inventory menu in the toolbar."
      : "Sheets created but webhook failed.\n\nCheck the WEBHOOK_URL in the script."
    );

  } catch (e) {
    ui.alert("Setup Error\n\n" + e.message);
    Logger.log("runSetup error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  SCAN FOR OUT OF STOCK
// ================================================================

function scanDelist() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Load posted products — this defines our scope
    var posted = _safeLoadPosted();
    var postedSkus = Object.keys(posted);
    if (postedSkus.length === 0) {
      ui.alert("📦 Posted Products is empty.\n\nNo products to monitor.\nUpload products to Shopify first.");
      return;
    }

    // Build lookup set
    var postedSet = {};
    for (var p = 0; p < postedSkus.length; p++) postedSet[postedSkus[p]] = true;

    // Load already-delisted
    var delisted = _safeLoadDelisted();

    // Prepare queue
    var queue = _ensureSheet(SN.DELIST_QUEUE, DELIST_Q_HEADERS, "#C62828", false);
    _safeClearQueue(queue);

    var results = [];
    var now = new Date().toLocaleString();
    var supplierNames = Object.keys(SUPPLIERS);

    for (var s = 0; s < supplierNames.length; s++) {
      var sheetName = supplierNames[s];
      var cfg = SUPPLIERS[sheetName];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < cfg.DATA_START_ROW || lastCol < 1) continue;

      var h = _safeHeaders(sheet, cfg.HEADER_ROW);
      var invIdx  = h.indexOf(cfg.INVENTORY_COL);
      var skuIdx  = h.indexOf(cfg.SKU_COL);
      var nameIdx = h.indexOf(cfg.NAME_COL);
      var upcIdx  = h.indexOf(cfg.UPC_COL);
      if (invIdx === -1 || skuIdx === -1) continue;

      var data = sheet.getRange(cfg.DATA_START_ROW, 1, lastRow - cfg.DATA_START_ROW + 1, lastCol).getValues();

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var sku = _str(row[skuIdx]);
        if (!sku) continue;
        if (!postedSet[sku]) continue;      // Not on our store — skip
        if (delisted[sku]) continue;         // Already delisted — skip

        if (_inv(row[invIdx]) === 0) {
          results.push([
            cfg.SUPPLIER,
            sku,
            nameIdx !== -1 ? _str(row[nameIdx]) : "",
            upcIdx  !== -1 ? _str(row[upcIdx])  : "",
            posted[sku] || "",
            now
          ]);
        }
      }
    }

    // Write results
    if (results.length > 0) {
      queue.getRange(2, 1, results.length, DELIST_Q_HEADERS.length).setValues(results);
    }
    SpreadsheetApp.flush();

    _safeRefreshDash();
    ss.setActiveSheet(queue);

    ui.alert(
      "Scan Complete\n\n" +
      "Monitored: " + postedSkus.length + " posted products\n" +
      "Out of stock: " + results.length + "\n\n" +
      (results.length > 0
        ? "Review them in ⬇️ Delist Queue.\nThen → 🏠 Inventory → 📤 Submit Delist."
        : "All posted products are in stock.")
    );

  } catch (e) {
    ui.alert("Scan Error\n\n" + e.message);
    Logger.log("scanDelist error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  SUBMIT DELIST
// ================================================================

function submitDelist() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var queue = ss.getSheetByName(SN.DELIST_QUEUE);
    if (!queue) { ui.alert("Run Setup first."); return; }

    var count = _safeQueueCount(queue);
    if (count === 0) { ui.alert("⬇️ Delist Queue is empty.\n\nRun Scan for Out of Stock first."); return; }

    var lastRow = queue.getLastRow();
    var data = queue.getRange(2, 1, lastRow - 1, DELIST_Q_HEADERS.length).getValues();

    var products = [];
    var logRows  = [];
    var now = new Date().toLocaleString();

    for (var i = 0; i < data.length; i++) {
      var sku = _str(data[i][1]);
      if (!sku) continue;
      var obj = {
        supplier:     _str(data[i][0]),
        sku:          sku,
        productName:  _str(data[i][2]),
        upc:          _str(data[i][3]),
        productTitle: _str(data[i][4])
      };
      products.push(obj);
      logRows.push([obj.supplier, sku, obj.productName, obj.upc, obj.productTitle, now]);
    }

    if (products.length === 0) { ui.alert("No valid products in queue."); return; }

    var confirm = ui.alert(
      "Submit Delist?",
      products.length + " products will be set to Draft on Shopify.\nThey'll be moved to 📕 Delist History.\n\nProceed?",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    var ok = _post({
      action:       "delist",
      productCount: products.length,
      products:     products,
      timestamp:    new Date().toISOString()
    });

    if (!ok) { ui.alert("Failed to send.\n\nQueue unchanged.\nCheck your internet connection."); return; }

    // Success — log and clear
    var hist = _ensureSheet(SN.DELIST_HISTORY, DELIST_H_HEADERS, "#B71C1C", true);
    var lr = hist.getLastRow();
    hist.getRange(lr + 1, 1, logRows.length, DELIST_H_HEADERS.length).setValues(logRows);
    _safeClearQueue(queue);
    SpreadsheetApp.flush();

    _safeRefreshDash();
    var dash = ss.getSheetByName(SN.DASHBOARD);
    if (dash) ss.setActiveSheet(dash);

    ui.alert("Delist Submitted\n\n" + products.length + " products sent.\nMoved to 📕 Delist History.");

  } catch (e) {
    ui.alert("Submit Error\n\n" + e.message);
    Logger.log("submitDelist error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  SCAN FOR RESTOCKED
// ================================================================

function scanRelist() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var delisted = _safeLoadDelisted();
    var posted   = _safeLoadPosted();
    var skuList  = Object.keys(delisted);

    if (skuList.length === 0) {
      ui.alert("No delisted products to check.\n\nNothing has been delisted yet.");
      return;
    }

    var queue = _ensureSheet(SN.RELIST_QUEUE, RELIST_Q_HEADERS, "#2E7D32", false);
    _safeClearQueue(queue);

    var results = [];
    var now = new Date().toLocaleString();
    var supplierNames = Object.keys(SUPPLIERS);

    for (var s = 0; s < supplierNames.length; s++) {
      var sheetName = supplierNames[s];
      var cfg = SUPPLIERS[sheetName];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < cfg.DATA_START_ROW || lastCol < 1) continue;

      var h = _safeHeaders(sheet, cfg.HEADER_ROW);
      var invIdx  = h.indexOf(cfg.INVENTORY_COL);
      var skuIdx  = h.indexOf(cfg.SKU_COL);
      var nameIdx = h.indexOf(cfg.NAME_COL);
      var upcIdx  = h.indexOf(cfg.UPC_COL);
      if (invIdx === -1 || skuIdx === -1) continue;

      var data = sheet.getRange(cfg.DATA_START_ROW, 1, lastRow - cfg.DATA_START_ROW + 1, lastCol).getValues();

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var sku = _str(row[skuIdx]);
        if (!sku) continue;
        if (!delisted[sku]) continue;   // Only check currently delisted

        var stock = _inv(row[invIdx]);
        if (stock > 0) {
          results.push([
            cfg.SUPPLIER,
            sku,
            nameIdx !== -1 ? _str(row[nameIdx]) : "",
            upcIdx  !== -1 ? _str(row[upcIdx])  : "",
            posted[sku] || "",
            stock,
            now
          ]);
        }
      }
    }

    if (results.length > 0) {
      queue.getRange(2, 1, results.length, RELIST_Q_HEADERS.length).setValues(results);
    }
    SpreadsheetApp.flush();

    _safeRefreshDash();
    ss.setActiveSheet(queue);

    ui.alert(
      "Restock Scan Complete\n\n" +
      "Checked: " + skuList.length + " delisted products\n" +
      "Back in stock: " + results.length + "\n\n" +
      (results.length > 0
        ? "Review them in ⬆️ Relist Queue.\nThen → 🏠 Inventory → 📤 Submit Relist."
        : "None restocked yet.")
    );

  } catch (e) {
    ui.alert("Scan Error\n\n" + e.message);
    Logger.log("scanRelist error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  SUBMIT RELIST
// ================================================================

function submitRelist() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var queue = ss.getSheetByName(SN.RELIST_QUEUE);
    if (!queue) { ui.alert("Run Setup first."); return; }

    var count = _safeQueueCount(queue);
    if (count === 0) { ui.alert("⬆️ Relist Queue is empty.\n\nRun Scan for Restocked first."); return; }

    var lastRow = queue.getLastRow();
    var data = queue.getRange(2, 1, lastRow - 1, RELIST_Q_HEADERS.length).getValues();

    var products = [];
    var logRows  = [];
    var now = new Date().toLocaleString();

    for (var i = 0; i < data.length; i++) {
      var sku = _str(data[i][1]);
      if (!sku) continue;
      var obj = {
        supplier:     _str(data[i][0]),
        sku:          sku,
        productName:  _str(data[i][2]),
        upc:          _str(data[i][3]),
        productTitle: _str(data[i][4]),
        currentStock: data[i][5]
      };
      products.push(obj);
      logRows.push([obj.supplier, sku, obj.productName, obj.upc, obj.productTitle, now]);
    }

    if (products.length === 0) { ui.alert("No valid products in queue."); return; }

    var confirm = ui.alert(
      "Submit Relist?",
      products.length + " products will be reactivated on Shopify.\nThey'll be moved to 📗 Relist History.\n\nProceed?",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    var ok = _post({
      action:       "relist",
      productCount: products.length,
      products:     products,
      timestamp:    new Date().toISOString()
    });

    if (!ok) { ui.alert("Failed to send.\n\nQueue unchanged.\nCheck your internet connection."); return; }

    var hist = _ensureSheet(SN.RELIST_HISTORY, RELIST_H_HEADERS, "#1B5E20", true);
    var lr = hist.getLastRow();
    hist.getRange(lr + 1, 1, logRows.length, RELIST_H_HEADERS.length).setValues(logRows);
    _safeClearQueue(queue);
    SpreadsheetApp.flush();

    _safeRefreshDash();
    var dash = ss.getSheetByName(SN.DASHBOARD);
    if (dash) ss.setActiveSheet(dash);

    ui.alert("Relist Submitted\n\n" + products.length + " products sent.\nMoved to 📗 Relist History.");

  } catch (e) {
    ui.alert("Submit Error\n\n" + e.message);
    Logger.log("submitRelist error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  DATA LOADERS (safe — never throw)
// ================================================================

/** Returns { sku: true } for currently delisted SKUs using count logic */
function _safeLoadDelisted() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dc = {};
    var rc = {};

    var dh = ss.getSheetByName(SN.DELIST_HISTORY);
    if (dh && dh.getLastRow() > 1 && dh.getLastColumn() >= 2) {
      var dd = dh.getRange(2, 2, dh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < dd.length; i++) {
        var s = _str(dd[i][0]);
        if (s) dc[s] = (dc[s] || 0) + 1;
      }
    }

    var rh = ss.getSheetByName(SN.RELIST_HISTORY);
    if (rh && rh.getLastRow() > 1 && rh.getLastColumn() >= 2) {
      var rd = rh.getRange(2, 2, rh.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < rd.length; j++) {
        var r = _str(rd[j][0]);
        if (r) rc[r] = (rc[r] || 0) + 1;
      }
    }

    var map = {};
    var all = Object.keys(dc);
    for (var k = 0; k < all.length; k++) {
      if ((dc[all[k]] || 0) > (rc[all[k]] || 0)) map[all[k]] = true;
    }
    return map;
  } catch (e) {
    Logger.log("_safeLoadDelisted error: " + e.message);
    return {};
  }
}

/** Returns { sku: "Product Title" } from Posted Products */
function _safeLoadPosted() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SN.POSTED);
    if (!sheet) return {};
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    if (lr < 2 || lc < 1) return {};

    var h = _safeHeaders(sheet, 1);
    var skuIdx   = h.indexOf("SKU");
    var titleIdx = h.indexOf("Product Title");
    if (skuIdx === -1) return {};

    var data = sheet.getRange(2, 1, lr - 1, lc).getValues();
    var map = {};
    for (var i = 0; i < data.length; i++) {
      var sku = _str(data[i][skuIdx]);
      if (sku) map[sku] = titleIdx !== -1 ? _str(data[i][titleIdx]) : "";
    }
    return map;
  } catch (e) {
    Logger.log("_safeLoadPosted error: " + e.message);
    return {};
  }
}


// ================================================================
//  DASHBOARD
// ================================================================

function _buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(SN.DASHBOARD);
  if (old) ss.deleteSheet(old);

  var d = ss.insertSheet(SN.DASHBOARD, 0);
  d.setColumnWidth(1, 20);  d.setColumnWidth(2, 260);
  d.setColumnWidth(3, 160); d.setColumnWidth(4, 20);
  d.setColumnWidth(5, 260); d.setColumnWidth(6, 160);
  d.setColumnWidth(7, 20);

  // Title
  d.getRange("B1:F1").merge().setValue("FILIGREE INVENTORY MANAGER")
    .setFontSize(18).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground("#1a1a2e").setHorizontalAlignment("center").setVerticalAlignment("middle");
  d.setRowHeight(1, 55);
  d.getRange("B2:F2").merge().setValue("Only products in 📦 Posted Products are monitored")
    .setFontSize(10).setFontColor("#888888").setFontStyle("italic").setHorizontalAlignment("center");
  d.setRowHeight(3, 8);

  // Delist section
  d.getRange("B4:C4").merge().setValue("⬇️  DELIST")
    .setFontSize(12).setFontWeight("bold").setBackground("#FFEBEE").setFontColor("#B71C1C");
  var dl = ["Pending review:", "Total delisted:", "Last scan:", "Last submitted:"];
  for (var i = 0; i < dl.length; i++) {
    d.getRange("B" + (5 + i)).setValue(dl[i]).setFontWeight("bold").setFontSize(10);
    d.getRange("C" + (5 + i)).setValue("—").setFontSize(10);
  }
  d.getRange("C5").setFontSize(18).setFontWeight("bold").setFontColor("#D32F2F");
  d.getRange("C6").setFontSize(18).setFontWeight("bold").setFontColor("#B71C1C");

  // Relist section
  d.getRange("E4:F4").merge().setValue("⬆️  RELIST")
    .setFontSize(12).setFontWeight("bold").setBackground("#E8F5E9").setFontColor("#1B5E20");
  var rl = ["Pending review:", "Total relisted:", "Last scan:", "Last submitted:"];
  for (var j = 0; j < rl.length; j++) {
    d.getRange("E" + (5 + j)).setValue(rl[j]).setFontWeight("bold").setFontSize(10);
    d.getRange("F" + (5 + j)).setValue("—").setFontSize(10);
  }
  d.getRange("F5").setFontSize(18).setFontWeight("bold").setFontColor("#2E7D32");
  d.getRange("F6").setFontSize(18).setFontWeight("bold").setFontColor("#1B5E20");

  d.setRowHeight(9, 8);

  // Overview
  d.getRange("B10:F10").merge().setValue("📊  OVERVIEW")
    .setFontSize(12).setFontWeight("bold").setBackground("#E3F2FD").setFontColor("#0D47A1");
  d.getRange("B11").setValue("Products on store:").setFontWeight("bold").setFontSize(10);
  d.getRange("C11").setValue("—").setFontSize(14).setFontWeight("bold").setFontColor("#4527A0");
  d.getRange("B12").setValue("Currently delisted:").setFontWeight("bold").setFontSize(10);
  d.getRange("C12").setValue("—").setFontSize(14).setFontWeight("bold").setFontColor("#D32F2F");
  d.getRange("E11").setValue("Kalalou Raw products:").setFontWeight("bold").setFontSize(10);
  d.getRange("F11").setValue("—").setFontSize(10);
  d.getRange("E12").setValue("Melrose Raw products:").setFontWeight("bold").setFontSize(10);
  d.getRange("F12").setValue("—").setFontSize(10);

  d.setRowHeight(13, 8);

  // Workflow
  d.getRange("B14:F14").merge().setValue("HOW IT WORKS")
    .setFontSize(12).setFontWeight("bold").setBackground("#FFF3E0").setFontColor("#E65100");
  var steps = [
    "DELIST:  Scan for Out of Stock  →  Review Queue  →  Submit Delist",
    "RELIST:  Scan for Restocked  →  Review Queue  →  Submit Relist",
    "",
    "Only SKUs in 📦 Posted Products are checked.",
    "History sheets prevent duplicates — processed items are never sent again."
  ];
  for (var k = 0; k < steps.length; k++) {
    d.getRange("B" + (15 + k) + ":F" + (15 + k)).merge().setValue(steps[k]).setFontSize(10);
    if (k < 2) d.getRange("B" + (15 + k)).setFontWeight("bold");
  }

  d.setHiddenGridlines(true);
  d.setFrozenRows(2);
  return d;
}

/** Refreshes dashboard — wrapped in try-catch so it never breaks anything */
function _safeRefreshDash() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var d = ss.getSheetByName(SN.DASHBOARD);
    if (!d) return;

    var dq = ss.getSheetByName(SN.DELIST_QUEUE);
    var dqCount = dq ? _safeQueueCount(dq) : 0;
    d.getRange("C5").setValue(dqCount);
    d.getRange("C5").setFontColor(dqCount > 0 ? "#D32F2F" : "#2E7D32")
                     .setBackground(dqCount > 0 ? "#FFEBEE" : "#E8F5E9");

    var dh = ss.getSheetByName(SN.DELIST_HISTORY);
    d.getRange("C6").setValue(dh && dh.getLastRow() > 1 ? dh.getLastRow() - 1 : 0);

    var rq = ss.getSheetByName(SN.RELIST_QUEUE);
    var rqCount = rq ? _safeQueueCount(rq) : 0;
    d.getRange("F5").setValue(rqCount);
    d.getRange("F5").setFontColor(rqCount > 0 ? "#2E7D32" : "#666666")
                     .setBackground(rqCount > 0 ? "#E8F5E9" : "#FFFFFF");

    var rh = ss.getSheetByName(SN.RELIST_HISTORY);
    d.getRange("F6").setValue(rh && rh.getLastRow() > 1 ? rh.getLastRow() - 1 : 0);

    var pp = ss.getSheetByName(SN.POSTED);
    d.getRange("C11").setValue(pp && pp.getLastRow() > 1 ? pp.getLastRow() - 1 : 0);

    var del = _safeLoadDelisted();
    d.getRange("C12").setValue(Object.keys(del).length);

    var kal = ss.getSheetByName("Kalalou Raw");
    d.getRange("F11").setValue(kal && kal.getLastRow() > 1 ? kal.getLastRow() - 1 : 0);
    var mel = ss.getSheetByName("Melrose Raw");
    d.getRange("F12").setValue(mel && mel.getLastRow() > 1 ? mel.getLastRow() - 1 : 0);

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("_safeRefreshDash error: " + e.message);
  }
}


// ================================================================
//  SHEET BUILDER (safe — never duplicates)
// ================================================================

function _ensureSheet(name, headers, color, moveToEnd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(name);
  if (existing) return existing;

  var sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground(color).setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  for (var c = 0; c < headers.length; c++) {
    var w = 140;
    var label = headers[c];
    if (label === "SKU")           w = 180;
    if (label === "Product Name")  w = 280;
    if (label === "Product Title") w = 300;
    if (label === "Supplier")      w = 100;
    if (label === "UPC")           w = 140;
    if (label === "Found On" || label === "Sent On") w = 170;
    if (label === "Current Stock") w = 110;
    sheet.setColumnWidth(c + 1, w);
  }

  if (moveToEnd) {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  return sheet;
}


// ================================================================
//  SAFE HELPERS (never throw)
// ================================================================

function _safeHeaders(sheet, row) {
  try {
    var lc = sheet.getLastColumn();
    if (lc < 1) return [];
    var raw = sheet.getRange(row, 1, 1, lc).getValues()[0];
    var out = [];
    for (var i = 0; i < raw.length; i++) out.push(_str(raw[i]));
    return out;
  } catch (e) { return []; }
}

function _safeClearQueue(queue) {
  try {
    var lr = queue.getLastRow();
    var lc = queue.getLastColumn();
    if (lr > 1 && lc > 0) {
      queue.getRange(2, 1, lr - 1, lc).clearContent().clearFormat();
    }
  } catch (e) {
    Logger.log("_safeClearQueue error: " + e.message);
  }
}

function _safeQueueCount(queue) {
  try {
    var lr = queue.getLastRow();
    if (lr < 2) return 0;
    var lc = queue.getLastColumn();
    if (lc < 2) return 0;
    var skus = queue.getRange(2, 2, lr - 1, 1).getValues();
    var count = 0;
    for (var i = 0; i < skus.length; i++) {
      if (_str(skus[i][0]) !== "") count++;
    }
    return count;
  } catch (e) { return 0; }
}

/** Converts any value to trimmed string. Never returns null/undefined. */
function _str(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

/** Parses inventory value. Returns 0 for out-of-stock, -1 for blank/invalid, number otherwise. */
function _inv(val) {
  if (val === null || val === undefined || val === "") return -1;
  var s = _str(val).toLowerCase();
  if (s === "0" || s === "out" || s === "oos" || s === "out of stock" || s === "sold out") return 0;
  var n = Number(s);
  return isNaN(n) ? -1 : n;
}

/** Sends webhook with timeout. Returns true/false. Never throws. */
function _post(payload) {
  try {
    var payloadStr = JSON.stringify(payload);

    // Safety: check payload isn't absurdly large (>5MB)
    if (payloadStr.length > 5000000) {
      Logger.log("Payload too large: " + payloadStr.length + " bytes");
      return false;
    }

    var resp = UrlFetchApp.fetch(WEBHOOK_URL, {
      method:             "post",
      contentType:        "application/json",
      payload:            payloadStr,
      muteHttpExceptions: true,
      followRedirects:    true,
      validateHttpsCertificates: true
    });

    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) return true;

    Logger.log("Webhook HTTP " + code + ": " + resp.getContentText().substring(0, 500));
    return false;
  } catch (e) {
    Logger.log("Webhook error: " + e.message);
    return false;
  }
}


// ================================================================
//  SEND TOP 15 FOR UPLOAD — skips already-posted SKUs
// ================================================================

function sendTopProducts() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName("Top 15");
    if (!sheet) { ui.alert("'Top 15' sheet not found.\n\nRun scoring first."); return; }
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) { ui.alert("'Top 15' sheet is empty.\n\nRun scoring first."); return; }

    // Load already-posted SKUs from Posted Products
    var alreadyPosted = {};
    var postedSheet = ss.getSheetByName(SN.POSTED);
    if (postedSheet && postedSheet.getLastRow() > 1 && postedSheet.getLastColumn() >= 1) {
      var ph = postedSheet.getRange(1, 1, 1, postedSheet.getLastColumn()).getValues()[0];
      var psi = -1;
      for (var pi = 0; pi < ph.length; pi++) {
        if (_str(ph[pi]) === "SKU") { psi = pi; break; }
      }
      if (psi !== -1) {
        var pd = postedSheet.getRange(2, psi + 1, postedSheet.getLastRow() - 1, 1).getValues();
        for (var pk = 0; pk < pd.length; pk++) {
          var ps = _str(pd[pk][0]);
          if (ps) alreadyPosted[ps] = true;
        }
      }
    }

    // Read Top 15 headers and data
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Build column index from headers
    var idx = {};
    for (var h = 0; h < headers.length; h++) {
      idx[_str(headers[h])] = h;
    }

    var products = [];
    var skippedCount = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var sku = idx["SKU"] !== undefined ? _str(row[idx["SKU"]]) : "";
      if (!sku) continue;

      // Skip if already posted
      if (alreadyPosted[sku]) { skippedCount++; continue; }

      products.push({
        supplier:    idx["Supplier"]       !== undefined ? _str(row[idx["Supplier"]])       : "",
        sku:         sku,
        productName: idx["Product Name"]   !== undefined ? _str(row[idx["Product Name"]])   : "",
        cost:        idx["Cost"]           !== undefined ? row[idx["Cost"]]                 : "",
        msrp:        idx["MSRP"]           !== undefined ? row[idx["MSRP"]]                 : "",
        marginPct:   idx["Margin %"]       !== undefined ? row[idx["Margin %"]]             : "",
        absMargin:   idx["Abs $ Margin"]   !== undefined ? row[idx["Abs $ Margin"]]         : "",
        stock:       idx["Stock"]          !== undefined ? row[idx["Stock"]]                : "",
        category:    idx["Category"]       !== undefined ? _str(row[idx["Category"]])       : "",
        theme:       idx["Theme"]          !== undefined ? _str(row[idx["Theme"]])          : "",
        material:    idx["Material"]       !== undefined ? _str(row[idx["Material"]])       : "",
        color:       idx["Color"]          !== undefined ? _str(row[idx["Color"]])          : "",
        setSize:     idx["Set Size"]       !== undefined ? _str(row[idx["Set Size"]])       : "",
        totalScore:  idx["TOTAL"]          !== undefined ? row[idx["TOTAL"]]                : "",
        tier:        idx["TIER"]           !== undefined ? _str(row[idx["TIER"]])           : "",
        pool:        idx["Pool"]           !== undefined ? _str(row[idx["Pool"]])           : ""
      });
    }

    // All already posted
    if (products.length === 0 && skippedCount > 0) {
      ui.alert("All " + skippedCount + " products are already posted.\n\nNothing new to upload.");
      return;
    }
    if (products.length === 0) {
      ui.alert("No valid products found in Top 15.");
      return;
    }

    // Confirm
    var msg = products.length + " new products will be sent for upload.";
    if (skippedCount > 0) msg += "\n" + skippedCount + " already-posted products skipped.";

    var confirm = ui.alert("Send for Upload?", msg + "\n\nProceed?", ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;

    var ok = _postUpload({
      action:       "upload_products",
      productCount: products.length,
      products:     products,
      timestamp:    new Date().toISOString()
    });

    if (ok) {
      ui.alert(
        "Sent for Upload\n\n" +
        products.length + " new products → Make.com" +
        (skippedCount > 0 ? "\n" + skippedCount + " already-posted skipped" : "")
      );
    } else {
      ui.alert("Failed to send.\n\nCheck your internet connection.");
    }

  } catch (e) {
    ui.alert("Upload Error\n\n" + e.message);
    Logger.log("sendTopProducts error: " + e.message + "\n" + e.stack);
  }
}


// ================================================================
//  UPLOAD WEBHOOK HELPER
// ================================================================

function _postUpload(payload) {
  try {
    var payloadStr = JSON.stringify(payload);
    if (payloadStr.length > 5000000) {
      Logger.log("Upload payload too large: " + payloadStr.length + " bytes");
      return false;
    }
    var resp = UrlFetchApp.fetch(UPLOAD_WEBHOOK_URL, {
      method:             "post",
      contentType:        "application/json",
      payload:            payloadStr,
      muteHttpExceptions: true,
      followRedirects:    true,
      validateHttpsCertificates: true
    });
    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) return true;
    Logger.log("Upload webhook HTTP " + code + ": " + resp.getContentText().substring(0, 500));
    return false;
  } catch (e) {
    Logger.log("Upload webhook error: " + e.message);
    return false;
  }
}
