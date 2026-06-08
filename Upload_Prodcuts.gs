// ╔══════════════════════════════════════════════════════════════════╗
// ║  FILIGREE PRODUCT UPLOADER                                       ║
// ║                                                                  ║
// ║  Reads the "Top 15" sheet and sends all products                ║
// ║  to Make.com in ONE webhook call.                               ║
// ║                                                                  ║
// ║  Run  ▶ setupUploader   once.                                   ║
// ║  Then use the 🚀 Product Uploader menu.                        ║
// ╚══════════════════════════════════════════════════════════════════╝


// ================================================================
//  CONFIG
// ================================================================

var UPLOAD_WEBHOOK_URL = "https://hook.us2.make.com/n8nwu5ljer8l3jmsxkyvdvqiskbir2ae";

var SOURCE_SHEET_NAME = "Top 15";
var HEADER_ROW        = 1;
var DATA_START_ROW    = 2;

// Column names — must match headers exactly
var COLUMNS = {
  SUPPLIER:    "Supplier",
  SKU:         "SKU",
  PRODUCT:     "Product Name",
  COST:        "Cost",
  MSRP:        "MSRP",
  MARGIN_PCT:  "Margin %",
  ABS_MARGIN:  "Abs $ Margin",
  STOCK:       "Stock",
  CATEGORY:    "Category",
  THEME:       "Theme",
  MATERIAL:    "Material",
  COLOR:       "Color",
  SET_SIZE:    "Set Size",
  ELIGIBLE:    "Eligible",
  REASON:      "Reason",
  TOTAL:       "TOTAL",
  TIER:        "TIER"
};


// ================================================================
//  MENU
// ================================================================

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu("🚀 Product Uploader")
      .addItem("▶ Setup",                    "setupUploader")
      .addSeparator()
      .addItem("🚀 Send Products to Make.com", "sendProducts")
      .addItem("🧪 Test Webhook",              "testUploadWebhook")
      .addToUi();
  } catch (e) {
    // UI not available — skip
  }
}


// ================================================================
//  SETUP
// ================================================================

function setupUploader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Verify sheet exists
  var sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sheet) {
    ui.alert("❌ Sheet '" + SOURCE_SHEET_NAME + "' not found.\n\nMake sure the tab name matches exactly.");
    return;
  }

  // Verify columns
  var headers = _getHeaders(sheet);
  var missing = [];
  var colNames = Object.keys(COLUMNS);

  for (var i = 0; i < colNames.length; i++) {
    var colName = COLUMNS[colNames[i]];
    if (headers.indexOf(colName) === -1) {
      missing.push(colName);
    }
  }

  if (missing.length > 0) {
    ui.alert(
      "❌ Missing columns in '" + SOURCE_SHEET_NAME + "':\n\n" +
      missing.join(", ") +
      "\n\nHeaders found:\n" + headers.slice(0, 20).join(", ")
    );
    return;
  }

  // Count rows
  var rowCount = Math.max(0, sheet.getLastRow() - 1);

  // Test webhook
  var webhookOk = _sendUploadWebhook({
    action: "test",
    source: "setup",
    message: "Product Uploader connected",
    timestamp: new Date().toISOString()
  });

  if (webhookOk) {
    ui.alert(
      "✔ Setup Complete!\n\n" +
      "• Sheet: '" + SOURCE_SHEET_NAME + "' found (" + rowCount + " products)\n" +
      "• All columns verified\n" +
      "• Webhook connected\n\n" +
      "Use the 🚀 Product Uploader menu\n→ Send Products to Make.com"
    );
  } else {
    ui.alert(
      "⚠ Sheet is fine, but webhook test FAILED.\n\nCheck the UPLOAD_WEBHOOK_URL."
    );
  }
}


// ================================================================
//  SEND PRODUCTS
// ================================================================

function sendProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sheet) {
    ui.alert("❌ Sheet '" + SOURCE_SHEET_NAME + "' not found.");
    return;
  }

  var headers = _getHeaders(sheet);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < DATA_START_ROW) {
    ui.alert("📋 No products in '" + SOURCE_SHEET_NAME + "'.");
    return;
  }

  // Find column indexes
  var idx = {};
  var colNames = Object.keys(COLUMNS);
  for (var c = 0; c < colNames.length; c++) {
    idx[colNames[c]] = headers.indexOf(COLUMNS[colNames[c]]);
  }

  // Read all data in ONE batch
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol).getValues();
  var products = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    // Skip empty rows (no SKU)
    var sku = idx.SKU !== -1 ? String(row[idx.SKU] || "").trim() : "";
    if (!sku) continue;

    var product = {
      supplier:    idx.SUPPLIER   !== -1 ? String(row[idx.SUPPLIER]   || "").trim() : "",
      sku:         sku,
      productName: idx.PRODUCT    !== -1 ? String(row[idx.PRODUCT]    || "").trim() : "",
      cost:        idx.COST       !== -1 ? row[idx.COST]       : "",
      msrp:        idx.MSRP       !== -1 ? row[idx.MSRP]       : "",
      marginPct:   idx.MARGIN_PCT !== -1 ? row[idx.MARGIN_PCT] : "",
      absMargin:   idx.ABS_MARGIN !== -1 ? row[idx.ABS_MARGIN] : "",
      stock:       idx.STOCK      !== -1 ? row[idx.STOCK]      : "",
      category:    idx.CATEGORY   !== -1 ? String(row[idx.CATEGORY]   || "").trim() : "",
      theme:       idx.THEME      !== -1 ? String(row[idx.THEME]      || "").trim() : "",
      material:    idx.MATERIAL   !== -1 ? String(row[idx.MATERIAL]   || "").trim() : "",
      color:       idx.COLOR      !== -1 ? String(row[idx.COLOR]      || "").trim() : "",
      setSize:     idx.SET_SIZE   !== -1 ? row[idx.SET_SIZE]   : "",
      eligible:    idx.ELIGIBLE   !== -1 ? String(row[idx.ELIGIBLE]   || "").trim() : "",
      reason:      idx.REASON     !== -1 ? String(row[idx.REASON]     || "").trim() : "",
      totalScore:  idx.TOTAL      !== -1 ? row[idx.TOTAL]      : "",
      tier:        idx.TIER       !== -1 ? String(row[idx.TIER]       || "").trim() : ""
    };

    products.push(product);
  }

  if (products.length === 0) {
    ui.alert("📋 No valid products found (all rows empty or no SKU).");
    return;
  }

  // Confirm
  var confirm = ui.alert(
    "🚀 Send Products to Make.com?",
    products.length + " products from '" + SOURCE_SHEET_NAME + "'\nwill be sent in ONE webhook call.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Send ONE webhook
  var payload = {
    action:       "upload_products",
    source:       "top_15_sheet",
    productCount: products.length,
    products:     products,
    timestamp:    new Date().toISOString()
  };

  var ok = _sendUploadWebhook(payload);

  if (ok) {
    ui.alert(
      "✔ Products Sent!\n\n" +
      products.length + " products → Make.com\n\n" +
      "Check Make.com to verify the workflow started."
    );
  } else {
    ui.alert("❌ Webhook FAILED.\n\nCheck the Execution Log for details.");
  }
}


// ================================================================
//  TEST WEBHOOK
// ================================================================

function testUploadWebhook() {
  var ok = _sendUploadWebhook({
    action:    "test",
    source:    "manual_test",
    message:   "Product Upload webhook test",
    timestamp: new Date().toISOString()
  });

  if (ok) {
    SpreadsheetApp.getUi().alert("✔ Webhook working!\nCheck Make.com.");
  } else {
    SpreadsheetApp.getUi().alert("❌ Webhook failed.\nCheck UPLOAD_WEBHOOK_URL.");
  }
}


// ================================================================
//  HELPERS
// ================================================================

function _getHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  var raw = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var result = [];
  for (var i = 0; i < raw.length; i++) {
    result.push(String(raw[i]).trim());
  }
  return result;
}

function _sendUploadWebhook(payload) {
  try {
    var options = {
      method:             "post",
      contentType:        "application/json",
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(UPLOAD_WEBHOOK_URL, options);
    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      Logger.log("✔ Upload webhook sent (HTTP " + code + ")");
      return true;
    }
    Logger.log("✖ Upload webhook HTTP " + code + ": " + resp.getContentText().substring(0, 300));
    return false;
  } catch (e) {
    Logger.log("✖ Upload webhook error: " + e.message);
    return false;
  }
}
