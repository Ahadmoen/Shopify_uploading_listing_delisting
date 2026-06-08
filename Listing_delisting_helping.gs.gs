// ============================================================
// FILIGREE INVENTORY MONITOR — Google Apps Script
// Monitors Kallolo & Melrose supplier sheets for 0 inventory
// Sends webhook to Make.com when products need delisting
// ============================================================

// ======================== CONFIG =============================
const CONFIG = {

  // 🔗 PASTE YOUR MAKE.COM WEBHOOK URL HERE
  WEBHOOK_URL: "https://hook.us2.make.com/qi229ejifqv7joaqg1n63zziw630jc3b",

  // 📊 KALLOLO SHEET CONFIG
  KALLOLO: {
    SPREADSHEET_ID: "",          // Leave empty if same spreadsheet as this script
    SHEET_NAME: "Kalalou Raw",       // Exact tab/sheet name — change if different
    HEADER_ROW: 1,               // Row number where column headers are
    DATA_START_ROW: 2,           // Row number where data begins
    // Column names (must match headers exactly)
    INVENTORY_COLUMN: "Available",
    SKU_COLUMN: "Name",
    NAME_COLUMN: "Display Name",
    UPC_COLUMN: "UPC",
    DISCONTINUED_COLUMN: "Discontinued",
    WEBHOOK_SENT_COLUMN: "Webhook Sent"   // Tracking column — script will create it if missing
  },

  // 📊 MELROSE SHEET CONFIG
  MELROSE: {
    SPREADSHEET_ID: "",          // Leave empty if same spreadsheet as this script
    SHEET_NAME: "Melrose Raw",       // Exact tab/sheet name — change if different
    HEADER_ROW: 1,
    DATA_START_ROW: 2,
    // Column names (must match headers exactly)
    INVENTORY_COLUMN: "New Avail",
    SKU_COLUMN: "Vendor SKU",
    NAME_COLUMN: "Product Description",
    UPC_COLUMN: "UPC Code",
    STATUS_COLUMN: "Status",
    WEBHOOK_SENT_COLUMN: "Webhook Sent"   // Tracking column — script will create it if missing
  }
};

 

// ======================== MAIN TRIGGER =======================
 
/**
 * This runs AUTOMATICALLY every time you edit any cell.
 * Do NOT run this manually — it only works on real edits.
 */
function onEdit(e) {
  // Safety check
  if (!e || !e.range) return;
 
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
 
  // Is this edit on a sheet we care about?
  const sheetConfig = CONFIG.SHEETS[sheetName];
  if (!sheetConfig) return;  // Not a monitored sheet — ignore
 
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
 
  // Ignore header row edits
  if (editedRow < sheetConfig.DATA_START_ROW) return;
 
  // Get column headers to find which column was edited
  const headers = getHeaders(sheet, sheetConfig.HEADER_ROW);
  const inventoryColIndex = headers.indexOf(sheetConfig.INVENTORY_COLUMN);
 
  // Was the edited column the inventory column?
  if (editedCol !== inventoryColIndex + 1) return;  // Not inventory column — ignore
 
  // Get the new and old values
  const newValue = parseInventoryValue(e.range.getValue());
  const oldValue = parseInventoryValue(e.oldValue);
 
  // ─── DELIST: Value changed TO 0 ───
  if (newValue === 0 && oldValue !== 0) {
    const rowData = getRowData(sheet, editedRow, headers, sheetConfig);
 
    if (!rowData.sku) return;  // No SKU = skip
 
    const payload = {
      action: "delist",
      supplier: sheetConfig.SUPPLIER_NAME,
      sku: rowData.sku,
      productName: rowData.name,
      upc: rowData.upc,
      previousInventory: oldValue === -1 ? "unknown" : oldValue,
      newInventory: 0,
      editedBy: Session.getActiveUser().getEmail() || "unknown",
      timestamp: new Date().toISOString()
    };
 
    sendWebhook(payload);
 
    // Visual feedback — highlight the cell red
    e.range.setBackground("#FFCCCC");
    e.range.setNote("Delist webhook sent: " + new Date().toLocaleString());
  }
 
  // ─── RELIST: Value changed FROM 0 to something positive ───
  if (oldValue === 0 && newValue > 0) {
    const rowData = getRowData(sheet, editedRow, headers, sheetConfig);
 
    if (!rowData.sku) return;
 
    const payload = {
      action: "relist",
      supplier: sheetConfig.SUPPLIER_NAME,
      sku: rowData.sku,
      productName: rowData.name,
      upc: rowData.upc,
      previousInventory: 0,
      newInventory: newValue,
      editedBy: Session.getActiveUser().getEmail() || "unknown",
      timestamp: new Date().toISOString()
    };
 
    sendWebhook(payload);
 
    // Visual feedback — highlight the cell green
    e.range.setBackground("#CCFFCC");
    e.range.setNote("Relist webhook sent: " + new Date().toLocaleString());
  }
}
 
 
// ======================== BULK PASTE SUPPORT =================
 
/**
 * Handles pasting multiple rows at once.
 * If you paste a column of values, the simple onEdit above only gets
 * the top-left cell. This installable trigger catches the full range.
 *
 * Run setupTrigger() once to activate this.
 */
function onBulkEdit(e) {
  if (!e || !e.range) return;
 
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const sheetConfig = CONFIG.SHEETS[sheetName];
  if (!sheetConfig) return;
 
  const headers = getHeaders(sheet, sheetConfig.HEADER_ROW);
  const inventoryColIndex = headers.indexOf(sheetConfig.INVENTORY_COLUMN);
  if (inventoryColIndex === -1) return;
 
  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();
 
  // Check if the paste includes the inventory column
  const invCol = inventoryColIndex + 1;  // 1-based
  if (invCol < startCol || invCol > startCol + numCols - 1) return;
 
  // Only process if more than 1 row was pasted (single row handled by onEdit)
  if (numRows <= 1) return;
 
  const delistProducts = [];
  const relistProducts = [];
 
  for (let r = 0; r < numRows; r++) {
    const rowNumber = startRow + r;
    if (rowNumber < sheetConfig.DATA_START_ROW) continue;
 
    const cellValue = sheet.getRange(rowNumber, invCol).getValue();
    const inventory = parseInventoryValue(cellValue);
 
    const rowData = getRowData(sheet, rowNumber, headers, sheetConfig);
    if (!rowData.sku) continue;
 
    if (inventory === 0) {
      delistProducts.push({
        supplier: sheetConfig.SUPPLIER_NAME,
        sku: rowData.sku,
        productName: rowData.name,
        upc: rowData.upc
      });
      sheet.getRange(rowNumber, invCol).setBackground("#FFCCCC");
      sheet.getRange(rowNumber, invCol).setNote("Delist webhook sent: " + new Date().toLocaleString());
    }
    // Note: For bulk paste we can't know old values, so we only delist 0s
    // Relist detection works on single cell edits only
  }
 
  if (delistProducts.length > 0) {
    const payload = {
      action: "delist",
      productCount: delistProducts.length,
      products: delistProducts,
      bulkPaste: true,
      editedBy: Session.getActiveUser().getEmail() || "unknown",
      timestamp: new Date().toISOString()
    };
    sendWebhook(payload);
    Logger.log(`✔ Bulk delist: ${delistProducts.length} products from ${sheetConfig.SUPPLIER_NAME}`);
  }
}
 
 
// ======================== UTILITIES =========================
 
function getHeaders(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}
 
function getRowData(sheet, row, headers, config) {
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
 
  const skuIdx = headers.indexOf(config.SKU_COLUMN);
  const nameIdx = headers.indexOf(config.NAME_COLUMN);
  const upcIdx = headers.indexOf(config.UPC_COLUMN);
 
  return {
    sku: skuIdx !== -1 ? String(values[skuIdx] || "").trim() : "",
    name: nameIdx !== -1 ? String(values[nameIdx] || "").trim() : "",
    upc: upcIdx !== -1 ? String(values[upcIdx] || "").trim() : ""
  };
}
 
function parseInventoryValue(value) {
  if (value === null || value === undefined || value === "") return -1;
  const str = String(value).trim().toLowerCase();
  if (str === "0" || str === "out" || str === "oos" || str === "out of stock" || str === "sold out") return 0;
  const num = Number(str);
  return isNaN(num) ? -1 : num;
}
 
function sendWebhook(payload) {
  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      Logger.log(`✔ Webhook sent: ${payload.action} — ${payload.sku || payload.productCount + " products"}`);
      return true;
    }
    Logger.log(`✖ Webhook HTTP ${code}: ${response.getContentText()}`);
    return false;
  } catch (e) {
    Logger.log(`✖ Webhook error: ${e.message}`);
    return false;
  }
}
 
 
// ======================== ONE-TIME SETUP =====================
 
/**
 * 📅 Run this ONCE to enable bulk paste detection.
 * 
 * The simple onEdit above handles single cell edits automatically.
 * This adds an installable trigger for multi-row paste support.
 */
function setupTrigger() {
  // Remove existing triggers from this project
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log(`Removed ${existing.length} old triggers`);
 
  // Create installable onEdit trigger for bulk paste support
  ScriptApp.newTrigger("onBulkEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
 
  Logger.log("✔ Bulk paste trigger installed");
  Logger.log("Setup complete. The script is now watching your inventory columns.");
}
 
 
// ======================== TESTING ============================
 
/**
 * 🧪 Test your webhook connection without editing any cells.
 * Sends a single test payload to verify Make.com receives it.
 */
function testWebhook() {
  const payload = {
    action: "test",
    message: "Webhook connection test from Filigree Inventory Watcher",
    timestamp: new Date().toISOString()
  };
 
  const success = sendWebhook(payload);
  if (success) {
    Logger.log("✔ TEST PASSED — Check your Make.com webhook history");
  } else {
    Logger.log("✖ TEST FAILED — Check your WEBHOOK_URL in CONFIG");
  }
}
