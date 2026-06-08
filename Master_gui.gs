// ================================================================
//  MASTER MENU — All menus in one place
//  This is the ONLY file with onOpen()
// ================================================================

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu("🏠 Inventory")
      // .addItem("⚙️  Initial Setup",            "runSetup")
      // .addSeparator()
      .addItem("🔍  Scan for Out of Stock",     "scanDelist")
      .addItem("📤  Submit Product for Delisting",              "submitDelist")
      .addSeparator()
      .addItem("🔄  Scan for Restocked",        "scanRelist")
      .addItem("📤  Submit Product for Relisting",              "submitRelist")
      .addToUi();

    ui.createMenu("🏆 Scoring")
      .addItem("Run Scoring → Generate All Tabs",   "scoreProducts")
      .addSeparator()
      .addItem("🚀 Upload Products to Shopify",         "sendTopProducts")
      .addToUi();

  } catch (e) {}
}
