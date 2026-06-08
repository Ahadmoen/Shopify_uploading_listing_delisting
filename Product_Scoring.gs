// ── PRODUCT SCORING SYSTEM v3.5 ───────────────────────────────────────────────
// v3.5: Keyword force-include via prompt
// Kalalou Top 15 = 5 furniture + up to 5 keyword (accessory/bigticket) + remaining merit
// Melrose Top 15 = up to 5 keyword + remaining merit
// Pool column shows "accessory | keyword match" etc for keyword matched products
// No new columns — same 28 column structure

function scoreProducts() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  const kaSheet     = ss.getSheetByName("Kalalou Raw");
  const melSheet    = ss.getSheetByName("Melrose Raw");
  const outSheet    = ss.getSheetByName("Top 15");
  const kaOutSheet  = ss.getSheetByName("Kalalou Top");
  const melOutSheet = ss.getSheetByName("Melrose Top");

  if (!kaSheet || !melSheet || !outSheet || !kaOutSheet || !melOutSheet) {
    ui.alert("Missing tabs. Need: 'Kalalou Raw', 'Melrose Raw', 'Top 15', 'Kalalou Top', 'Melrose Top'");
    return;
  }

  // ── Keyword prompt ────────────────────────────────────────────────────────
  var keywordResp = ui.prompt(
    "🔍 Keyword Force-Include (Optional)",
    "Enter keywords to force-include products (comma separated).\nExample: bee, butterfly, mushroom\n\nLeave blank to skip.",
    ui.ButtonSet.OK_CANCEL
  );
  if (keywordResp.getSelectedButton() === ui.Button.CANCEL) return;

  var keywords = [];
  var rawInput = keywordResp.getResponseText().trim();
  if (rawInput !== "") {
    keywords = rawInput.split(",")
      .map(function(k){ return k.trim().toLowerCase(); })
      .filter(function(k){ return k.length > 0; });
  }

  const kaData = kaSheet.getDataRange().getValues();
  const kaHdrs = kaData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });
  const KA = {
    sku          : kaHdrs.indexOf("Name"),
    name         : kaHdrs.indexOf("Display Name"),
    cost         : kaHdrs.findIndex(function(h){ return h.includes("Carton Price") || h.includes("Sold in Box/ Carton Price"); }),
    msrp         : kaHdrs.findIndex(function(h){ return h.includes("Carton MAP") || h.includes("MAP"); }),
    stock        : kaHdrs.indexOf("Available"),
    freightOnly  : kaHdrs.indexOf("Freight Only"),
    discontinued : kaHdrs.indexOf("Discontinued"),
    func         : kaHdrs.indexOf("Function"),
    theme        : kaHdrs.indexOf("Theme"),
    material     : kaHdrs.indexOf("Material"),
    color        : kaHdrs.indexOf("Color"),
    copy         : kaHdrs.indexOf("Product Copy"),
    construction : kaHdrs.indexOf("Construction"),
    boxQty       : kaHdrs.indexOf("Sold in Box/Qty of"),
  };

  const melData = melSheet.getDataRange().getValues();
  const melHdrs = melData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });
  const MEL = {
    sku      : melHdrs.indexOf("Vendor SKU"),
    name     : melHdrs.indexOf("Web Title"),
    cost     : melHdrs.indexOf("Dropship Freight Included Price"),
    msrp     : melHdrs.indexOf("MSRP"),
    stock    : melHdrs.findIndex(function(h){ return h === "Avail" || h === "New Avail"; }),
    category : melHdrs.indexOf("Product Category"),
    style    : melHdrs.indexOf("Style"),
    material : melHdrs.indexOf("Item Materials"),
    color    : melHdrs.indexOf("Color"),
    copy     : melHdrs.indexOf("Marketing Copy"),
    setOf    : melHdrs.indexOf("Set of"),
  };

  var FURN_KW = ["furniture","chair","bench","stool","table","shelf","shelving",
                 "cabinet","console","display furniture","accent chair",
                 "dining chair","side table","coffee table","accent table","console table"];

  function isFurniture(func) {
    var f = (func || "").toString().toLowerCase();
    return FURN_KW.some(function(k){ return f.indexOf(k) !== -1; });
  }

  function getPool(func, msrp) {
    if (isFurniture(func)) return "furniture";
    if (msrp >= 150) return "bigticket";
    return "accessory";
  }

  function matchesKeyword(productName) {
    if (keywords.length === 0) return false;
    var n = (productName || "").toString().toLowerCase();
    return keywords.some(function(k){ return n.indexOf(k) !== -1; });
  }

  function calcMarginPct(cost, msrp) {
    if (!msrp || msrp === 0) return 0;
    return Math.round((msrp - cost) / msrp * 100 * 10) / 10;
  }

  function grossMarginScore(cost, msrp) {
    if (!msrp || msrp === 0) return 0;
    var pct = Math.max(0, Math.min(65, (msrp - cost) / msrp * 100));
    return Math.round(pct / 65 * 18 * 100) / 100;
  }

  function absDollarScore(cost, msrp) {
    var d = msrp - cost;
    if (d >= 80) return 12;
    if (d >= 60) return 10;
    if (d >= 40) return 8;
    if (d >= 25) return 6;
    if (d >= 15) return 4;
    return 2;
  }

  function priceScore(msrp, pool) {
    if (pool === "furniture") {
      if (msrp >= 150 && msrp <= 800) return 12;
      if ((msrp >= 80 && msrp < 150) || (msrp > 800 && msrp <= 1200)) return 9;
      if (msrp > 1200) return 5;
      return 6;
    } else if (pool === "bigticket") {
      if (msrp >= 150 && msrp <= 400) return 12;
      if ((msrp >= 100 && msrp < 150) || (msrp > 400 && msrp <= 600)) return 9;
      if (msrp > 600) return 5;
      return 6;
    } else {
      if (msrp >= 40  && msrp <= 150) return 12;
      if (msrp >= 25  && msrp <  40)  return 9;
      if (msrp > 150  && msrp <= 250) return 9;
      if (msrp >= 15  && msrp <  25)  return 5;
      if (msrp > 250  && msrp <= 400) return 5;
      return 2;
    }
  }

  function availScore(stock, pool) {
    if (pool === "furniture") {
      if (stock >= 20) return 6;
      if (stock >= 10) return 4;
      return 2;
    } else {
      if (stock >= 100) return 10;
      if (stock >= 50)  return 8;
      if (stock >= 20)  return 6;
      if (stock >= 10)  return 4;
      return 2;
    }
  }

  function reorderScore(stock, pool) {
    if (pool === "furniture") return 3;
    if (stock >= 300) return 10;
    if (stock >= 150) return 8;
    if (stock >= 75)  return 5;
    if (stock >= 30)  return 3;
    return 1;
  }

  function categoryScore(cat) {
    if (!cat) return 3;
    var c = cat.toString().toLowerCase();
    var high = ["decor","vase","candleholder","planter","faux","wall dec","lantern",
                "mirror","tray","bowl","basket","wreath","garland","stem","spray",
                "botanica","sculptural","furniture","shelf","shelving","display",
                "cabinet","console"];
    var med  = ["garden","novelty","serving","tabletop","candle","pots",
                "chair","bench","stool","table","lamp"];
    if (high.some(function(k){ return c.indexOf(k) !== -1; })) return 10;
    if (med.some(function(k){  return c.indexOf(k) !== -1; })) return 6;
    return 3;
  }

  function setSizeScoreByName(productName) {
    var name = (productName || "").toString().toLowerCase();
    if (name.match(/set of (4|5|6|7|8|9|\d{2})/)) return 8;
    if (name.match(/set of 3/))               return 7;
    if (name.match(/set of 2/))               return 6;
    if (name.match(/bundle|pack of|pair of/)) return 5;
    return 5;
  }

  function setSizeScoreByField(setOfVal, productName) {
    if (setOfVal && !isNaN(parseFloat(setOfVal))) {
      var qty = parseFloat(setOfVal);
      if (qty >= 4) return 8;
      if (qty === 3) return 7;
      if (qty === 2) return 6;
    }
    return setSizeScoreByName(productName);
  }

  function multiFuncScore(category, productName) {
    var combined = ((category || "") + " " + (productName || "")).toString().toLowerCase();
    var pairs = [
      ["planter","vase"],["tray","serving"],["storage","display"],
      ["basket","planter"],["vase","planter"],["bowl","tray"],
      ["lantern","planter"],["shelf","display"],["holder","vase"],
      ["console","shelf"],["bench","storage"],["table","shelf"]
    ];
    for (var i = 0; i < pairs.length; i++) {
      if (combined.indexOf(pairs[i][0]) !== -1 && combined.indexOf(pairs[i][1]) !== -1) return 7;
    }
    var versatile = ["multi","multipurpose","versatile","2-in-1","dual"];
    if (versatile.some(function(k){ return combined.indexOf(k) !== -1; })) return 7;
    var broad = ["tray","basket","bowl","planter","lantern","shelf","display","console"];
    if (broad.some(function(k){ return combined.indexOf(k) !== -1; })) return 4;
    return 1;
  }

  function handmadeScore(copy, construction, material, productName) {
    var combined = [copy, construction, material, productName]
      .map(function(v){ return (v || "").toString().toLowerCase(); }).join(" ");
    var strong = ["handcrafted","hand-crafted","hand crafted","hand-woven","hand woven",
                  "handwoven","hand-painted","hand painted","handpainted",
                  "skilled artisan","artisan","hand-made","handmade"];
    var medium = ["natural material","natural fiber","rattan","seagrass","jute",
                  "wicker","bamboo","reclaimed","recycled glass","recycled metal",
                  "hand-blown","terracotta","clay","teak","mango wood","acacia",
                  "recycled wood","iron"];
    if (strong.some(function(k){ return combined.indexOf(k) !== -1; })) return 8;
    if (medium.some(function(k){ return combined.indexOf(k) !== -1; })) return 5;
    return 1;
  }

  function trendScore(productName, category, theme, material) {
    var combined = [productName, category, theme, material]
      .map(function(v){ return (v || "").toString().toLowerCase(); }).join(" ");
    var kw = ["coastal","sculptural","textured","organic","terracotta","rattan",
              "seagrass","wabi","japandi","boho","arch","curved","mushroom","dried",
              "pampas","linen","stone","marble","travertine","boucle","checkered",
              "abstract","earthy","neutral","minimalist","industrial","transitional",
              "mid-century","contemporary","modern"];
    var hits = kw.filter(function(k){ return combined.indexOf(k) !== -1; }).length;
    if (hits >= 3) return 5;
    if (hits === 2) return 4;
    if (hits === 1) return 2;
    return 0;
  }

  function getTier(score) {
    if (score >= 72) return "🟢 STRONG BUY";
    if (score >= 58) return "🟡 CONSIDER";
    return "🔴 LOW PRIORITY";
  }

  function scoreKalalou(cost, msrp, stock, func, theme, mat, name, copy, cons, pool) {
    var mSc  = grossMarginScore(cost, msrp);
    var abSc = absDollarScore(cost, msrp);
    var pSc  = priceScore(msrp, pool);
    var aSc  = availScore(stock, pool);
    var rSc  = reorderScore(stock, pool);
    var cSc  = categoryScore(func);
    var sSc  = setSizeScoreByName(name);
    var mfSc = multiFuncScore(func, name);
    var hSc  = handmadeScore(copy, cons, mat, name);
    var tSc  = trendScore(name, func, theme, mat);
    var tot  = Math.round((mSc+abSc+pSc+aSc+rSc+cSc+sSc+mfSc+hSc+tSc)*100)/100;
    return [mSc, abSc, pSc, aSc, rSc, cSc, sSc, mfSc, hSc, tSc, tot];
  }

  function scoreMelrose(cost, msrp, stock, cat, style, mat, name, copy, setOf, pool) {
    var mSc  = grossMarginScore(cost, msrp);
    var abSc = absDollarScore(cost, msrp);
    var pSc  = priceScore(msrp, pool);
    var aSc  = availScore(stock, pool);
    var rSc  = reorderScore(stock, pool);
    var cSc  = categoryScore(cat);
    var sSc  = setSizeScoreByField(setOf, name);
    var mfSc = multiFuncScore(cat, name);
    var hSc  = handmadeScore(copy, "", mat, name);
    var tSc  = trendScore(name, cat, style, mat);
    var tot  = Math.round((mSc+abSc+pSc+aSc+rSc+cSc+sSc+mfSc+hSc+tSc)*100)/100;
    return [mSc, abSc, pSc, aSc, rSc, cSc, sSc, mfSc, hSc, tSc, tot];
  }

  function writeSheet(ws, headers, rows) {
    ws.clearContents();
    ws.clearFormats();
    var nc = headers.length;
    ws.getRange(1, 1, 1, nc).setValues([headers]).setFontWeight("bold");
    ws.setRowHeight(1, 22);
    ws.setFrozenRows(1);
    if (rows.length === 0) return;
    ws.getRange(2, 1, rows.length, nc).setValues(rows);
    ws.getRange(2, 4, rows.length, 1).setNumberFormat("#,##0.00");
    ws.getRange(2, 5, rows.length, 1).setNumberFormat("#,##0.00");
    ws.getRange(2, 6, rows.length, 1).setNumberFormat("0.0");
    ws.getRange(2, 7, rows.length, 1).setNumberFormat("#,##0.00");
    ws.getRange(2, 8, rows.length, 1).setNumberFormat("#,##0");
    ws.setColumnWidth(3, 266);
    for (var i = 0; i < rows.length; i++) ws.setRowHeight(i + 2, 18);
  }

  // ── Score Kalalou ─────────────────────────────────────────────────────────
  var kaScored = [];

  for (var i = 1; i < kaData.length; i++) {
    var row  = kaData[i];
    var sku  = (row[KA.sku]  || "").toString();
    var name = (row[KA.name] || "").toString();
    if (!sku && !name) continue;

    var cost  = parseFloat(row[KA.cost]) || 0;
    var msrp  = parseFloat(row[KA.msrp]) || 0;
    var stock = parseFloat(row[KA.stock]) || 0;
    var disc  = (row[KA.discontinued] || "").toString().trim().toUpperCase();
    var frt   = (row[KA.freightOnly]  || "").toString().trim().toUpperCase();
    var func  = (row[KA.func]         || "").toString();
    var theme = (row[KA.theme]        || "").toString();
    var mat   = (row[KA.material]     || "").toString();
    var col   = (row[KA.color]        || "").toString();
    var copy  = (KA.copy >= 0         ? row[KA.copy]         : "") || "";
    var cons  = (KA.construction >= 0 ? row[KA.construction] : "") || "";
    var pool  = getPool(func, msrp);
    var furn  = (pool === "furniture");

    var discFail  = disc === "YES";
    var frtFail   = !furn && frt === "YES";
    var stockFail = furn ? stock < 5 : stock <= 0;
    var costFail  = cost <= 0;
    var msrpFail  = msrp <= 0;
    var eligible  = !discFail && !frtFail && !stockFail && !costFail && !msrpFail;
    var eligLabel = eligible ? "✅ INCLUDE" : "❌ EXCLUDE";

    var reason = "";
    if      (discFail)  reason = "Discontinued";
    else if (frtFail)   reason = "Freight Only";
    else if (stockFail) reason = furn ? "Stock < 5 units" : "Zero Stock";
    else if (costFail)  reason = "Missing Cost";
    else if (msrpFail)  reason = "Missing MSRP";

    var sc = eligible
      ? scoreKalalou(cost, msrp, stock, func, theme, mat, name, copy, cons, pool)
      : ["","","","","","","","","","",""];

    kaScored.push([
      "Kalalou", sku, name.substring(0, 60),
      cost, msrp, calcMarginPct(cost, msrp),
      eligible ? (msrp - cost) : "",
      stock, func, theme, mat, col,
      name.substring(0, 60),
      eligLabel, reason,
      sc[0],sc[1],sc[2],sc[3],sc[4],
      sc[5],sc[6],sc[7],sc[8],sc[9],
      sc[10], sc[10] !== "" ? getTier(sc[10]) : "",
      pool
    ]);
  }

  kaScored.sort(function(a, b) {
    if (a[13]==="✅ INCLUDE" && b[13]!=="✅ INCLUDE") return -1;
    if (a[13]!=="✅ INCLUDE" && b[13]==="✅ INCLUDE") return 1;
    return (b[25]||0) - (a[25]||0);
  });

  // ── Score Melrose ─────────────────────────────────────────────────────────
  var melScored = [];

  for (var j = 1; j < melData.length; j++) {
    var mrow  = melData[j];
    var msku  = (mrow[MEL.sku]  || "").toString();
    var mname = (mrow[MEL.name] || "").toString();
    if (!msku && !mname) continue;

    var mcost  = parseFloat(mrow[MEL.cost])  || 0;
    var mmsrp  = parseFloat(mrow[MEL.msrp])  || 0;
    var mstock = parseFloat(mrow[MEL.stock]) || 0;
    var mcat   = (mrow[MEL.category] || "").toString();
    var mstyle = (mrow[MEL.style]    || "").toString();
    var mmat   = (mrow[MEL.material] || "").toString();
    var mcol   = (mrow[MEL.color]    || "").toString();
    var mcopy  = (MEL.copy  >= 0 ? mrow[MEL.copy]  : "") || "";
    var msetOf = (MEL.setOf >= 0 ? mrow[MEL.setOf] : 1)  || 1;
    var mpool  = mmsrp >= 150 ? "bigticket" : "accessory";
    var melig  = mstock > 0 && mcost > 0 && mmsrp > 0;
    var meligL = melig ? "✅ INCLUDE" : "❌ EXCLUDE";
    var mreason = !mstock ? "Zero Stock" : !mcost ? "Missing Cost" : !mmsrp ? "Missing MSRP" : "";

    var msc = melig
      ? scoreMelrose(mcost, mmsrp, mstock, mcat, mstyle, mmat, mname, mcopy, msetOf, mpool)
      : ["","","","","","","","","","",""];

    melScored.push([
      "Melrose", msku, mname.substring(0, 60),
      mcost, mmsrp, calcMarginPct(mcost, mmsrp),
      melig ? (mmsrp - mcost) : "",
      mstock, mcat, mstyle, mmat, mcol,
      mname.substring(0, 60),
      meligL, mreason,
      msc[0],msc[1],msc[2],msc[3],msc[4],
      msc[5],msc[6],msc[7],msc[8],msc[9],
      msc[10], msc[10] !== "" ? getTier(msc[10]) : "",
      mpool
    ]);
  }

  melScored.sort(function(a, b) {
    if (a[13]==="✅ INCLUDE" && b[13]!=="✅ INCLUDE") return -1;
    if (a[13]!=="✅ INCLUDE" && b[13]==="✅ INCLUDE") return 1;
    return (b[25]||0) - (a[25]||0);
  });

  // ── Build Kalalou Top 15 ──────────────────────────────────────────────────
  var kaEligible  = kaScored.filter(function(r){ return r[13]==="✅ INCLUDE"; });
  var kaFurniture = kaEligible.filter(function(r){ return r[27]==="furniture"; });
  var kaOther     = kaEligible.filter(function(r){ return r[27]!=="furniture"; });

  kaFurniture.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });
  kaOther.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  var kaKeyword = kaOther.filter(function(r){ return matchesKeyword(r[2]); });
  var kaMerit   = kaOther.filter(function(r){ return !matchesKeyword(r[2]); });

  kaKeyword.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });
  kaMerit.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  var kaTop5Furn    = kaFurniture.slice(0, 5);
  var kaTop5Keyword = kaKeyword.slice(0, 5);
  var meritSlots    = 15 - kaTop5Furn.length - kaTop5Keyword.length;
  var kaTopMerit    = kaMerit.slice(0, meritSlots);

  kaTop5Keyword = kaTop5Keyword.map(function(r){
    var row = r.slice();
    row[27] = row[27] + " | keyword match";
    return row;
  });

  var kaTop15 = kaTop5Furn.concat(kaTop5Keyword).concat(kaTopMerit);
  kaTop15.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  // ── Build Melrose Top 15 ──────────────────────────────────────────────────
  var melEligible = melScored.filter(function(r){ return r[13]==="✅ INCLUDE"; });
  melEligible.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  var melKeyword = melEligible.filter(function(r){ return matchesKeyword(r[2]); });
  var melMerit   = melEligible.filter(function(r){ return !matchesKeyword(r[2]); });

  melKeyword.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });
  melMerit.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  var melTop5Keyword = melKeyword.slice(0, 5);
  var melMeritSlots  = 15 - melTop5Keyword.length;
  var melTopMerit    = melMerit.slice(0, melMeritSlots);

  melTop5Keyword = melTop5Keyword.map(function(r){
    var row = r.slice();
    row[27] = row[27] + " | keyword match";
    return row;
  });

  var melTop15 = melTop5Keyword.concat(melTopMerit);
  melTop15.sort(function(a,b){ return (b[25]||0)-(a[25]||0); });

  // ── Combined Top 15 tab ───────────────────────────────────────────────────
  var top15 = kaTop15.concat(melTop15);

  var hdrs = [
    "Supplier","SKU","Product Name","Cost","MSRP","Margin %","Abs $ Margin","Stock",
    "Category","Theme","Material","Color","Set Size",
    "Eligible","Reason",
    "Margin Sc","Abs$ Sc","Price Sc","Avail Sc","Reorder Sc",
    "Cat Sc","Set Sc","MultiFunc Sc","Handmade Sc","Trend Sc",
    "TOTAL","TIER","Pool"
  ];

  writeSheet(kaOutSheet,  hdrs, kaScored);
  writeSheet(melOutSheet, hdrs, melScored);
  writeSheet(outSheet,    hdrs, top15);

  var kwSummary = keywords.length > 0
    ? "\nKeywords: " + keywords.join(", ") + "\n" +
      "  Kalalou keyword matches: " + kaTop5Keyword.length + "\n" +
      "  Melrose keyword matches: " + melTop5Keyword.length
    : "\nNo keywords — pure merit scoring.";

  ui.alert(
    "✅ Done! Scoring v3.5\n\n" +
    "Kalalou Top 15:\n" +
    "  🪑 Furniture: " + kaTop5Furn.length + "\n" +
    "  🔍 Keyword matches: " + kaTop5Keyword.length + "\n" +
    "  🏠 Merit: " + kaTopMerit.length + "\n\n" +
    "Melrose Top 15:\n" +
    "  🔍 Keyword matches: " + melTop5Keyword.length + "\n" +
    "  🏠 Merit: " + melTopMerit.length +
    kwSummary
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏆 Product Scoring")
    .addItem("Run Scoring → Generate All Tabs", "scoreProducts")
    .addToUi();
}

function debugKalalou() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kaSheet = ss.getSheetByName("Kalalou Raw");
  const kaData = kaSheet.getDataRange().getValues();
  const kaHdrs = kaData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });
  
  const costIdx = kaHdrs.findIndex(function(h){ return h.includes("Carton Price") || h.includes("Sold in Box/ Carton Price"); });
  const msrpIdx = kaHdrs.findIndex(function(h){ return h.includes("Carton MAP") || h.includes("MAP"); });
  const stockIdx = kaHdrs.indexOf("Available");
  const discIdx = kaHdrs.indexOf("Discontinued");
  const nameIdx = kaHdrs.indexOf("Name");
  
  SpreadsheetApp.getUi().alert(
    "Name col: " + nameIdx + "\n" +
    "Cost col: " + costIdx + "\n" +
    "MSRP col: " + msrpIdx + "\n" +
    "Stock col: " + stockIdx + "\n" +
    "Disc col: " + discIdx + "\n\n" +
    "Cost header: " + kaHdrs[costIdx] + "\n" +
    "MSRP header: " + kaHdrs[msrpIdx]
  );
}

function debugKalalouData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kaSheet = ss.getSheetByName("Kalalou Raw");
  const kaData = kaSheet.getDataRange().getValues();
  const kaHdrs = kaData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });

  const KA = {
    sku          : kaHdrs.indexOf("Name"),
    name         : kaHdrs.indexOf("Display Name"),
    cost         : kaHdrs.findIndex(function(h){ return h.includes("Carton Price") || h.includes("Sold in Box/ Carton Price"); }),
    msrp         : kaHdrs.findIndex(function(h){ return h.includes("Carton MAP") || h.includes("MAP"); }),
    stock        : kaHdrs.indexOf("Available"),
    freightOnly  : kaHdrs.indexOf("Freight Only"),
    discontinued : kaHdrs.indexOf("Discontinued"),
  };

  var results = "";
  var count = 0;

  for (var i = 1; i < kaData.length && count < 5; i++) {
    var row = kaData[i];
    var sku  = (row[KA.sku]  || "").toString();
    var name = (row[KA.name] || "").toString();
    if (!sku && !name) continue;

    var cost  = parseFloat(row[KA.cost]) || 0;
    var msrp  = parseFloat(row[KA.msrp]) || 0;
    var stock = parseFloat(row[KA.stock]) || 0;
    var disc  = (row[KA.discontinued] || "").toString().trim().toUpperCase();
    var frt   = (row[KA.freightOnly]  || "").toString().trim().toUpperCase();

    results += "SKU: " + sku + "\n";
    results += "Cost: " + cost + " | MSRP: " + msrp + " | Stock: " + stock + "\n";
    results += "Disc: " + disc + " | Freight: " + frt + "\n\n";
    count++;
  }

  SpreadsheetApp.getUi().alert(results);
}


function debugA0290() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kaSheet = ss.getSheetByName("Kalalou Raw");
  const kaData = kaSheet.getDataRange().getValues();
  const kaHdrs = kaData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });

  const KA = {
    sku          : kaHdrs.indexOf("Name"),
    name         : kaHdrs.indexOf("Display Name"),
    cost         : kaHdrs.findIndex(function(h){ return h.includes("Carton Price") || h.includes("Sold in Box/ Carton Price"); }),
    msrp         : kaHdrs.findIndex(function(h){ return h.includes("Carton MAP") || h.includes("MAP"); }),
    stock        : kaHdrs.indexOf("Available"),
    freightOnly  : kaHdrs.indexOf("Freight Only"),
    discontinued : kaHdrs.indexOf("Discontinued"),
    func         : kaHdrs.indexOf("Function"),
  };

  for (var i = 1; i < kaData.length; i++) {
    var row = kaData[i];
    var sku = (row[KA.sku] || "").toString();
    if (sku !== "A0290") continue;

    var cost  = parseFloat(row[KA.cost]) || 0;
    var msrp  = parseFloat(row[KA.msrp]) || 0;
    var stock = parseFloat(row[KA.stock]) || 0;
    var disc  = (row[KA.discontinued] || "").toString().trim().toUpperCase();
    var frt   = (row[KA.freightOnly]  || "").toString().trim().toUpperCase();
    var func  = (row[KA.func] || "").toString();

    var discFail  = disc === "YES";
    var frtFail   = frt === "YES";
    var stockFail = stock <= 0;
    var costFail  = cost <= 0;
    var msrpFail  = msrp <= 0;

    SpreadsheetApp.getUi().alert(
      "SKU: " + sku + "\n" +
      "Cost: " + cost + " | costFail: " + costFail + "\n" +
      "MSRP: " + msrp + " | msrpFail: " + msrpFail + "\n" +
      "Stock: " + stock + " | stockFail: " + stockFail + "\n" +
      "Disc: " + disc + " | discFail: " + discFail + "\n" +
      "Freight: " + frt + " | frtFail: " + frtFail + "\n" +
      "Function: " + func + "\n" +
      "ELIGIBLE: " + (!discFail && !frtFail && !stockFail && !costFail && !msrpFail)
    );
    return;
  }
}


function debugKalalouTop15() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kaSheet = ss.getSheetByName("Kalalou Raw");
  const kaData = kaSheet.getDataRange().getValues();
  const kaHdrs = kaData[0].map(function(h){ return h.toString().replace(/\n/g, " ").trim(); });

  const KA = {
    sku          : kaHdrs.indexOf("Name"),
    name         : kaHdrs.indexOf("Display Name"),
    cost         : kaHdrs.findIndex(function(h){ return h.includes("Carton Price") || h.includes("Sold in Box/ Carton Price"); }),
    msrp         : kaHdrs.findIndex(function(h){ return h.includes("Carton MAP") || h.includes("MAP"); }),
    stock        : kaHdrs.indexOf("Available"),
    freightOnly  : kaHdrs.indexOf("Freight Only"),
    discontinued : kaHdrs.indexOf("Discontinued"),
    func         : kaHdrs.indexOf("Function"),
  };

  var eligibleCount = 0;
  var furnitureCount = 0;
  var sampleScore = "";

  for (var i = 1; i < kaData.length; i++) {
    var row = kaData[i];
    var sku  = (row[KA.sku]  || "").toString();
    var name = (row[KA.name] || "").toString();
    if (!sku && !name) continue;

    var cost  = parseFloat(row[KA.cost]) || 0;
    var msrp  = parseFloat(row[KA.msrp]) || 0;
    var stock = parseFloat(row[KA.stock]) || 0;
    var disc  = (row[KA.discontinued] || "").toString().trim().toUpperCase();
    var frt   = (row[KA.freightOnly]  || "").toString().trim().toUpperCase();
    var func  = (row[KA.func] || "").toString().toLowerCase();

    var isFurn = ["furniture","chair","bench","stool","table","shelf","shelving",
                  "cabinet","console"].some(function(k){ return func.indexOf(k) !== -1; });
    var pool = isFurn ? "furniture" : (msrp >= 150 ? "bigticket" : "accessory");

    var eligible = disc !== "YES" && 
                   !((!isFurn) && frt === "YES") && 
                   (isFurn ? stock >= 5 : stock > 0) && 
                   cost > 0 && msrp > 0;

    if (eligible) {
      eligibleCount++;
      if (isFurn) furnitureCount++;
      if (eligibleCount <= 3) {
        sampleScore += "SKU: " + sku + " | Pool: " + pool + " | Score will be calculated\n";
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    "Total eligible: " + eligibleCount + "\n" +
    "Furniture: " + furnitureCount + "\n\n" +
    "Sample eligible products:\n" + sampleScore
  );
}
