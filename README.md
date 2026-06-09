# Shopify Product Lifecycle Management — Upload, Delist & Relist Automation

End-to-end Shopify product lifecycle automation powered by **Google Sheets + Google Apps Script + Make.com + Shopify API**. From scoring and uploading the best-fit supplier products, to delisting out-of-stock items, and relisting them when inventory returns — all controlled from a single Google Sheet dashboard with App Script buttons.

Optimized to run within Make.com's free-plan credit limits by batching all products in a single webhook call.

---

## Repository Structure

| File | Type | Purpose |
| --- | --- | --- |
| `Master_gui.gs` | Apps Script | Builds the top-bar menu in Google Sheets and binds buttons to their respective functions |
| `Product_Scoring.gs` | Apps Script | Main scoring algorithm — selects the top 12 / 15 / 30 products from both supplier sheets based on store-specific keywords and criteria |
| `Upload_Prodcuts.gs` | Apps Script | Webhook configuration + sends the batched product payload (all 30 in one request) to Make.com |
| `Listing_delisting_helping.gs` | Apps Script | Scans for out-of-stock products and triggers the delist webhook |
| `ReListing_Delisting.gs` | Apps Script | Handles relisting logic and writes status back to the Master Update Sheet |
| `product_uploading_sheet.json` | Make.com Blueprint | Importable scenario for the product upload workflow |
| `product_listing_delisting.json` | Make.com Blueprint | Importable scenario for the delist + relist workflow |
| `Shopify Scoring Sheet.xlsx` | Spreadsheet | Reference copy of the master Google Sheet (tabs, columns, structure) |
| `scoring_user_guide2.docx` | Doc | End-user guide for running the scoring + upload flow |

---

## Business Purpose

Eliminates manual Shopify intervention. Suppliers drop raw weekly data into a Google Sheet, the scoring script picks the best products for the store, Make.com pushes them into Shopify with content, images, variants, price, and inventory — and the same dashboard handles delisting and relisting based on stock status. Supports multiple suppliers out of the box.

---

## How It's Triggered

Everything runs from buttons added to the Google Sheet's top bar by `Master_gui.gs`:

| Button | Action |
| --- | --- |
| **Run Scoring** | Runs `Product_Scoring.gs` across both supplier sheets |
| **Upload Products to Shopify** | `Upload_Prodcuts.gs` fires the upload webhook to Make.com |
| **Out of Stock Scan** | `Listing_delisting_helping.gs` scans for out-of-stock products |
| **Submit Product for Delisting** | Triggers the Make.com delist workflow |
| **Submit for Relisting** | `ReListing_Delisting.gs` triggers the Make.com relist workflow |

---

## Google Sheet Structure

Master Google Sheet:
[Link to Sheet](https://docs.google.com/spreadsheets/d/1lSHxA3EYMjwkXE_aoFG3bn_8NPMjEL5J5ato9ejySIw/edit?usp=sharing)

| Tab | Purpose |
| --- | --- |
| **Supplier 1 / Supplier 2** | Raw weekly supplier data pasted here each week (Kelunuro / Mellaro) |
| **Top Sheet (Top 15 / 30)** | Scored and selected products ready to upload |
| **Master Update Sheet** | Full log of every action — uploaded, delisted, relisted |
| **Product Sending Sheet** | Webhook config used by `Upload_Prodcuts.gs` |
| **Master Menu** | Controls all buttons and navigates between sections |

---

## Data Flow

```
Supplier Sheet (raw data)
    ↓
Product_Scoring.gs  (scoring algorithm)
    ↓
Top Sheet (Top 30 products)
    ↓
Upload_Prodcuts.gs  (single batched webhook POST)
    ↓
Make.com  →  Route by Supplier  →  Build Content  →  Image Routes (1/2/3 URLs)
    ↓
Shopify API  →  Create Product  →  Create Variant (Price + SKU)  →  Sleep  →  Update Inventory  →  Set Available
    ↓
Master Update Sheet (logs)
```

---

## Workflow 1 — Product Scoring & Upload

1. **Paste supplier data** into the Supplier 1 / Supplier 2 tabs.
2. **Click "Run Scoring"** — `Product_Scoring.gs` applies keyword + criteria filters and populates the Top Sheet with the 30 best-fit products.
3. **Click "Upload Products to Shopify"** — `Upload_Prodcuts.gs` fires a single webhook POST containing all 30 products to Make.com (batched to minimize Make.com credit usage).
4. **Make.com receives + routes** by supplier (Route 1 = Supplier 1, Route 2 = Supplier 2). Unused routes can be removed if only one supplier is active.
5. **Build product content** — title, body HTML, color details, and image URLs (3 parallel sub-routes handle products with 1, 2, or 3 images).
6. **Create Product in Shopify** — title, body HTML, vendor, images, and metafields (customizable per store: height, material, ideal_for, etc.).
7. **Create Variant** — adds price and SKU as a separate step. Shopify's API does not allow price during initial product creation; this is a known platform limitation.
8. **Sleep → Get Product → Update Inventory** — wait ensures Shopify has registered the product, then the variant ID is fetched, inventory is updated with quantity + location, and the product is set to Available.
9. **Log to Master Update Sheet** — upload status, product ID, timestamp.

---

## Workflow 2 — Delisting (Out of Stock)

1. **Click "Out of Stock Scan"** — `Listing_delisting_helping.gs` checks the sheet for products flagged out-of-stock.
2. **Review** the identified products.
3. **Click "Submit Product for Delisting"** — fires the delist webhook to Make.com.
4. **Make.com archives** the product in Shopify (archived, not deleted — so it can be relisted later).
5. **Log** to the Master Update Sheet.

---

## Workflow 3 — Relisting

1. **Review** delisted products in the relisting section of the sheet.
2. **Click "Submit for Relisting"** — `ReListing_Delisting.gs` fires the relist webhook to Make.com.
3. **Make.com activates** the archived product in Shopify, setting it Available again.
4. **Log** to the Master Update Sheet.

---

## Setup

### 1. Apps Script

1. Open the master Google Sheet → **Extensions → Apps Script**.
2. Create the following script files and paste contents from this repo:
   - `Master_gui.gs`
   - `Product_Scoring.gs`
   - `Upload_Prodcuts.gs`
   - `Listing_delisting_helping.gs`
   - `ReListing_Delisting.gs`
3. Update the webhook URLs inside `Upload_Prodcuts.gs`, `Listing_delisting_helping.gs`, and `ReListing_Delisting.gs` to point to your Make.com scenarios.
4. Save and reload the sheet — the menu defined in `Master_gui.gs` will appear in the top bar.

### 2. Make.com Scenarios

1. In Make.com, create two new scenarios.
2. Import the blueprints:
   - `product_uploading_sheet.json` → upload scenario
   - `product_listing_delisting.json` → delist / relist scenario
3. Reconnect the Google Sheets and Shopify modules with your own credentials.
4. Copy each scenario's webhook URL into the corresponding Apps Script file.

### 3. Shopify App (One-Time)

Required when connecting Make.com to a new Shopify store (post-January 2026 Shopify policy update).

1. Shopify Admin → **Apps → App Settings**.
2. **Develop Apps → Build Apps in Dev Dashboard**.
3. Go to `dev.shopify.com/dashboard/[store-id]/apps`.
4. Create a new app and configure scopes:
   - `read_products`, `write_products`
   - `read_inventory`, `write_inventory`
   - `read_locations`
   - *(Add additional scopes per store requirements)*
5. **Install the app** on the store.
6. Copy the **API Key & Secret**.
7. In Make.com → Shopify module → Connection settings, paste the credentials and select required scopes.
8. Save and test the connection.

---

## Customization

`Product_Scoring.gs` is the main file to adjust per store:

- Keyword filters
- Allowed product types
- Selection criteria / scoring weights
- Number of products selected (12 / 15 / 30)

Metafields applied during product creation are configured inside the Make.com upload scenario (`product_uploading_sheet.json`).

---

## Error Handling

| Failure | Cause | Handling |
| --- | --- | --- |
| Image URL mismatch | Product has 1, 2, or 3 image URLs | 3 parallel image routes in Make.com cover all cases |
| Price not added | Shopify API doesn't accept price on create | Separate Create Variant step handles price + SKU |
| Inventory not updated | Product not yet registered in Shopify | Sleep module added before inventory call |
| API rate limit | Too many Shopify API calls | Single-batch webhook (30 products / 1 call) minimizes requests |
| Invalid credentials | Expired or revoked Shopify app token | Reconnect in Make.com Shopify module; reinstall app if needed |
| Supplier sheet format change | Column positions shifted | Update column mappings in `Product_Scoring.gs` |
| Google Sheets auth expired | Make.com loses Sheets connection | Reconnect the Google Sheets module in Make.com |

**Recovery procedure:**

1. Check the **Master Update Sheet** logs for failed or incomplete entries.
2. Open **Make.com execution history** for the relevant scenario and inspect the error block.
3. Verify Shopify API credentials and scopes in the Make.com connection.
4. Re-run the specific step manually via the sheet button.
5. If still failing, check the Shopify app install status at `dev.shopify.com`.

---

## Security Notes

- Shopify API credentials live **only** in Make.com's secure connection vault — never in the Google Sheet.
- Restrict Google Sheet access to specific emails — never public.
- Apps Script deployment access set to **"Only myself"** or specific org users.
- Follow **least-privilege** on API scopes — only grant what's required.
- Rotate Shopify API tokens periodically and update Make.com connections.

---

## Maintenance

**Weekly**

- Paste new supplier sheets into Supplier 1 / Supplier 2 tabs before running scoring.
- Verify supplier sheet column structure matches what `Product_Scoring.gs` expects.

**Monthly**

- Review Make.com execution logs for recurring errors.
- Check Shopify app token validity at `dev.shopify.com`.
- Audit the Master Update Sheet for anomalies (unexpected delist/relist counts).

**Common Fixes**

| Issue | Fix |
| --- | --- |
| Make.com stops mid-scenario | Reconnect Google Sheets or Shopify under module connections |
| Scoring returns no results | Keyword filters in `Product_Scoring.gs` may be too restrictive |
| Products uploaded but no price | Variant creation step failed — check Make.com logs |
| Inventory not showing after upload | Location ID in inventory module doesn't match store location |
| Delist not working | Confirm webhook URL in `Listing_delisting_helping.gs` points to the right Make.com scenario |

---

## Documentation

| Resource | Link |
| --- | --- |
| Shopify API Docs | [dev.shopify.com](https://dev.shopify.com) |
| Make.com Docs | [make.com/help](https://www.make.com/en/help) |
| Google Apps Script Docs | [developers.google.com/apps-script](https://developers.google.com/apps-script) |
| User Guide | `scoring_user_guide2.docx` |

---

## Future Improvements

- **AI-based product scoring** — replace keyword filtering in `Product_Scoring.gs` with an AI model that scores based on store performance history.
- **Automated supplier sheet import** — pull supplier data directly via API or email-attachment parser, removing the manual paste step.
- **Real-time inventory sync** — replace the manual "Out of Stock Scan" with a scheduled or webhook-triggered inventory check.
- **Supplier performance analytics** — track which supplier's products sell faster and have fewer delists.
- **Automated weekly reporting** — email or Slack summary of upload / delist / relist counts per cycle.
- **Multi-location inventory** — extend inventory update logic to handle multiple Shopify warehouse locations.
