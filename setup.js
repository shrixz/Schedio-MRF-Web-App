// =============================================================================
// SETUP — idempotent. Safe to run repeatedly. Creates any missing sheets and
// seeds default dropdown values. Never deletes or overwrites existing rows.
//
// It also "heals" the header row of sheets that ALREADY exist: it appends any
// missing trailing columns and fills blank header cells, so you can fix column
// drift by just re-running setup instead of editing sheets by hand. It will
// NOT overwrite a header cell that is already filled with different text (that
// might be a deliberate rename) — such mismatches are reported, not changed.
// No data rows are ever touched.
// =============================================================================
//
// SHEET LAYOUT
//   Row-data sheets (one row = one record):
//     - BOQ Database, Project Database, Employee Database, Supplier Database
//   Transactional / log sheets (one row = one event):
//     - MRF Pending Queue, MRF Submission Logs, Inventory Logs, Supplier Price-in,
//       PO Generation Queue, Accounting Queue, Receiving Logs, BOQ Logs,
//       Justification Logs, Payment Logs, Expense Logs, Expense Activity Logs
//   Dropdown source (ONE sheet for all picklist values):
//     - Dropdowns  (Category | Value | Meta | Active | Sort Order)
//
// To add a new dropdown anywhere in the app, append rows to "Dropdowns" with a
// fresh Category name and read them back via getDropdownValues("YourCategory").
// =============================================================================

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetSchemas = {
    // --- Row-data sheets ---
    // BOQ Database is 16 columns wide. Columns M–P (Source / Drive Link / Uploader /
    // Date Uploaded) were originally written only by Accounting's "Add Item" flow;
    // they are now also populated by the BOQ Upload Portal ingestion path. Keeping
    // the header in sync prevents column drift in downstream readers like
    // getBOQDataForProject().
    "BOQ Database":         ["Project Title", "Phase", "Scope", "Sub Scope", "Sub-Sub Scope", "Item Description", "Unit", "Qty", "Reserved 1", "Reserved 2", "Labor Cost", "Material Cost", "Source", "Drive Link", "Uploader", "Date Uploaded"],
    // Project Database gained a "Total Contract Price" column when the BOQ Upload
    // Portal was integrated — the value is captured at upload time and used
    // downstream by Payments/Supplier-Payments tabs.
    // Cols I/J/K (Total Allocated Fund / Total Expenses / Remaining Balance) back
    // the Petty Cash module: the fund now lives here instead of a separate
    // "PettyCash Projects" sheet, so the project name isn't duplicated. Allocated
    // (I) is the stored input (bumped by replenishments); Expenses (J) and Balance
    // (K) are written back by loadPettyCashProjects_ from the PettyCash Expenses ledger.
    "Project Database":     ["Project Title", "Owner", "Address", "Date", "Bidder", "Source Link", "Date Uploaded", "Total Contract Price", "Total Allocated Fund", "Total Expenses", "Remaining Balance"],
    "Employee Database":    ["Name", "Email", "Position", "Assigned Projects", "Password", "Salt"],
    "Supplier Database":    ["Company Name", "Nature of Company", "Email", "Viber Number", "Contact Person"],

    // Per-project payment-term breakdown captured at BOQ upload time
    // (Milestone % / Payment % / Uploaded At / Uploader).
    "Payment Terms Database": ["Project Title", "Milestone %", "Payment %", "Date Uploaded", "Uploader"],

    // --- Transactional / log sheets ---
    // First column is a Sheets checkbox column used by the approver queue UI
    // (submitMRF writes `false` into it; processBatchApproval ignores it). Header
    // is left blank so the row reads cleanly when exported, but the column slot
    // still exists.
    "MRF Pending Queue":    ["", "Timestamp", "ID", "Project", "Requestor", "Phase", "Item", "Qty", "Unit", "Remarks", "Status", "Document Type"],
    "MRF Submission Logs":  ["Date Req", "MRF Number", "Project Code", "Requestor", "Phase", "Item Name", "Scope", "Sub Scope", "Sub-Sub Scope", "Request Qty", "Unit", "Req Remarks", "Current Status", "Approver", "Approved Qty", "Appr Remarks", "Date of Decision", "Photo Link", "Document Link", "Requestor Email", "Document Type"],
    "Inventory Logs":       ["Date", "Project Title", "Phase", "Item Name", "Scope", "Sub Scope", "Sub-Sub Scope", "Qty Out", "Unit"],
    "Supplier Price-in":    ["Timestamp", "MRF Number", "Item Name", "Supplier", "Price", "Encoded By", "Payment Terms", "Payment Date"],
    "PO Generation Queue":  ["Timestamp", "MRF Number", "Project Code", "Requestor", "Phase", "Item Name", "Approved Qty", "Unit", "Remarks", "Supplier", "Price", "Payment Terms", "Payment Date"],
    "Accounting Queue":     ["Timestamp", "MRF Number", "Project", "Supplier", "PO Document Link", "Status"],
    "Receiving Logs":       ["Timestamp", "MRF ID", "Item Name", "Received Qty", "Remarks"],
    "BOQ Logs":             ["Timestamp", "Added By", "Action", "Project", "Phase", "Scope", "Sub Scope", "Sub-Sub Scope", "Item Description", "Unit", "Qty", "Mat Cost", "Lab Cost", "Reason"],
    "Justification Logs":   ["Timestamp", "MRF ID", "Project", "Requestor", "Approver Email", "Item Name", "Question / Note", "Requestor Reply", "Status"],
    "Payment Logs":         ["Purchase Order", "Payment Term", "%", "Supplier", "Invoiced Amount", "Payment Due Date", "Bank", "Check Number", "Payment Amount"],
    "Expense Logs":         ["Timestamp", "Expense ID", "Category", "Project Name", "Company Name", "Expense Type", "Description", "Amount", "Status", "Submitted By", "Submitter Email", "Notes", "Paid By", "Paid Date", "Payment Method", "Receipt Link"],
    "Expense Activity Logs":["Timestamp", "Expense ID", "Action", "Performed By", "Old Status", "New Status", "Notes"],

    // --- Petty Cash module sheets (ported from Finance Portal) ---
    //   The per-project fund now lives in the "Project Database" sheet (cols I/J/K
    //   above) — there is no separate "PettyCash Projects" sheet anymore, so the
    //   project name isn't duplicated across two sheets.
    //   PettyCash Expenses     = individual petty cash spends; cap of ₱5,000 per submission enforced server-side.
    //   PettyCash Replenishments = requests from employees (Pending → Approved/Denied)
    //                              and direct logs from Accounting (Approved on create).
    "PettyCash Expenses":        ["Timestamp", "Doc Ref", "User", "Project", "Line Item", "Amount", "Balance After", "Receipt URL"],
    "PettyCash Replenishments":  ["Timestamp", "Req ID", "Requestor Name", "Requestor Email", "Project ID", "Project Name", "Amount", "Status", "Receipt URL"],

    // --- Single source for ALL dropdown values ---
    "Dropdowns":            ["Category", "Value", "Meta", "Active", "Sort Order"]
  };

  // Default dropdown rows seeded only when the Dropdowns sheet is newly created.
  // Meta is interpreted per Category: for "Expense Type", it's the comma-separated
  // list of expense Categories (Project / Office) the type is valid for; for other
  // categories it's currently unused but reserved for future use.
  const defaultDropdowns = [
    // Banks
    ["Bank", "BDO",           "", "Yes", 1],
    ["Bank", "BPI",           "", "Yes", 2],
    ["Bank", "Metrobank",     "", "Yes", 3],
    ["Bank", "Landbank",      "", "Yes", 4],
    ["Bank", "UnionBank",     "", "Yes", 5],
    ["Bank", "Security Bank", "", "Yes", 6],
    ["Bank", "PNB",           "", "Yes", 7],
    ["Bank", "RCBC",          "", "Yes", 8],
    ["Bank", "China Bank",    "", "Yes", 9],
    ["Bank", "EastWest Bank", "", "Yes", 10],

    // Expense Types (Meta = allowed expense categories)
    ["Expense Type", "Labor",          "Project,Office", "Yes", 1],
    ["Expense Type", "Material",       "Project",        "Yes", 2],
    ["Expense Type", "Equipment",      "Project,Office", "Yes", 3],
    ["Expense Type", "General Office", "Office",         "Yes", 4],
    ["Expense Type", "Utilities",      "Project,Office", "Yes", 5],
    ["Expense Type", "Transportation", "Project,Office", "Yes", 6],
    ["Expense Type", "Meals",          "Project,Office", "Yes", 7],
    ["Expense Type", "Supplies",       "Project,Office", "Yes", 8],
    ["Expense Type", "Other",          "Project,Office", "Yes", 9],

    // Expense Statuses (suggested values shown in dropdowns; Accounting can also enter custom ones)
    ["Expense Status", "Pending",        "", "Yes", 1],
    ["Expense Status", "Approved",       "", "Yes", 2],
    ["Expense Status", "Paid",           "", "Yes", 3],
    ["Expense Status", "Partially Paid", "", "Yes", 4],
    ["Expense Status", "On Hold",        "", "Yes", 5],
    ["Expense Status", "Cancelled",      "", "Yes", 6],
    ["Expense Status", "Rejected",       "", "Yes", 7],

    // Payment Methods (used by Expense Logging — Mark Paid modal)
    ["Payment Method", "Cash",          "", "Yes", 1],
    ["Payment Method", "Check",         "", "Yes", 2],
    ["Payment Method", "Bank Transfer", "", "Yes", 3],
    ["Payment Method", "Credit Card",   "", "Yes", 4],
    ["Payment Method", "Petty Cash",    "", "Yes", 5],
    ["Payment Method", "Other",         "", "Yes", 6]
  ];

  // --- Create any missing sheet, or heal the header of an existing one ---
  let createdCount = 0;
  let healedCols = 0;
  const mismatches = [];
  Object.keys(sheetSchemas).forEach(function (sheetName) {
    const headers = sheetSchemas[sheetName];
    const existing = ss.getSheetByName(sheetName);

    if (existing) {
      // Sheet is already there — bring its header row up to schema without
      // disturbing any data. Reports drift instead of clobbering renames.
      healedCols += healSheetHeaders_(existing, headers, sheetName, mismatches);
      return;
    }

    const sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    createdCount++;

    // Seed Dropdowns on first creation so the app has working picklists immediately.
    if (sheetName === "Dropdowns") {
      sheet.getRange(2, 1, defaultDropdowns.length, defaultDropdowns[0].length).setValues(defaultDropdowns);
    }
  });

  // --- One-time migration: if legacy "Bank Database" / "Expense Types" sheets exist
  //     and the user hasn't yet populated Dropdowns with those categories, copy the
  //     legacy rows over so nothing is lost. We leave the legacy sheets in place;
  //     the user can delete them manually when they're confident the migration is good.
  try { migrateLegacyDropdownSheets_(ss); } catch (e) { Logger.log("Migration skipped: " + e); }

  let msg = "Setup complete.\n" +
            "Sheets created this run: " + createdCount + "\n" +
            "Header cells added to existing sheets: " + healedCols;
  if (mismatches.length) {
    msg += "\n\nHeader mismatches left UNTOUCHED (review/fix manually if these are drift, " +
           "or ignore if they're intentional renames):\n- " + mismatches.join("\n- ");
  }
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg); // No UI context (run from script editor / trigger) — log instead.
  }
}

// Bring an EXISTING sheet's header row up to the expected schema WITHOUT touching
// data: widen the grid if needed, fill blank header cells, and add headers for
// columns beyond the current width. A header cell that is already filled but
// DIFFERENT from the expected text is left alone and recorded in `mismatches` —
// we never clobber what might be a deliberate rename. Schema headers that are
// intentionally blank (e.g. the approver-queue checkbox column) are skipped.
// Returns the number of header cells written.
function healSheetHeaders_(sheet, headers, sheetName, mismatches) {
  // Ensure the sheet has at least as many columns as the schema needs.
  const maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length) {
    sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
  }

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let written = 0;
  for (let i = 0; i < headers.length; i++) {
    const want = (headers[i] === null || headers[i] === undefined) ? "" : headers[i].toString().trim();
    if (want === "") continue; // schema intentionally leaves this header blank

    const have = (current[i] === null || current[i] === undefined) ? "" : current[i].toString().trim();
    if (have === "") {
      // Blank slot — a missing trailing column or an empty header cell. Fill it.
      sheet.getRange(1, i + 1).setValue(headers[i]).setFontWeight("bold");
      written++;
    } else if (have.toLowerCase() !== want.toLowerCase()) {
      // Filled but different — possible drift or a deliberate rename. Don't touch.
      mismatches.push(sheetName + " col " + columnLetter_(i + 1) + ': has "' + have + '", expected "' + want + '"');
    }
  }
  if (written > 0) sheet.setFrozenRows(1);
  return written;
}

// Column number -> spreadsheet letter (1 -> "A", 27 -> "AA"). For readable
// drift messages.
function columnLetter_(col) {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - m) / 26);
  }
  return s;
}

// Copy rows from legacy "Bank Database" and "Expense Types" sheets into the new
// unified "Dropdowns" sheet. Only adds rows whose Value isn't already present
// under that Category, so re-running setup is safe.
function migrateLegacyDropdownSheets_(ss) {
  const dropdownsSheet = ss.getSheetByName("Dropdowns");
  if (!dropdownsSheet) return;

  // Build a {Category|Value -> true} index of what's already in Dropdowns.
  const existingKey = {};
  if (dropdownsSheet.getLastRow() >= 2) {
    const cur = dropdownsSheet.getRange(2, 1, dropdownsSheet.getLastRow() - 1, 2).getValues();
    cur.forEach(function (r) {
      const cat = (r[0] || "").toString().trim().toLowerCase();
      const val = (r[1] || "").toString().trim().toLowerCase();
      if (cat && val) existingKey[cat + "|" + val] = true;
    });
  }

  const rowsToAppend = [];

  // Legacy Bank Database -> Category=Bank
  const banksSheet = ss.getSheetByName("Bank Database");
  if (banksSheet && banksSheet.getLastRow() >= 2) {
    const data = banksSheet.getRange(2, 1, banksSheet.getLastRow() - 1, 3).getValues();
    data.forEach(function (r) {
      const name = (r[0] || "").toString().trim();
      if (!name) return;
      const key = "bank|" + name.toLowerCase();
      if (existingKey[key]) return;
      const active = (r[1] === "" || r[1] === undefined || r[1] === null) ? "Yes" : r[1];
      const sort = (r[2] === "" || r[2] === undefined || r[2] === null) ? rowsToAppend.length + 100 : r[2];
      rowsToAppend.push(["Bank", name, "", active, sort]);
      existingKey[key] = true;
    });
  }

  // Legacy Expense Types -> Category=Expense Type, Meta=Allowed Categories
  const typesSheet = ss.getSheetByName("Expense Types");
  if (typesSheet && typesSheet.getLastRow() >= 2) {
    const data = typesSheet.getRange(2, 1, typesSheet.getLastRow() - 1, 4).getValues();
    data.forEach(function (r) {
      const name = (r[0] || "").toString().trim();
      if (!name) return;
      const key = "expense type|" + name.toLowerCase();
      if (existingKey[key]) return;
      const meta = (r[1] || "").toString().trim();
      const active = (r[2] === "" || r[2] === undefined || r[2] === null) ? "Yes" : r[2];
      const sort = (r[3] === "" || r[3] === undefined || r[3] === null) ? rowsToAppend.length + 100 : r[3];
      rowsToAppend.push(["Expense Type", name, meta, active, sort]);
      existingKey[key] = true;
    });
  }

  if (rowsToAppend.length) {
    dropdownsSheet.getRange(dropdownsSheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
  }
}
