const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  BOQ: "BOQ Database",
  EMPLOYEES: "Employee Database",
  QUEUE: "MRF Pending Queue",
  LOGS: "MRF Submission Logs",
  INVENTORY: "Inventory Logs",
  PROJECTS: "Project Database",
  SUPPLIER_PRICE_IN: "Supplier Price-in",
  PO_QUEUE: "PO Generation Queue",
  SUPPLIERS: "Supplier Database",
  ACCOUNTING: "Accounting Queue",
  RECEIVING_LOGS: "Receiving Logs",
  BOQ_LOGS: "BOQ Logs",
  JUSTIFICATIONS: "Justification Logs",
  PAYMENT_LOGS: "Payment Logs",
  PAYMENT_TERMS: "Payment Terms Database",
  EXPENSE_LOGS: "Expense Logs",
  EXPENSE_ACTIVITY: "Expense Activity Logs",
  // Client Payments (collections / accounts receivable) and its audit trail.
  CLIENT_PAYMENTS: "Client Payments",
  CLIENT_PAYMENT_LOGS: "Client Payment Logs",
  // Petty Cash module. Per-project allocations now live in Project Database
  // cols I/J/K (see setup.js); the projectId carried by petty cash payloads is
  // the Project Database row index — i.e. row 2 = projectId 2 — so do not
  // reorder Project Database rows. Two ledger sheets remain:
  PETTY_CASH_EXPENSES:       "PettyCash Expenses",
  PETTY_CASH_REPLENISHMENTS: "PettyCash Replenishments",
  // Single sheet that backs every dropdown in the app. Columns:
  //   Category | Value | Meta | Active | Sort Order
  // To add a new picklist, add rows with a new Category and read them back via
  // getDropdownValues("YourCategory"). See setup.gs for the default seed list.
  DROPDOWNS: "Dropdowns"
};

// Drive folder IDs. Both values can be overridden per-deployment via Script
// Properties (Project Settings → Script Properties in the Apps Script editor):
//   - MRF_ROOT_FOLDER_ID   → root for all auto-generated docs (MRFs, POs, etc.)
//   - MRF_LOGOS_FOLDER_ID  → folder containing logo1.png / logo2.png
// The constants below are the original defaults and stay in effect when no
// override is set, so existing deployments don't break.
const FOLDERS = {
  LOGOS: "1y50Y2Cnivhgox-5PI3D2D7dLlBC4a7tz",
  ROOT:  "1Aju0oGNCz2-U1zDuPaMJCcVtSUIzP3fY"
};

function getRootFolderId_() {
  try {
    const override = PropertiesService.getScriptProperties().getProperty("MRF_ROOT_FOLDER_ID");
    return (override && override.trim()) ? override.trim() : FOLDERS.ROOT;
  } catch (e) {
    return FOLDERS.ROOT;
  }
}

function getLogosFolderId_() {
  try {
    const override = PropertiesService.getScriptProperties().getProperty("MRF_LOGOS_FOLDER_ID");
    return (override && override.trim()) ? override.trim() : FOLDERS.LOGOS;
  } catch (e) {
    return FOLDERS.LOGOS;
  }
}

// ==========================================
// --- AUTHENTICATION & LOGIN LOGIC ---
// ==========================================

function loginUser(identifier, password) {
  try {
    const data = SS.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
    const input = identifier.toString().trim().toLowerCase();
    
    // Find user (ignoring header)
    const rowIndex = data.findIndex((row, idx) => idx > 0 && (
      (row[0] && row[0].toString().trim().toLowerCase() === input) || 
      (row[1] && row[1].toString().trim().toLowerCase() === input)
    ));

    if (rowIndex === -1) return { error: "User account not found." };

    const name = data[rowIndex][0];
    const email = data[rowIndex][1];
    const role = data[rowIndex][2] ? data[rowIndex][2].toString().toLowerCase().trim() : "employee";
    const assignedProjects = data[rowIndex][3] ? data[rowIndex][3].toString().trim() : "";
    const storedPass = data[rowIndex][4];
    const salt = data[rowIndex][5];
    
    const isNewUser = (!salt || salt.toString().trim() === "");
    const storedPassStr = (storedPass === null || storedPass === undefined) ? "" : storedPass.toString();
    const submittedPass = (password === null || password === undefined) ? "" : password.toString();

    // "TMP:" sentinel: written by recoverPassword. Validate the hash against
    // the substring, but force isNew=true so the UI shows the set-password
    // modal. finalizePassword overwrites with a normal hash on first login.
    if (storedPassStr.indexOf("TMP:") === 0 && !isNewUser) {
      const expected = storedPassStr.substring(4);
      if (hashPassword(submittedPass, salt) === expected) {
        return { success: true, name, email, role, assignedProjects, isNew: true };
      }
      return { error: "Incorrect password." };
    }

    if (isNewUser) {
      // Require BOTH a non-empty stored value AND an exact match. Previously a
      // blank stored password accepted any submitted password — closed by the
      // explicit storedPassStr length guard.
      if (storedPassStr.length > 0 && submittedPass === storedPassStr) {
        return { success: true, name, email, role, assignedProjects, isNew: true };
      }
    } else {
      if (hashPassword(submittedPass, salt) === storedPassStr) {
        return { success: true, name, email, role, assignedProjects, isNew: false };
      }
    }
    return { error: "Incorrect password." };
  } catch (e) { return { error: e.toString() }; }
}

function finalizePassword(emailOrName, newPassword) {
  const sheet = SS.getSheetByName(SHEETS.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  const input = emailOrName.toString().trim().toLowerCase();
  
  const rowIndex = data.findIndex((r, idx) => idx > 0 && (
    r[0].toString().trim().toLowerCase() === input || r[1].toString().trim().toLowerCase() === input
  ));
  
  if (rowIndex === -1) return "User verification failed.";

  const salt = Utilities.getUuid(); 
  const hashedPassword = hashPassword(newPassword, salt);
  
  sheet.getRange(rowIndex + 1, 5, 1, 2).setValues([[hashedPassword, salt]]);
  return "SUCCESS";
}

function recoverPassword(identifier) {
  try {
    const sheet = SS.getSheetByName(SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    const input = identifier.toString().trim().toLowerCase();

    const rowIndex = data.findIndex((row, idx) => idx > 0 && (
      (row[0] && row[0].toString().trim().toLowerCase() === input) || 
      (row[1] && row[1].toString().trim().toLowerCase() === input)
    ));

    if (rowIndex === -1) return "❌ Name or Email not found.";

    const fullName = data[rowIndex][0];
    const email = data[rowIndex][1];
    const tempPass = Math.random().toString(36).slice(-8).toUpperCase();

    // Try to send the email FIRST. Only write the reset to the sheet if the
    // notification actually went out — otherwise we'd lock the user out and
    // leak the temp password back to whoever called the endpoint.
    let sent = false;
    try {
      if (email) {
        GmailApp.sendEmail(
          email,
          "MRF System - Password Reset",
          `Hello ${fullName},\n\nYour password has been reset.\n\nTemporary Password: ${tempPass}\n\nPlease log in and set a new password.`
        );
        sent = true;
      }
    } catch (mailErr) {
      console.error("Password reset email failed: " + mailErr.toString());
    }

    if (!sent) {
      return "⚠️ Could not send the reset email. Please contact an administrator.";
    }

    // Store the temp password hashed with a fresh salt and a "TMP:" prefix.
    // The prefix tells loginUser to accept the value and treat the session as
    // isNew (so the new-password modal still appears on first login) — and
    // finalizePassword() overwrites it with a normal hash. This avoids
    // leaving the sheet cell in cleartext, which previously combined with
    // the open web app to expose temp passwords in the sheet.
    const salt = Utilities.getUuid();
    const hashed = hashPassword(tempPass, salt);
    sheet.getRange(rowIndex + 1, 5, 1, 2).setValues([["TMP:" + hashed, salt]]);
    return "✅ SUCCESS: A temporary password has been sent to your registered email address.";
  } catch (e) { return "❌ Error: " + e.toString(); }
}

function hashPassword(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

// ==========================================
// --- CORE UTILITIES & ROUTING ---
// ==========================================

function getFileDateSuffix() {
  const d = new Date();
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return `${y}${m}${day}`;
}

function safeDriveAction(action) {
  for (let i = 0; i < 3; i++) {
    try { return action(); } catch (e) {
      if (i === 2) throw e;
      Utilities.sleep(1000);
    }
  }
}

// Serialize critical sections that read-modify-write the spreadsheet so two
// concurrent users (e.g., two approvers acting at the same time) can't clobber
// each other. Apps Script LockService is per-script execution; we use the
// document-scoped lock keyed to this spreadsheet.
//
// The wrapped function returns whatever it would normally return; if the lock
// can't be acquired within `timeoutMs`, we surface the same "Error: …" string
// shape every batch endpoint already produces so the UI's `indexOf('Error')`
// check still works.
function withDocumentLock_(fn, timeoutMs) {
  const lock = LockService.getDocumentLock();
  const wait = timeoutMs || 15000;
  try {
    if (!lock.tryLock(wait)) {
      return "Error: Another operation is in progress. Please retry in a moment.";
    }
  } catch (e) {
    return "Error: Lock acquisition failed (" + e.message + ").";
  }
  try {
    const out = fn();
    return out;
  } catch (e) {
    // Convert any uncaught throw inside the critical section into the same
    // "Error: ..." string shape callers already match against with
    // .indexOf('Error'). Without this, callers received `undefined` and a
    // ".indexOf is not a function" stack trace.
    return "Error: " + (e && e.message ? e.message : e);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getOrCreateDynamicFolder(projectName, phaseName, typeName) {
  const cache = CacheService.getScriptCache();
  const pathKey = `FLD_${projectName}_${phaseName}_${typeName}`.replace(/[^a-zA-Z0-9_]/g, '');
  
  const cachedId = cache.get(pathKey);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch(e) { cache.remove(pathKey); }
  }

  const MAIN_ROOT_ID = getRootFolderId_();
  let parentFolder = DriveApp.getFolderById(MAIN_ROOT_ID);

  const pathNames = [
    projectName ? projectName.toString().trim() : "General",
    phaseName ? phaseName.toString().trim() : "General",
    typeName ? typeName.toString().trim() : "General"
  ];

  let currentPath = "ROOT";
  for (let i = 0; i < pathNames.length; i++) {
    let name = pathNames[i];
    currentPath += "_" + name.replace(/[^a-zA-Z0-9_]/g, '');
    let intermediateId = cache.get(currentPath);
    let found = false;
    
    if (intermediateId) {
        try { 
            parentFolder = DriveApp.getFolderById(intermediateId); 
            found = true;
        } catch(e) { cache.remove(currentPath); }
    }
    
    if (!found) {
        safeDriveAction(() => {
            let folders = parentFolder.getFoldersByName(name);
            if (folders.hasNext()) parentFolder = folders.next();
            else parentFolder = parentFolder.createFolder(name);
        });
        cache.put(currentPath, parentFolder.getId(), 21600);
    }
  }
  
  cache.put(pathKey, parentFolder.getId(), 21600);
  return parentFolder;
}

function getFileIdFromUrl(url) {
  if (!url) return null;
  try {
    if (url.includes('id=')) return url.split('id=')[1].split('&')[0];
    if (url.includes('/d/')) return url.split('/d/')[1].split('/')[0];
  } catch(e) {}
  return null;
}

function formatFullDate(date) {
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

function formatCurrency(num) {
  return parseFloat(num).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function getLogoDisplayUrl(fileName) {
  const props = PropertiesService.getScriptProperties();
  const propKey = "LOGO_" + fileName.replace(/[^a-zA-Z0-9_]/g, '');
  const savedUrl = props.getProperty(propKey);
  if (savedUrl) return savedUrl; 
  try {
    const folder = DriveApp.getFolderById(getLogosFolderId_());
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";
      props.setProperty(propKey, url); 
      return url;
    }
  } catch (e) { return ""; }
  return "";
}

function getLogoBase64(fileName) {
  const cache = CacheService.getScriptCache();
  const key = "B64_" + fileName.replace(/[^a-zA-Z0-9_]/g, '');
  let b64 = cache.get(key);
  if (b64) return b64;
  try {
    const folder = DriveApp.getFolderById(getLogosFolderId_());
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      b64 = Utilities.base64Encode(files.next().getBlob().getBytes());
      cache.put(key, b64, 21600); 
      return b64;
    }
  } catch(e) {}
  return "";
}

function doGet(e) {
  if (e.parameter.receiveToken) {
    // Validate the token before rendering. Strip anything but [a-zA-Z0-9-] so
    // a malicious token can't smuggle characters into the HTML template (the
    // <?= ?> scriptlet already HTML-escapes, but we keep this tight anyway).
    const rawToken = (e.parameter.receiveToken || "").toString().trim();
    const safeToken = rawToken.replace(/[^A-Za-z0-9\-]/g, "");
    const mappedMrfId = resolveReceiveToken_(safeToken);
    let template = HtmlService.createTemplateFromFile('Receive');
    template.receiveToken = mappedMrfId ? safeToken : "";
    template.mrfReference = mappedMrfId || "";
    template.tokenValid = !!mappedMrfId;
    template.logo1Url = getLogoDisplayUrl("logo1.png");
    template.logo2Url = getLogoDisplayUrl("logo2.png");
    return template.evaluate()
      .setTitle('Receiving Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ONLY serve Index.html. All other views are fetched securely after login.
  let template = HtmlService.createTemplateFromFile('Index');
  template.logo1Url = getLogoDisplayUrl("logo1.png");
  template.logo2Url = getLogoDisplayUrl("logo2.png");
  template.mode = 'request';
  template.scriptUrl = ScriptApp.getService().getUrl(); 
  
  return template.evaluate()
    .setTitle('Construction MRF System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- SECURE ACCOUNT-BASED RENDERING ---
function getSecurePortalHtml(email) {
  try {
    const data = SS.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
    const input = email.toString().trim().toLowerCase();
    
    // Find user by email directly from DB. Guard against blank email cells —
    // a single missing col B used to crash this whole render.
    const userRow = data.find((row, idx) => idx > 0 && row[1] && row[1].toString().trim().toLowerCase() === input);
    if (!userRow) return "";
    
    const role = (userRow[2] || "").toString().toLowerCase().trim();
    let template;
    
    if (role.includes('finance') || role.includes('accounting')) {
      template = HtmlService.createTemplateFromFile('Accounting');
    } else if (role.includes('encoder') || role.includes('purchasing')) {
      template = HtmlService.createTemplateFromFile('Encode');
    } else {
      return ""; // Unauthorized role
    }
    
    template.logo1Url = getLogoDisplayUrl("logo1.png");
    template.logo2Url = getLogoDisplayUrl("logo2.png");
    // Securely pass down script Url to sub-templates to match logout fix mechanism
    template.scriptUrl = ScriptApp.getService().getUrl();
    // Pass logged-in user identity so the sub-portal can render the user bar + logout
    // server-side (no fragile document.write injection).
    template.currentUserName = (userRow[0] || "").toString();
    template.currentUserRole = (userRow[2] || "").toString();
    return template.evaluate().getContent();
  } catch(e) {
    return "";
  }
}

// ==========================================
// --- DATA LOADING & FILTERING ---
// ==========================================

function getInitialUIData(userIdentifier) {
  const data = {
    projects: [],
    pendingJustifications: [],
    boqMap: {} // Restored your lightning-fast cache approach
  };

  // 1. Authenticate and verify assigned projects
  const empSheet = SS.getSheetByName(SHEETS.EMPLOYEES);
  let allowedProjects = [];
  let userName = "";

  if (empSheet && userIdentifier) {
    const empData = empSheet.getDataRange().getValues();
    const input = userIdentifier.toString().trim().toLowerCase();
    
    for (let i = 1; i < empData.length; i++) {
      const nameCell  = empData[i][0] ? empData[i][0].toString().trim().toLowerCase() : "";
      const emailCell = empData[i][1] ? empData[i][1].toString().trim().toLowerCase() : "";
      if (nameCell === input || emailCell === input) {
         userName = empData[i][0];
         let pStr = empData[i][3] ? empData[i][3].toString().trim() : "";
         if (pStr !== "") allowedProjects = pStr.split(/[,;\n]/).map(s => s.trim()).filter(s => s);
         break;
      }
    }
  }

  // 2. Build boqMap exactly like the old code for instantaneous dropdowns
  const projSheet = SS.getSheetByName(SHEETS.PROJECTS);
  const boqSheet = SS.getSheetByName(SHEETS.BOQ);
  const projects = new Set();

  if (projSheet) {
    const pData = projSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
       let val = pData[i][0] ? pData[i][0].toString().trim() : "";
       if (val && val.toLowerCase() !== "project name" && val.toLowerCase() !== "project title") {
           projects.add(val);
       }
    }
  }

  if (boqSheet) {
    const bData = boqSheet.getDataRange().getValues();
    for (let i = 2; i < bData.length; i++) {
      const p = bData[i][0] ? bData[i][0].toString().trim() : "";
      const ph = bData[i][1] ? bData[i][1].toString().trim() : "";
      const sc = bData[i][2] ? bData[i][2].toString().trim() : "";

      if (p) {
        projects.add(p);
        if (!data.boqMap[p]) data.boqMap[p] = {};
        if (ph) {
          if (!data.boqMap[p][ph]) data.boqMap[p][ph] = new Set(["Consumables"]);
          if (sc) data.boqMap[p][ph].add(sc);
        }
      }
    }
  }

  // Filter projects based on employee's allowed list (case/whitespace-insensitive)
  if (allowedProjects.length > 0) {
     const allowedSet = new Set(allowedProjects.map(s => s.toLowerCase()));
     data.projects = [...projects].filter(p => allowedSet.has(p.toLowerCase())).sort();
  } else {
     data.projects = [...projects].sort();
  }

  // Convert map sets back into clean arrays
  for (let p in data.boqMap) {
    for (let ph in data.boqMap[p]) {
      data.boqMap[p][ph] = [...data.boqMap[p][ph]].sort();
    }
  }

  // 3. Fetch Pending Justifications for this Requestor
  const justSheet = SS.getSheetByName(SHEETS.JUSTIFICATIONS);
  if (justSheet && userName) {
      const jData = justSheet.getDataRange().getValues();
      for (let j = 1; j < jData.length; j++) {
         const reqCell    = jData[j][3] ? jData[j][3].toString().trim() : "";
         const statusCell = jData[j][8] ? jData[j][8].toString().trim() : "";
         if (reqCell === userName && statusCell === "Pending Reply") {
            data.pendingJustifications.push({
               date: formatFullDate(new Date(jData[j][0])),
               mrfId: jData[j][1],
               project: jData[j][2],
               itemName: jData[j][5],
               question: jData[j][6]
            });
         }
      }
  }

  return data;
}

// These two functions remain untouched in case they are used elsewhere, 
// but the frontend now uses the BOQ Map locally.
function getPhasesForProject(proj) {
  const sheet = SS.getSheetByName(SHEETS.BOQ);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const phases = new Set();
  const sProj = proj.toString().toLowerCase().trim();
  for (let i = 2; i < data.length; i++) { 
    if (data[i][0] && data[i][0].toString().toLowerCase().trim() === sProj) {
      if (data[i][1]) phases.add(data[i][1].toString().trim());
    }
  }
  return [...phases].sort();
}

function getScopesForPhase(proj, phase) {
  const sheet = SS.getSheetByName(SHEETS.BOQ);
  let scopes = new Set();
  scopes.add("Consumables");
  if (!sheet) return [...scopes].sort();
  const data = sheet.getDataRange().getValues();
  const sProj = proj.toString().toLowerCase().trim();
  const sPhase = phase.toString().toLowerCase().trim();
  
  for (let i = 2; i < data.length; i++) { 
    if (data[i][0] && data[i][0].toString().toLowerCase().trim() === sProj &&
        data[i][1] && data[i][1].toString().toLowerCase().trim() === sPhase) {
      if (data[i][2]) scopes.add(data[i][2].toString().trim());
    }
  }
  return [...scopes].sort();
}

function getItemsForScope(proj, phase, scope) {
  const boqSheet = SS.getSheetByName(SHEETS.BOQ);
  const invSheet = SS.getSheetByName(SHEETS.INVENTORY);
  if (!boqSheet) return [];
  
  const boqData = boqSheet.getDataRange().getValues();
  const invData = invSheet.getDataRange().getValues();

  const sProject = proj.toString().trim().toLowerCase();
  const sPhase = phase.toString().trim().toLowerCase();
  const sScope = scope.toString().trim().toLowerCase();

  let items = {};

  function getCost(val) {
    if (val === "" || val === null || val === undefined || val.toString().toLowerCase().includes("not included")) return "No Data";
    let num = parseFloat(val);
    return isNaN(num) ? "No Data" : num;
  }

  for (let i = 2; i < boqData.length; i++) { 
    if (boqData[i][0] && boqData[i][0].toString().trim().toLowerCase() === sProject && 
        boqData[i][1] && boqData[i][1].toString().trim().toLowerCase() === sPhase &&
        boqData[i][2] && boqData[i][2].toString().trim().toLowerCase() === sScope) {
      
      const subScope = boqData[i][3];
      const subSubScope = boqData[i][4];
      const itemName = boqData[i][5];
      const unit = boqData[i][6];
      const allocated = parseFloat(boqData[i][7]) || 0;
      
      let labCost = getCost(boqData[i][10]);
      let matCost = getCost(boqData[i][11]);
      let totCost = "No Data";

      if (labCost !== "No Data" || matCost !== "No Data") {
        totCost = (labCost !== "No Data" ? labCost : 0) + (matCost !== "No Data" ? matCost : 0);
      }
      
      const key = `${subScope}-${subSubScope}-${itemName}`;

      if (!items[key]) {
        items[key] = {
          scope: boqData[i][2],
          subScope: subScope,
          subSubScope: subSubScope,
          itemName: itemName,
          unit: unit,
          allocated: allocated, 
          matCost: matCost,
          labCost: labCost,
          totCost: totCost
        };
      } else {
        items[key].allocated += allocated;
        if (typeof items[key].matCost === 'number' && typeof matCost === 'number') items[key].matCost += matCost;
        if (typeof items[key].labCost === 'number' && typeof labCost === 'number') items[key].labCost += labCost;
        if (typeof items[key].totCost === 'number' && typeof totCost === 'number') items[key].totCost += totCost;
      }
    }
  }

  for (let j = 1; j < invData.length; j++) {
    if (invData[j][1] && invData[j][1].toString().trim().toLowerCase() === sProject && 
        invData[j][2] && invData[j][2].toString().trim().toLowerCase() === sPhase && 
        invData[j][4] && invData[j][4].toString().trim().toLowerCase() === sScope) {

      const subScope = invData[j][5];
      const subSubScope = invData[j][6];
      const itemName = invData[j][3];
      const qtyOut = parseFloat(invData[j][7]) || 0;
      const unitUsed = invData[j][8] ? invData[j][8].toString().trim() : "";
      const key = `${subScope}-${subSubScope}-${itemName}`;

      if (items[key]) {
        if (unitUsed === "Labor Budget" && typeof items[key].labCost === 'number') items[key].labCost -= qtyOut;
        else if (unitUsed === "Material Budget" && typeof items[key].matCost === 'number') items[key].matCost -= qtyOut;
        else if (unitUsed === "Total Budget" && typeof items[key].totCost === 'number') items[key].totCost -= qtyOut;
        else items[key].allocated -= qtyOut; 
      }
    }
  }

  // Coerce to string before localeCompare — sheet cells like a numeric
  // sub-scope code (e.g., 1.1) return as Number from getValues() and don't
  // have a .localeCompare method.
  function _s(v) { return (v === null || v === undefined) ? "" : v.toString(); }
  return Object.values(items).sort((a, b) => {
    if (a.subScope !== b.subScope) return _s(a.subScope).localeCompare(_s(b.subScope));
    if (a.subSubScope !== b.subSubScope) return _s(a.subSubScope).localeCompare(_s(b.subSubScope));
    return _s(a.itemName).localeCompare(_s(b.itemName));
  });
}

// ==========================================
// --- REQUEST SUBMISSION ---
// ==========================================

function submitMRF(payload, fileObj) {
  try {
    const queue = SS.getSheetByName(SHEETS.QUEUE);
    const logs = SS.getSheetByName(SHEETS.LOGS);
    
    if (!queue || !logs) throw new Error("System sheets missing.");

    const activeUserEmail = payload.email; 
    const requestId = (payload.docType === 'Additional Request Outside BOQ' ? 'ARF-' : 'MRF-') + new Date().getTime().toString().slice(-5);
    const timestamp = new Date();
    
    let photoLink = "";
    let actualAttachment = null;
    
    if (fileObj) {
      try {
        const phaseName = (payload.items && payload.items[0] && payload.items[0].phase) ? payload.items[0].phase : "General";
        const folder = getOrCreateDynamicFolder(payload.project, phaseName, "Uploaded Photos");
        
        const extension = fileObj.fileName.split('.').pop() || "pdf";
        const newFileName = `${requestId}_${getFileDateSuffix()}.${extension}`;
        
        actualAttachment = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, newFileName);
        const photoFile = safeDriveAction(() => folder.createFile(actualAttachment));
        photoLink = photoFile.getUrl();
      } catch (e) { console.error("File upload failed: " + e.toString()); }
    }

    const projectLogValue = payload.project;
    let address = "N/A";
    
    const projSheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (projSheet) {
      const projData = projSheet.getDataRange().getValues();
      for(let i = projData.length - 1; i >= 1; i--) {
        if(projData[i][0].toString().trim() === payload.project.trim()) { 
            address = projData[i][2] || "N/A"; 
            break; 
        }
      }
    }
    
    const mrfPdfData = generateDocPDF(payload, requestId, timestamp, null, address, "MATERIAL REQUEST FORM", payload.requestor);
    const mrfPdfUrl = mrfPdfData.url;
    
    let queueUpdates = [];
    let logUpdates = [];

    payload.items.forEach(item => {
      const isAdditional = payload.docType === 'Additional Request Outside BOQ';
      const status = isAdditional ? 'Pending Additional' : 'Pending';
      
      queueUpdates.push([
        false, timestamp, requestId, payload.project, payload.requestor, 
        item.phase, item.name, item.qty, item.unit, item.remarks, status, payload.docType
      ]);
      
      logUpdates.push([
        timestamp, requestId, projectLogValue, payload.requestor, item.phase, item.name, 
        item.scope || "", item.subScope || "", item.subSubScope || "", item.qty, item.unit, 
        item.remarks, status, "", "", "", "", photoLink, mrfPdfUrl, activeUserEmail, payload.docType
      ]);
    });

    if (queueUpdates.length > 0) {
      const queueStartRow = queue.getLastRow() + 1;
      queue.getRange(queueStartRow, 1, queueUpdates.length, queueUpdates[0].length).setValues(queueUpdates);
      queue.getRange(queueStartRow, 1, queueUpdates.length, 1).insertCheckboxes();
    }
    if (logUpdates.length > 0) {
      logs.getRange(logs.getLastRow() + 1, 1, logUpdates.length, logUpdates[0].length).setValues(logUpdates);
    }

    try {
      const empData = SS.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
      const approvers = empData.filter(row => row[2] && row[2].toString().toLowerCase().trim() === "mrf approver").map(row => row[1]);

      if (approvers.length > 0) {
        const appUrl = ScriptApp.getService().getUrl(); 
        GmailApp.sendEmail(approvers.join(","), "Pending Approval: " + requestId + " - " + payload.project, "A new request has been submitted by " + payload.requestor + ".\n\nPlease review here: " + appUrl);
      }
      
      if (activeUserEmail) {
        let reqAttachments = [mrfPdfData.blob]; 
        if (actualAttachment) reqAttachments.push(actualAttachment); 

        GmailApp.sendEmail(activeUserEmail, "MRF Copy: " + requestId + " - " + payload.project, 
          "Hello " + payload.requestor + ",\n\nAttached is your Material Request " + requestId + " for Project: " + payload.project + ".\n\nThank you.",
          { attachments: reqAttachments });
      }
    } catch (e) {}

    return { success: true, id: requestId };
  } catch (err) { 
    return { success: false, error: err.toString() }; 
  }
}

// Look up a person's job title from Employee Database col C (Position).
// Returns "" if not found, so callers can fall back gracefully. Case- and
// whitespace-insensitive match against col A (Name).
function getEmployeePosition_(name) {
  if (!name) return "";
  try {
    const sheet = SS.getSheetByName(SHEETS.EMPLOYEES);
    if (!sheet) return "";
    const data = sheet.getDataRange().getValues();
    const key = name.toString().trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const cell = (data[i][0] || "").toString().trim().toLowerCase();
      if (cell === key) return (data[i][2] || "").toString().trim();
    }
  } catch (e) {}
  return "";
}

function generateDocPDF(payload, id, timestamp, photoBlob, address, title, signatoryName, hasPrices = false, supplierName = "", isForSupplier = false) {
  let groupedItems = [];
  payload.items.forEach(it => {
    let existing = groupedItems.find(g => 
      g.name.toString().trim().toLowerCase() === it.name.toString().trim().toLowerCase() &&
      (g.unit || "").toString().trim().toLowerCase() === (it.unit || "").toString().trim().toLowerCase()
    );
    if (existing) {
      let existingQty = parseFloat(existing.finalQty !== undefined ? existing.finalQty : existing.qty) || 0;
      let newQty = parseFloat(it.finalQty !== undefined ? it.finalQty : it.qty) || 0;
      existing.qty = existingQty + newQty;
      existing.finalQty = existing.qty; 
    } else {
      groupedItems.push({...it});
    }
  });

  const costUnits = ["labor budget", "material budget", "total budget", "labor cost", "material cost", "total cost", "labor", "material", "lot"];
  
  const isAllBudget = groupedItems.length > 0 && groupedItems.every(it => 
    costUnits.includes((it.unit || "").toString().trim().toLowerCase()) || 
    (it.unit || "").toString().trim() === "-"
  );

  let grandTotal = 0;
  let headerHtml = "";
  
  const isPOorRFQ = title.includes("PURCHASE ORDER") || title.includes("QUOTATION");
  const hideUnitAndQty = isAllBudget && isPOorRFQ;
  const hideRemarks = isPOorRFQ;
  
  let qtyHeader = "QUANTITY";
  if (!hasPrices) { 
    const isCost = groupedItems.some(it => costUnits.includes((it.unit || "").trim().toLowerCase()) || (it.unit || "").toString().trim() === "-");
    if (isCost) qtyHeader = "COST";
  }

  if (hideUnitAndQty) {
    if (hasPrices) {
      headerHtml = `<tr><th style="width:75%;">PRODUCT NAME / DESCRIPTION</th><th style="width:25%;">TOTAL AMOUNT</th></tr>`;
    } else {
      if (hideRemarks) {
        headerHtml = `<tr><th style="width:100%;">PRODUCT NAME / DESCRIPTION</th></tr>`;
      } else {
        headerHtml = `<tr><th style="width:70%;">PRODUCT NAME / DESCRIPTION</th><th style="width:30%;">REMARKS</th></tr>`;
      }
    }
  } else {
    if (hasPrices) {
       headerHtml = `<tr><th style="width:10%;">${qtyHeader}</th><th style="width:10%;">UNIT</th><th style="width:45%;">PRODUCT NAME / DESCRIPTION</th><th style="width:15%;">UNIT PRICE</th><th style="width:20%;">TOTAL AMOUNT</th></tr>`;
    } else {
       if (hideRemarks) {
          headerHtml = `<tr><th style="width:15%;">${qtyHeader}</th><th style="width:15%;">UNIT</th><th style="width:70%;">PRODUCT NAME / DESCRIPTION</th></tr>`;
       } else {
          headerHtml = `<tr><th style="width:10%;">${qtyHeader}</th><th style="width:10%;">UNIT</th><th style="width:50%;">PRODUCT NAME / DESCRIPTION</th><th style="width:30%;">REMARKS</th></tr>`;
       }
    }
  }

  let rowsHtml = groupedItems.map(it => {
    const qty = parseFloat(it.finalQty) || parseFloat(it.qty) || 0;
    let displayQty = qty.toLocaleString(undefined, {maximumFractionDigits: 2});
    const price = parseFloat(it.price) || 0;
    
    let displayUnit = it.unit || '';
    if (displayUnit === 'Labor Budget') displayUnit = 'Labor';
    else if (displayUnit === 'Material Budget') displayUnit = 'Material';
    else if (displayUnit === 'Total Budget') displayUnit = 'Lot';

    const isBudget = costUnits.includes((it.unit || "").toString().trim().toLowerCase()) || (it.unit || "").toString().trim() === "-";

    let amount = qty * price;
    let displayPrice = formatCurrency(price);
    
    if (isBudget) {
        displayQty = "-"; displayUnit = "-";
        if (hasPrices) { amount = price; displayPrice = "-"; }
    }

    grandTotal += amount;
    
    if (hideUnitAndQty) {
       if (hasPrices) {
           return `<tr><td>${it.name}</td><td style="text-align:center;">${formatCurrency(amount)}</td></tr>`;
       } else {
           if (hideRemarks) {
               return `<tr><td>${it.name}</td></tr>`;
           } else {
               return `<tr><td>${it.name}</td><td>${it.remarks || ''}</td></tr>`;
           }
       }
    } else {
       if (hasPrices) {
           return `<tr><td style="text-align:center;">${displayQty}</td><td style="text-align:center;">${displayUnit}</td><td>${it.name}</td><td style="text-align:center;">${displayPrice}</td><td style="text-align:center;">${formatCurrency(amount)}</td></tr>`;
       } else {
           if (hideRemarks) {
               return `<tr><td style="text-align:center;">${displayQty}</td><td style="text-align:center;">${displayUnit}</td><td>${it.name}</td></tr>`;
           } else {
               return `<tr><td style="text-align:center;">${displayQty}</td><td style="text-align:center;">${displayUnit}</td><td>${it.name}</td><td>${it.remarks || ''}</td></tr>`;
           }
       }
    }
  }).join('');

  let footerHtml = "";
  if (hasPrices && !isForSupplier) {
    if (hideUnitAndQty) {
        footerHtml = `<tfoot><tr><td align="right" style="padding:8px;"><b>GRAND TOTAL</b></td><td align="center" style="background:#f0f0f0;color:#000000;font-weight:900;font-size:15px;"><b>${formatCurrency(grandTotal)}</b></td></tr></tfoot>`;
    } else {
        footerHtml = `<tfoot><tr><td colspan="4" align="right" style="padding:8px;"><b>GRAND TOTAL</b></td><td align="center" style="background:#f0f0f0;color:#000000;font-weight:900;font-size:15px;"><b>${formatCurrency(grandTotal)}</b></td></tr></tfoot>`;
    }
  }

  let photoSection = "";
  if (photoBlob) {
    photoSection = `<div style="margin-top:30px; border-top: 2px solid #2c3e50; padding-top: 20px;"><h3 style="color:#2c3e50; text-align:center;">PHOTO ATTACHMENT</h3><div style="text-align:center;"><img src="data:${photoBlob.getContentType()};base64,${Utilities.base64Encode(photoBlob.getBytes())}" style="max-width:450px; border: 1px solid #ddd; border-radius: 8px; padding: 5px;"/></div></div>`;
  }

  // Signature block: name on top, the person's actual Position underneath
  // (looked up from Employee Database). Falls back to a blank second line so
  // we never print the generic "Approver" / "Requestor" labels.
  const signatoryPosition = getEmployeePosition_(signatoryName);
  let signatureBlock = `<div style="margin-top: 40px; width: 250px; text-align: center; float: right;"><div style="border-bottom: 1.5px solid #000; padding-bottom: 5px; font-weight: bold; text-transform: uppercase;">${signatoryName}</div><div style="font-size: 13px; margin-top: 5px;">${signatoryPosition || "&nbsp;"}</div></div><div style="clear: both;"></div>`;

  const subjectField = payload.items[0].scope || payload.items[0].phase || "N/A";
  
  let supplierHtml = "";
  if (supplierName !== "") {
    supplierHtml = `<br><b>SUPPLIER:</b> ${supplierName}`;
    if (payload.paymentTerms || payload.paymentDate) {
        supplierHtml += `<br><br><b>PAYMENT SCHEDULE:</b>`;
        let termsArr = (payload.paymentTerms || "").split(" | ");
        let datesArr = (payload.paymentDate || "").split(" | ");
        
        supplierHtml += `<ul style="margin-top: 5px; margin-bottom: 0; padding-left: 20px; font-size: 13px;">`;
        for(let i = 0; i < Math.max(termsArr.length, datesArr.length); i++) {
            let t = termsArr[i] || "N/A";
            let d = datesArr[i] || "TBD";
            let dStr = d;
            
            if (d !== "TBD" && d !== "") {
                let pd = new Date(d);
                if (!isNaN(pd.getTime())) dStr = formatFullDate(pd);
            }
            supplierHtml += `<li>${t} — <b>Due:</b> ${dStr}</li>`;
        }
        supplierHtml += `</ul>`;
    }
  }

  let logo1Html = "";
  let logo2Html = "";
  if (title.includes("PURCHASE ORDER")) {
      const b64_1 = getLogoBase64("logo1.png");
      const b64_2 = getLogoBase64("logo2.png");
      if (b64_1) logo1Html = `<img src="data:image/png;base64,${b64_1}" style="max-height:60px; max-width:100%;"/>`;
      if (b64_2) logo2Html = `<img src="data:image/png;base64,${b64_2}" style="max-height:60px; max-width:100%;"/>`;
  }

  const html = `<html><body style="font-family:sans-serif;padding:30px;color:#2c3e50;">
    <div style="background:white; padding:10px; border-bottom:4px solid #212529; width:100%;">
      <table style="width:100%; border-collapse:collapse;">
        <tr>
          <td style="width:20%; text-align:left; vertical-align:top;">${logo1Html}</td>
          <td style="width:60%; text-align:center; vertical-align:bottom; padding-top:15px;"><h1 style="margin:0; text-transform:uppercase; color:#212529;">${title}</h1></td>
          <td style="width:20%; text-align:right; vertical-align:top;">${logo2Html}</td>
        </tr>
      </table>
    </div>
    <table style="width:100%;margin-top:20px;font-size:14px;"><tr><td><b>PROJECT:</b> ${payload.project}<br><b>ADDRESS:</b> ${address}<br><b>SUBJECT:</b> ${subjectField}${supplierHtml}</td><td align="right" valign="top" style="white-space: nowrap;"><b>NO:</b> ${id}<br><b>DATE:</b> ${formatFullDate(timestamp)}</td></tr></table>
    <table border="1" style="width:100%;border-collapse:collapse;margin-top:20px;"><thead style="background:#f8f9fa;">${headerHtml}</thead><tbody>${rowsHtml}</tbody>${footerHtml}</table>
    ${signatureBlock}${photoSection}
  </body></html>`;

  const finalPdfName = `${id}_${getFileDateSuffix()}.pdf`;
  const pdfBlob = safeDriveAction(() => HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(finalPdfName));
  
  const phaseName = (payload.items && payload.items[0] && payload.items[0].phase) ? payload.items[0].phase : "General";
  let docTypeFolder = "Material Request Forms";
  if (title.includes("PURCHASE ORDER")) docTypeFolder = "Purchase Orders";
  else if (title.includes("QUOTATION")) docTypeFolder = "Request for Quotation";
  else if (title.includes("APPROVED")) docTypeFolder = "Approved MRFs";

  const target = getOrCreateDynamicFolder(payload.project, phaseName, docTypeFolder);
  const createdFile = safeDriveAction(() => target.createFile(pdfBlob));
  
  return { url: createdFile.getUrl(), blob: pdfBlob };
}

// ==========================================
// --- APPROVAL & JUSTIFICATION LOGIC ---
// ==========================================

function getPendingQueueGrouped() {
  const sheet = SS.getSheetByName(SHEETS.QUEUE);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  const groups = {};
  
  data.forEach(row => {
    const status = row[10] ? row[10].toString().trim().toLowerCase() : "";
    if (status === "pending" || status === "pending additional") {
      const id = row[2] ? row[2].toString() : "";
      if (!groups[id]) {
        let d = new Date(row[1]);
        let dateStr = !isNaN(d.getTime()) ? formatFullDate(d) : row[1].toString();
        groups[id] = { id: id, project: row[3] ? row[3].toString() : "", requestor: row[4] ? row[4].toString() : "", date: dateStr, items: [], docType: row[11] ? row[11].toString() : "" };
      }
      const itemName = row[6] ? row[6].toString() : "";
      const itemQty = parseFloat(row[7]) || 0;
      const itemUnit = row[8] ? row[8].toString() : "";
      
      let existing = groups[id].items.find(i => i.name === itemName && (i.unit || "") === itemUnit);
      if (existing) {
         existing.qty += itemQty;
      } else {
         groups[id].items.push({ phase: row[5] ? row[5].toString() : "", name: itemName, qty: itemQty, unit: itemUnit, remarks: row[9] ? row[9].toString() : "" });
      }
    }
  });
  return Object.values(groups).reverse();
}

function getLogMetadata(requestIds) {
  const logData = SS.getSheetByName(SHEETS.LOGS).getDataRange().getValues();
  const justSheet = SS.getSheetByName(SHEETS.JUSTIFICATIONS);
  let justData = [];
  if (justSheet) justData = justSheet.getDataRange().getValues();

  let metadata = [];
  logData.forEach(row => {
    if (requestIds.includes(row[1].toString())) {
      let mrfId = row[1].toString();
      let itemName = row[5].toString();

      // Find latest justification for this specific item
      let latestJustification = "";
      let latestQuestion = "";
      for (let j = justData.length - 1; j >= 1; j--) {
        if (justData[j][1].toString() === mrfId && justData[j][5].toString() === itemName) {
          latestQuestion = justData[j][6].toString();
          latestJustification = justData[j][7].toString(); // The reply
          break;
        }
      }

      metadata.push({ 
        id: mrfId, 
        name: itemName, 
        scope: row[6].toString(), 
        subScope: row[7].toString(), 
        subSubScope: row[8].toString(),
        question: latestQuestion,
        justification: latestJustification
      });
    }
  });
  return metadata;
}

function processBatchApproval(requestId, itemUpdates, approverName) {
  return withDocumentLock_(function () {
  try {
    const now = new Date();
    const appName = approverName || "Approver";
    
    const empData = SS.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
    let encoderEmails = [];
    for(let e = empData.length - 1; e >= 1; e--) {
       const em = empData[e][1] ? empData[e][1].toString().toLowerCase().trim() : "";
       const rl = empData[e][2] ? empData[e][2].toString().toLowerCase().trim() : "";
       if (rl === "encoder" && em !== "") encoderEmails.push(em);
    }
    
    const projData = SS.getSheetByName(SHEETS.PROJECTS).getDataRange().getValues();
    const projMap = {};
    for (let p = 1; p < projData.length; p++) {
       projMap[projData[p][0].toString().trim()] = projData[p][2] || "N/A";
    }

    const logSheet = SS.getSheetByName(SHEETS.LOGS);
    const invSheet = SS.getSheetByName(SHEETS.INVENTORY);
    const queue = SS.getSheetByName(SHEETS.QUEUE);
    
    const logData = logSheet.getDataRange().getValues();
    let invUpdates = [];
    let requestorEmail = ""; 
    let declinedItemsSummary = []; 
    let projectTitle = itemUpdates[0].project;
    let requestorOriginalUpload = null;

    itemUpdates.forEach(update => {
      const finalQty = parseFloat(update.finalQty) || 0;
      if (finalQty <= 0) declinedItemsSummary.push(`- ${update.name}`);
      
      let matchedIndices = [];
      const sReqId = requestId.toString().trim().toLowerCase();
      const sName = update.name.toString().trim().toLowerCase();
      const sUnit = (update.unit || "").toString().trim().toLowerCase();

      for (let i = logData.length - 1; i >= 1; i--) {
        const rowReqId = logData[i][1].toString().trim().toLowerCase();
        const rowName = logData[i][5].toString().trim().toLowerCase();
        const rowUnit = (logData[i][10] || "").toString().trim().toLowerCase();
        const rowStatus = logData[i][12].toString().trim().toLowerCase();

        if (rowReqId === sReqId && rowName === sName && rowUnit === sUnit && rowStatus.includes("pending")) {
          matchedIndices.push(i);
        }
      }

      if (matchedIndices.length > 0) {
        let remainingQtyToApprove = finalQty;
        let allocations = matchedIndices.map(() => 0);
        
        for (let k = 0; k < matchedIndices.length; k++) {
          let idx = matchedIndices[k];
          let originalQty = parseFloat(logData[idx][9]) || 0; 
          
          if (remainingQtyToApprove >= originalQty) {
            allocations[k] = originalQty;
            remainingQtyToApprove -= originalQty;
          } else {
            allocations[k] = remainingQtyToApprove;
            remainingQtyToApprove = 0;
          }
        }
        // Surplus over the total originally requested is dumped into the
        // first allocation slot — preserves the historical behavior in case
        // anyone relies on it. We log a warning so it shows up in Stackdriver
        // if it ever surfaces in practice.
        if (remainingQtyToApprove > 0) {
          console.warn("processBatchApproval: approved qty for '" + sName + "' exceeded total requested by " + remainingQtyToApprove + "; allocated to first matched row.");
          allocations[0] += remainingQtyToApprove;
        }

        for (let k = 0; k < matchedIndices.length; k++) {
          let idx = matchedIndices[k];
          let approvedForThisRow = allocations[k];
          
          if (!requestorEmail && logData[idx][19]) requestorEmail = logData[idx][19].toString().trim(); 
          if (!requestorOriginalUpload && logData[idx][17]) {
             try {
                let fId = getFileIdFromUrl(logData[idx][17].toString());
                if (fId) requestorOriginalUpload = safeDriveAction(() => DriveApp.getFileById(fId));
             } catch(e){}
          }

          if (approvedForThisRow > 0) invUpdates.push([now, logData[idx][2], logData[idx][4], logData[idx][5], logData[idx][6], logData[idx][7], logData[idx][8], approvedForThisRow, logData[idx][10]]);

          const status = approvedForThisRow > 0 ? "Approved" : "Declined";
          logData[idx][12] = status;
          logData[idx][13] = appName;
          logData[idx][14] = approvedForThisRow;
          logData[idx][15] = update.appRemarks || "";
          logData[idx][16] = now;
        }
      }
    });

    // Skip row 1 on write-back — it's the header row, never modified by the
    // loop above, and including it risks stomping a future header change.
    if (logData.length > 1) {
      logSheet.getRange(2, 1, logData.length - 1, logData[0].length).setValues(logData.slice(1));
    }
    if (invUpdates.length > 0) invSheet.getRange(invSheet.getLastRow() + 1, 1, invUpdates.length, invUpdates[0].length).setValues(invUpdates);

    const approvedItems = itemUpdates.filter(it => (parseFloat(it.finalQty) || 0) > 0);
    let pdfAttachment = null;
    let suppPdfAttachment = null;
    
    if (approvedItems.length > 0) {
      const address = projMap[projectTitle.trim()] || "N/A";
      const approvedPdfData = generateDocPDF({project: projectTitle, items: approvedItems}, requestId, now, null, address, "APPROVED MATERIAL REQUEST", appName, false, "", false);
      const supplierPdfData = generateDocPDF({project: projectTitle, items: approvedItems}, requestId, now, null, address, "REQUEST FOR QUOTATION", appName, false, "", true);

      pdfAttachment = approvedPdfData.blob; 
      suppPdfAttachment = supplierPdfData.blob;
    }

    if (requestorEmail !== "") {
      let emailBody = `Hello,\n\nYour Material Request ${requestId} has been reviewed by ${appName}.\n\n`;
      if (declinedItemsSummary.length > 0) emailBody += `NOTE: The following items were DECLINED:\n${declinedItemsSummary.join("\n")}\n\n`; 
      emailBody += "Approved items have been processed and are now pending for supplier and price encoding.";
      
      let reqAttachments = [];
      if (pdfAttachment) reqAttachments.push(pdfAttachment);
      if (requestorOriginalUpload) reqAttachments.push(requestorOriginalUpload);
      
      try { GmailApp.sendEmail(requestorEmail, "MRF Processed: " + requestId + " - " + projectTitle, emailBody, { attachments: reqAttachments }); } catch(e){}
    }
    
    if (encoderEmails.length > 0) {
      const encodeUrl = ScriptApp.getService().getUrl();
      try { GmailApp.sendEmail(encoderEmails.join(","), "Pending Encoding: " + requestId + " - " + projectTitle, `MRF ${requestId} has been approved and requires supplier/price encoding.\n\nPlease proceed here: ${encodeUrl}`); } catch(e){}
    }
    
    if (approvedItems.length > 0) {
      const supplierSheet = SS.getSheetByName(SHEETS.SUPPLIERS);
      let supplierAttachments = [];
      if (suppPdfAttachment) supplierAttachments.push(suppPdfAttachment);
      if (requestorOriginalUpload) supplierAttachments.push(requestorOriginalUpload); 

      if (supplierSheet) {
        const supplierData = supplierSheet.getDataRange().getValues();
        supplierData.slice(1).forEach(row => {
          const supplierEmail = row[2];
          const contactPerson = row[4] || "Supplier"; 
          if (supplierEmail) {
            try { GmailApp.sendEmail(supplierEmail, "Quotation Request: " + requestId, 
              "Hello " + contactPerson + ",\n\nPlease provide your best quotation for the items listed in the attached approved MRF for Project: " + projectTitle + ".\nPlease refer to the attachment for detailed quantities and specifications.\n\nThank you.",
              { attachments: supplierAttachments }); } catch(e){}
          }
        });
      }
    }

    const qRows = queue.getDataRange().getValues();
    for (let j = qRows.length - 1; j >= 1; j--) { 
      if (qRows[j][2].toString() === requestId) queue.deleteRow(j + 1); 
    }

    return "Processed.";
  } catch (e) { return "Error: " + e.message; }
  });
}

function processBatchDecline(requestId, itemRemarks, approverName) {
  return withDocumentLock_(function () {
  try {
    const logs = SS.getSheetByName(SHEETS.LOGS);
    const queue = SS.getSheetByName(SHEETS.QUEUE);
    const now = new Date();
    const appName = approverName || "Approver";

    const logData = logs.getDataRange().getValues();
    let requestorEmail = "";
    const remarksMap = {};
    itemRemarks.forEach(it => { remarksMap[it.name] = it.remarks; });

    for (let i = logData.length - 1; i >= 1; i--) {
      if (logData[i][1].toString() === requestId) {
        if (!requestorEmail && logData[i][19]) requestorEmail = logData[i][19].toString().trim();
        const itemName = logData[i][5].toString();
        
        logData[i][12] = "Declined";
        logData[i][13] = appName;
        logData[i][14] = 0;
        logData[i][15] = remarksMap[itemName] || "";
        logData[i][16] = now;
      }
    }
    if (logData.length > 1) {
      logs.getRange(2, 1, logData.length - 1, logData[0].length).setValues(logData.slice(1));
    }

    if (requestorEmail !== "") {
      const reason = itemRemarks.length > 0 ? itemRemarks[0].remarks : "No reason provided.";
      try{ GmailApp.sendEmail(requestorEmail, "MRF Declined: " + requestId, "Hello,\n\nYour request " + requestId + " has been declined.\n\nReason: " + reason + "\n\nIf you have any questions, please contact the approver."); }catch(e){}
    }
    
    const qRows = queue.getDataRange().getValues();
    for (let j = qRows.length - 1; j >= 1; j--) { 
      if (qRows[j][2].toString() === requestId) queue.deleteRow(j + 1); 
    }

    return "Batch Declined.";
  } catch (e) { return "Error: " + e.message; }
  });
}

function processBatchJustify(requestId, itemQuestions, approverEmail) {
  return withDocumentLock_(function () {
  try {
    const logs = SS.getSheetByName(SHEETS.LOGS);
    const justSheet = SS.getSheetByName(SHEETS.JUSTIFICATIONS);
    const queue = SS.getSheetByName(SHEETS.QUEUE);
    
    if (!justSheet) return "Error: Justification sheet missing. Admin must run setup script.";

    const now = new Date();
    const appEmail = approverEmail || "Approver";
    
    const logData = logs.getDataRange().getValues();
    let project = "";
    let requestor = "";

    const questionsMap = {};
    itemQuestions.forEach(it => questionsMap[it.name] = it.question);

    for (let i = logData.length - 1; i >= 1; i--) {
      if (logData[i][1].toString() === requestId && questionsMap[logData[i][5].toString()]) {
         logData[i][12] = "Pending Justification";
         project = logData[i][2];
         requestor = logData[i][3];
      }
    }
    if (logData.length > 1) {
      logs.getRange(2, 1, logData.length - 1, logData[0].length).setValues(logData.slice(1));
    }

    let jUpdates = [];
    for (let itemName in questionsMap) {
       jUpdates.push([now, requestId, project, requestor, appEmail, itemName, questionsMap[itemName], "", "Pending Reply"]);
    }
    if (jUpdates.length > 0) justSheet.getRange(justSheet.getLastRow() + 1, 1, jUpdates.length, jUpdates[0].length).setValues(jUpdates);

    const qRows = queue.getDataRange().getValues();
    for (let j = qRows.length - 1; j >= 1; j--) {
      if (qRows[j][2].toString() === requestId && questionsMap[qRows[j][6].toString()]) {
         qRows[j][10] = "Pending Justification";
      }
    }
    if (qRows.length > 1) {
      queue.getRange(2, 1, qRows.length - 1, qRows[0].length).setValues(qRows.slice(1));
    }

    return "Requested further justification.";
  } catch (e) { return "Error: " + e.message; }
  });
}

function submitJustificationReply(repliesArray) {
   return withDocumentLock_(function () {
   try {
     const justSheet = SS.getSheetByName(SHEETS.JUSTIFICATIONS);
     const logs = SS.getSheetByName(SHEETS.LOGS);
     const queue = SS.getSheetByName(SHEETS.QUEUE);

     if (!justSheet) return "Error: Justification sheet missing.";
     if (!repliesArray || !repliesArray.length) return "Error: No replies provided.";

     // Reject empty replies — they would otherwise flip status to "Resolved" with
     // no actual content, leaving the approver to chase down an empty answer.
     for (let k = 0; k < repliesArray.length; k++) {
       const reply = (repliesArray[k].reply || "").toString().trim();
       if (!reply) return "Error: Reply for '" + (repliesArray[k].itemName || "an item") + "' cannot be empty.";
     }

     const jData = justSheet.getDataRange().getValues();
     const lData = logs.getDataRange().getValues();
     const qData = queue.getDataRange().getValues();
     
     let emailsToSend = {}; // Store { "approver@email.com": { mrfId: "MRF-123", items: [] } }
     let matched = 0;

     repliesArray.forEach(upd => {
        // Update Justification Database
        for(let i = jData.length - 1; i >= 1; i--) {
           if(jData[i][1].toString() === upd.mrfId && jData[i][5].toString() === upd.itemName && jData[i][8] === "Pending Reply") {
              jData[i][7] = upd.reply;
              jData[i][8] = "Resolved";
              matched++;

              let appEmail = jData[i][4].toString();
              if (appEmail) {
                  if (!emailsToSend[appEmail]) emailsToSend[appEmail] = {};
                  if (!emailsToSend[appEmail][upd.mrfId]) emailsToSend[appEmail][upd.mrfId] = [];
                  emailsToSend[appEmail][upd.mrfId].push(`- ${upd.itemName}: ${upd.reply}`);
              }
           }
        }
        // Change MRF Status back to Pending for Approver to see again
        for(let i = lData.length - 1; i >= 1; i--) {
           if(lData[i][1].toString() === upd.mrfId && lData[i][5].toString() === upd.itemName && lData[i][12] === "Pending Justification") {
              lData[i][12] = "Pending";
           }
        }
        for(let i = qData.length - 1; i >= 1; i--) {
           if(qData[i][2].toString() === upd.mrfId && qData[i][6].toString() === upd.itemName && qData[i][10] === "Pending Justification") {
              qData[i][10] = "Pending";
           }
        }
     });

     // If no Pending Reply row matched, log a warning (Stackdriver) but
     // preserve the historical "Replies submitted successfully." response
     // so we don't surprise the UI with a new error path on edge cases.
     if (matched === 0) {
       console.warn("submitJustificationReply: no Pending Reply rows matched the submitted items.");
     }

     // Skip the header row on each write-back to avoid stomping headers if
     // any future loop accidentally touches index 0.
     if (jData.length > 1) justSheet.getRange(2, 1, jData.length - 1, jData[0].length).setValues(jData.slice(1));
     if (lData.length > 1) logs.getRange(2, 1, lData.length - 1, lData[0].length).setValues(lData.slice(1));
     if (qData.length > 1) queue.getRange(2, 1, qData.length - 1, qData[0].length).setValues(qData.slice(1));

     // Trigger automatic email to approver
     for (let email in emailsToSend) {
        for (let mId in emailsToSend[email]) {
            let msg = `Hello,\n\nThe requestor has replied to your request for justification on ${mId}.\n\nReplies:\n${emailsToSend[email][mId].join("\n")}\n\nPlease log in to your Approver Portal to proceed.`;
            try { GmailApp.sendEmail(email, "Justification Received: " + mId, msg); } catch(e){}
        }
     }

     return "Replies submitted successfully.";
   } catch(e) { return "Error: " + e.message; }
   });
}

// ==========================================
// --- ENCODER PORTAL ENDPOINTS ---
// These three feed the Encode Supplier & Price portal.
//   getSupplierList()      — populates the supplier dropdown
//   getApprovedMRFItems()  — returns MRFs ready to be encoded (approved + not yet in PO_QUEUE)
//   processEncoding(...)   — writes the encoder's supplier/price/payment-term entries to PO_QUEUE
// ==========================================

function getSupplierList() {
  try {
    const sheet = SS.getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const out = [];
    const seen = {};
    for (let i = 0; i < data.length; i++) {
      const name = (data[i][0] || "").toString().trim();
      if (!name) continue;
      const key = name.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      out.push(name);
    }
    out.sort();
    return out;
  } catch (e) {
    return [];
  }
}

// Returns approved-but-not-yet-encoded MRFs grouped by MRF Number.
// "Already encoded" = appears in PO Generation Queue.
function getApprovedMRFItems() {
  try {
    const logSheet = SS.getSheetByName(SHEETS.LOGS);
    if (!logSheet || logSheet.getLastRow() < 2) return [];

    // Index of MRF Numbers already present in PO_QUEUE (= already encoded).
    const encodedMRFs = {};
    const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
    if (poSheet && poSheet.getLastRow() >= 2) {
      const poIds = poSheet.getRange(2, 2, poSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < poIds.length; i++) {
        const id = (poIds[i][0] || "").toString().trim().toLowerCase();
        if (id) encodedMRFs[id] = true;
      }
    }

    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 21).getValues();
    const groupMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = (row[12] || "").toString().trim().toLowerCase();
      const mrfId = (row[1] || "").toString().trim();
      const approvedQty = parseFloat(row[14]) || 0;

      if (status !== "approved") continue;
      if (approvedQty <= 0) continue;
      if (encodedMRFs[mrfId.toLowerCase()]) continue;

      if (!groupMap[mrfId]) {
        const dateReq = row[16] || row[0]; // Date of Decision preferred, else Date Req
        let dateStr = "";
        if (dateReq) {
          const d = new Date(dateReq);
          dateStr = isNaN(d.getTime()) ? dateReq.toString() : formatFullDate(d);
        }
        groupMap[mrfId] = {
          id: mrfId,
          date: dateStr,
          project: (row[2] || "").toString(),
          requestor: (row[3] || "").toString(),
          items: []
        };
      }

      groupMap[mrfId].items.push({
        name: (row[5] || "").toString(),
        qty: approvedQty,
        unit: (row[10] || "").toString(),
        scope: (row[6] || "").toString(),
        phase: (row[4] || "").toString(),
        subScope: (row[7] || "").toString(),
        subSubScope: (row[8] || "").toString()
      });
    }

    const out = [];
    for (const k in groupMap) out.push(groupMap[k]);
    out.sort(function (a, b) { return a.id.localeCompare(b.id); });
    return out;
  } catch (e) {
    return [];
  }
}

// Receives [{ mrfId, supplierName, items: [{ itemName, originalUnit, price, qty,
// unit, termDesc, termDate }] }]. Writes one row per item to PO_QUEUE and to
// Supplier Price-in. Looks up project/requestor/phase/remarks for each item
// from MRF Submission Logs so PO_QUEUE rows carry the full context.
function processEncoding(payloadUpdates) {
  return withDocumentLock_(function () {
  try {
    if (!payloadUpdates || !payloadUpdates.length) return "Error: No data to encode.";

    const logSheet = SS.getSheetByName(SHEETS.LOGS);
    const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
    const priceSheet = SS.getSheetByName(SHEETS.SUPPLIER_PRICE_IN);
    const accSheet = SS.getSheetByName(SHEETS.ACCOUNTING);
    const projSheet = SS.getSheetByName(SHEETS.PROJECTS);
    const suppSheet = SS.getSheetByName(SHEETS.SUPPLIERS);
    if (!logSheet || !poSheet) return "Error: Required sheets missing. Run Setup.";

    // Collected status messages from the per-supplier email send loop. Surfaced
    // in the return value so the encoder UI can tell the user why a supplier
    // didn't get an email (most common cause: missing/invalid email in Supplier DB).
    const emailNotes = [];

    const logData = logSheet.getDataRange().getValues();
    const now = new Date();

    // Project → address lookup (used for the PO PDF header).
    const projAddr = {};
    if (projSheet) {
      const pData = projSheet.getDataRange().getValues();
      for (let i = 1; i < pData.length; i++) {
        const name = (pData[i][0] || "").toString().trim();
        if (name) projAddr[name] = (pData[i][2] || "N/A").toString();
      }
    }

    payloadUpdates.forEach(function (update) {
      const mrfId = (update.mrfId || "").toString().trim();
      const supplier = (update.supplierName || "").toString().trim();
      if (!mrfId || !supplier || !update.items || !update.items.length) return;

      // Look up MRF context (project / requestor / per-item phase + remarks)
      // from MRF Submission Logs. Multiple log rows can share an item name; we
      // take the first match for each.
      let project = "";
      let requestor = "";
      let approver = "";   // MRF Submission Logs col N (index 13) — set by processBatchApproval
      const itemMeta = {};
      for (let i = 1; i < logData.length; i++) {
        if ((logData[i][1] || "").toString().trim() !== mrfId) continue;
        const rowStatus = (logData[i][12] || "").toString().trim().toLowerCase();
        if (!project)   project   = (logData[i][2] || "").toString();
        if (!requestor) requestor = (logData[i][3] || "").toString();
        if (!approver)  approver  = (logData[i][13] || "").toString().trim();
        // Only pick item metadata from APPROVED rows. Previously the first
        // matching row won — which could be a declined row with the same
        // item name, leaving the PO PDF with stale phase/remarks.
        if (rowStatus !== "approved") continue;
        const itemKey = (logData[i][5] || "").toString().trim().toLowerCase();
        if (itemKey && !itemMeta[itemKey]) {
          itemMeta[itemKey] = {
            phase:   (logData[i][4]  || "").toString(),
            remarks: (logData[i][11] || "").toString()
          };
        }
      }

      // Items collected for the PO PDF. We also capture the first non-blank
      // payment term/date pair so generateDocPDF can render the payment
      // schedule block in the PDF header.
      const poItems = [];
      let firstTermDesc = "";
      let firstTermDate = "";

      update.items.forEach(function (item) {
        const meta = itemMeta[(item.itemName || "").toString().trim().toLowerCase()] || { phase: "", remarks: "" };
        const qty = parseFloat(item.qty) || 0;
        const price = parseFloat(item.price) || 0;

        // PO Generation Queue: Timestamp | MRF# | Project | Requestor | Phase |
        //                     Item | Qty | Unit | Remarks | Supplier | Price |
        //                     Payment Terms | Payment Date
        poSheet.appendRow([
          now, mrfId, project, requestor, meta.phase,
          item.itemName || "",
          qty,
          item.unit || "",
          meta.remarks,
          supplier,
          price,
          item.termDesc || "",
          item.termDate || ""
        ]);

        // Supplier Price-in: Timestamp | MRF# | Item | Supplier | Price |
        //                    Encoded By | Payment Terms | Payment Date
        if (priceSheet) {
          priceSheet.appendRow([
            now, mrfId, item.itemName || "", supplier, price,
            "Encoder",
            item.termDesc || "",
            item.termDate || ""
          ]);
        }

        poItems.push({
          name: item.itemName || "",
          qty: qty,
          unit: item.unit || "",
          price: price,
          remarks: meta.remarks,
          phase: meta.phase
        });
        if (!firstTermDesc && item.termDesc) firstTermDesc = item.termDesc;
        if (!firstTermDate && item.termDate) firstTermDate = item.termDate;
      });

      // --- Generate the PO PDF (PURCHASE ORDER, with prices, addressed to supplier).
      let poLink = "";
      let poBlobForEmail = null;
      try {
        const address = projAddr[project.trim()] || "N/A";
        const poId = poDisplayId_(mrfId);
        // Signatory on the PO is the approver who released this MRF (their
        // Position is looked up from Employee Database inside generateDocPDF
        // and printed under the name). Fall back to "Accounting" only if the
        // log row is missing the approver name (e.g., legacy / imported MRFs).
        const signatory = approver || "Accounting";
        const poPdfData = generateDocPDF(
          { project: project, items: poItems, paymentTerms: firstTermDesc, paymentDate: firstTermDate },
          poId, now, null, address, "PURCHASE ORDER", signatory,
          true,      // hasPrices — show unit price + grand total
          supplier,  // supplierName — renders supplier + payment-schedule block in header
          false      // isForSupplier — keep grand total visible (it's the buyer-side copy)
        );
        poLink = poPdfData.url;
        // generateDocPDF returns { url, blob? } — capture the blob if present so
        // we can attach without a second Drive round-trip below.
        if (poPdfData && poPdfData.blob) poBlobForEmail = poPdfData.blob;
      } catch (e) {
        console.error("PO PDF generation failed: " + e.toString());
      }

      // --- Email the supplier with the PO attached. Best-effort — never fail
      //     the encoding flow if the email pipeline is unhappy, but record a
      //     status note so the encoder UI can show why no email was sent.
      try {
        if (!suppSheet) {
          emailNotes.push('Supplier Database sheet missing — no email sent for "' + supplier + '"');
        } else {
          const sData = suppSheet.getDataRange().getValues();
          let supplierEmail = "";
          let contactPerson = "Supplier";
          let matched = false;
          for (let i = 1; i < sData.length; i++) {
            if ((sData[i][0] || "").toString().trim().toLowerCase() === supplier.toLowerCase()) {
              supplierEmail = (sData[i][2] || "").toString().trim();
              contactPerson = (sData[i][4] || "Supplier").toString();
              matched = true;
              break;
            }
          }
          if (!matched) {
            emailNotes.push('Supplier "' + supplier + '" not found in Supplier Database — no email sent');
          } else if (!supplierEmail) {
            emailNotes.push('Supplier "' + supplier + '" has no Email (col C) in Supplier Database — no email sent');
          } else if (!poLink && !poBlobForEmail) {
            emailNotes.push('PO PDF generation failed for "' + supplier + '" — no email sent (see logs)');
          } else {
            const attachments = [];
            if (poBlobForEmail) {
              attachments.push(poBlobForEmail);
            } else {
              const poFileId = getFileIdFromUrl(poLink);
              if (poFileId) {
                try { attachments.push(safeDriveAction(function () { return DriveApp.getFileById(poFileId).getBlob(); })); } catch (e) {}
              }
            }
            const termsLine = firstTermDesc ? ("\n\nPayment terms: " + firstTermDesc + (firstTermDate ? " (target dates: " + firstTermDate + ")" : "")) : "";
            const body =
              "Hello " + contactPerson + ",\n\n" +
              "Please find attached the Purchase Order for Project: " + project + "." +
              termsLine + "\n\n" +
              "Once payment has been processed by our Accounting team, we will follow up with the deposit slip.\n\n" +
              "Thank you.";
            try {
              GmailApp.sendEmail(
                supplierEmail,
                "Purchase Order Issued: " + poDisplayId_(mrfId) + " - " + project,
                body,
                { attachments: attachments }
              );
            } catch (sendErr) {
              emailNotes.push('Sending PO to "' + supplier + '" failed: ' + sendErr.toString());
              console.error("PO supplier email failed: " + sendErr.toString());
            }
          }
        }
      } catch (lookupErr) {
        emailNotes.push('Supplier email lookup failed for "' + supplier + '": ' + lookupErr.toString());
        console.error("PO supplier email lookup failed: " + lookupErr.toString());
      }

      // --- Append (or refresh) the Accounting Queue row for this (PO, supplier).
      // Columns: Timestamp | MRF# | Project | Supplier | PO Doc Link | Status
      if (accSheet) {
        try {
          const accData = accSheet.getDataRange().getValues();
          let existingRow = 0;
          for (let i = 1; i < accData.length; i++) {
            const sameId = (accData[i][1] || "").toString().trim().toLowerCase() === mrfId.toLowerCase();
            const sameSupp = (accData[i][3] || "").toString().trim().toLowerCase() === supplier.toLowerCase();
            if (sameId && sameSupp) { existingRow = i + 1; break; }
          }
          if (existingRow) {
            // Don't overwrite a row that's already been deposited.
            const curStatus = (accData[existingRow - 1][5] || "").toString().trim().toLowerCase();
            if (curStatus !== "deposited" && poLink) accSheet.getRange(existingRow, 5).setValue(poLink);
          } else {
            accSheet.appendRow([now, mrfId, project, supplier, poLink, "Pending Deposit"]);
          }
        } catch (e) {
          console.error("Accounting Queue write failed: " + e.toString());
        }
      }

      // Seed Payment Logs rows for this PO so the Payments tab is populated immediately
      // (instead of lazily on first read of getSupplierPaymentSchedule).
      try { ensurePaymentLogRowsForPO_(mrfId); } catch (e) {}
    });

    let msg = "Encoding submitted successfully.";
    if (emailNotes.length) {
      // Encoding itself worked; the only thing that may not have happened is
      // the supplier email. Surface the reason(s) so the user can fix the data.
      msg += "\n\nSupplier email notes:\n- " + emailNotes.join("\n- ");
    }
    return msg;
  } catch (e) {
    return "Error: " + e.message;
  }
  });
}

// ==========================================
// --- SUPPLIER PAYMENTS LOGIC (NEW) ---
// ==========================================

function getPOPaymentStatus() {
  const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
  const rcvSheet = SS.getSheetByName(SHEETS.RECEIVING_LOGS);
  const paySheet = SS.getSheetByName(SHEETS.PAYMENT_LOGS || "Payment Logs"); 

  if (!poSheet || !rcvSheet || !paySheet) return { error: "Required sheets missing. Run Setup." };

  const poData = poSheet.getDataRange().getValues();
  const rcvData = rcvSheet.getDataRange().getValues();
  const payData = paySheet.getDataRange().getValues();

  // 1. Group PO Generation Queue by MRF Number (key is lowercased for robust matching)
  let poMap = {};
  let displayIdMap = {}; // key -> original-cased ID for display
  for (let i = 1; i < poData.length; i++) {
    let rawMrfId = poData[i][1].toString().trim();
    if (!rawMrfId) continue;
    let mrfId = rawMrfId.toLowerCase();
    if (!displayIdMap[mrfId]) displayIdMap[mrfId] = rawMrfId;
    let project = poData[i][2].toString();
    let itemName = poData[i][5].toString().trim();
    let qtyRaw = poData[i][6];
    let unit = poData[i][7].toString().toLowerCase();
    let supplier = poData[i][9].toString();
    let price = parseFloat(poData[i][10]) || 0;

    let isBudget = ["labor budget", "material budget", "total budget", "labor cost", "material cost", "total cost", "labor", "material", "lot"].includes(unit);
    let qty = isBudget ? 1 : (parseFloat(qtyRaw) || 0);
    let amount = isBudget ? price : (qty * price);

    if (!poMap[mrfId]) {
      poMap[mrfId] = {
        mrfId: rawMrfId,
        project: project,
        supplier: supplier,
        totalAmount: 0,
        items: {},
        paidAmount: 0,
        firstDelivery: null,
        fullDelivery: null,
        deliveryStatus: "Pending Receipt"
      };
    }
    poMap[mrfId].totalAmount += amount;
    
    if (!poMap[mrfId].items[itemName]) {
      poMap[mrfId].items[itemName] = { req: 0, rcv: 0 };
    }
    poMap[mrfId].items[itemName].req += qty;
  }

  // 2. Fetch Received Items per MRF (case-insensitive match)
  let rcvTimestamps = {};
  for (let i = 1; i < rcvData.length; i++) {
    let ts = new Date(rcvData[i][0]);
    let mrfId = rcvData[i][1].toString().trim().toLowerCase();
    let itemName = rcvData[i][2].toString().trim();
    let rcvQty = parseFloat(rcvData[i][3]) || 0;

    if (poMap[mrfId]) {
      if (!rcvTimestamps[mrfId]) rcvTimestamps[mrfId] = [];
      if (!isNaN(ts.getTime())) rcvTimestamps[mrfId].push(ts);

      if (poMap[mrfId].items[itemName]) {
        poMap[mrfId].items[itemName].rcv += rcvQty;
      }
    }
  }

  // 3. Fetch Paid Logged Payments per MRF (case-insensitive match).
  //    Payment Logs schema (see PAYMENT_LOG_HEADERS):
  //      0:Purchase Order | 1:Term | 2:% | 3:Supplier | 4:Invoiced
  //      5:Due Date | 6:Bank | 7:Check# | 8:Payment Amount
  //    Match on Purchase Order (col 0, may be PO-xxxx or MRF-xxxx) and sum col 8.
  for (let i = 1; i < payData.length; i++) {
    let payPo = (payData[i][0] || "").toString().trim().toLowerCase();
    if (!payPo) continue;
    // Reverse the PO→MRF display rewrite so we can join against poMap keys.
    let mrfKey = payPo.replace(/^po-/i, "mrf-");
    let amount = parseFloat(payData[i][8]) || 0;
    if (poMap[mrfKey]) {
      poMap[mrfKey].paidAmount += amount;
    } else if (poMap[payPo]) {
      poMap[payPo].paidAmount += amount;
    }
  }

  // 4. Calculate Final Statuses for the UI
  let result = [];
  for (let mrfId in poMap) {
    let po = poMap[mrfId];
    let allFull = true;
    let anyRcv = false;

    for (let item in po.items) {
      if (po.items[item].req > 0) { 
        if (po.items[item].rcv > 0) anyRcv = true;
        if (po.items[item].rcv < po.items[item].req) allFull = false;
      }
    }

    if (rcvTimestamps[mrfId] && rcvTimestamps[mrfId].length > 0) {
      rcvTimestamps[mrfId].sort((a, b) => a - b);
      po.firstDelivery = rcvTimestamps[mrfId][0];
      po.fullDelivery = rcvTimestamps[mrfId][rcvTimestamps[mrfId].length - 1];
    }

    if (allFull && Object.keys(po.items).length > 0) po.deliveryStatus = "Completed Delivery";
    else if (anyRcv) po.deliveryStatus = "Partial Delivered";
    else po.deliveryStatus = "Pending Receipt";

    // Timeline labels:
    //  - Completed Delivery → show First + Full delivery dates
    //  - Partial Delivered  → show First delivery + Latest receipt date
    //  - Pending Receipt    → both N/A
    po.firstDelivery = po.firstDelivery ? formatFullDate(po.firstDelivery) : "N/A";
    if (po.deliveryStatus === "Pending Receipt" || !po.fullDelivery) {
      po.fullDelivery = "N/A";
      po.fullDeliveryLabel = "Full";
    } else if (po.deliveryStatus === "Partial Delivered") {
      po.fullDelivery = formatFullDate(po.fullDelivery);
      po.fullDeliveryLabel = "Latest";
    } else {
      po.fullDelivery = formatFullDate(po.fullDelivery);
      po.fullDeliveryLabel = "Full";
    }

    result.push(po);
  }

  return result.reverse(); 
}

// =============================================================================
// DROPDOWNS — single sheet feeds every picklist (Bank, Expense Type, Status,
// Payment Method, ...). Columns: Category | Value | Meta | Active | Sort Order.
// =============================================================================

// Auto-create the sheet on first call so first-time setups don't have to run
// setup() before the dropdowns start working. Seeds nothing here — setup() owns
// the default seeding so the source of truth stays in one place.
function getOrCreateDropdownsSheet_() {
  let sheet = SS.getSheetByName(SHEETS.DROPDOWNS);
  if (!sheet) {
    sheet = SS.insertSheet(SHEETS.DROPDOWNS);
    const headers = ["Category", "Value", "Meta", "Active", "Sort Order"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Return active rows for a given Category, ordered by Sort Order. Each entry:
//   { value, meta, sortOrder }
// Use this everywhere dropdowns need data — never read the sheet directly.
function getDropdownRows(category) {
  try {
    const sheet = getOrCreateDropdownsSheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    const wantCat = (category || "").toString().trim().toLowerCase();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      const cat = (data[i][0] || "").toString().trim().toLowerCase();
      if (cat !== wantCat) continue;
      const value = (data[i][1] || "").toString().trim();
      if (!value) continue;
      const rawActive = (data[i][3] === undefined || data[i][3] === null) ? "yes" : data[i][3].toString().trim().toLowerCase();
      const isActive = (rawActive === "" || rawActive === "yes" || rawActive === "true" || rawActive === "1" || data[i][3] === true);
      if (!isActive) continue;
      const sortOrder = parseFloat(data[i][4]);
      out.push({
        value: value,
        meta: (data[i][2] || "").toString().trim(),
        sortOrder: isNaN(sortOrder) ? 999 : sortOrder
      });
    }
    out.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
    return out;
  } catch (e) {
    return [];
  }
}

// Convenience: just the values (string array), for plain dropdowns with no metadata.
function getDropdownValues(category) {
  return getDropdownRows(category).map(function (r) { return r.value; });
}

// Bank dropdown — used by the Supplier Payments → Log Payment modal.
function getBankList() {
  return getDropdownValues("Bank");
}

// ==========================================
// --- PAYMENT LOGS (one row per payment term per PO) ---
// Columns: Purchase Order | Payment Term | % | Supplier | Invoiced Amount
//          | Payment Due Date | Bank | Check Number | Payment Amount
// ==========================================

const PAYMENT_LOG_HEADERS = ["Purchase Order", "Payment Term", "%", "Supplier", "Invoiced Amount", "Payment Due Date", "Bank", "Check Number", "Payment Amount"];

function getOrCreatePaymentLogsSheet_() {
  const name = SHEETS.PAYMENT_LOGS || "Payment Logs";
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.getRange(1, 1, 1, PAYMENT_LOG_HEADERS.length).setValues([PAYMENT_LOG_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Parse a leading or embedded percentage from a term description.
// "30% DP" → 30; "Down Payment 50%" → 50; "On Delivery" → null
function parseTermPercent_(termDesc) {
  if (!termDesc) return null;
  const m = termDesc.toString().match(/(\d+(?:\.\d+)?)\s*%/);
  return m ? parseFloat(m[1]) : null;
}

// Convert a raw PO ID (e.g., "MRF-0123") into the display form ("PO-0123" / "PO-…").
function poDisplayId_(rawId) {
  return (rawId || "").toString().replace(/^MRF/i, "PO").replace(/^ARF/i, "PO");
}

// Ensure Payment Logs has one row per (PO, Supplier, Term) for the given raw
// MRF id. Correctly handles:
//   - Per Supplier encoder mode: every item under that supplier shares one
//     term string → one set of Payment Logs rows for the supplier.
//   - Per Item encoder mode: items under the same supplier may carry different
//     term strings → one set of Payment Logs rows per distinct term string,
//     keyed on the subtotal of the items that share that string.
//   - Multi-supplier MRFs: each supplier gets its own payment stream (no longer
//     pools every supplier's items into a single invoice total).
// Dedupe key = (display PO id, supplier, term). Re-running this is safe.
function ensurePaymentLogRowsForPO_(rawMrfId) {
  const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
  const payLogs = getOrCreatePaymentLogsSheet_();
  if (!poSheet) return;

  const poData = poSheet.getDataRange().getValues();
  const targetLower = (rawMrfId || "").toString().trim().toLowerCase();
  const displayPo = poDisplayId_(rawMrfId);
  const BUDGET_UNITS = ["labor budget", "material budget", "total budget", "labor cost", "material cost", "total cost", "labor", "material", "lot"];

  // Group PO_QUEUE rows by Supplier, then by exact terms string. Each unique
  // (supplier, terms) bucket accumulates the subtotal of items that share that
  // schedule. The dates string is taken from the first row in the bucket
  // (matching the existing UI contract that "Payment Date" is paired by index
  // with "Payment Terms" — both are pipe-joined strings).
  const bySupplier = {};
  for (let i = 1; i < poData.length; i++) {
    if (!poData[i][1] || poData[i][1].toString().trim().toLowerCase() !== targetLower) continue;
    const supplier = (poData[i][9] || "").toString().trim();
    if (!supplier) continue;
    const supKey = supplier.toLowerCase();
    if (!bySupplier[supKey]) bySupplier[supKey] = { supplier: supplier, buckets: {} };

    const unit = (poData[i][7] || "").toString().toLowerCase();
    const isBudget = BUDGET_UNITS.indexOf(unit) !== -1;
    const qty = isBudget ? 1 : (parseFloat(poData[i][6]) || 0);
    const price = parseFloat(poData[i][10]) || 0;
    const lineTotal = isBudget ? price : qty * price;

    const termsStr = (poData[i][11] || "").toString();
    const datesStr = (poData[i][12] || "").toString();
    if (!termsStr.trim()) continue; // Item has no schedule — skip; nothing to invoice against.

    if (!bySupplier[supKey].buckets[termsStr]) {
      bySupplier[supKey].buckets[termsStr] = { datesStr: datesStr, subtotal: 0 };
    }
    bySupplier[supKey].buckets[termsStr].subtotal += lineTotal;
  }

  // Index existing Payment Logs rows so we don't duplicate on re-runs. Older rows
  // may have been stored under either the raw MRF id or the display PO id —
  // accept either form when deduping.
  const existing = {};
  const lastRow = payLogs.getLastRow();
  if (lastRow >= 2) {
    const logData = payLogs.getRange(2, 1, lastRow - 1, PAYMENT_LOG_HEADERS.length).getValues();
    for (let r = 0; r < logData.length; r++) {
      const po = (logData[r][0] || "").toString().trim().toLowerCase();
      const supp = (logData[r][3] || "").toString().trim().toLowerCase();
      const term = (logData[r][1] || "").toString().trim().toLowerCase();
      if (!po || !term) continue;
      if (po === displayPo.toLowerCase() || po === targetLower) {
        existing[displayPo.toLowerCase() + "|" + supp + "|" + term] = true;
      }
    }
  }

  // Build rows to append. % is parsed from the term string when present,
  // otherwise distributed evenly across unmarked terms within the same schedule.
  const newRows = [];
  for (const supKey in bySupplier) {
    const sup = bySupplier[supKey];
    for (const termsStr in sup.buckets) {
      const bucket = sup.buckets[termsStr];
      const termsArr = termsStr.split(" | ").map(function (s) { return s.trim(); }).filter(Boolean);
      const datesArr = bucket.datesStr.split(" | ").map(function (s) { return s.trim(); });

      const percents = termsArr.map(parseTermPercent_);
      const explicitTotal = percents.reduce(function (a, p) { return a + (p || 0); }, 0);
      const noPercentCount = percents.filter(function (p) { return p === null; }).length;
      const remaining = Math.max(0, 100 - explicitTotal);
      const evenShare = noPercentCount > 0 ? (remaining / noPercentCount) : 0;
      const resolvedPercents = percents.map(function (p) { return (p === null) ? evenShare : p; });

      // First pass: compute the 2-decimal invoiced amount for every term.
      // Then absorb any rounding residual (subtotal - sum) into the term with
      // the largest amount, so the sum equals the supplier subtotal exactly.
      // Example: ₱100,000 split 3 ways at 33.33% each → 33,330 + 33,330 + 33,340.
      const expectedTotalCents = Math.round(bucket.subtotal * 100);
      let invoicedCents = termsArr.map(function (_, i) {
        const pct = resolvedPercents[i] || 0;
        return Math.round(bucket.subtotal * pct); // = subtotal * pct/100 in cents
      });
      let diffCents = expectedTotalCents - invoicedCents.reduce(function (a, b) { return a + b; }, 0);
      if (diffCents !== 0) {
        let maxIdx = 0;
        for (let i = 1; i < invoicedCents.length; i++) {
          if (invoicedCents[i] > invoicedCents[maxIdx]) maxIdx = i;
        }
        invoicedCents[maxIdx] += diffCents;
      }

      for (let i = 0; i < termsArr.length; i++) {
        const term = termsArr[i];
        const dedupeKey = displayPo.toLowerCase() + "|" + supKey + "|" + term.toLowerCase();
        if (existing[dedupeKey]) continue;
        existing[dedupeKey] = true;

        const pct = resolvedPercents[i] || 0;
        const invoiced = invoicedCents[i] / 100;
        newRows.push([
          displayPo,           // Purchase Order
          term,                // Payment Term
          pct,                 // %
          sup.supplier,        // Supplier
          invoiced,            // Invoiced Amount
          datesArr[i] || "",   // Payment Due Date
          "",                  // Bank
          "",                  // Check Number
          ""                   // Payment Amount
        ]);
      }
    }
  }

  if (newRows.length) {
    payLogs.getRange(payLogs.getLastRow() + 1, 1, newRows.length, PAYMENT_LOG_HEADERS.length).setValues(newRows);
  }
}

// One row per PO/Term combo — used to render the Supplier Payments tab.
function getSupplierPaymentSchedule() {
  try {
    // Make sure every PO in PO_QUEUE has its term rows seeded in Payment Logs.
    const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
    if (poSheet && poSheet.getLastRow() >= 2) {
      const poData = poSheet.getRange(2, 2, poSheet.getLastRow() - 1, 1).getValues();
      const seen = {};
      for (let i = 0; i < poData.length; i++) {
        const id = (poData[i][0] || "").toString().trim();
        if (id && !seen[id.toLowerCase()]) {
          seen[id.toLowerCase()] = true;
          ensurePaymentLogRowsForPO_(id);
        }
      }
    }

    const sheet = getOrCreatePaymentLogsSheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PAYMENT_LOG_HEADERS.length).getValues();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      const po = (data[i][0] || "").toString().trim();
      if (!po) continue;
      let due = data[i][5];
      let dueStr = "";
      if (due) {
        const d = new Date(due);
        dueStr = isNaN(d.getTime()) ? due.toString() : formatFullDate(d);
      }
      out.push({
        row: i + 2,
        purchaseOrder: po,
        paymentTerm: (data[i][1] || "").toString(),
        percent: parseFloat(data[i][2]) || 0,
        supplier: (data[i][3] || "").toString(),
        invoicedAmount: parseFloat(data[i][4]) || 0,
        dueDate: dueStr,
        dueDateRaw: due ? new Date(due).getTime() : null,
        bank: (data[i][6] || "").toString(),
        checkNumber: (data[i][7] || "").toString(),
        paymentAmount: parseFloat(data[i][8]) || 0,
        isPaid: !!((data[i][6] || "").toString().trim() && (data[i][7] || "").toString().trim() && parseFloat(data[i][8]) > 0)
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

// Update a Payment Logs row identified by sheet row index. Bank/Check/Payment Amount
// are the only fields the user supplies; everything else was seeded from PO_QUEUE.
function logSupplierPayment(payload) {
  try {
    if (!payload || !payload.row) return "Error: Missing row reference.";
    const sheet = getOrCreatePaymentLogsSheet_();
    const row = parseInt(payload.row, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) return "Error: Invalid row reference.";

    const bank = (payload.bank || "").toString().trim();
    const checkNumber = (payload.checkNumber || "").toString().trim();
    const paymentAmount = parseFloat((payload.paymentAmount || "0").toString().replace(/,/g, ""));
    if (!bank) return "Error: Bank is required.";
    if (!checkNumber) return "Error: Check Number is required.";
    if (isNaN(paymentAmount) || paymentAmount <= 0) return "Error: Payment Amount must be a positive number.";

    sheet.getRange(row, 7).setValue(bank);          // Bank
    sheet.getRange(row, 8).setValue(checkNumber);   // Check Number
    sheet.getRange(row, 9).setValue(paymentAmount); // Payment Amount
    return "Payment Logged Successfully.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

// Backwards-compat: keep the old name pointing at the new implementation in case
// any older HTML still calls submitSupplierPayment.
function submitSupplierPayment(payload) {
  return logSupplierPayment(payload);
}

function getScriptAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ==========================================
// --- RECEIVING PORTAL ENDPOINTS ---
// Backs Receive.html. The requestor (employee who submitted the MRF) opens a
// link of the form:
//   <scriptUrl>?receiveToken=<random-uuid>
// and we route to that template via doGet(). These endpoints feed it.
//
// Each token is a random UUID generated at deposit-slip time and stored in
// Script Properties (key "RCV_<token>" → "<mrfId>") so the token itself
// can't be guessed from the MRF id. The receiving link is emailed to the
// REQUESTOR (not the supplier) — the supplier still gets PO+deposit slip,
// they just don't get the receiving page.
// ==========================================

function createReceiveToken_(mrfId) {
  const id = (mrfId || "").toString().trim();
  if (!id) return "";
  try {
    const props = PropertiesService.getScriptProperties();
    // Reuse an existing token for this MRF if we already minted one — keeps
    // links stable if the requestor re-receives or accounting re-issues the
    // email for the same PO.
    const all = props.getProperties();
    for (const k in all) {
      if (k.indexOf("RCV_") === 0 && all[k] === id) {
        return k.substring(4);
      }
    }
    const token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty("RCV_" + token, id);
    return token;
  } catch (e) {
    return "";
  }
}

function resolveReceiveToken_(token) {
  const t = (token || "").toString().trim();
  if (!t) return "";
  try {
    const mapped = PropertiesService.getScriptProperties().getProperty("RCV_" + t);
    if (mapped) return mapped;
  } catch (e) {}
  // Back-compat: if a legacy "MRF-XXXXX" / "ARF-XXXXX" token is still floating
  // around in an old email, accept it so receiving doesn't silently break.
  // New tokens are 32-char hex, no dash — easy to tell apart.
  if (/^(MRF|ARF)-/i.test(t)) return t;
  return "";
}

// Units that mean "no discrete count" — quantity-tracking doesn't make sense
// for them, so we collapse poQty/received to 1/0 and treat them as a single
// "lot" line in the receiving table. Kept in sync with the same list used in
// getPOPaymentStatus and ensurePaymentLogRowsForPO_.
const BUDGET_UNITS_ = ["labor budget", "material budget", "total budget", "labor cost", "material cost", "total cost", "labor", "material", "lot"];

function isBudgetUnit_(unit) {
  return BUDGET_UNITS_.indexOf((unit || "").toString().toLowerCase().trim()) !== -1;
}

// Returns { items, history } or { error } for the Receiving portal.
//   items:   [{ name, unit, poQty, remaining }]
//   history: [{ date, name, qty, remarks }]
function getItemsForReceiving(token) {
  try {
    const mrfId = resolveReceiveToken_(token);
    if (!mrfId) return { error: "Invalid or expired receiving link." };

    const poSheet = SS.getSheetByName(SHEETS.PO_QUEUE);
    const rcvSheet = SS.getSheetByName(SHEETS.RECEIVING_LOGS);
    if (!poSheet) return { error: "PO Generation Queue sheet missing. Run Setup." };

    const target = mrfId.toLowerCase();

    // 1. Aggregate ordered quantities per item from PO_QUEUE. Multiple suppliers
    //    can share an MRF — for the receiving view we just want the total qty
    //    expected to be delivered against the MRF, so we sum across all rows.
    const items = {};   // itemName → { unit, poQty }
    let foundAny = false;
    const poData = poSheet.getDataRange().getValues();
    for (let i = 1; i < poData.length; i++) {
      if (!poData[i][1] || poData[i][1].toString().trim().toLowerCase() !== target) continue;
      foundAny = true;
      const name = (poData[i][5] || "").toString().trim();
      if (!name) continue;
      const unit = (poData[i][7] || "").toString();
      const qty  = parseFloat(poData[i][6]) || 0;
      if (!items[name]) items[name] = { unit: unit, poQty: 0 };
      items[name].poQty += qty;
    }

    if (!foundAny) return { error: "MRF not found: " + mrfId };

    // 2. Subtract anything already received, and build the history list.
    const history = [];
    if (rcvSheet && rcvSheet.getLastRow() >= 2) {
      const rData = rcvSheet.getDataRange().getValues();
      for (let i = 1; i < rData.length; i++) {
        if (!rData[i][1] || rData[i][1].toString().trim().toLowerCase() !== target) continue;
        const name = (rData[i][2] || "").toString().trim();
        const qty  = parseFloat(rData[i][3]) || 0;
        const remarks = (rData[i][4] || "").toString();
        // poQty is the ORDERED total; remaining is computed below from the received tally.
        const ts = rData[i][0];
        let dateStr = "";
        if (ts) { const d = new Date(ts); dateStr = isNaN(d.getTime()) ? ts.toString() : formatFullDate(d); }
        history.push({ date: dateStr, name: name, qty: qty, remarks: remarks });
      }
    }

    // Tally received quantities per item separately so the math is clean.
    const received = {};
    if (rcvSheet && rcvSheet.getLastRow() >= 2) {
      const rData = rcvSheet.getRange(2, 1, rcvSheet.getLastRow() - 1, 5).getValues();
      for (let i = 0; i < rData.length; i++) {
        if (!rData[i][1] || rData[i][1].toString().trim().toLowerCase() !== target) continue;
        const name = (rData[i][2] || "").toString().trim();
        received[name] = (received[name] || 0) + (parseFloat(rData[i][3]) || 0);
      }
    }

    const itemsOut = [];
    for (const name in items) {
      const ordered = items[name].poQty;
      const rcvd    = received[name] || 0;
      const remaining = Math.max(0, ordered - rcvd);
      itemsOut.push({ name: name, unit: items[name].unit, poQty: ordered, remaining: remaining });
    }
    itemsOut.sort(function (a, b) { return a.name.localeCompare(b.name); });

    return { items: itemsOut, history: history };
  } catch (e) {
    return { error: e.toString() };
  }
}

// Writes one Receiving Logs row per item with rcvdQty > 0. Optionally saves the
// delivery-receipt file under <Project>/Delivery Receipts/<Supplier>/. Returns a
// plain string ("Success: …" / "Error: …") to match how Receive.html parses it
// (it checks for the substring "Error").
function submitReceivedItems(token, finalItems, fileData) {
  try {
    const mrfId = resolveReceiveToken_(token);
    if (!mrfId) return "Error: Invalid or expired receiving link.";
    if (!finalItems || !finalItems.length) return "Error: No items submitted.";

    const rcvSheet = SS.getSheetByName(SHEETS.RECEIVING_LOGS);
    const poSheet  = SS.getSheetByName(SHEETS.PO_QUEUE);
    if (!rcvSheet) return "Error: Receiving Logs sheet missing. Run Setup.";
    if (!poSheet) return "Error: PO Generation Queue sheet missing. Run Setup.";

    // Look up project + supplier for the file path (and to validate the token
    // actually maps to a real PO).
    let project = "";
    let supplier = "";
    let foundAny = false;
    const target = mrfId.toLowerCase();
    const poData = poSheet.getDataRange().getValues();
    for (let i = 1; i < poData.length; i++) {
      if (!poData[i][1] || poData[i][1].toString().trim().toLowerCase() !== target) continue;
      foundAny = true;
      if (!project)  project  = (poData[i][2] || "").toString();
      if (!supplier) supplier = (poData[i][9] || "").toString();
      if (project && supplier) break;
    }
    if (!foundAny) return "Error: MRF not found: " + mrfId;

    // Save the Delivery Receipt (best-effort).
    if (fileData && fileData.data) {
      try {
        const folder = getOrCreateDynamicFolder(project || "General", "Delivery Receipts", supplier || "Unfiled");
        const ext = (fileData.fileName || "").split('.').pop() || "pdf";
        const fileName = "DR_" + mrfId + "_" + getFileDateSuffix() + "." + ext;
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, fileName);
        safeDriveAction(function () { return folder.createFile(blob); });
      } catch (e) {
        console.error("Delivery Receipt upload failed: " + e.toString());
      }
    }

    // Receiving Logs: Timestamp | MRF ID | Item Name | Received Qty | Remarks
    const now = new Date();
    const rows = [];
    for (let i = 0; i < finalItems.length; i++) {
      const qty = parseFloat(finalItems[i].rcvdQty) || 0;
      if (qty <= 0) continue;
      rows.push([
        now,
        mrfId,
        (finalItems[i].name || "").toString(),
        qty,
        (finalItems[i].remarks || "").toString()
      ]);
    }
    if (!rows.length) return "Error: No items with a received quantity greater than 0.";

    rcvSheet.getRange(rcvSheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    return "Success: Logged " + rows.length + " item(s) as received.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

// ==========================================
// --- ACCOUNTING ENDPOINTS ---
// ==========================================

function getAccountingQueue() {
  try {
    const sheet = SS.getSheetByName(SHEETS.ACCOUNTING);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let queue = [];
    for (let i = 1; i < data.length; i++) {
      const status = (data[i][5] || "").toString().trim().toLowerCase();
      // Case-insensitive — accept any casing of "deposited" (manual edits, mixed-case status, etc.)
      if (status !== 'deposited') {
        queue.push({
          row: i + 1,
          timestamp: data[i][0] ? formatFullDate(new Date(data[i][0])) : "",
          mrfId: data[i][1] ? data[i][1].toString() : "",
          project: data[i][2] ? data[i][2].toString() : "",
          supplier: data[i][3] ? data[i][3].toString() : "",
          poLink: data[i][4] ? data[i][4].toString() : "",
          status: data[i][5] ? data[i][5].toString() : ""
        });
      }
    }
    return queue.reverse();
  } catch(e) { return []; }
}

// Finalize a Pending Deposit Slip:
//   1. Save the uploaded slip to Drive under <Project>/Deposit Slips/<Supplier>/
//   2. Stamp the Accounting Queue row: Status="Deposited", col G=deposit slip link,
//      col H=who deposited, col I=timestamp
//   3. Email the supplier with PO + deposit slip attached (best-effort)
function processAccountingUpload(rowIdx, mrfId, supplierName, fileObj, depositedBy) {
  return withDocumentLock_(function () {
  try {
    const sheet = SS.getSheetByName(SHEETS.ACCOUNTING);
    if (!sheet) return "Error: Accounting Queue sheet missing. Run Setup.";
    const row = parseInt(rowIdx, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) return "Error: Invalid row reference.";

    // Read the existing row to pick up project + PO link (already stamped on encode).
    const rowData = sheet.getRange(row, 1, 1, 6).getValues()[0];
    const projectName = (rowData[2] || "").toString();
    const supplier = (supplierName || rowData[3] || "").toString();
    const poLink = (rowData[4] || "").toString();

    // --- 1. Save the slip file (if provided)
    let depositLink = "";
    if (fileObj && fileObj.data) {
      try {
        const folder = getOrCreateDynamicFolder(projectName || "General", "Deposit Slips", supplier || "Unfiled");
        const ext = (fileObj.fileName || "").split('.').pop() || "pdf";
        const fileName = "DEP_" + mrfId + "_" + getFileDateSuffix() + "." + ext;
        const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, fileName);
        const created = safeDriveAction(function () { return folder.createFile(blob); });
        depositLink = created.getUrl();
      } catch (e) {
        console.error("Deposit slip upload failed: " + e.toString());
      }
    }

    // --- 2. Stamp the row. We extend past column F (Status) to track the slip link,
    //        depositor, and timestamp without touching the schema for older rows.
    //        F=Status, G=Deposit Slip Link, H=Deposited By, I=Deposit Timestamp
    sheet.getRange(row, 6).setValue("Deposited");
    if (depositLink) sheet.getRange(row, 7).setValue(depositLink);
    sheet.getRange(row, 8).setValue((depositedBy || "").toString());
    sheet.getRange(row, 9).setValue(new Date());

    // --- 3. Email the supplier (PO + deposit slip). Best-effort — never fail the
    //        operation if the email pipeline is unhappy, but DO surface the
    //        outcome so accounting knows when the supplier was not notified.
    //        If shane@your-domain is in the supplier row, that means the
    //        Supplier Database has the wrong email — not a code bug.
    let emailStatus = "ok";   // ok | no_match | no_email | send_failed
    try {
      const suppSheet = SS.getSheetByName(SHEETS.SUPPLIERS);
      if (suppSheet) {
        const sData = suppSheet.getDataRange().getValues();
        let supplierEmail = "";
        let contactPerson = "Supplier";
        let matched = false;
        for (let i = 1; i < sData.length; i++) {
          if ((sData[i][0] || "").toString().trim().toLowerCase() === supplier.toLowerCase()) {
            supplierEmail = (sData[i][2] || "").toString().trim();
            contactPerson = (sData[i][4] || "Supplier").toString();
            matched = true;
            break;
          }
        }
        if (!matched) emailStatus = "no_match";
        else if (!supplierEmail) emailStatus = "no_email";
        if (supplierEmail) {
          const attachments = [];
          const poFileId = getFileIdFromUrl(poLink);
          if (poFileId) {
            try { attachments.push(safeDriveAction(function () { return DriveApp.getFileById(poFileId).getBlob(); })); } catch (e) {}
          }
          const depFileId = getFileIdFromUrl(depositLink);
          if (depFileId) {
            try { attachments.push(safeDriveAction(function () { return DriveApp.getFileById(depFileId).getBlob(); })); } catch (e) {}
          }

          // Supplier email: PO + deposit slip only. The receiving link is sent
          // to the requestor in a separate email below — suppliers should not
          // be the ones logging deliveries against our inventory.
          const body =
            "Hello " + contactPerson + ",\n\n" +
            "Please find attached the Purchase Order and proof of deposit for Project: " + projectName + ".\n\n" +
            "Kindly confirm receipt and proceed with delivery as agreed.\n\n" +
            "Thank you.";

          try {
            GmailApp.sendEmail(
              supplierEmail,
              "Purchase Order + Deposit Slip: " + poDisplayId_(mrfId) + " - " + projectName,
              body,
              { attachments: attachments }
            );
          } catch (sendErr) {
            emailStatus = "send_failed";
            console.error("Supplier email send failed: " + sendErr.toString());
          }
        }
      }
    } catch (e) {
      emailStatus = "send_failed";
      console.error("Supplier email failed: " + e.toString());
    }

    // --- 3b. Email the REQUESTOR with the receiving link.
    //     The requestor (employee who submitted the original MRF) is the one
    //     who logs delivered items — not the supplier. We look up their email
    //     from MRF Submission Logs (col T = userEmail captured at submitMRF).
    //     Best-effort: never fail the operation if this email can't go out.
    try {
      const logsSheet = SS.getSheetByName(SHEETS.LOGS);
      if (logsSheet) {
        const logData = logsSheet.getDataRange().getValues();
        let requestorEmail = "";
        let requestorName = "";
        for (let i = 1; i < logData.length; i++) {
          if ((logData[i][1] || "").toString().trim().toLowerCase() === mrfId.toLowerCase()) {
            requestorEmail = (logData[i][19] || "").toString().trim(); // col T = email captured at submission
            requestorName  = (logData[i][3]  || "").toString().trim();
            if (requestorEmail) break;
          }
        }
        if (requestorEmail) {
          const token = createReceiveToken_(mrfId);
          let receiveUrl = "";
          try {
            const base = ScriptApp.getService().getUrl();
            if (base && token) receiveUrl = base + (base.indexOf('?') > -1 ? '&' : '?') + 'receiveToken=' + encodeURIComponent(token);
          } catch (e) {}

          if (receiveUrl) {
            const body =
              "Hello " + (requestorName || "there") + ",\n\n" +
              "Your purchase order " + poDisplayId_(mrfId) + " for Project: " + projectName +
              " has been paid and the supplier (" + supplier + ") has been notified.\n\n" +
              "Once the goods arrive, please log the received items here:\n" + receiveUrl + "\n\n" +
              "Thank you.";
            try {
              GmailApp.sendEmail(
                requestorEmail,
                "Receiving Link: " + poDisplayId_(mrfId) + " - " + projectName,
                body
              );
            } catch (sendErr) {
              console.error("Requestor receiving-link email failed: " + sendErr.toString());
            }
          }
        }
      }
    } catch (e) {
      console.error("Requestor email lookup failed: " + e.toString());
    }

    // Build a clear status string so the UI can tell accounting whether the
    // supplier was actually emailed. The deposit slip is already recorded
    // regardless — this is purely about the notification side-effect.
    let suffix = "";
    switch (emailStatus) {
      case "no_match":
        suffix = ' Note: supplier "' + supplier + '" was not found in Supplier Database — no email sent. Add the supplier (with a valid Email in col C) and resend if needed.';
        break;
      case "no_email":
        suffix = ' Note: supplier "' + supplier + '" has no email address in Supplier Database — no email sent. Fill col C for that row and resend if needed.';
        break;
      case "send_failed":
        suffix = " Note: deposit slip recorded but the supplier email could not be sent. See the script logs for details.";
        break;
    }
    return "Successfully processed." + suffix;
  } catch (e) {
    return "Error: " + e.message;
  }
  });
}

// Mark one or more Payment Logs rows as settled-by-deposit. Stamps a sentinel
// Bank/Check ("Deposit Made" / "DEP-xxxxxx") and the full invoiced amount so the
// row also reads as Paid in the Supplier Payments tab. Accounting can edit the
// row later via the Log Payment modal to replace the sentinel with real check
// details.
function markPaymentDeposited(rowIndices) {
  try {
    const sheet = getOrCreatePaymentLogsSheet_();
    if (!rowIndices) return "Error: No rows provided.";
    const indices = Array.isArray(rowIndices) ? rowIndices : [rowIndices];
    const now = new Date();
    let updated = 0;

    indices.forEach(function (idx) {
      const row = parseInt(idx, 10);
      if (!row || row < 2 || row > sheet.getLastRow()) return;
      const existing = sheet.getRange(row, 1, 1, PAYMENT_LOG_HEADERS.length).getValues()[0];
      const invoiced = parseFloat(existing[4]) || 0;
      const bank = (existing[6] || "").toString().trim();
      const check = (existing[7] || "").toString().trim();
      const amt = parseFloat(existing[8]) || 0;
      if (bank && check && amt > 0) return; // Already settled — leave alone.

      sheet.getRange(row, 7).setValue("Deposit Made");
      sheet.getRange(row, 8).setValue("DEP-" + now.getTime().toString().slice(-6) + "-" + row);
      sheet.getRange(row, 9).setValue(invoiced);
      updated++;
    });

    return updated > 0
      ? "Marked " + updated + " term(s) as deposited."
      : "No rows updated (already settled or invalid).";
  } catch (e) {
    return "Error: " + e.message;
  }
}

function getProjectList() {
  try {
    const sheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const projs = new Set();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() !== "project title") {
        projs.add(data[i][0].toString().trim());
      }
    }
    return [...projs].sort();
  } catch(e) { return []; }
}

// ---------------------------------------------------------
// --- FIXED: BOQ ADJUSTMENT ENDPOINTS ---
// ---------------------------------------------------------

function getBOQDataForProject(proj) {
  try {
    const sheet = SS.getSheetByName(SHEETS.BOQ);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let results = [];
    const sProj = proj.toString().trim().toLowerCase();
    
    // Assuming Row 1 and 2 might be headers, we start at i=1 but check structure
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === sProj) {
        results.push({
          rowIdx: i + 1, // Store the exact row number so we can edit it later
          phase: data[i][1] ? data[i][1].toString() : "",
          scope: data[i][2] ? data[i][2].toString() : "",
          subScope: data[i][3] ? data[i][3].toString() : "",
          subSubScope: data[i][4] ? data[i][4].toString() : "",
          item: data[i][5] ? data[i][5].toString() : "",
          unit: data[i][6] ? data[i][6].toString() : "",
          qty: data[i][7] || 0,
          labCost: data[i][10] || 0,
          matCost: data[i][11] || 0
        });
      }
    }
    return results;
  } catch (e) {
    return [];
  }
}

function updateBOQItemCost(rowIdx, newMat, newLab, newUnit, reason) {
  try {
    const boqSheet = SS.getSheetByName(SHEETS.BOQ);
    const logSheet = SS.getSheetByName(SHEETS.BOQ_LOGS);
    if (!boqSheet) return "Error: BOQ Database sheet missing.";
    
    // Fetch existing row data for logging purposes before overwriting
    const rowData = boqSheet.getRange(rowIdx, 1, 1, 12).getValues()[0];
    
    // Update the BOQ Sheet directly at the targeted row
    // Column G (7) = Unit | Column K (11) = Labor Cost | Column L (12) = Material Cost
    // Strip thousand separators before writing — same reason as addBOQItem.
    boqSheet.getRange(rowIdx, 7).setValue(newUnit);
    boqSheet.getRange(rowIdx, 11).setValue(toNumber_(newLab));
    boqSheet.getRange(rowIdx, 12).setValue(toNumber_(newMat));
    
    // Record the adjustment in the BOQ Logs
    if (logSheet) {
      logSheet.appendRow([
        new Date(),
        "Accounting System", // Added By
        "Adjustment", // Action
        rowData[0], // Project
        rowData[1], // Phase
        rowData[2], // Scope
        rowData[3], // Sub Scope
        rowData[4], // Sub-Sub Scope
        rowData[5], // Item Description
        newUnit,    // Unit
        rowData[7], // Qty (Unchanged by Accounting)
        newMat,     // Mat Cost
        newLab,     // Lab Cost
        reason      // Reason
      ]);
    }
    
    return "BOQ Updated Successfully.";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// Coerce a possibly comma-formatted currency/qty string ("1,500.00") into a
// plain Number. Sheet cells produced from JS-formatted strings used to land
// as text, and downstream parseFloat("1,500.00") would silently return 1.
function toNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const cleaned = v.toString().replace(/,/g, "").trim();
  if (cleaned === "") return 0;
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function addBOQItem(payload) {
  try {
    const boqSheet = SS.getSheetByName(SHEETS.BOQ);
    const logSheet = SS.getSheetByName(SHEETS.BOQ_LOGS);
    if (!boqSheet) return "Error: BOQ Database sheet missing.";

    const now = new Date();
    const uploader = (payload.uploader || "Accounting Portal").toString();

    // Numeric fields can arrive as formatted strings ("1,500.00") from the
    // currency-formatted inputs. Coerce them before write so the cells land
    // as real numbers, otherwise parseFloat downstream loses the thousands.
    const qtyNum = toNumber_(payload.qty);
    const labNum = toNumber_(payload.lab);
    const matNum = toNumber_(payload.mat);

    // Construct the new row spanning columns A–P (16 cols).
    //   A–H: identity (project / phase / scope / sub-scope / sub-sub-scope / item / unit / qty)
    //   I, J: legacy reserved
    //   K, L: labor + material cost
    //   M:    source marker — kept for parity with rows seeded by the older Accounting flow
    //   N:    Google Drive Link (empty when no file was uploaded — user can fill later)
    //   O:    Uploader (logged-in user's display name)
    //   P:    Date upload
    const newRow = [
      payload.project,           // A
      payload.phase,             // B
      payload.scope,             // C
      payload.subScope,          // D
      payload.subSubScope,       // E
      payload.desc,              // F
      payload.unit,              // G
      qtyNum,                    // H
      "",                        // I (reserved)
      "",                        // J (reserved)
      labNum,                    // K
      matNum,                    // L
      "Added via Accounting Portal", // M — source marker
      payload.driveLink || "",   // N — Google Drive Link
      uploader,                  // O — Uploader
      now                        // P — Date upload
    ];

    boqSheet.appendRow(newRow);

    // Record the addition in the BOQ Logs (use the actual uploader, not a generic label)
    if (logSheet) {
      logSheet.appendRow([
        now,
        uploader,
        "Added New Item",
        payload.project,
        payload.phase,
        payload.scope,
        payload.subScope,
        payload.subSubScope,
        payload.desc,
        payload.unit,
        qtyNum,
        matNum,
        labNum,
        payload.reason
      ]);
    }

    return "Item Added Successfully.";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// ==========================================
// --- EXPENSE LOGGING (NEW FEATURE) ---
// ==========================================

function getOrCreateExpenseSheet_() {
  let sheet = SS.getSheetByName(SHEETS.EXPENSE_LOGS);
  if (!sheet) {
    sheet = SS.insertSheet(SHEETS.EXPENSE_LOGS);
    const headers = ["Timestamp", "Expense ID", "Category", "Project Name", "Company Name", "Expense Type", "Description", "Amount", "Status", "Submitted By", "Submitter Email", "Notes", "Paid By", "Paid Date", "Payment Method", "Receipt Link"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Expense Type dropdown — reads from the unified Dropdowns sheet.
// Each entry: { name, allowedCategories: ["Project" | "Office"], sortOrder }
// Meta column on the Dropdowns sheet stores the allowed-categories CSV.
function getExpenseTypes() {
  return getDropdownRows("Expense Type").map(function (r) {
    const cats = (r.meta || "").split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    return {
      name: r.value,
      allowedCategories: cats.length ? cats : ["Project", "Office"],
      sortOrder: r.sortOrder
    };
  });
}

// --- EXPENSE ACTIVITY LOG (Audit trail tied to Expense Doc Ref ID) ---
function getOrCreateExpenseActivitySheet_() {
  let sheet = SS.getSheetByName(SHEETS.EXPENSE_ACTIVITY);
  if (!sheet) {
    sheet = SS.insertSheet(SHEETS.EXPENSE_ACTIVITY);
    const headers = ["Timestamp", "Expense ID", "Action", "Performed By", "Old Status", "New Status", "Notes"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logExpenseActivity_(expenseId, action, performedBy, oldStatus, newStatus, notes) {
  try {
    const sheet = getOrCreateExpenseActivitySheet_();
    sheet.appendRow([new Date(), expenseId || "", action || "", performedBy || "", oldStatus || "", newStatus || "", notes || ""]);
  } catch (e) {
    // Audit logging is best-effort; never break the user-facing operation.
  }
}

function getExpenseActivity(expenseId) {
  try {
    if (!expenseId) return [];
    const sheet = getOrCreateExpenseActivitySheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    const id = expenseId.toString().trim().toLowerCase();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      if ((data[i][1] || "").toString().trim().toLowerCase() === id) {
        let ts = data[i][0];
        let tsStr = "";
        if (ts) {
          let d = new Date(ts);
          tsStr = isNaN(d.getTime()) ? ts.toString() : formatFullDate(d);
        }
        out.push({
          timestamp: tsStr,
          expenseId: data[i][1] ? data[i][1].toString() : "",
          action: data[i][2] ? data[i][2].toString() : "",
          performedBy: data[i][3] ? data[i][3].toString() : "",
          oldStatus: data[i][4] ? data[i][4].toString() : "",
          newStatus: data[i][5] ? data[i][5].toString() : "",
          notes: data[i][6] ? data[i][6].toString() : ""
        });
      }
    }
    return out.reverse();
  } catch (e) {
    return [];
  }
}

function submitExpenseLog(payload, fileObj) {
  try {
    if (!payload) return { success: false, error: "Missing payload." };

    const category = (payload.category || "").toString().trim();
    if (category !== "Project" && category !== "Office") {
      return { success: false, error: "Category must be 'Project' or 'Office'." };
    }

    const projectName = (payload.projectName || "").toString().trim();
    const companyName = (payload.companyName || "").toString().trim();

    // ENFORCE MUTUAL EXCLUSIVITY (Business Logic Validation)
    if (projectName && companyName) {
      return { success: false, error: "An expense cannot be tied to BOTH a Project and a Company. Choose one." };
    }
    if (category === "Project" && !projectName) {
      return { success: false, error: "Project Name is required for Project Expense." };
    }
    if (category === "Office" && !companyName) {
      return { success: false, error: "Company Name is required for Office Expense." };
    }

    const expenseType = (payload.expenseType || "").toString().trim();
    if (!expenseType) return { success: false, error: "Expense Type is required." };

    // DB-driven validation: type must exist & be active, and must allow this category.
    // Admins control this via Dropdowns sheet (Category = "Expense Type"), no code changes needed.
    const availableTypes = getExpenseTypes();
    const typeDef = availableTypes.find(function (t) { return t.name.toLowerCase() === expenseType.toLowerCase(); });
    if (!typeDef) {
      return { success: false, error: "Invalid Expense Type. Choose one from the available list." };
    }
    if (typeDef.allowedCategories.indexOf(category) === -1) {
      return { success: false, error: '"' + typeDef.name + '" is not allowed for ' + category + ' expenses.' };
    }

    const description = (payload.description || "").toString().trim();
    if (!description) return { success: false, error: "Description is required." };

    const amount = parseFloat((payload.amount || "0").toString().replace(/,/g, ""));
    if (isNaN(amount) || amount <= 0) return { success: false, error: "Amount must be a positive number." };

    const sheet = getOrCreateExpenseSheet_();
    const now = new Date();
    const expenseId = "EXP-" + now.getTime().toString().slice(-6);

    // Optional receipt upload
    let receiptLink = "";
    if (fileObj && fileObj.data) {
      try {
        const folderName = category === "Project" ? projectName : companyName;
        const folder = getOrCreateDynamicFolder(folderName || "General", "Expenses", "Receipts");
        const extension = (fileObj.fileName || "").split('.').pop() || "pdf";
        const newFileName = `${expenseId}_${getFileDateSuffix()}.${extension}`;
        const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, newFileName);
        const created = safeDriveAction(() => folder.createFile(blob));
        receiptLink = created.getUrl();
      } catch (e) {
        // Receipt upload is best-effort; log silently
        console.error("Receipt upload failed: " + e.toString());
      }
    }

    sheet.appendRow([
      now,                                  // Timestamp
      expenseId,                            // Expense ID
      category,                             // Category (Project/Office)
      projectName,                          // Project Name
      companyName,                          // Company Name
      expenseType,                          // Expense Type
      description,                          // Description
      amount,                               // Amount
      "Pending",                            // Status (default)
      payload.submittedBy || "",            // Submitted By
      payload.submitterEmail || "",         // Submitter Email
      payload.notes || "",                  // Notes
      "",                                   // Paid By
      "",                                   // Paid Date
      "",                                   // Payment Method
      receiptLink                           // Receipt Link
    ]);

    // Audit trail entry linked to the Expense Doc Ref ID.
    const tag = category === "Project" ? ("Project: " + projectName) : ("Company: " + companyName);
    logExpenseActivity_(expenseId, "Created", payload.submittedBy || "", "", "Pending", tag + " | " + expenseType + " | ₱" + amount);

    return { success: true, id: expenseId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getMyExpenses(userIdentifier) {
  try {
    if (!userIdentifier) return [];
    const sheet = getOrCreateExpenseSheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
    const id = userIdentifier.toString().trim().toLowerCase();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      const submittedBy = (data[i][9] || "").toString().trim().toLowerCase();
      const email = (data[i][10] || "").toString().trim().toLowerCase();
      if (submittedBy === id || email === id) {
        let ts = data[i][0];
        let tsStr = "";
        if (ts) {
          let d = new Date(ts);
          tsStr = isNaN(d.getTime()) ? ts.toString() : formatFullDate(d);
        }
        out.push({
          row: i + 2,
          timestamp: tsStr,
          expenseId: data[i][1] ? data[i][1].toString() : "",
          category: data[i][2] ? data[i][2].toString() : "",
          projectName: data[i][3] ? data[i][3].toString() : "",
          companyName: data[i][4] ? data[i][4].toString() : "",
          expenseType: data[i][5] ? data[i][5].toString() : "",
          description: data[i][6] ? data[i][6].toString() : "",
          amount: parseFloat(data[i][7]) || 0,
          status: data[i][8] ? data[i][8].toString() : "Pending",
          notes: data[i][11] ? data[i][11].toString() : "",
          paidBy: data[i][12] ? data[i][12].toString() : "",
          paidDate: data[i][13] ? (function(v){ let d = new Date(v); return isNaN(d.getTime()) ? v.toString() : formatFullDate(d); })(data[i][13]) : "",
          paymentMethod: data[i][14] ? data[i][14].toString() : "",
          receiptLink: data[i][15] ? data[i][15].toString() : ""
        });
      }
    }
    return out.reverse();
  } catch (e) {
    return [];
  }
}

function getAllExpenses() {
  try {
    const sheet = getOrCreateExpenseSheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      let ts = data[i][0];
      let tsStr = "";
      if (ts) {
        let d = new Date(ts);
        tsStr = isNaN(d.getTime()) ? ts.toString() : formatFullDate(d);
      }
      out.push({
        row: i + 2,
        timestamp: tsStr,
        expenseId: data[i][1] ? data[i][1].toString() : "",
        category: data[i][2] ? data[i][2].toString() : "",
        projectName: data[i][3] ? data[i][3].toString() : "",
        companyName: data[i][4] ? data[i][4].toString() : "",
        expenseType: data[i][5] ? data[i][5].toString() : "",
        description: data[i][6] ? data[i][6].toString() : "",
        amount: parseFloat(data[i][7]) || 0,
        status: data[i][8] ? data[i][8].toString() : "Pending",
        submittedBy: data[i][9] ? data[i][9].toString() : "",
        submitterEmail: data[i][10] ? data[i][10].toString() : "",
        notes: data[i][11] ? data[i][11].toString() : "",
        paidBy: data[i][12] ? data[i][12].toString() : "",
        paidDate: data[i][13] ? (function(v){ let d = new Date(v); return isNaN(d.getTime()) ? v.toString() : formatFullDate(d); })(data[i][13]) : "",
        paymentMethod: data[i][14] ? data[i][14].toString() : "",
        receiptLink: data[i][15] ? data[i][15].toString() : ""
      });
    }
    return out.reverse();
  } catch (e) {
    return [];
  }
}

function updateExpenseStatus(rowIdx, newStatus, paidBy, paymentMethod) {
  try {
    const sheet = getOrCreateExpenseSheet_();
    const row = parseInt(rowIdx, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) return "Error: Invalid row reference.";

    // Accounting authorizes the status — accept any non-empty string. The two
    // well-known values ("Paid", "Pending") have special handling for payment
    // metadata; everything else (e.g., "Partially Paid", "Cancelled", "On Hold")
    // is stored as-is without touching the paid metadata.
    const status = (newStatus || "").toString().trim();
    if (!status) return "Error: Status cannot be empty.";

    const existing = sheet.getRange(row, 1, 1, 16).getValues()[0];
    const expenseId = (existing[1] || "").toString();
    const oldStatus = (existing[8] || "Pending").toString();

    sheet.getRange(row, 9).setValue(status); // Status column

    let action;
    let notesStr;
    if (status === "Paid") {
      sheet.getRange(row, 13).setValue(paidBy || "");      // Paid By
      sheet.getRange(row, 14).setValue(new Date());        // Paid Date
      sheet.getRange(row, 15).setValue(paymentMethod || ""); // Payment Method
      action = "Marked Paid";
      notesStr = "Method: " + (paymentMethod || "N/A");
    } else if (status === "Pending") {
      sheet.getRange(row, 13, 1, 3).setValues([["", "", ""]]);
      action = "Reverted to Pending";
      notesStr = "Payment metadata cleared.";
    } else {
      // Custom statuses: do not overwrite paid metadata; record who set it.
      action = "Status changed";
      notesStr = "Set to '" + status + "' by " + (paidBy || "unknown");
    }

    logExpenseActivity_(expenseId, action, paidBy || "", oldStatus, status, notesStr);

    return "Status updated to " + status + ".";
  } catch (e) {
    return "Error: " + e.message;
  }
}

// Accounting-side expense logging. Unlike submitExpenseLog (which forces status = "Pending"),
// this lets Accounting set any status at creation time, including Paid/Partially Paid/etc.
function submitAccountingExpense(payload, fileObj) {
  try {
    if (!payload) return { success: false, error: "Missing payload." };

    const category = (payload.category || "").toString().trim();
    if (category !== "Project" && category !== "Office") {
      return { success: false, error: "Category must be 'Project' or 'Office'." };
    }

    const projectName = (payload.projectName || "").toString().trim();
    const companyName = (payload.companyName || "").toString().trim();
    if (projectName && companyName) {
      return { success: false, error: "An expense cannot be tied to BOTH a Project and a Company." };
    }
    if (category === "Project" && !projectName) return { success: false, error: "Project Name is required for Project Expense." };
    if (category === "Office" && !companyName) return { success: false, error: "Company Name is required for Office Expense." };

    const expenseType = (payload.expenseType || "").toString().trim();
    if (!expenseType) return { success: false, error: "Expense Type is required." };

    const availableTypes = getExpenseTypes();
    const typeDef = availableTypes.find(function (t) { return t.name.toLowerCase() === expenseType.toLowerCase(); });
    if (!typeDef) return { success: false, error: "Invalid Expense Type." };
    if (typeDef.allowedCategories.indexOf(category) === -1) {
      return { success: false, error: '"' + typeDef.name + '" is not allowed for ' + category + ' expenses.' };
    }

    const description = (payload.description || "").toString().trim();
    if (!description) return { success: false, error: "Description is required." };

    const amount = parseFloat((payload.amount || "0").toString().replace(/,/g, ""));
    if (isNaN(amount) || amount <= 0) return { success: false, error: "Amount must be a positive number." };

    // Accounting may set ANY status (no Pending/Paid lock).
    const status = (payload.status || "Pending").toString().trim() || "Pending";

    const sheet = getOrCreateExpenseSheet_();
    const now = new Date();
    const expenseId = "EXP-" + now.getTime().toString().slice(-6);

    let receiptLink = "";
    if (fileObj && fileObj.data) {
      try {
        const folderName = category === "Project" ? projectName : companyName;
        const folder = getOrCreateDynamicFolder(folderName || "General", "Expenses", "Receipts");
        const extension = (fileObj.fileName || "").split('.').pop() || "pdf";
        const newFileName = expenseId + "_" + getFileDateSuffix() + "." + extension;
        const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, newFileName);
        const created = safeDriveAction(function () { return folder.createFile(blob); });
        receiptLink = created.getUrl();
      } catch (e) {
        console.error("Receipt upload failed: " + e.toString());
      }
    }

    // If Accounting marked it Paid at creation, capture paid metadata too.
    const isPaid = (status === "Paid");
    sheet.appendRow([
      now,
      expenseId,
      category,
      projectName,
      companyName,
      expenseType,
      description,
      amount,
      status,
      payload.submittedBy || "",
      payload.submitterEmail || "",
      payload.notes || "",
      isPaid ? (payload.paidBy || payload.submittedBy || "") : "",
      isPaid ? now : "",
      isPaid ? (payload.paymentMethod || "") : "",
      receiptLink
    ]);

    const tag = category === "Project" ? ("Project: " + projectName) : ("Company: " + companyName);
    logExpenseActivity_(expenseId, "Created (Accounting)", payload.submittedBy || "", "", status, tag + " | " + expenseType + " | ₱" + amount);

    return { success: true, id: expenseId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =============================================================================
// --- PETTY CASH MODULE (ported from Finance Portal) ---
// Self-contained block — does not modify any existing logic. Storage:
//   Project Database cols I/J/K  (per-project petty cash fund — see setup.js).
//                                 Row index in Project Database = projectId —
//                                 do not reorder Project Database rows.
//   PettyCash Expenses           (one row per line item)
//   PettyCash Replenishments     (Pending → Approved/Denied; or Direct = Approved on create)
//
// Permission model:
//   - Employees only see projects in their "Assigned Projects" column (Employee DB).
//   - Roles containing 'accounting' or 'finance' see all projects and approve requests.
// =============================================================================

const PETTY_CASH_LIMIT = 5000;  // Combined per-submission cap, in pesos.

function isAccountingRole_(role) {
  const r = (role || "").toString().toLowerCase();
  return r.indexOf("accounting") !== -1 || r.indexOf("finance") !== -1;
}

function lookupEmployee_(identifier) {
  if (!identifier) return null;
  const sheet = SS.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const input = identifier.toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const nm = (data[i][0] || "").toString().trim().toLowerCase();
    const em = (data[i][1] || "").toString().trim().toLowerCase();
    if (nm === input || em === input) {
      return {
        name: data[i][0] || "",
        email: data[i][1] || "",
        role: (data[i][2] || "").toString().toLowerCase().trim(),
        assignedProjects: (data[i][3] || "").toString().trim()
      };
    }
  }
  return null;
}

// Returns lowercase set of project names the employee is allowed to see. Empty
// set means "no restriction" — used for accounting/finance roles, who see all.
function getAllowedPettyCashProjects_(emp) {
  if (!emp) return null;
  if (isAccountingRole_(emp.role)) return null;
  if (!emp.assignedProjects) return new Set();
  return new Set(
    emp.assignedProjects
      .split(/[,;\n]/)
      .map(s => s.trim().toLowerCase())
      .filter(s => s)
  );
}

// Safety cap on petty cash sheet reads. A stray value or leftover formatting far
// down a sheet inflates getLastRow() to hundreds of thousands of rows; reading
// that whole range can stall the call for 30s+ (the bug that left Petty Cash
// stuck on the loading screen). Real petty cash sheets are nowhere near this big.
const PC_MAX_ROWS = 20000;

// Reads rows 2..N of a sheet, bounded by PC_MAX_ROWS. Returns [] for a missing
// or header-only sheet. Centralises the runaway-getLastRow() guard.
function pcReadRows_(sheet, startCol, numCols) {
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, PC_MAX_ROWS);
  return sheet.getRange(2, startCol, n, numCols).getValues();
}

// Reads the petty cash fund column from Project Database (cols A, I) and
// computes spent/balance live from the PettyCash Expenses ledger. Spent and
// balance are NEVER read from the sheet — the ledger is the only source of
// truth, so a stale or accidentally-edited cell on Project Database cannot
// desynchronize the math.
//
// A project is "petty cash enabled" iff col I (Petty Cash Allocated) is non-blank.
// Rows with a blank col I are excluded, matching the legacy behavior where the
// existence of a PettyCash Projects row signalled enablement.
//
// id returned to clients is the Project Database row index (1-based), which is
// also the value submitPettyCashExpense and the replenishment writers use to
// locate the row when bumping/reading col I.
function loadPettyCashProjects_() {
  const pSheet = SS.getSheetByName(SHEETS.PROJECTS);
  if (!pSheet) return [];

  // Map of projectName(lowercased) → total spent (bounded read).
  const eSheet = SS.getSheetByName(SHEETS.PETTY_CASH_EXPENSES);
  const spentMap = {};
  pcReadRows_(eSheet, 4, 3).forEach(r => { // cols D..F = Project, Item, Amount
    const key = (r[0] || "").toString().trim().toLowerCase();
    const amt = Number(r[2]) || 0;
    if (key) spentMap[key] = (spentMap[key] || 0) + amt;
  });

  // Read cols A..I (1..9). We only use A (name) and I (allocated).
  return pcReadRows_(pSheet, 1, 9).map((r, idx) => {
    const name = (r[0] || "").toString().trim();
    if (!name) return null;
    const allocCell = r[8]; // col I, zero-indexed
    if (allocCell === "" || allocCell === null || allocCell === undefined) return null;
    const allocated = Number(allocCell) || 0;
    const spent = spentMap[name.toLowerCase()] || 0;
    return {
      id: idx + 2,             // row index in Project Database
      name: name,
      allocated: allocated,
      spent: spent,
      balance: allocated - spent
    };
  }).filter(p => p !== null);
}

// Fast diagnostic — returns immediately without loading/serializing the data.
// Confirms the google.script.run bridge is alive, whether the active user is in
// the Employee Database, and the size of every sheet Petty Cash touches. If the
// portal ever stalls again, this pinpoints which sheet/condition is at fault.
function pettyCashPing(userIdentifier) {
  const out = { ok: true, time: new Date().toISOString(), maxRows: PC_MAX_ROWS };
  try {
    const emp = lookupEmployee_(userIdentifier);
    out.identifier = (userIdentifier || "").toString();
    out.userFound = !!emp;
    out.role = emp ? emp.role : "";
  } catch (e) { out.lookupError = e.toString(); }

  [["expenses", SHEETS.PETTY_CASH_EXPENSES],
   ["replenishments", SHEETS.PETTY_CASH_REPLENISHMENTS],
   ["projectDb", SHEETS.PROJECTS],
   ["employees", SHEETS.EMPLOYEES]].forEach(pair => {
    try {
      const sh = SS.getSheetByName(pair[1]);
      out[pair[0] + "Exists"] = !!sh;
      out[pair[0] + "LastRow"] = sh ? sh.getLastRow() : 0;
    } catch (e) { out[pair[0] + "Error"] = e.toString(); }
  });
  return out;
}

// Single entry point the front-end calls on portal load. Returns everything
// the petty cash UI needs in one trip. Role-aware filtering happens here so
// the front-end can render without further lookups.
function getPettyCashData(userIdentifier) {
  const diag = {};
  try {
    const emp = lookupEmployee_(userIdentifier);
    if (!emp) {
      return {
        projects: [], expenses: [], replenishments: [], allProjects: [], allKnownProjectNames: [], role: "",
        error: 'Your account ("' + (userIdentifier || "") + '") was not found in the Employee Database. '
             + 'Add it (Name or Email must match) and try again.'
      };
    }

    // Each section is isolated: a single slow/broken sheet degrades to empty data
    // for that section instead of hanging or failing the whole request. Every
    // read is bounded by pcReadRows_ so a runaway getLastRow() can't stall us.
    let allProjects = [];
    try { allProjects = loadPettyCashProjects_(); }
    catch (e) { diag.projectsError = e.toString(); }

    const allowed = getAllowedPettyCashProjects_(emp);
    const visibleProjects = (allowed === null)
      ? allProjects
      : allProjects.filter(p => allowed.has(p.name.toLowerCase()));

    // Build each row with EVERY cell explicitly converted to a primitive
    // (dates → ISO string, amounts → number, text → string) — the same defensive
    // shaping getAccountingQueue() uses on the tabs that already work, so no raw
    // Date/blank/object cell can sneak into the response.
    const cell = function (v) {
      if (v === null || v === undefined) return "";
      return (v instanceof Date) ? v.toISOString() : v.toString();
    };
    const num = function (v) { return Number(v) || 0; };

    // Non-accounting users only see their OWN rows. Accounting/finance roles
    // see everyone's rows so they can audit/approve. Previously the server
    // returned everyone's rows and the client filtered — which leaked petty
    // cash spend across employees over the wire.
    const isAcct = isAccountingRole_(emp.role);
    const ownerKey = (emp.name || "").toString().trim().toLowerCase();
    const ownerEmailKey = (emp.email || "").toString().trim().toLowerCase();

    let expenses = [];
    try {
      const eSheet = SS.getSheetByName(SHEETS.PETTY_CASH_EXPENSES);
      const rows = pcReadRows_(eSheet, 1, 8);
      expenses = rows.filter(function (r) {
        if (isAcct) return true;
        const rowUser = (r[2] || "").toString().trim().toLowerCase();
        return rowUser === ownerKey;
      }).reverse().map(function (r) {
        return [cell(r[0]), cell(r[1]), cell(r[2]), cell(r[3]), cell(r[4]), num(r[5]), num(r[6]), cell(r[7])];
      });
    } catch (e) { diag.expensesError = e.toString(); }

    let replenishments = [];
    try {
      const rSheet = SS.getSheetByName(SHEETS.PETTY_CASH_REPLENISHMENTS);
      const rows = pcReadRows_(rSheet, 1, 9);
      replenishments = rows.filter(function (r) {
        if (isAcct) return true;
        const reqName  = (r[2] || "").toString().trim().toLowerCase();
        const reqEmail = (r[3] || "").toString().trim().toLowerCase();
        return reqName === ownerKey || (ownerEmailKey && reqEmail === ownerEmailKey);
      }).reverse().map(function (r) {
        return [cell(r[0]), cell(r[1]), cell(r[2]), cell(r[3]), cell(r[4]), cell(r[5]), num(r[6]), cell(r[7]), cell(r[8])];
      });
    } catch (e) { diag.replenishmentsError = e.toString(); }

    // The Direct Replenishment dropdown is sourced ONLY from the PettyCash
    // Projects sheet (col A) — projects are mapped into that sheet manually,
    // since it is the only place a running balance can be maintained. We no
    // longer pull names from the master Project Database here (it was extra
    // sheet-read latency on every load for a list the UI no longer uses).
    let allKnownProjectNames = [];

    return {
      projects: visibleProjects,
      allProjects: isAccountingRole_(emp.role) ? allProjects : visibleProjects,
      allKnownProjectNames: allKnownProjectNames,
      expenses: expenses,
      replenishments: replenishments,
      role: emp.role,
      userName: emp.name,
      userEmail: emp.email,
      isAccounting: isAccountingRole_(emp.role),
      _diag: diag
    };
  } catch (e) {
    return { projects: [], expenses: [], replenishments: [], allProjects: [], allKnownProjectNames: [], role: "", error: e.toString(), _diag: diag };
  }
}

// STRING wrapper around getPettyCashData — this is what the front-end actually
// calls. Returning a JSON string instead of a live object sidesteps every
// google.script.run object-serialization quirk (raw Date cells, mixed types,
// NaN/Infinity in number columns) that can make a call silently hang with
// NEITHER the success nor the failure handler firing — the exact symptom Petty
// Cash was hitting while every other tab worked. JSON.stringify flattens Dates
// to strings and NaN to null, and a plain string always serializes cleanly.
// The front-end JSON.parse()s the result.
function getPettyCashDataString(userIdentifier) {
  try {
    return JSON.stringify(getPettyCashData(userIdentifier));
  } catch (e) {
    return JSON.stringify({
      projects: [], expenses: [], replenishments: [], allProjects: [], allKnownProjectNames: [],
      role: "", error: "Serialization failed: " + e.toString()
    });
  }
}

// Run this straight from the Apps Script editor (Run ▸ TEST_pettyCash) to prove
// whether the SERVER side works independently of the web-app bridge. It logs how
// long the call takes and how much data comes back. If this completes in the
// editor but the portal still hangs, the problem is the deployment/bridge, not
// the data; redeploy a new version. If it errors/stalls here, the log says why.
function TEST_pettyCash() {
  const who = Session.getActiveUser().getEmail();
  const t0 = new Date().getTime();
  const json = getPettyCashDataString(who);
  const ms = new Date().getTime() - t0;
  const obj = JSON.parse(json);
  Logger.log("user=%s  took=%sms  bytes=%s", who, ms, json.length);
  Logger.log("projects=%s expenses=%s replenishments=%s error=%s diag=%s",
    (obj.projects || []).length, (obj.expenses || []).length,
    (obj.replenishments || []).length, obj.error || "(none)", JSON.stringify(obj._diag || {}));
  return json;
}

// Notification helpers — best-effort. Silently no-op if Gmail is unavailable
// (e.g. user hasn't authorized GmailApp scope yet).
function notifyAccountingPC_(type, user, amount) {
  try {
    const url = ScriptApp.getService().getUrl();
    const empSheet = SS.getSheetByName(SHEETS.EMPLOYEES);
    if (!empSheet) return;
    const data = empSheet.getDataRange().getValues();
    const emails = [];
    for (let i = 1; i < data.length; i++) {
      const role = (data[i][2] || "").toString().toLowerCase();
      if (role.indexOf("accounting") !== -1 || role.indexOf("finance") !== -1) {
        if (data[i][1]) emails.push(data[i][1].toString());
      }
    }
    if (!emails.length) return;
    const subject = `New ${type} Request - ${user}`;
    const body = `A new ${type} request for ₱${Number(amount).toLocaleString()} has been submitted by ${user}.\n\nReview it in the portal: ${url}`;
    GmailApp.sendEmail(emails.join(','), subject, body);
  } catch (e) { /* ignore */ }
}

function notifyUserPC_(email, type, status, amount, projectName, attachedFile) {
  try {
    if (!email) return;
    const url = ScriptApp.getService().getUrl();
    const reqAmount = Number(amount).toLocaleString('en-US', {minimumFractionDigits: 2});
    const subject = `Update on your ${type}: ${status}`;
    let body = `Hello,\n\nYour ${type} request has been ${status}.\n\nAmount: PHP ${reqAmount}\nProject: ${projectName}\n\nPortal: ${url}\n`;
    let options = {};
    if (attachedFile) {
      options.attachments = [attachedFile.getAs(MimeType.PDF)];
      body += `\nThe approval document is attached.`;
    }
    GmailApp.sendEmail(email, subject, body, options);
  } catch (e) { /* ignore */ }
}

// Employee submits one or more petty cash line items against a project, with
// a receipt attached. Combined total of all line items must be ≤ PETTY_CASH_LIMIT
// and ≤ project remaining balance.
function submitPettyCashExpense(payload) {
  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, error: "Server busy — retry in a moment." };

    if (!payload || !payload.projectId || !payload.items || !payload.items.length) {
      return { success: false, error: "Missing project or line items." };
    }

    let totalAmount = 0;
    payload.items.forEach(i => totalAmount += Number(i.amount) || 0);
    if (totalAmount <= 0) return { success: false, error: "Total expense amount must be greater than zero." };
    if (totalAmount > PETTY_CASH_LIMIT) {
      return { success: false, error: `Combined transactions over ₱${PETTY_CASH_LIMIT.toLocaleString()} are not permitted in Petty Cash.` };
    }

    // Project Database is the source of truth. projectId = Project Database row.
    const pSheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!pSheet) return { success: false, error: "Project Database sheet missing — run setup()." };

    // Resolve project via row index — must exist.
    const projectId = Number(payload.projectId);
    if (!projectId || projectId < 2 || projectId > pSheet.getLastRow()) {
      return { success: false, error: "Invalid project." };
    }
    const projectName = (pSheet.getRange(projectId, 1).getValue() || "").toString().trim();
    if (!projectName) return { success: false, error: "Project not found." };

    // Compute current balance from authoritative source (sheet allocated − ledger total)
    const allProjects = loadPettyCashProjects_();
    const proj = allProjects.find(p => p.id === projectId);
    if (!proj) return { success: false, error: "Project not found." };
    let currentBalance = proj.balance;

    if (totalAmount > currentBalance) {
      return { success: false, error: "Insufficient petty cash balance for this project." };
    }

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd");
    const docRef = "PC-" + dateStr + "-" + Math.floor(1000 + Math.random() * 9000);

    // Upload receipt — receipt is required (matches Finance Portal behavior).
    let receiptUrl = "No Receipt";
    if (payload.fileData) {
      try {
        const folder = getOrCreateDynamicFolder(projectName, "Petty Cash", "Receipts");
        const ext = (payload.fileName || "").split('.').pop() || "pdf";
        let baseName = docRef;
        let finalName = baseName + "." + ext;
        let counter = 2;
        while (folder.getFilesByName(finalName).hasNext()) {
          finalName = baseName + "_" + counter + "." + ext;
          counter++;
        }
        const dataPart = payload.fileData.indexOf(',') !== -1
          ? payload.fileData.split(',')[1]
          : payload.fileData;
        const blob = Utilities.newBlob(Utilities.base64Decode(dataPart), payload.mimeType, finalName);
        receiptUrl = safeDriveAction(() => folder.createFile(blob)).getUrl();
      } catch (uploadErr) {
        return { success: false, error: "Receipt upload failed: " + uploadErr.toString() };
      }
    } else {
      return { success: false, error: "Receipt attachment is required for petty cash logging." };
    }

    const eSheet = SS.getSheetByName(SHEETS.PETTY_CASH_EXPENSES);
    const rowsToInsert = [];
    payload.items.forEach(item => {
      currentBalance -= Number(item.amount);
      rowsToInsert.push([
        timestamp, docRef, payload.user, projectName,
        item.item, Number(item.amount), currentBalance, receiptUrl
      ]);
    });

    if (rowsToInsert.length) {
      eSheet.getRange(eSheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length)
            .setValues(rowsToInsert);
    }

    return { success: true, docRef: docRef };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Employee files a replenishment request — accounting must approve.
function requestPettyCashReplenishment(payload) {
  try {
    if (!payload || !payload.projectId || !payload.projectName) {
      return { success: false, error: "Missing project information." };
    }

    const amountStr = (payload.amount || "").toString();
    if (/[^0-9.,]/.test(amountStr)) {
      return { success: false, error: "Invalid characters in amount. Only digits, commas, and a decimal point are allowed." };
    }
    const cleanAmount = Number(amountStr.replace(/,/g, ''));
    if (isNaN(cleanAmount) || cleanAmount <= 0) {
      return { success: false, error: "Amount must be greater than zero." };
    }

    const rSheet = SS.getSheetByName(SHEETS.PETTY_CASH_REPLENISHMENTS);
    const reqId = "REP-" + new Date().getTime().toString().slice(-6);
    rSheet.appendRow([
      new Date(), reqId, payload.user, payload.email,
      payload.projectId, payload.projectName, cleanAmount, 'Pending', ''
    ]);

    notifyAccountingPC_('Petty Cash Replenishment', payload.user, cleanAmount);
    return { success: true, reqId: reqId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Accounting bypass — log a direct replenishment that's auto-approved.
// Project Database is the master list of projects; this function refuses to
// replenish a project that isn't registered there. If the project exists but
// has never had petty cash enabled (col I blank), this call effectively enables
// it by writing the replenishment amount into col I.
function directPettyCashReplenish(payload) {
  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, error: "Server busy — retry in a moment." };

    if (!payload || !payload.projectName) {
      return { success: false, error: "Missing project name." };
    }
    const amountStr = (payload.amount || "").toString();
    if (/[^0-9.,]/.test(amountStr)) {
      return { success: false, error: "Invalid characters in amount." };
    }
    const cleanAmount = Number(amountStr.replace(/,/g, ''));
    if (isNaN(cleanAmount) || cleanAmount <= 0) {
      return { success: false, error: "Amount must be greater than zero." };
    }

    const pSheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!pSheet) return { success: false, error: "Project Database sheet missing — run setup()." };

    // Resolve project row in Project Database — by id if supplied, else by name.
    // Unlike before, we do NOT auto-create master project records here: a project
    // must already be registered in Project Database (via the BOQ Upload Portal or
    // a manual row) before petty cash can be replenished.
    const wantedName = payload.projectName.toString().trim();
    let projectId = Number(payload.projectId);
    if (!projectId || projectId < 2 || projectId > pSheet.getLastRow()) {
      const lookupRange = pSheet.getLastRow() >= 2
        ? pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 1).getValues()
        : [];
      const wantedKey = wantedName.toLowerCase();
      const foundIdx = lookupRange.findIndex(r => (r[0] || "").toString().trim().toLowerCase() === wantedKey);
      if (foundIdx === -1) {
        return { success: false, error: 'Project "' + wantedName + '" not found in Project Database. Register the project first.' };
      }
      projectId = foundIdx + 2;
    }

    // Col I (index 9) holds Petty Cash Allocated. Blank = first-time enable.
    const allocCell = pSheet.getRange(projectId, 9);
    const currentAlloc = Number(allocCell.getValue()) || 0;
    allocCell.setValue(currentAlloc + cleanAmount);

    const reqId = "DIR-" + new Date().getTime().toString().slice(-6);
    const reqAmount = cleanAmount.toLocaleString('en-US', {minimumFractionDigits: 2});

    // Best-effort PDF receipt — failure here doesn't abort the replenishment.
    let attachedFile = null;
    let fileUrl = "-";
    try {
      const folder = getOrCreateDynamicFolder(payload.projectName, "Petty Cash", "Replenishments");
      let baseName = reqId;
      let finalName = baseName + ".pdf";
      let counter = 2;
      while (folder.getFilesByName(finalName).hasNext()) {
        finalName = baseName + "_" + counter + ".pdf";
        counter++;
      }
      const html = `<div style="font-family: Arial, sans-serif; padding: 40px;">
        <h2>Direct Replenishment Confirmation</h2>
        <p><strong>Document ID:</strong> ${reqId}</p>
        <p><strong>Status:</strong> DIRECT LOG (APPROVED)</p>
        <hr>
        <p><strong>Amount:</strong> PHP ${reqAmount}</p>
        <p><strong>Project:</strong> ${payload.projectName}</p>
        <p><strong>Processed By:</strong> ${payload.user}</p>
      </div>`;
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(finalName);
      attachedFile = safeDriveAction(() => folder.createFile(blob));
      fileUrl = attachedFile.getUrl();
    } catch (pdfErr) {
      console.error("Direct replenishment PDF generation failed: " + pdfErr.toString());
    }

    const rSheet = SS.getSheetByName(SHEETS.PETTY_CASH_REPLENISHMENTS);
    rSheet.appendRow([
      new Date(), reqId, payload.user, payload.email,
      projectId, payload.projectName, cleanAmount, 'Approved', fileUrl
    ]);

    notifyUserPC_(payload.email, 'Direct Fund Replenishment', 'Logged & Approved', cleanAmount, payload.projectName, attachedFile);
    return { success: true, reqId: reqId };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Accounting batch action — approves or denies a list of pending replenishment
// request IDs. Approved requests bump the project allocation and email the
// requestor a PDF; denied ones are flagged and the requestor is emailed.
function processBatchPettyCashReplenish(reqIds, action, actorIdentifier) {
  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(20000)) return { success: false, error: "Server busy — retry in a moment." };
    if (action !== 'Approved' && action !== 'Denied') {
      return { success: false, error: "Invalid action." };
    }
    // actorIdentifier is accepted but not enforced — kept for future audit-trail use.

    const rSheet = SS.getSheetByName(SHEETS.PETTY_CASH_REPLENISHMENTS);
    const pSheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!rSheet || !pSheet) return { success: false, error: "Required sheets missing — run setup()." };

    const data = rSheet.getDataRange().getValues();

    (reqIds || []).forEach(reqId => {
      const rowIndex = data.findIndex(r => r[1] === reqId);
      if (rowIndex === -1 || data[rowIndex][7] !== 'Pending') return;

      rSheet.getRange(rowIndex + 1, 8).setValue(action);

      const requestorName  = data[rowIndex][2];
      const requestorEmail = data[rowIndex][3];
      const projectId      = Number(data[rowIndex][4]);
      const projectName    = data[rowIndex][5];
      const rawAmount      = Number(data[rowIndex][6]);
      const reqAmount      = rawAmount.toLocaleString('en-US', {minimumFractionDigits: 2});

      let attachedFile = null;
      let fileUrl = "-";

      if (action === 'Approved') {
        // Col I (9) of Project Database = Petty Cash Allocated.
        if (projectId && projectId >= 2 && projectId <= pSheet.getLastRow()) {
          const allocCell = pSheet.getRange(projectId, 9);
          const currentAlloc = Number(allocCell.getValue()) || 0;
          allocCell.setValue(currentAlloc + rawAmount);
        }

        try {
          const folder = getOrCreateDynamicFolder(projectName, "Petty Cash", "Replenishments");
          let baseName = reqId;
          let finalName = baseName + ".pdf";
          let counter = 2;
          while (folder.getFilesByName(finalName).hasNext()) {
            finalName = baseName + "_" + counter + ".pdf";
            counter++;
          }
          const html = `<div style="font-family: Arial, sans-serif; padding: 40px;">
            <h2>Approved Replenishment</h2>
            <p><strong>Document ID:</strong> ${reqId}</p>
            <p><strong>Status:</strong> APPROVED</p>
            <hr>
            <p><strong>Amount:</strong> PHP ${reqAmount}</p>
            <p><strong>Project:</strong> ${projectName}</p>
            <p><strong>Requestor:</strong> ${requestorName}</p>
          </div>`;
          const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(finalName);
          attachedFile = safeDriveAction(() => folder.createFile(blob));
          fileUrl = attachedFile.getUrl();
        } catch (pdfErr) {
          console.error("Replenishment PDF generation failed: " + pdfErr.toString());
        }
      }

      rSheet.getRange(rowIndex + 1, 9).setValue(fileUrl);
      notifyUserPC_(requestorEmail, 'Petty Cash Replenishment', action, rawAmount, projectName, attachedFile);
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// =====================================================================
// --- BOQ UPLOAD PORTAL (Accounting → BOQ Upload tab) ---
// Front-end picks an Excel/CSV BOQ workbook, the TCP, and the payment-term
// breakdown. processBoqPortalUpload converts the file to a temporary Google
// Sheet, parses every non-SUMMARY tab using the legacy "phase/scope/sub-scope"
// detection logic, and writes:
//   - one Project Database row (if not already registered)
//   - one Payment Terms Database row per milestone
//   - one BOQ Database row per detected line item
// The database is the ACTIVE spreadsheet (SS) — no external DB URL.
// =====================================================================

function processBoqPortalUpload(obj) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    const resource = {
      name: "Temp_BOQ_Upload_" + obj.fileName,
      mimeType: MimeType.GOOGLE_SHEETS
    };
    const tempFile = Drive.Files.create(resource, blob);
    const tempSheetId = tempFile.id;

    const tempSs = SpreadsheetApp.openById(tempSheetId);

    // Resolve uploader once and pass through so the worker doesn't need another
    // Session lookup (Session.getActiveUser().getEmail() returns "" for anonymous
    // web app users — fall back to the display name passed from the front end).
    const uploader = (Session.getActiveUser().getEmail() || obj.uploader || "Web Portal User").toString();

    const resultMessage = ingestBoqWorkbookIntoActiveDb_(tempSs, obj.tcp, obj.terms, uploader);

    // Cleanup: trash the temp file so the user's Drive isn't polluted.
    try { DriveApp.getFileById(tempSheetId).setTrashed(true); } catch (e) {}

    return resultMessage;
  } catch (e) {
    throw new Error("File processing failed: " + e.message);
  }
}

// Scans sheets in the uploaded workbook for the project metadata block
// (B3..B7). SUMMARY is skipped; the first sheet with a usable B3 wins.
function getBoqProjectMetadata_(ss) {
  const sheets = ss.getSheets();
  const meta = { title: "", owner: "", address: "", date: "", bidder: "" };
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (sheet.getName().toUpperCase() === "SUMMARY") continue;
    const title = sheet.getRange("B3").getValue();
    if (title && title !== "" && title !== "#REF!") {
      meta.title   = title;
      meta.owner   = sheet.getRange("B4").getValue();
      meta.address = sheet.getRange("B5").getValue();
      meta.date    = sheet.getRange("B6").getValue();
      meta.bidder  = sheet.getRange("B7").getValue();
      break;
    }
  }
  return meta;
}

function ingestBoqWorkbookIntoActiveDb_(ss, tcp, paymentTerms, uploader) {
  const dbSheet     = SS.getSheetByName(SHEETS.BOQ);
  const projDbSheet = SS.getSheetByName(SHEETS.PROJECTS);
  const termsDbSheet = SS.getSheetByName(SHEETS.PAYMENT_TERMS);

  if (!dbSheet || !projDbSheet || !termsDbSheet) {
    return "❌ Required database sheets are missing. Run setup() first.";
  }

  const meta = getBoqProjectMetadata_(ss);
  if (!meta.title) {
    return "❌ Error: Could not find Project Title in Row 3 (Col B) of any non-SUMMARY tab.";
  }

  // --- PART A: PROJECT REGISTRATION (skip if duplicate) ---
  let isDuplicate = false;
  const lastProjRow = projDbSheet.getLastRow();
  if (lastProjRow > 1) {
    const existingProjData = projDbSheet.getRange(2, 1, lastProjRow - 1, 5).getValues();
    for (let j = 0; j < existingProjData.length; j++) {
      if (existingProjData[j][0] == meta.title &&
          existingProjData[j][1] == meta.owner &&
          existingProjData[j][2] == meta.address &&
          existingProjData[j][4] == meta.bidder) {
        isDuplicate = true;
        break;
      }
    }
  }

  const uploadDate = new Date();
  let regStatus = "";
  if (!isDuplicate) {
    projDbSheet.appendRow([
      meta.title,
      meta.owner,
      meta.address,
      meta.date,
      meta.bidder,
      "Uploaded via Web Portal",
      uploadDate,
      tcp
    ]);
    regStatus = "✅ Project registered. ";
  } else {
    regStatus = "ℹ️ Project already existed. ";
  }

  // --- Save Payment Terms ---
  if (paymentTerms && paymentTerms.length > 0) {
    const termsDataToInsert = paymentTerms.map(t => [
      meta.title,
      t.milestone + "%",
      t.payment + "%",
      uploadDate,
      uploader
    ]);
    termsDbSheet.getRange(termsDbSheet.getLastRow() + 1, 1, termsDataToInsert.length, 5)
                .setValues(termsDataToInsert);
  }

  // --- PART B: BOQ ROW INGESTION (16-column layout, identical to BOQ.js) ---
  //
  // Column map mirrors the BOQ.js "BOQ Database" output cell-for-cell:
  //   A Project Title | B Sheet/Tab Name | C Phase | D Scope | E Sub Scope
  //   F Item (col C) | G Unit (col D) | H Qty (col E) | I col F | J col G
  //   K col H | L col I | M col J | N Source | O Uploader | P Date
  //
  // Workbook tabs start with metadata in rows 1-12 and BOQ line items from
  // row 13 onward across columns A-J. The classification rules below detect
  // phase / scope / sub-scope header rows by looking at numbering and
  // empty cost cells — identical to the BOQ.js logic.
  const rowsToInsert = [];

  ss.getSheets().forEach(sheet => {
    // CSV/single-tab imports name the tab after the temp file
    // ("Temp_BOQ_Upload_<name>.csv"). Strip our temp prefix and any
    // extension so column B holds only the raw tab title. Multi-tab .xlsx
    // workbooks keep their real tab names, so this is a no-op for them.
    const sheetName = sheet.getName()
      .replace(/^Temp_BOQ_Upload_/, "")
      .replace(/\.(csv|xlsx|xls)$/i, "");
    if (sheetName.toUpperCase() === "SUMMARY") return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 13) return;

    const projectTitle = sheet.getRange("B3").getValue();

    const data = sheet.getRange(13, 1, lastRow - 12, 10).getValues();

    let currentPhase = "";
    let currentScope = "";
    let currentSubScope = "";
    const sourceLink = "Uploaded via Web Portal";

    for (let i = 0; i < data.length; i++) {
      const colA = String(data[i][0]).trim();
      const colB = String(data[i][1]).trim();
      const colC = String(data[i][2]).trim();
      const colD_Unit = String(data[i][3]).trim();
      const colE_Qty  = String(data[i][4]).trim();
      const hasData = (colD_Unit !== "" || colE_Qty !== "");

      const isPhaseNumeric  = (colA !== "" && /^\d+$/.test(colA));
      const isPhaseTextOnly = (colA === "" && colB !== "" && !hasData && !/^\d+\.\d+/.test(colB) && !colB.includes("#REF!"));

      if (isPhaseNumeric || isPhaseTextOnly) {
        currentPhase = colB;
        currentScope = "";
        currentSubScope = "";
        continue;
      }

      const isRef = colB.includes("#REF!");
      const isSingleDecimal = /^\d+\.\d+$/.test(colB);

      if (!hasData) {
        if (isRef || isSingleDecimal) {
          currentScope = colC;
          currentSubScope = "";
        } else if (/^\d+\.\d+\.\d+$/.test(colB)) {
          currentSubScope = colC;
        }
      }

      if (hasData && colC !== "") {
        // Identical column mapping to BOQ.js — tab name goes in B, which
        // shifts Phase/Scope/Sub-Scope into C/D/E.
        rowsToInsert.push([
          projectTitle,        // A — Project Title
          sheetName,           // B — Sheet/Tab Name
          currentPhase,        // C — Phase
          currentScope,        // D — Scope
          currentSubScope,     // E — Sub Scope
          colC,                // F — Item Description (source col C)
          data[i][3],          // G — Unit (source col D)
          data[i][4],          // H — Qty (source col E)
          data[i][5],          // I — source col F
          data[i][6],          // J — source col G
          data[i][7],          // K — source col H
          data[i][8],          // L — source col I
          "",                  // M — left blank (no longer pulled from source col J)
          sourceLink,          // N — Source
          uploader,            // O — Uploader
          uploadDate           // P — Date Uploaded
        ]);
      }
    }
  });

  if (rowsToInsert.length > 0) {
    dbSheet.getRange(dbSheet.getLastRow() + 1, 1, rowsToInsert.length, 16).setValues(rowsToInsert);
    return regStatus + "\n🚀 Ingested " + rowsToInsert.length + " BOQ rows.";
  }
  return regStatus + "\n⚠️ No BOQ rows found.";
}

// =============================================================================
// CLIENT PAYMENTS (Collections / Accounts Receivable)
// -----------------------------------------------------------------------------
// Records money received FROM clients, tracked against each project's Total
// Contract Price (Project Database col H) and its milestone schedule (Payment
// Terms Database). The "Client Payments" sheet is the collections ledger; one
// row per payment. Milestone Pending/Partial/Fully-Paid status and the
// project-level Billed→Collected→Outstanding roll-ups are COMPUTED at read time
// from the ledger — nothing is double-stored. Milestones can be edited here and
// every edit/void is appended to "Client Payment Logs" with before/after values.
//
// Following the petty-cash lesson, getClientPaymentsData returns only formatted
// strings/numbers (never raw Date cells) so google.script.run never hangs on a
// serialization quirk.
// =============================================================================

const CLIENT_PAYMENT_HEADERS = ["Timestamp", "Payment ID", "Project Title", "Milestone Row", "Milestone", "Billing Ref", "Amount Due", "Amount Received", "Date Received", "Payment Method", "Bank Name", "Deposited-To Account", "Check Number", "Check Date", "Reference No", "OR Number", "Received By", "Remarks", "Status"];
const CLIENT_PAYMENT_LOG_HEADERS = ["Timestamp", "Entity", "Reference", "Action", "Field", "Old Value", "New Value", "Performed By", "Notes"];

function getOrCreateClientPaymentsSheet_() {
  let sheet = SS.getSheetByName(SHEETS.CLIENT_PAYMENTS);
  if (!sheet) {
    sheet = SS.insertSheet(SHEETS.CLIENT_PAYMENTS);
    sheet.getRange(1, 1, 1, CLIENT_PAYMENT_HEADERS.length).setValues([CLIENT_PAYMENT_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateClientPaymentLogSheet_() {
  let sheet = SS.getSheetByName(SHEETS.CLIENT_PAYMENT_LOGS);
  if (!sheet) {
    sheet = SS.insertSheet(SHEETS.CLIENT_PAYMENT_LOGS);
    sheet.getRange(1, 1, 1, CLIENT_PAYMENT_LOG_HEADERS.length).setValues([CLIENT_PAYMENT_LOG_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Best-effort audit write; never breaks the user-facing operation.
function logClientPaymentActivity_(entity, reference, action, field, oldVal, newVal, performedBy, notes) {
  try {
    const sheet = getOrCreateClientPaymentLogSheet_();
    sheet.appendRow([new Date(), entity || "", reference || "", action || "", field || "", oldVal === undefined || oldVal === null ? "" : oldVal, newVal === undefined || newVal === null ? "" : newVal, performedBy || "", notes || ""]);
  } catch (e) { /* audit is best-effort */ }
}

// Format a sheet cell that may hold a Date (or a string) into a display string.
function cpFormatDateCell_(v) {
  if (v === "" || v === null || v === undefined) return "";
  const d = new Date(v);
  return isNaN(d.getTime()) ? v.toString() : formatFullDate(d);
}

// Coerce "20%", "20", 20, "1,000.00" → a number. Blank/garbage → 0.
function cpToNumber_(v) {
  if (v === "" || v === null || v === undefined) return 0;
  const n = parseFloat(v.toString().replace(/[%,]/g, "").trim());
  return isNaN(n) ? 0 : n;
}

// Total Contract Price for a project (Project Database col H = index 7).
function getProjectTCP_(projectTitle) {
  const sheet = SS.getSheetByName(SHEETS.PROJECTS);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const want = projectTitle.toString().trim().toLowerCase();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === want) return cpToNumber_(data[i][7]);
  }
  return 0;
}

// Milestone rows for a project from Payment Terms Database, with their sheet row
// index (the stable key a payment links to — editing a milestone's % never moves
// its row, matching the row-index convention used elsewhere in this app).
function getProjectMilestones_(projectTitle) {
  const sheet = SS.getSheetByName(SHEETS.PAYMENT_TERMS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const want = projectTitle.toString().trim().toLowerCase();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === want) {
      out.push({ row: i + 2, milestonePct: cpToNumber_(data[i][1]), paymentPct: cpToNumber_(data[i][2]) });
    }
  }
  return out;
}

// Main read endpoint for the Client Payments tab.
//   - Always returns { projectList, summary }.
//   - When `project` is supplied, also returns { project, tcp, milestones, payments }.
// summary rolls up across ALL projects: total contract value, total collected
// (active payments only), total outstanding.
function getClientPaymentsData(project) {
  try {
    const projectList = getProjectList();

    // Read the whole collections ledger once.
    const sheet = getOrCreateClientPaymentsSheet_();
    const payments = [];
    if (sheet.getLastRow() >= 2) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CLIENT_PAYMENT_HEADERS.length).getValues();
      for (let i = 0; i < data.length; i++) {
        if (!(data[i][1] || "").toString().trim() && !(data[i][2] || "").toString().trim()) continue;
        payments.push({
          row: i + 2,
          timestamp: cpFormatDateCell_(data[i][0]),
          paymentId: (data[i][1] || "").toString(),
          project: (data[i][2] || "").toString(),
          milestoneRow: parseInt(data[i][3], 10) || 0,
          milestone: (data[i][4] || "").toString(),
          billingRef: (data[i][5] || "").toString(),
          amountDue: cpToNumber_(data[i][6]),
          amountReceived: cpToNumber_(data[i][7]),
          dateReceived: cpFormatDateCell_(data[i][8]),
          paymentMethod: (data[i][9] || "").toString(),
          bankName: (data[i][10] || "").toString(),
          depositedTo: (data[i][11] || "").toString(),
          checkNumber: (data[i][12] || "").toString(),
          checkDate: cpFormatDateCell_(data[i][13]),
          referenceNo: (data[i][14] || "").toString(),
          orNumber: (data[i][15] || "").toString(),
          receivedBy: (data[i][16] || "").toString(),
          remarks: (data[i][17] || "").toString(),
          status: (data[i][18] || "Active").toString()
        });
      }
    }

    // Global roll-up across all projects.
    let totalContractValue = 0;
    const projSheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (projSheet && projSheet.getLastRow() >= 2) {
      const pd = projSheet.getRange(2, 1, projSheet.getLastRow() - 1, 8).getValues();
      for (let i = 0; i < pd.length; i++) {
        if ((pd[i][0] || "").toString().trim()) totalContractValue += cpToNumber_(pd[i][7]);
      }
    }
    let totalCollected = 0;
    payments.forEach(function (p) { if (p.status !== "Voided") totalCollected += p.amountReceived; });

    const result = {
      projectList: projectList,
      summary: {
        totalContractValue: totalContractValue,
        totalCollected: totalCollected,
        totalOutstanding: totalContractValue - totalCollected
      }
    };

    if (project) {
      const proj = project.toString().trim();
      const tcp = getProjectTCP_(proj);
      const ms = getProjectMilestones_(proj);
      const projPayments = payments.filter(function (p) { return p.project.trim().toLowerCase() === proj.toLowerCase(); });

      const milestones = ms.map(function (m, idx) {
        const due = tcp * (m.paymentPct / 100);
        let collected = 0;
        projPayments.forEach(function (p) { if (p.status !== "Voided" && p.milestoneRow === m.row) collected += p.amountReceived; });
        let st = "Pending";
        if (due > 0 && collected >= due - 0.005) st = "Fully Paid";
        else if (collected > 0) st = "Partial";
        return {
          row: m.row,
          index: idx + 1,
          label: "Milestone " + (idx + 1),
          milestonePct: m.milestonePct,
          paymentPct: m.paymentPct,
          amountDue: due,
          collected: collected,
          balance: due - collected,
          status: st
        };
      });

      result.project = proj;
      result.tcp = tcp;
      result.milestones = milestones;
      result.payments = projPayments;
    }

    return result;
  } catch (e) {
    return { projectList: [], summary: { totalContractValue: 0, totalCollected: 0, totalOutstanding: 0 }, error: e.toString() };
  }
}

// Record one client payment. Returns { success, id } or { success:false, error }.
function recordClientPayment(payload) {
  try {
    if (!payload) return { success: false, error: "Missing payload." };
    const project = (payload.project || "").toString().trim();
    if (!project) return { success: false, error: "Project is required." };

    const amountReceived = cpToNumber_(payload.amountReceived);
    if (amountReceived <= 0) return { success: false, error: "Amount Received must be a positive number." };

    const method = (payload.paymentMethod || "").toString().trim();
    if (!method) return { success: false, error: "Payment Method is required." };

    if (!payload.dateReceived) return { success: false, error: "Date Received is required." };
    const dateReceived = new Date(payload.dateReceived);
    if (isNaN(dateReceived.getTime())) return { success: false, error: "Date Received is invalid." };

    const sheet = getOrCreateClientPaymentsSheet_();
    const now = new Date();
    const paymentId = "CP-" + now.getTime().toString().slice(-6);
    const checkDate = payload.checkDate ? new Date(payload.checkDate) : "";

    sheet.appendRow([
      now,                                              // Timestamp
      paymentId,                                        // Payment ID
      project,                                          // Project Title
      parseInt(payload.milestoneRow, 10) || "",         // Milestone Row (blank = general)
      (payload.milestone || "").toString(),             // Milestone (label snapshot)
      (payload.billingRef || "").toString(),            // Billing Ref
      cpToNumber_(payload.amountDue),                   // Amount Due (snapshot, editable)
      amountReceived,                                   // Amount Received
      dateReceived,                                     // Date Received
      method,                                           // Payment Method
      (payload.bankName || "").toString(),              // Bank Name
      (payload.depositedTo || "").toString(),           // Deposited-To Account
      (payload.checkNumber || "").toString(),           // Check Number
      (checkDate && !isNaN(checkDate.getTime())) ? checkDate : "", // Check Date
      (payload.referenceNo || "").toString(),           // Reference No
      (payload.orNumber || "").toString(),              // OR Number
      (payload.receivedBy || "").toString(),            // Received By
      (payload.remarks || "").toString(),               // Remarks
      "Active"                                           // Status
    ]);

    logClientPaymentActivity_("Payment", paymentId, "Recorded", "Amount Received", "", amountReceived, payload.receivedBy || "", project + (payload.milestone ? (" | " + payload.milestone) : ""));
    return { success: true, id: paymentId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Edit a project's milestone (Payment Terms Database row). Writes back only the
// changed fields and logs before/after values + who edited.
function updateClientMilestone(payload) {
  try {
    if (!payload || !payload.row) return { success: false, error: "Missing milestone row reference." };
    const sheet = SS.getSheetByName(SHEETS.PAYMENT_TERMS);
    if (!sheet) return { success: false, error: "Payment Terms Database not found." };
    const row = parseInt(payload.row, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) return { success: false, error: "Invalid milestone row reference." };

    const existing = sheet.getRange(row, 1, 1, 3).getValues()[0];
    const project = (existing[0] || "").toString();
    const oldMilestone = (existing[1] || "").toString();
    const oldPayment = (existing[2] || "").toString();
    const performedBy = (payload.performedBy || "").toString();

    const newMilestone = (payload.milestonePct === undefined || payload.milestonePct === null || payload.milestonePct.toString().trim() === "") ? oldMilestone : payload.milestonePct.toString().trim();
    const newPayment = (payload.paymentPct === undefined || payload.paymentPct === null || payload.paymentPct.toString().trim() === "") ? oldPayment : payload.paymentPct.toString().trim();

    let changed = false;
    if (newMilestone !== oldMilestone) {
      sheet.getRange(row, 2).setValue(newMilestone);
      logClientPaymentActivity_("Milestone", project + " (row " + row + ")", "Edited", "Milestone %", oldMilestone, newMilestone, performedBy, "");
      changed = true;
    }
    if (newPayment !== oldPayment) {
      sheet.getRange(row, 3).setValue(newPayment);
      logClientPaymentActivity_("Milestone", project + " (row " + row + ")", "Edited", "Payment %", oldPayment, newPayment, performedBy, "");
      changed = true;
    }
    return { success: true, changed: changed };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Void a client payment (kept on the sheet, flagged so it drops out of totals).
function voidClientPayment(payload) {
  try {
    if (!payload || !payload.row) return { success: false, error: "Missing row reference." };
    const sheet = getOrCreateClientPaymentsSheet_();
    const row = parseInt(payload.row, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) return { success: false, error: "Invalid row reference." };

    const existing = sheet.getRange(row, 1, 1, CLIENT_PAYMENT_HEADERS.length).getValues()[0];
    const paymentId = (existing[1] || "").toString();
    const oldStatus = (existing[18] || "Active").toString();
    if (oldStatus === "Voided") return { success: false, error: "This payment is already voided." };

    sheet.getRange(row, CLIENT_PAYMENT_HEADERS.length).setValue("Voided"); // Status is the last column
    logClientPaymentActivity_("Payment", paymentId, "Voided", "Status", oldStatus, "Voided", payload.performedBy || "", payload.reason || "");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Audit-trail viewer for the Client Payments tab. Optionally filter by project
// (matches the "Reference" column prefix or any payment that belongs to it).
function getClientPaymentAuditLog(project) {
  try {
    const sheet = getOrCreateClientPaymentLogSheet_();
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CLIENT_PAYMENT_LOG_HEADERS.length).getValues();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      out.push({
        timestamp: cpFormatDateCell_(data[i][0]),
        entity: (data[i][1] || "").toString(),
        reference: (data[i][2] || "").toString(),
        action: (data[i][3] || "").toString(),
        field: (data[i][4] || "").toString(),
        oldValue: (data[i][5] === null || data[i][5] === undefined) ? "" : data[i][5].toString(),
        newValue: (data[i][6] === null || data[i][6] === undefined) ? "" : data[i][6].toString(),
        performedBy: (data[i][7] || "").toString(),
        notes: (data[i][8] || "").toString()
      });
    }
    return out.reverse();
  } catch (e) {
    return [];
  }
}

// =============================================================================
// PROJECT & MILESTONE EDITOR (Accounting → BOQ Upload → "Edit Project & Milestones")
// Lets Accounting edit a project already in Project Database (Owner / Address /
// Bidder / Total Contract Price) and add milestones to projects that have none
// yet. Strictly additive — reuses getProjectList, getProjectTCP_,
// getProjectMilestones_, updateClientMilestone and the logClientPaymentActivity_
// audit writer (so every change lands in the "Client Payment Logs" sheet with
// before/after values). Returns only formatted primitives — never raw Date
// cells — so google.script.run never hangs (the petty-cash serialization lesson).
// =============================================================================

// Read endpoint for the editor: a project's editable metadata + TCP + its
// existing milestone rows. { found:false } when the project isn't in the DB.
function getProjectEditData(project) {
  const empty = { found: false, row: 0, project: "", owner: "", address: "", bidder: "", tcp: 0, milestones: [] };
  try {
    if (!project) return empty;
    const sheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!sheet || sheet.getLastRow() < 2) return empty;
    const want = project.toString().trim().toLowerCase();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    let found = null, foundRow = 0;
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim().toLowerCase() === want) {
        found = data[i];
        foundRow = i + 2;
        break;
      }
    }
    if (!found) return empty;

    const ms = getProjectMilestones_((found[0] || project).toString());
    const milestones = ms.map(function (m, idx) {
      return { row: m.row, index: idx + 1, milestonePct: m.milestonePct, paymentPct: m.paymentPct };
    });

    return {
      found: true,
      row: foundRow,
      project: (found[0] || "").toString(),
      owner: (found[1] || "").toString(),
      address: (found[2] || "").toString(),
      bidder: (found[4] || "").toString(),
      tcp: cpToNumber_(found[7]),
      milestones: milestones
    };
  } catch (e) {
    empty.error = e.toString();
    return empty;
  }
}

// Edit an existing project's Owner / Address / Bidder / Total Contract Price.
// Project Title is the key linking BOQ rows, milestones and payments, so it is
// NOT editable here. Writes only changed cells (cols B/C/E/H — never touches the
// petty-cash cols I/J/K) and logs each changed field before/after.
function updateProjectDetails(payload) {
  try {
    if (!payload || !payload.project) return { success: false, error: "Project is required." };
    const sheet = SS.getSheetByName(SHEETS.PROJECTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Project Database not found." };

    const want = payload.project.toString().trim().toLowerCase();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    let row = -1, existing = null;
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim().toLowerCase() === want) { row = i + 2; existing = data[i]; break; }
    }
    if (row < 0) return { success: false, error: "Project not found in Project Database." };

    const projTitle = (existing[0] || "").toString();
    const performedBy = (payload.performedBy || "").toString();
    let changed = false;

    // Text metadata fields: Owner (col B/2), Address (col C/3), Bidder (col E/5).
    const fields = [
      { col: 2, key: "owner",   label: "Owner"   },
      { col: 3, key: "address", label: "Address" },
      { col: 5, key: "bidder",  label: "Bidder"  }
    ];
    fields.forEach(function (f) {
      if (payload[f.key] === undefined || payload[f.key] === null) return;
      const oldVal = (existing[f.col - 1] || "").toString();
      const newVal = payload[f.key].toString();
      if (newVal !== oldVal) {
        sheet.getRange(row, f.col).setValue(newVal);
        logClientPaymentActivity_("Project", projTitle, "Edited", f.label, oldVal, newVal, performedBy, "");
        changed = true;
      }
    });

    // Total Contract Price (col H/8).
    if (payload.tcp !== undefined && payload.tcp !== null && payload.tcp.toString().trim() !== "") {
      const oldTcp = cpToNumber_(existing[7]);
      const newTcp = cpToNumber_(payload.tcp);
      if (Math.abs(newTcp - oldTcp) > 0.005) {
        sheet.getRange(row, 8).setValue(newTcp);
        logClientPaymentActivity_("Project", projTitle, "Edited", "Total Contract Price", oldTcp, newTcp, performedBy, "");
        changed = true;
      }
    }

    return { success: true, changed: changed };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Append one or more milestone rows to Payment Terms Database for a project that
// has none yet (or to extend an existing schedule). Rows are written in the same
// "20%" string format the BOQ upload uses, so the Client Payments tab reads them
// back identically. Each addition is logged.
function addProjectMilestones(payload) {
  try {
    if (!payload || !payload.project) return { success: false, error: "Project is required." };
    const list = (payload.milestones && payload.milestones.length) ? payload.milestones : [];
    if (!list.length) return { success: false, error: "No milestones to add." };

    const sheet = SS.getSheetByName(SHEETS.PAYMENT_TERMS);
    if (!sheet) return { success: false, error: "Payment Terms Database not found." };

    const project = payload.project.toString().trim();
    const performedBy = (payload.performedBy || "").toString();
    const now = new Date();
    const rows = [];
    for (let i = 0; i < list.length; i++) {
      const mPct = cpToNumber_(list[i].milestone);
      const pPct = cpToNumber_(list[i].payment);
      if (pPct <= 0) return { success: false, error: "Each milestone needs a Payment % greater than 0." };
      rows.push([project, mPct + "%", pPct + "%", now, performedBy || "Accounting Portal"]);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    for (let i = 0; i < rows.length; i++) {
      logClientPaymentActivity_("Milestone", project, "Added", "Milestone / Payment %", "", rows[i][1] + " / " + rows[i][2], performedBy, "");
    }

    return { success: true, added: rows.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}