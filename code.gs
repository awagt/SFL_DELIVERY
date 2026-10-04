/* =====================================================================
   SFL Dispatch Tracker — Google Sheets backend (v5 — clean rebuild)

   This script is pinned directly to YOUR spreadsheet by its ID (see
   SPREADSHEET_ID below), so it works no matter how the Apps Script
   project itself was created — it does not rely on being "bound" to
   the sheet. This was the actual cause of the earlier "not working" /
   404 problems: a script that only works via getActiveSpreadsheet()
   silently fails the moment it's deployed as a standalone Web App.

   There is NO hidden database layer anymore. The 5 tabs in your
   spreadsheet ARE the database — every read and write goes straight
   to them:

     "SFL Dispatch"      — Order Number, Batch ID, Dispatch Date,
                            Dispatcher, Delivery Agent, Dispatch
                            Location, Status
                            (one row per order; several rows can share
                            a Batch ID when scanned together)
     "SFL Returns"       — Order Number, SKU, QTY, Return Date,
                            Receiver, Return Agent, Matched to Dispatch
                            (one row per SKU returned)
     "USER"              — EMP ID, NAME, USERNAME, PIN, USER ACCESS,
                            USER STATUS
                            (USER ACCESS is "admin", "staff", or
                            "guest" — case doesn't matter)
     "DELIVERY AGENT"    — DRIVER (master list, one name per row)
     "DISPATCH LOCATION" — LOCATION (master list, one name per row)

   Plus one new tab this script manages for you automatically:

     "LOG" — Timestamp, User, Role, Action, Details
             (the audit trail shown on the app's admin-only Log tab)

   PERMISSIONS (enforced by the app's UI, same as before):
     - admin: add, view, edit, delete — everything
     - staff: add new dispatches/returns, view everything — cannot
              edit or delete
     - guest: view only — cannot add, edit, or delete anything

   Login: the app sends a username + password to the "login" action
   below. This script checks it against the USER sheet and replies
   with only {success, name, role} — the PIN itself is never sent
   back to the browser, and nothing is ever hardcoded in the HTML file.
   ===================================================================== */

const SPREADSHEET_ID = '17R9O_GjF_aMJswEJ-fDMDzV8WZgjfOydQ2Vc7X7XsbI';

const DISPATCH_SHEET = 'SFL Dispatch';
const DISPATCH_HEADERS = ['Order Number', 'Batch ID', 'Dispatch Date', 'Dispatcher', 'Delivery Agent', 'Dispatch Location', 'Status'];

const RETURNS_SHEET = 'SFL Returns';
const RETURNS_HEADERS = ['Order Number', 'SKU', 'QTY', 'Return Date', 'Receiver', 'Return Agent', 'Matched to Dispatch'];

const USER_SHEET = 'USER';
const USER_HEADERS = ['EMP ID', 'NAME', 'USERNAME', 'PIN', 'USER ACCESS', 'USER STATUS'];

const AGENT_SHEET = 'DELIVERY AGENT';
const AGENT_HEADERS = ['DRIVER'];

const LOCATION_SHEET = 'DISPATCH LOCATION';
const LOCATION_HEADERS = ['LOCATION'];

const LOG_SHEET = 'LOG';
const LOG_HEADERS = ['Timestamp', 'User', 'Role', 'Action', 'Details'];

const MAGENTO_SHEET = 'MAGENTO DATA';
const MAGENTO_HEADERS = ['Data', 'Imported At'];

function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet(name, headers) {
  const spreadsheet = ss();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/* ---------- Routing ---------- */

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'login') return jsonOutput(checkLogin(e.parameter.username, e.parameter.password));
    if (action === 'listUsers') return jsonOutput({ items: listUsersPublic() });
    if (action === 'listDispatch') return jsonOutput({ items: listDispatch() });
    if (action === 'listReturns') return jsonOutput({ items: listReturns() });
    if (action === 'listAgents') return jsonOutput({ items: listSingleColumn(AGENT_SHEET, AGENT_HEADERS) });
    if (action === 'listLocations') return jsonOutput({ items: listSingleColumn(LOCATION_SHEET, LOCATION_HEADERS) });
    if (action === 'listLog') return jsonOutput({ items: listLog() });
    if (action === 'getMagentoIndex') return jsonOutput(getMagentoIndex());
    return jsonOutput({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'addDispatch') addDispatch(data.batch);
    else if (data.action === 'updateDispatch') updateDispatch(data.batchId, data.patch);
    else if (data.action === 'deleteDispatch') deleteDispatch(data.batchId);
    else if (data.action === 'addReturn') addReturn(data.ret);
    else if (data.action === 'updateReturn') updateReturn(data.identity, data.patch);
    else if (data.action === 'deleteReturn') deleteReturn(data.identity);
    else if (data.action === 'upsertUser') upsertUser(data.user);
    else if (data.action === 'deleteUser') removeRowsByColumnValue(USER_SHEET, 'USERNAME', data.username);
    else if (data.action === 'addLog') addLog(data.entry);
    else if (data.action === 'setMagentoIndex') setMagentoIndex(data.data, data.importedAt);
    else return jsonOutput({ error: 'Unknown action: ' + data.action });
    return jsonOutput({ status: 'ok' });
  } catch (err) {
    return jsonOutput({ error: String(err) });
  }
}

/* ---------- Shared helpers ---------- */

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach(function (h, c) { obj[h] = data[i][c]; });
    rows.push(obj);
  }
  return rows;
}

function removeRowsByColumnValue(sheetName, headerName, value) {
  const sheet = ss().getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return;
  const colIndex = data[0].indexOf(headerName);
  if (colIndex === -1) return;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][colIndex] === value) sheet.deleteRow(i + 1);
  }
}

function rememberIfNew(sheetName, headers, value) {
  if (!value) return;
  const sheet = getOrCreateSheet(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === value) return; // already known
  }
  sheet.appendRow([value]);
}

function listSingleColumn(sheetName, headers) {
  const sheet = getOrCreateSheet(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) items.push(data[i][0]);
  }
  return items;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- SFL Dispatch ---------- */

function listDispatch() {
  const sheet = getOrCreateSheet(DISPATCH_SHEET, DISPATCH_HEADERS);
  return sheetToObjects(sheet).map(function (r) {
    return {
      orderNumber: r['Order Number'],
      batchId: r['Batch ID'],
      dispatchDate: r['Dispatch Date'],
      dispatcher: r['Dispatcher'],
      agent: r['Delivery Agent'],
      location: r['Dispatch Location'],
      status: r['Status']
    };
  });
}

function addDispatch(batch) {
  // batch: { batchId, dispatchDate, dispatcher, agent, location, status, orders: [...] }
  const sheet = getOrCreateSheet(DISPATCH_SHEET, DISPATCH_HEADERS);
  (batch.orders || []).forEach(function (order) {
    sheet.appendRow([order, batch.batchId, batch.dispatchDate, batch.dispatcher, batch.agent, batch.location, batch.status]);
  });
  rememberIfNew(AGENT_SHEET, AGENT_HEADERS, batch.agent);
  rememberIfNew(LOCATION_SHEET, LOCATION_HEADERS, batch.location);
}

function updateDispatch(batchId, patch) {
  // patch: object of { 'Dispatcher': 'x', 'Delivery Agent': 'y', 'Status': 'delivered', ... }
  const sheet = getOrCreateSheet(DISPATCH_SHEET, DISPATCH_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const batchCol = headers.indexOf('Batch ID');
  for (let i = 1; i < data.length; i++) {
    if (data[i][batchCol] === batchId) {
      Object.keys(patch).forEach(function (key) {
        const col = headers.indexOf(key);
        if (col > -1) sheet.getRange(i + 1, col + 1).setValue(patch[key]);
      });
    }
  }
  if (patch['Delivery Agent']) rememberIfNew(AGENT_SHEET, AGENT_HEADERS, patch['Delivery Agent']);
  if (patch['Dispatch Location']) rememberIfNew(LOCATION_SHEET, LOCATION_HEADERS, patch['Dispatch Location']);
}

function deleteDispatch(batchId) {
  removeRowsByColumnValue(DISPATCH_SHEET, 'Batch ID', batchId);
}

/* ---------- SFL Returns ---------- */

function listReturns() {
  const sheet = getOrCreateSheet(RETURNS_SHEET, RETURNS_HEADERS);
  return sheetToObjects(sheet).map(function (r) {
    return {
      orderNumber: r['Order Number'],
      sku: r['SKU'],
      qty: r['QTY'],
      returnDate: r['Return Date'],
      receiver: r['Receiver'],
      returnAgent: r['Return Agent'],
      matched: r['Matched to Dispatch'] === 'Yes'
    };
  });
}

function addReturn(ret) {
  // ret: { orderNumber, returnDate, receiver, returnAgent, matched, skus: [{sku, qty}] }
  const sheet = getOrCreateSheet(RETURNS_SHEET, RETURNS_HEADERS);
  (ret.skus || []).forEach(function (entry) {
    sheet.appendRow([ret.orderNumber, entry.sku, entry.qty, ret.returnDate, ret.receiver, ret.returnAgent, ret.matched ? 'Yes' : 'No']);
  });
}

function findReturnRows(identity) {
  // identity: { orderNumber, returnDate, receiver } — the three columns that
  // together identify "one logged return event" since there's no ID column.
  const sheet = getOrCreateSheet(RETURNS_SHEET, RETURNS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = { order: headers.indexOf('Order Number'), date: headers.indexOf('Return Date'), receiver: headers.indexOf('Receiver') };
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col.order]) === String(identity.orderNumber) &&
        String(data[i][col.date]) === String(identity.returnDate) &&
        String(data[i][col.receiver]) === String(identity.receiver)) {
      rows.push(i + 1); // 1-based sheet row number
    }
  }
  return { sheet, headers, rows };
}

function updateReturn(identity, patch) {
  // patch: { 'Receiver': 'x', 'Return Agent': 'y', 'Return Date': 'z' }
  const found = findReturnRows(identity);
  found.rows.forEach(function (rowNum) {
    Object.keys(patch).forEach(function (key) {
      const col = found.headers.indexOf(key);
      if (col > -1) found.sheet.getRange(rowNum, col + 1).setValue(patch[key]);
    });
  });
}

function deleteReturn(identity) {
  const found = findReturnRows(identity);
  // delete bottom-up so row numbers stay valid while removing
  found.rows.sort(function (a, b) { return b - a; }).forEach(function (rowNum) {
    found.sheet.deleteRow(rowNum);
  });
}

/* ---------- Accounts & login (USER sheet is the ONLY source of truth) ---------- */

function getOrCreateUserSheet() {
  const sheet = getOrCreateSheet(USER_SHEET, USER_HEADERS);
  // Seed one Admin account the very first time this tab is empty, so there's
  // always a way to log in. Change this PIN from inside the sheet any time —
  // this only ever runs once, when the sheet has no data rows yet.
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['EMP001', 'Wajid', 'Wajid92', '123123', 'admin', 'active']);
  }
  return sheet;
}

function checkLogin(username, password) {
  if (!username || !password) return { success: false, reason: 'Missing username or password' };
  const sheet = getOrCreateUserSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {
    name: headers.indexOf('NAME'), username: headers.indexOf('USERNAME'),
    pin: headers.indexOf('PIN'), role: headers.indexOf('USER ACCESS'), status: headers.indexOf('USER STATUS')
  };
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[col.username]).toLowerCase() === String(username).toLowerCase()) {
      if (String(row[col.status]).toLowerCase() !== 'active') return { success: false, reason: 'Account is not active' };
      if (String(row[col.pin]) !== String(password)) return { success: false, reason: 'Incorrect password' };
      return { success: true, name: row[col.name], username: row[col.username], role: String(row[col.role]).toLowerCase() };
    }
  }
  return { success: false, reason: 'No such account' };
}

function listUsersPublic() {
  const sheet = getOrCreateUserSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = { name: headers.indexOf('NAME'), username: headers.indexOf('USERNAME'), role: headers.indexOf('USER ACCESS'), status: headers.indexOf('USER STATUS') };
  const items = [];
  for (let i = 1; i < data.length; i++) {
    items.push({ name: data[i][col.name], username: data[i][col.username], role: String(data[i][col.role]).toLowerCase(), status: data[i][col.status] });
  }
  return items;
}

function upsertUser(user) {
  // user: { name, username, password, role, status } — password === '' keeps the existing PIN
  const sheet = getOrCreateUserSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {
    empId: headers.indexOf('EMP ID'), name: headers.indexOf('NAME'), username: headers.indexOf('USERNAME'),
    pin: headers.indexOf('PIN'), role: headers.indexOf('USER ACCESS'), status: headers.indexOf('USER STATUS')
  };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col.username]).toLowerCase() === String(user.username).toLowerCase()) {
      sheet.getRange(i + 1, col.name + 1).setValue(user.name);
      sheet.getRange(i + 1, col.role + 1).setValue(user.role);
      sheet.getRange(i + 1, col.status + 1).setValue(user.status || 'active');
      if (user.password) sheet.getRange(i + 1, col.pin + 1).setValue(user.password);
      return;
    }
  }
  sheet.appendRow([nextEmpId(data, headers), user.name, user.username, user.password || '', user.role, user.status || 'active']);
}

function nextEmpId(data, headers) {
  const empCol = headers.indexOf('EMP ID');
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(String(data[i][empCol]).replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return 'EMP' + String(max + 1).padStart(3, '0');
}

/* ---------- Audit log ---------- */

function addLog(entry) {
  // entry: { user, role, action, details }
  const sheet = getOrCreateSheet(LOG_SHEET, LOG_HEADERS);
  sheet.appendRow([new Date(), entry.user, entry.role, entry.action, entry.details]);
}

function listLog() {
  const sheet = getOrCreateSheet(LOG_SHEET, LOG_HEADERS);
  return sheetToObjects(sheet).map(function (r) {
    return { timestamp: r['Timestamp'], user: r['User'], role: r['Role'], action: r['Action'], details: r['Details'] };
  });
}

/* ---------- Magento reference-data import (one JSON blob, its own tiny tab) ---------- */

function getMagentoIndex() {
  const sheet = getOrCreateSheet(MAGENTO_SHEET, MAGENTO_HEADERS);
  if (sheet.getLastRow() < 2) return { data: null, importedAt: null };
  const row = sheet.getRange(2, 1, 1, 2).getValues()[0];
  return { data: row[0] || null, importedAt: row[1] || null };
}

function setMagentoIndex(data, importedAt) {
  const sheet = getOrCreateSheet(MAGENTO_SHEET, MAGENTO_HEADERS);
  if (sheet.getLastRow() < 2) sheet.appendRow([data, importedAt]);
  else sheet.getRange(2, 1, 1, 2).setValues([[data, importedAt]]);
}
