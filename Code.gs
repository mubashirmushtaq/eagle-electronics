/**
 * Eagle Electronics - Repair Shop Management
 * Google Apps Script backend (Web App). The Google Sheet is the database.
 *
 * SHEETS REQUIRED (exact tab names, header row = row 1):
 *
 * JobCards:
 *   JobNo | Date | CustomerName | Phone | Phone2 | CNIC | DeviceType | BrandModel |
 *   SerialNo | Fault | Accessories | EstCharges | Advance | AtDelivery | Balance |
 *   Status | Technician | ReceivedBy | ExpectedDate | ReadyDate | DeliveredDate |
 *   DeliveryProof | CNICPhotoURL | Notes
 *
 * Payments:
 *   PaymentID | Date | JobNo | Amount | Type | ReceivedBy | Note
 *   (Type = Advance | Delivery | Other)
 *
 * Users:
 *   Username | Password | Name | Role | Active
 *   (Role = admin | staff, Active = TRUE/FALSE)
 *
 * Settings:
 *   Key | Value
 *   (ShopName, Address, Phone1, Phone2, Services, Currency, JobPrefix, NextJobNo)
 *
 * Deploy: Deploy > New deployment > Web app > Execute as: Me >
 *         Who has access: Anyone. Copy the /exec URL into index.html (API_URL).
 */

const SHEET_JOBCARDS = 'JobCards';
const SHEET_PAYMENTS = 'Payments';
const SHEET_USERS = 'Users';
const SHEET_SETTINGS = 'Settings';
const CNIC_DRIVE_FOLDER = 'EagleElectronics_CNIC';

// ---------- helpers ----------

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}

// Rows -> array of objects keyed by header names. Skips fully-empty rows.
function rowsToObjects(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(String);
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (row.every(c => c === '' || c === null)) continue;
    const o = {};
    headers.forEach((h, i) => { o[h] = row[i]; });
    out.push(o);
  }
  return out;
}

// 1-based row number of first data row where header col == value. -1 if none.
function findRow(sh, headerName, value) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = headers.indexOf(headerName) + 1;
  if (col < 1) throw new Error('Column not found: ' + headerName);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  const target = String(value);
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === target) return i + 2;
  }
  return -1;
}

function getSetting(key, fallback) {
  const sh = sheet(SHEET_SETTINGS);
  const r = findRow(sh, 'Key', key);
  if (r === -1) return fallback;
  return sh.getRange(r, 2).getValue();
}

function setSettingRaw(key, value) {
  const sh = sheet(SHEET_SETTINGS);
  const r = findRow(sh, 'Key', key);
  if (r === -1) sh.appendRow([key, value]);
  else sh.getRange(r, 2).setValue(value);
}

function isoNow() {
  return new Date().toISOString();
}

function num(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

// Recompute Balance for one job row: EstCharges - Advance - AtDelivery
function recomputeBalance(sh, row) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const est = num(sh.getRange(row, col('EstCharges')).getValue());
  const adv = num(sh.getRange(row, col('Advance')).getValue());
  const del = num(sh.getRange(row, col('AtDelivery')).getValue());
  sh.getRange(row, col('Balance')).setValue(Math.max(0, est - adv - del));
}

// ---------- GET: load everything after login ----------

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'load') {
      const settings = {};
      rowsToObjects(sheet(SHEET_SETTINGS)).forEach(r => { settings[String(r.Key)] = r.Value; });
      const users = rowsToObjects(sheet(SHEET_USERS))
        .filter(u => String(u.Active).toUpperCase() === 'TRUE')
        .map(u => ({ Username: u.Username, Name: u.Name, Role: u.Role }));
      return json({
        jobCards: rowsToObjects(sheet(SHEET_JOBCARDS)),
        payments: rowsToObjects(sheet(SHEET_PAYMENTS)),
        users: users,
        settings: settings
      });
    }
    return json({ error: 'Unknown action' });
  } catch (err) {
    return json({ error: String(err) });
  }
}

// ---------- POST ----------

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    switch (action) {
      case 'login':      return handleLogin(body);
      case 'createJob':  return handleCreateJob(body);
      case 'updateJob':  return handleUpdateJob(body);
      case 'setStatus':  return handleSetStatus(body);
      case 'addPayment': return handleAddPayment(body);
      case 'deliver':    return handleDeliver(body);
      case 'addUser':    return handleAddUser(body);
      case 'toggleUser': return handleToggleUser(body);
      case 'saveSettings': return handleSaveSettings(body);
      default: return json({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

function ok(extra) {
  const o = { status: 'success' };
  if (extra) for (const k in extra) o[k] = extra[k];
  return json(o);
}
function fail(msg) { return json({ status: 'error', message: msg }); }

// ----- login -----
function handleLogin(b) {
  const sh = sheet(SHEET_USERS);
  const r = findRow(sh, 'Username', String(b.username || '').trim());
  if (r === -1) return fail('Invalid username or password');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const pw = String(sh.getRange(r, col('Password')).getValue());
  const active = String(sh.getRange(r, col('Active')).getValue()).toUpperCase() === 'TRUE';
  if (!active || pw !== String(b.password || '')) return fail('Invalid username or password');
  return ok({ user: {
    username: sh.getRange(r, col('Username')).getValue(),
    name: sh.getRange(r, col('Name')).getValue(),
    role: sh.getRange(r, col('Role')).getValue()
  }});
}

// ----- create job card -----
const JOBCARD_HEADERS = ['JobNo','Date','CustomerName','Phone','Phone2','CNIC','DeviceType',
  'BrandModel','SerialNo','Fault','Accessories','EstCharges','Advance','AtDelivery','Balance',
  'Status','Technician','ReceivedBy','ExpectedDate','ReadyDate','DeliveredDate',
  'DeliveryProof','CNICPhotoURL','Notes'];

function handleCreateJob(b) {
  const j = b.job || {};
  if (!j.CustomerName || !j.Phone || !j.DeviceType) {
    return fail('Customer name, phone and device type are required');
  }
  const sh = sheet(SHEET_JOBCARDS);
  const prefix = String(getSetting('JobPrefix', 'EE-'));
  let next = parseInt(getSetting('NextJobNo', '1001'), 10);
  if (isNaN(next)) next = 1001;
  const jobNo = prefix + next;

  const row = JOBCARD_HEADERS.map(h => {
    switch (h) {
      case 'JobNo': return jobNo;
      case 'Date': return isoNow();
      case 'Status': return 'Received';
      case 'Advance': return num(j.Advance);
      case 'EstCharges': return num(j.EstCharges);
      case 'AtDelivery': return 0;
      case 'Balance': return Math.max(0, num(j.EstCharges) - num(j.Advance));
      default: return j[h] !== undefined ? j[h] : '';
    }
  });
  sh.appendRow(row);
  setSettingRaw('NextJobNo', String(next + 1));

  // Log the advance as a payment so revenue reports stay consistent
  if (num(j.Advance) > 0) {
    sheet(SHEET_PAYMENTS).appendRow([
      'P' + Date.now(), isoNow(), jobNo, num(j.Advance), 'Advance',
      j.ReceivedBy || '', 'Advance at intake'
    ]);
  }
  return ok({ jobNo: jobNo });
}

// ----- update job fields -----
const EDITABLE = ['CustomerName','Phone','Phone2','CNIC','DeviceType','BrandModel','SerialNo',
  'Fault','Accessories','EstCharges','Technician','ReceivedBy','ExpectedDate','Notes'];

function handleUpdateJob(b) {
  const sh = sheet(SHEET_JOBCARDS);
  const r = findRow(sh, 'JobNo', b.jobNo);
  if (r === -1) return fail('Job not found: ' + b.jobNo);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const u = b.updates || {};
  EDITABLE.forEach(k => {
    if (u[k] !== undefined && col(k) > 0) sh.getRange(r, col(k)).setValue(u[k]);
  });
  recomputeBalance(sh, r);
  return ok({});
}

// ----- status transitions -----
const VALID_STATUS = ['Received', 'In Repair', 'Ready', 'Delivered', 'Cancelled'];

function handleSetStatus(b) {
  if (VALID_STATUS.indexOf(b.status) === -1) return fail('Invalid status');
  if (b.status === 'Delivered') return fail('Use the delivery flow to mark a job delivered');
  const sh = sheet(SHEET_JOBCARDS);
  const r = findRow(sh, 'JobNo', b.jobNo);
  if (r === -1) return fail('Job not found: ' + b.jobNo);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const cur = String(sh.getRange(r, col('Status')).getValue());
  if (cur === 'Delivered') return fail('Job already delivered');
  sh.getRange(r, col('Status')).setValue(b.status);
  if (b.status === 'Ready' && col('ReadyDate') > 0) sh.getRange(r, col('ReadyDate')).setValue(isoNow());
  return ok({});
}

// ----- payments -----
function handleAddPayment(b) {
  const amount = num(b.amount);
  if (amount <= 0) return fail('Amount must be greater than zero');
  const type = ['Advance', 'Delivery', 'Other'].indexOf(b.type) >= 0 ? b.type : 'Other';
  const sh = sheet(SHEET_JOBCARDS);
  const r = findRow(sh, 'JobNo', b.jobNo);
  if (r === -1) return fail('Job not found: ' + b.jobNo);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const cur = String(sh.getRange(r, col('Status')).getValue());
  if (cur === 'Delivered' || cur === 'Cancelled') return fail('Cannot add payment to a closed job');

  sheet(SHEET_PAYMENTS).appendRow([
    'P' + Date.now(), isoNow(), b.jobNo, amount, type, b.receivedBy || '', b.note || ''
  ]);
  if (type === 'Advance') {
    sh.getRange(r, col('Advance')).setValue(num(sh.getRange(r, col('Advance')).getValue()) + amount);
  } else if (type === 'Delivery') {
    sh.getRange(r, col('AtDelivery')).setValue(num(sh.getRange(r, col('AtDelivery')).getValue()) + amount);
  }
  recomputeBalance(sh, r);
  return ok({});
}

// ----- delivery: card returned OR CNIC proof -----
function handleDeliver(b) {
  const sh = sheet(SHEET_JOBCARDS);
  const r = findRow(sh, 'JobNo', b.jobNo);
  if (r === -1) return fail('Job not found: ' + b.jobNo);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const cur = String(sh.getRange(r, col('Status')).getValue());
  if (cur === 'Delivered') return fail('Job already delivered');
  if (cur === 'Cancelled') return fail('Job is cancelled');

  const proof = b.proof === 'cnic' ? 'CNIC' : 'Card';
  let photoUrl = '';

  if (proof === 'CNIC') {
    if (!b.cnicNumber && !b.cnicPhotoBase64) {
      return fail('CNIC number or photo is required when the job card is not returned');
    }
    if (b.cnicNumber && col('CNIC') > 0) sh.getRange(r, col('CNIC')).setValue(String(b.cnicNumber));
    if (b.cnicPhotoBase64) {
      photoUrl = saveCnicPhoto(b.jobNo, b.cnicPhotoBase64, b.cnicPhotoName || 'cnic.jpg');
    }
  }

  // Optional final payment taken at delivery
  const finalAmt = num(b.finalAmount);
  if (finalAmt > 0) {
    sheet(SHEET_PAYMENTS).appendRow([
      'P' + Date.now(), isoNow(), b.jobNo, finalAmt, 'Delivery', b.receivedBy || '', 'Collected at delivery'
    ]);
    sh.getRange(r, col('AtDelivery')).setValue(num(sh.getRange(r, col('AtDelivery')).getValue()) + finalAmt);
  }

  sh.getRange(r, col('Status')).setValue('Delivered');
  if (col('DeliveredDate') > 0) sh.getRange(r, col('DeliveredDate')).setValue(isoNow());
  if (col('DeliveryProof') > 0) sh.getRange(r, col('DeliveryProof')).setValue(proof);
  if (photoUrl && col('CNICPhotoURL') > 0) sh.getRange(r, col('CNICPhotoURL')).setValue(photoUrl);
  recomputeBalance(sh, r);
  return ok({ photoUrl: photoUrl });
}

function saveCnicPhoto(jobNo, base64, fileName) {
  const clean = String(base64).replace(/^data:.*?;base64,/, '');
  const bytes = Utilities.base64Decode(clean);
  let mime = 'image/jpeg';
  const m = String(base64).match(/^data:(.*?);base64,/);
  if (m) mime = m[1];
  const blob = Utilities.newBlob(bytes, mime, jobNo + '_' + fileName);
  let folder;
  const folders = DriveApp.getFoldersByName(CNIC_DRIVE_FOLDER);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CNIC_DRIVE_FOLDER);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}

// ----- users (admin) -----
function handleAddUser(b) {
  if (!b.username || !b.password) return fail('Username and password required');
  const sh = sheet(SHEET_USERS);
  if (findRow(sh, 'Username', String(b.username).trim()) !== -1) return fail('Username already exists');
  sh.appendRow([String(b.username).trim(), String(b.password), b.name || '', b.role === 'admin' ? 'admin' : 'staff', 'TRUE']);
  return ok({});
}

function handleToggleUser(b) {
  const sh = sheet(SHEET_USERS);
  const r = findRow(sh, 'Username', b.username);
  if (r === -1) return fail('User not found');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => headers.indexOf(name) + 1;
  const cur = String(sh.getRange(r, col('Active')).getValue()).toUpperCase() === 'TRUE';
  sh.getRange(r, col('Active')).setValue(cur ? 'FALSE' : 'TRUE');
  return ok({ active: !cur });
}

// ----- settings (admin) -----
const SETTABLE = ['ShopName', 'Address', 'Phone1', 'Phone2', 'Services', 'Currency', 'JobPrefix'];

function handleSaveSettings(b) {
  const s = b.settings || {};
  SETTABLE.forEach(k => { if (s[k] !== undefined) setSettingRaw(k, String(s[k])); });
  return ok({});
}
