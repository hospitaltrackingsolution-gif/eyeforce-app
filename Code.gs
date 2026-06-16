/**
 * EyeForce Solutions — Backend (Code.gs)
 * -------------------------------------------------
 * Handles: page serving, authentication, daily data entry,
 * same-day edit rules + correction request workflow,
 * pipeline updates, dashboard data aggregation, and search.
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
// IMPORTANT: use the SPREADSHEET's timezone, not Session.getScriptTimeZone().
// The Apps Script project timezone setting can differ from the Spreadsheet's
// own timezone (File → Settings in Sheets). Google Sheets stores/reads back
// date cells relative to the SPREADSHEET timezone, so using anything else
// here causes "today's" entries to intermittently format one day off and
// fail every same-day match (My Entries / locking / corrections / dashboard
// "Today" range all rely on this).
const TZ = SS.getSpreadsheetTimeZone();

// Bumped with every meaningful release. Shown in the sidebar footer and the
// Settings → Diagnostics panel so you can confirm a "Deploy → Manage
// deployments → New version" actually took effect on the live URL — if the
// app still shows an old version number after redeploying, the deployment
// itself didn't update (see README).
const BUILD_VERSION = 'v2026-06-17.3';

// ---------- PAGE SERVING ----------

function doGet(e) {
  // ── PWA Web App Manifest for Android Chrome "Add to Home Screen" ──────
  if (e && e.parameter && e.parameter['manifest'] === '1') {
    try {
      const settings = getSettings();
      const appName  = (settings && settings['Hospital Name'])
                       ? 'EyeForce — ' + settings['Hospital Name']
                       : 'EyeForce Solutions';
      const logoUrl  = 'https://lh3.googleusercontent.com/d/1c5z4txQiWuvg1gNgXZZptq83swvbHxUD';
      const manifest = {
        name: appName,
        short_name: 'EyeForce',
        description: 'Hospital Practice & Business Intelligence by EyeForce Solutions',
        start_url: ScriptApp.getService().getUrl(),
        display: 'standalone',
        orientation: 'portrait-primary',
        background_color: '#15233B',
        theme_color: '#15233B',
        icons: [
          { src: logoUrl, sizes: '192x192', type: 'image/png', purpose: 'any maskable' },
          { src: logoUrl, sizes: '512x512', type: 'image/png', purpose: 'any maskable' }
        ]
      };
      return ContentService
        .createTextOutput(JSON.stringify(manifest))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── Normal hospital web app — serve the dashboard ─────────────────────
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('EyeForce Solutions')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ============================================================
 *  AUTHORIZATION HELPER — run this FIRST from the Apps Script
 *  editor (not from the web app URL) any time you update the
 *  code and the app goes blank or stops responding.
 *
 *  HOW TO USE:
 *  1. In the Apps Script editor, select "authorizeAndTest" from
 *     the function dropdown at the top.
 *  2. Click the ▶ Run button.
 *  3. A "Authorization required" popup will appear — click
 *     "Review permissions" → choose your Google account →
 *     click "Allow".
 *  4. The function logs a simple success message. Done.
 *  5. Now redeploy: Deploy → Manage deployments → pencil icon
 *     → New version → Deploy.
 *
 *  You MUST do this every time you add new functions that
 *  access Google services (Sheets, Gmail, Drive, etc.) —
 *  Apps Script will refuse to run any function at all until
 *  the new permissions are granted, which is why the whole
 *  app goes silent rather than just the new function failing.
 * ============================================================
 */
function authorizeAndTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const staffSheet = ss.getSheetByName('Staff');
  const staffRows = staffSheet ? staffSheet.getLastRow() - 1 : 0;
  Logger.log('✓ Authorization successful.');
  Logger.log('  Spreadsheet: ' + ss.getName());
  Logger.log('  Timezone: ' + tz);
  Logger.log('  Server today: ' + today);
  Logger.log('  Staff rows: ' + staffRows);
  Logger.log('  Build: ' + BUILD_VERSION);
  Logger.log('');
  Logger.log('Next step: Deploy → Manage deployments → pencil → New version → Deploy.');
  return 'OK';
}

// ---------- SETTINGS ----------

/**
 * Returns PriceMaster rates keyed by "Procedure|Sub-Type" so the
 * Counselling form can auto-fill Quote Given when Procedure + Category
 * (and optional Cash Package) are selected. Revenue fields stay
 * completely hidden from the counsellor — only Quote Given is shown.
 */
function getPriceMasterRates() {
  const sheet = SS.getSheetByName('PriceMaster');
  if (!sheet) return {};
  const { rows } = getSheetData_('PriceMaster');
  const rates = {};
  rows.forEach(r => {
    const item = r['Item / Procedure'] || r['Item'];
    const payType = r['Payment Type'] || r['Sub-Type / Package'];
    const rate = Number(r['Rate (₹) — fill in'] || r['Rate (₹)']) || 0;
    if (!item || !payType || !rate) return;
    rates[item + '|' + payType] = rate;
  });
  return rates;
}

/**
 * Looks up a single rate from the Price Master.
 * procedure: e.g. 'Cataract'
 * category: e.g. 'Cash', 'CGHS', 'ECHS', 'Other TPA'
 * cashPkg: e.g. 'Basic' (only used when category=Cash)
 * Returns 0 if not found or not set.
 */
function lookupRate_(rates, item, category, cashPkg) {
  if (!item || !category) return 0;
  const payType = (category === 'Cash' && cashPkg) ? 'Cash - ' + cashPkg : category;
  return rates[item + '|' + payType] || 0;
}

/**
 * Auto-computes OT revenue split by actual-collected vs expected-receivable.
 * Called server-side from submitEntry — results stored in backend-only columns,
 * never shown to OT staff.
 *
 * Returns { collected, cghs, cghsRate, echs, echsRate, tpaBilled, totalCases,
 *           totalCollected, totalExpected }
 */
function computeOTRevenue_(rates, data) {
  const proc = data['Procedure Type'] || '';

  const cashCases       = Number(data['Cash Cases'])              || 0;
  const cghsCases       = Number(data['CGHS Cases'])              || 0;
  const cghsUpgraded    = Number(data['CGHS Upgraded Cases'])     || 0;
  const echsCases       = Number(data['ECHS Cases'])              || 0;
  const echsUpgraded    = Number(data['ECHS Upgraded Cases'])     || 0;
  const delhiCases      = Number(data['Delhi Govt Cases'])        || 0;
  const delhiUpgraded   = Number(data['Delhi Govt Upgraded Cases'])|| 0;
  const tpaCases        = Number(data['TPA Cases'])               || 0;
  const tpaUpgraded     = Number(data['TPA Upgraded Cases'])      || 0;
  const tpaBilled       = Number(data['TPA Billed Amount (₹)'])   || 0;
  const cashCollected   = Number(data['Cash Collected (₹)'])      || 0;

  const cghsRate  = rates[proc + '|CGHS']       || 0;
  const echsRate  = rates[proc + '|ECHS']        || 0;
  const delhiRate = rates[proc + '|Delhi Govt']  || cghsRate; // Delhi Govt often = CGHS rate

  const cghsExpected  = cghsCases  * cghsRate;
  const echsExpected  = echsCases  * echsRate;
  const delhiExpected = delhiCases * delhiRate;
  // Private TPA expected = staff-entered billed amount
  const totalExpected = cghsExpected + echsExpected + delhiExpected + tpaBilled;
  const totalCases    = cashCases + cghsCases + echsCases + delhiCases + tpaCases;

  // Upgraded case counts are flagged so TPA dept knows to collect co-pay
  // from those patients separately. They do NOT affect revenue calculation here —
  // co-pay collected is logged via TPASettlements with Settlement Type = Patient Co-pay
  return {
    cghsRate, cghsExpected,
    echsRate, echsExpected,
    delhiRate, delhiExpected,
    totalCases,
    totalCollected: cashCollected,
    totalExpected,
    // Summary of upgraded cases — TPA dept uses this as their worklist
    upgradedSummary: {
      cghs: cghsUpgraded, echs: echsUpgraded,
      delhi: delhiUpgraded, tpa: tpaUpgraded,
      total: cghsUpgraded + echsUpgraded + delhiUpgraded + tpaUpgraded
    }
  };
}

/**
 * Auto-computes revenue for a Counselling per-patient converted row.
 */
function computeCounsellingRevenue_(rates, data) {
  if (String(data['Converted'] || '').toLowerCase() !== 'yes') return 0;
  const proc = data['Procedure'] || '';
  const cat = data['Category'] || '';
  const pkg = data['Cash Package'] || '';
  return lookupRate_(rates, proc, cat, pkg);
}

/**
 * Auto-computes revenue for Diagnostic from paid test counts × rates.
 * Uses '[Test] - Paid' columns only (free tests have no revenue).
 */
function computeDiagnosticRevenue_(rates, data) {
  const diagTests = ['Biometry','Topography','OCT','Cataract Work-up','LASIK Work-up',
    'Green Laser','Fundus Photography','Visual Field/Perimetry','B-Scan/A-Scan','Other Tests'];
  return diagTests.reduce((total, test) => {
    const paidUnits = Number(data[test + ' - Paid']) || 0;
    const rate = rates[test + '|Cash'] || rates[test + '|Diagnostic|Cash'] || 0;
    return total + (paidUnits * rate);
  }, 0);
}

function getInventoryMaster() {
  const { rows } = getSheetData_('InventoryMaster');
  return rows.map(r => ({
    name: r['Item Name'],
    category: r['Category'],
    unit: r['Unit'],
    reorderLevel: Number(r['Reorder Level']) || 0
  }));
}

function getSettings() {
  const sheet = SS.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    settings[key] = value;
  }
  // Split comma-lists into arrays for convenience
  ['Departments Enabled', 'Doctor List', 'Procedure List', 'TPA Category List',
   'Cash Package List', 'PRO List', 'Area List', 'Lead Category List',
   'Issued To List', 'Transaction Type List', 'Purchase Category List',
   'Vendor List', 'Payment Status List', 'Diagnostic Category List'].forEach(k => {
    if (settings[k]) settings[k] = String(settings[k]).split(',').map(s => s.trim()).filter(Boolean);
    else settings[k] = [];
  });
  return settings;
}

// ============================================================
//  ADMIN SETTINGS PAGE — dynamic, grouped, easy-to-edit Settings UI
//  (replaces having to hand-edit the raw Settings sheet for everything)
// ============================================================

const ALL_DEPARTMENTS = ['Reception','Counselling','OT','Diagnostic','Pharmacy','Optical','Marketing','HR','TPA','TPAClaims','TPASettlements','Operations','Store','Purchase','CallCentre','HRSummary','TPASummary'];
const LIST_SETTING_KEYS = ['Doctor List','Procedure List','Diagnostic Category List','TPA Category List',
  'Cash Package List','PRO List','Area List','Lead Category List','Issued To List','Transaction Type List',
  'Purchase Category List','Vendor List','Payment Status List'];

/**
 * Describes every Settings field as a grouped, typed form so the Admin
 * Settings screen can render itself dynamically — no raw-sheet editing
 * required for day-to-day configuration.
 */
const SETTINGS_SCHEMA = [
  { title: 'Hospital Info', fields: [
    { key: 'Hospital Name', type: 'text', help: 'Shown at the top of every page and on reports.' },
    { key: 'Hospital Code', type: 'text', help: 'Short internal code, shown next to the hospital name.' },
    { key: 'Active Status', type: 'select', options: ['Active','Inactive'], help: 'Set to Inactive to disable this hospital\'s app temporarily.' },
    { key: 'Logo URL', type: 'logo', help: 'Hospital logo shown top-left. Use a public image link (e.g. Google Drive "Anyone with the link" share, converted to a direct image URL).' }
  ]},
  { title: 'Branding', fields: [
    { key: 'EyeForce Logo URL', type: 'logo', help: 'Optional — EyeForce Solutions logo shown on the login screen. Leave blank to use the default mark.' }
  ]},
  { title: 'Departments & Mode Settings', fields: [
    { key: 'Departments Enabled', type: 'checkboxList', options: ALL_DEPARTMENTS, help: 'Untick a department to hide it from every staff member\'s menu (Admin/Doctor are unaffected).' },
    { key: 'Counselling Mode (Per-Patient / Daily Summary)', type: 'select', options: ['Per-Patient','Daily Summary'],
      help: 'Per-Patient = detailed entry per counselled patient (needs a dedicated counsellor). Daily Summary = one aggregate row per day by procedure category. Change anytime — old data is preserved.' },
    { key: 'HR Mode (Per-Person / Daily Summary)', type: 'select', options: ['Per-Person','Daily Summary'],
      help: 'Per-Person = one HR entry per staff member per day (full individual tracking). Daily Summary = one consolidated row per day with total counts. Change via Settings — both modes are fully supported.' },
    { key: 'TPA Mode (Detailed / Daily Summary)', type: 'select', options: ['Detailed','Daily Summary'],
      help: 'Detailed = individual claim + settlement tracking per case (full revenue traceability — recommended). Daily Summary = aggregate daily counts per category. Revenue logic works correctly in both modes.' }
  ]},
  { title: 'Master Lists — Dropdown Options', fields: LIST_SETTING_KEYS.map(k => ({ key: k, type: 'list', help: 'One option per line.' })) },
  { title: 'Daily & Monthly Reports', fields: [
    { key: 'Daily Report Emails - Admin', type: 'text', help: 'Comma-separated. Receives today\'s PDF report daily.' },
    { key: 'Daily Report Emails - Doctor', type: 'text', help: 'Comma-separated.' },
    { key: 'Monthly Report Emails - Admin', type: 'text', help: 'Comma-separated. Admin receives last month\'s complete PDF on the monthly trigger date.' },
    { key: 'Monthly Report Emails - Doctor', type: 'text', help: 'Comma-separated.' },
    { key: 'Monthly Report Send Day', type: 'select', options: ['1','2','3','4','5'], help: 'Day of each month to send the previous month\'s complete report.' },
    { key: 'WhatsApp - Admin Phone', type: 'text', help: 'Admin WhatsApp number with country code, no spaces or + (e.g. 919876543210). Daily report will be sent here.' },
    { key: 'WhatsApp - Doctor Phone', type: 'text', help: 'Doctor WhatsApp number with country code (e.g. 919876543210). Daily report will be sent here.' },
    { key: 'Notification Emails', type: 'text', help: 'Comma-separated. Get an email the moment any staff member submits a correction request.' }
  ]},
  { title: 'Other', fields: [
    { key: 'Google Review Link', type: 'text', help: 'Shown to staff so they can share it with patients for Google Reviews.' },
    { key: 'AI API Key', type: 'password', help: 'Anthropic API key — enables the AI Executive Analysis section and embeds it in the Admin\'s monthly PDF. Leave blank to disable.' }
  ]}
];

/**
 * Returns the schema above pre-filled with current values from the Settings
 * sheet — Admin only. List/checkboxList fields come back as arrays.
 */
function getSettingsForAdmin(user) {
  if (user.role !== 'Admin') return { success: false, message: 'Settings are available to Admin only.' };
  const sheet = SS.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const raw = {};
  for (let i = 1; i < data.length; i++) raw[data[i][0]] = data[i][1];

  const groups = SETTINGS_SCHEMA.map(group => ({
    title: group.title,
    fields: group.fields.map(f => {
      let value = raw[f.key] !== undefined ? raw[f.key] : '';
      if (f.type === 'list' || f.type === 'checkboxList') {
        value = String(value || '').split(',').map(s => s.trim()).filter(Boolean);
      }
      return { key: f.key, type: f.type, options: f.options || null, help: f.help || '', value: value };
    })
  }));
  return { success: true, groups };
}

/**
 * Saves a flat { 'Setting Key': value } object back to the Settings sheet.
 * Array values (list/checkboxList) are joined with commas. Unknown keys are
 * appended as new rows so the sheet can grow without breaking anything.
 * Admin only — every change is written to the AuditLog.
 */
function updateSettingsForAdmin(values, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Settings are available to Admin only.' };
  const sheet = SS.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const keyRow = {};
  for (let i = 1; i < data.length; i++) keyRow[data[i][0]] = i + 1; // 1-based sheet row

  let changedCount = 0;
  Object.keys(values).forEach(key => {
    let value = values[key];
    if (Array.isArray(value)) value = value.join(',');
    value = (value === undefined || value === null) ? '' : value;
    if (keyRow[key]) {
      sheet.getRange(keyRow[key], 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
      keyRow[key] = sheet.getLastRow();
    }
    changedCount++;
  });

  logAudit_(user.name, 'SETTINGS-UPDATE', 'Settings', 0, `${changedCount} field(s) updated by ${user.name}.`);
  return { success: true, message: `✓ Settings saved (${changedCount} field${changedCount === 1 ? '' : 's'} updated).` };
}

// ---------- AUTH ----------

function login(loginId, pin) {
  const sheet = SS.getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[3]).toLowerCase() === String(loginId).toLowerCase() && String(row[4]) === String(pin)) {
      if (String(row[5]) !== 'Active') {
        return { success: false, message: 'This account has been deactivated. Contact admin.' };
      }
      return {
        success: true,
        user: { name: row[0], department: row[1], role: row[2], loginId: row[3] }
      };
    }
  }
  return { success: false, message: 'Invalid login ID or PIN.' };
}

// ---------- HELPERS ----------

function todayStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function nowStr_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
}

/**
 * Converts a 'Date' cell value (which Sheets may store as either an actual
 * Date object OR plain text, depending on the cell's number format) into a
 * reliable 'yyyy-MM-dd' string — WITHOUT round-tripping through `new
 * Date(stringValue)`, which silently shifts by a day whenever the script's
 * assumed timezone doesn't exactly match how the string should be read.
 * This is the single source of truth for "what date is this row for?" used
 * everywhere (My Entries Today, locking, corrections, dashboard ranges).
 */
function dateToStr_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10); // already 'yyyy-MM-dd...' — use directly
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/**
 * Lightweight diagnostic the front-end calls on load to show "today" exactly
 * as the SERVER computes it (used for the top-bar date badge). If this ever
 * looks one day off from the device's own date, the Spreadsheet's timezone
 * (File → Settings → Time zone) doesn't match the hospital's local time —
 * fix that first, since every "today" feature in the app depends on it.
 */
function getServerToday() {
  const today = todayStr_();
  return {
    dateStr: today,
    label: Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy'),
    timezone: TZ,
    build: BUILD_VERSION
  };
}

/**
 * Admin-only diagnostics — shows EXACTLY what the server sees for each
 * daily-entry sheet's most recent row: the raw Date cell value/type, what
 * it parses to, whether that matches "today", and whether the Edit
 * Status / Submitted By columns exist at all. This turns "My entries
 * today is empty and I don't know why" into a concrete, shareable report
 * instead of guesswork — use it from Settings → Diagnostics.
 */
function getDiagnostics(user) {
  if (user.role !== 'Admin') return { success: false, message: 'Diagnostics are available to Admin only.' };
  const today = todayStr_();
  const sheetNames = ['Reception','Counselling','CounsellingSummary','OT','Diagnostic','Pharmacy','Optical','Marketing','HR','TPA','TPAClaims','TPASettlements','Operations','Store','Purchase','CallCentre','HRSummary','TPASummary'];
  const sheets = sheetNames.map(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) return { sheet: name, exists: false };
    const lastRow = sheet.getLastRow();
    const headers = lastRow >= 1 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    const result = {
      sheet: name, exists: true, lastRow,
      hasDateCol: headers.indexOf('Date') !== -1,
      hasTimestampCol: headers.indexOf('Timestamp') !== -1,
      hasEditStatusCol: headers.indexOf('Edit Status') !== -1,
      hasSubmittedByCol: headers.indexOf('Submitted By') !== -1
    };
    if (lastRow >= 2) {
      const lastVals = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const dateCol = headers.indexOf('Date');
      const statusCol = headers.indexOf('Edit Status');
      const submittedByCol = headers.indexOf('Submitted By');
      const rawDate = dateCol !== -1 ? lastVals[dateCol] : null;
      result.lastRowRawDate = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate);
      result.lastRowRawDateType = Object.prototype.toString.call(rawDate);
      result.lastRowParsedDate = dateToStr_(rawDate);
      result.matchesToday = dateToStr_(rawDate) === today;
      result.lastRowEditStatus = statusCol !== -1 ? String(lastVals[statusCol]) : null;
      result.lastRowSubmittedBy = submittedByCol !== -1 ? String(lastVals[submittedByCol]) : null;
    }
    return result;
  });

  // Staff sheet — surfaces department-name mismatches (the other common
  // cause of "My entries today" appearing empty for non-Admin/Doctor users:
  // canEditRow_ compares Staff!Department against the sheet name, and a
  // typo/extra space there will silently fail the comparison).
  const { rows: staff } = getSheetData_('Staff');
  const staffSummary = staff.map(s => ({ name: s['Staff Name'], department: s['Department'], role: s['Role'], status: s['Status'] }));

  return {
    success: true,
    build: BUILD_VERSION,
    timezone: TZ,
    today,
    sheets,
    staff: staffSummary,
    knownDepartments: ALL_DEPARTMENTS
  };
}

function getSheetData_(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) return { headers: [], rows: [] };  // missing sheet → empty, never throws
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return { headers: [], rows: [] };
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { headers: [], rows: [] };
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];
  const rows = data.slice(1).map((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj._row = idx + 2;
    return obj;
  });
  return { headers, rows };
}

// ---------- DAILY ENTRY SUBMISSION ----------

/**
 * Controls how "one entry per day" is enforced per sheet — this is the
 * single source of truth for duplication control across the whole app:
 *
 *  - Sheets listed here with an EMPTY key array are "one row per day,
 *    department-wide" (Reception, Diagnostic, Pharmacy, Optical, TPA,
 *    Operations, CounsellingSummary). Re-submitting the form for the same
 *    day UPDATES that single row in place — no duplicates possible, and
 *    this doubles as the correction mechanism (just resubmit with the
 *    corrected numbers).
 *  - Sheets listed with key fields (HR → Staff Name; Marketing → PRO Name;
 *    OT → Doctor Name + Procedure Type) are "one row per day PER KEY".
 *    Re-submitting for the same key on the same day updates that person's/
 *    doctor's row instead of adding a second one.
 *  - Sheets NOT listed here (Counselling, Store, Purchase) are append-only
 *    — every submission is a distinct event (a patient, a stock
 *    transaction, an invoice) and multiple per day is normal and expected.
 */
const UPSERT_KEYS = {
  Reception: [], Diagnostic: [], Pharmacy: [], Optical: [],
  TPA: [], TPASummary: [], Operations: [],
  CounsellingSummary: [], HRSummary: [], CallCentre: [],
  HR: ['Staff Name'], Marketing: ['PRO Name'], OT: ['Doctor Name', 'Procedure Type']
};

/**
 * data = { field1: value1, field2: value2, ... } matching sheet headers
 * (minus Timestamp/Date/Submitted By/Edit Status which are auto-filled).
 *
 * For sheets in UPSERT_KEYS, if a row for TODAY already matches the key
 * fields (or, for department-wide sheets, any row for today at all), that
 * row is UPDATED in place instead of a new one being appended — this is
 * what prevents duplicate daily reports. Everything else is appended.
 */
function submitEntry(sheetName, data, user) {
  try {
    return submitEntry_(sheetName, data, user);
  } catch(e) {
    Logger.log('submitEntry ERROR: ' + sheetName + ' — ' + e.message + '\n' + e.stack);
    return { success: false, message: '⚠ Server error saving ' + sheetName + ': ' + e.message + '. Please try again. If this repeats, run setupHospitalSheet() from Apps Script editor.' };
  }
}

function submitEntry_(sheetName, data, user) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    return {
      success: false,
      message: `⚠ Sheet "${sheetName}" not found. Please run setupHospitalSheet() from the Apps Script editor to create all required tabs, then reload the app.`
    };
  }
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return { success: false, message: `⚠ Sheet "${sheetName}" has no columns. Run setupHospitalSheet() to rebuild it.` };
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const desc = entryDescription_(sheetName, data);

  // Admin date override — for entering missed/past data on behalf of absent staff
  let dateStr = todayStr_();
  let tempAccessRow = null;
  if (data['__dateOverride__']) {
    const overrideDate = String(data['__dateOverride__']).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(overrideDate) && overrideDate < todayStr_()) {
      if (user.role === 'Admin') {
        // Admin can always back-date
        dateStr = overrideDate;
      } else {
        // Staff: check if they have an open TempAccess window for this date
        const { rows: accessRows } = getSheetData_('TempAccess');
        const window_ = accessRows.find(r =>
          r['Staff Name'] === user.name &&
          dateToStr_(r['Allowed Date']) === overrideDate &&
          r['Status'] === 'Open'
        );
        if (window_) {
          dateStr = overrideDate;
          tempAccessRow = window_._row; // will be closed after successful save
        } else {
          return { success: false, message: '⚠ You do not have permission to submit data for ' + overrideDate + '. Please ask Admin to open a data-entry window for that date.' };
        }
      }
    }
  }
  delete data['__dateOverride__'];
  const niceDate = Utilities.formatDate(new Date(dateStr + 'T12:00:00'), TZ, 'dd MMM yyyy');

  // ── Revenue auto-compute: ONLY for sheets that use Price Master ────────────
  // getPriceMasterRates() reads the entire PriceMaster sheet — calling it on
  // every department (including Operations, HR, etc.) was causing multi-second
  // hangs. Now only called when actually needed.
  // ── Department-specific validation ────────────────────────────────────────
  if (sheetName === 'Reception') {
    const others = Number(data['Others Cash Collected (₹)']) || 0;
    if (others > 0 && !String(data['Others Remarks'] || '').trim()) {
      return { success: false, message: '⚠ Please fill in "Others Remarks" — describe what the ₹' + others.toLocaleString('en-IN') + ' in Others Cash is for.' };
    }
  }

  const PRICE_MASTER_SHEETS = ['OT', 'Counselling', 'Diagnostic'];
  if (PRICE_MASTER_SHEETS.includes(sheetName)) {
    const rates = getPriceMasterRates();
    if (sheetName === 'OT') {
      const otRev = computeOTRevenue_(rates, data);
      data['CGHS Rate (₹)']             = otRev.cghsRate;
      data['CGHS Expected (₹)']         = otRev.cghsExpected;
      data['ECHS Rate (₹)']             = otRev.echsRate;
      data['ECHS Expected (₹)']         = otRev.echsExpected;
      data['Delhi Govt Rate (₹)']       = otRev.delhiRate;
      data['Delhi Govt Expected (₹)']   = otRev.delhiExpected;
      data['Total Cases']               = otRev.totalCases;
      data['Total Cash Collected (₹)']  = otRev.totalCollected;
      data['Total Expected Claims (₹)'] = otRev.totalExpected;
    } else if (sheetName === 'Counselling') {
      const rate = computeCounsellingRevenue_(rates, data);
      data['Price Master Rate (₹)']  = rate;
      data['Estimated Revenue (₹)']  = rate;
      if (data['Converted'] === 'Yes') {
        data['Pipeline Status'] = data['Surgery Date Given'] ? 'Scheduled' : 'Pending Date';
      } else {
        data['Pipeline Status'] = '';
      }
    } else if (sheetName === 'Diagnostic') {
      data['Diagnostic Revenue Auto (₹)'] = computeDiagnosticRevenue_(rates, data);
    }
  }

  // Counselling Pipeline Status is set inside the PRICE_MASTER_SHEETS block above.

  if (UPSERT_KEYS.hasOwnProperty(sheetName)) {
    const keys = UPSERT_KEYS[sheetName];
    const { rows } = getSheetData_(sheetName);
    const match = rows.find(r =>
      dateToStr_(r['Date']) === dateStr &&
      keys.every(k => String(r[k] || '').trim().toLowerCase() === String(data[k] || '').trim().toLowerCase())
    );
    if (match) {
      const newRow = headers.map(h => {
        if (h === 'Timestamp') return nowStr_();
        if (h === 'Submitted By') return user.name;
        if (h === 'Date' || h === 'Edit Status') return match[h]; // preserve original cell as-is
        return (data[h] !== undefined) ? data[h] : (match[h] !== undefined ? match[h] : '');
      });
      sheet.getRange(match._row, 1, 1, headers.length).setValues([newRow]);
      SpreadsheetApp.flush();
      logAudit_(user.name, 'UPDATE (resubmit)', sheetName, match._row, JSON.stringify(data));
      const updatedObj = {};
      headers.forEach((h, i) => updatedObj[h] = newRow[i]);
      updatedObj._row = match._row;
      return {
        success: true,
        message: `✓ ${desc} for ${niceDate} updated — today's entry corrected in place, not duplicated.`,
        date: dateStr, action: 'updated',
        todayRows: [updatedObj]
      };
    }
  }

  const row = headers.map(h => {
    if (h === 'Timestamp') return nowStr_();
    if (h === 'Date') return dateStr;
    if (h === 'Submitted By') return user.name;
    if (h === 'Edit Status') return 'Editable';
    return (data[h] !== undefined) ? data[h] : '';
  });
  sheet.appendRow(row);
  SpreadsheetApp.flush();
  const newRowNum = sheet.getLastRow();
  logAudit_(user.name, 'CREATE', sheetName, newRowNum, JSON.stringify(data));
  // Build a row object matching what getMyTodayEntries returns, so the
  // client can render "My entries today" immediately from this response
  // without a second round-trip (which has a race condition).
  const rowObj = {};
  headers.forEach((h, i) => rowObj[h] = row[i]);
  rowObj._row = newRowNum;

  // Close the TempAccess window now that staff has successfully submitted
  if (tempAccessRow) {
    try {
      const taSheet = SS.getSheetByName('TempAccess');
      if (taSheet) {
        const taHeaders = taSheet.getRange(1, 1, 1, taSheet.getLastColumn()).getValues()[0];
        const statusCol = taHeaders.indexOf('Status') + 1;
        if (statusCol > 0) { taSheet.getRange(tempAccessRow, statusCol).setValue('Used'); SpreadsheetApp.flush(); }
      }
    } catch(e) { /* don't fail the main submission if window closure fails */ }
  }

  const pastNote = dateStr !== todayStr_() ? ` (back-dated to ${niceDate})` : '';
  return {
    success: true,
    message: `✓ ${desc} for ${niceDate} submitted successfully.${pastNote}`,
    date: dateStr, action: 'created',
    todayRows: [rowObj]
  };
}

/**
 * Friendly department label for confirmation messages (e.g. "Counselling —
 * Daily Summary" instead of the raw sheet name "CounsellingSummary").
 */
function friendlyDeptName_(sheetName) {
  const map = {
    CounsellingSummary: 'Counselling — Daily Summary',
    OT: 'OT',
    HR: 'HR',
    TPA: 'TPA / Accounts'
  };
  return map[sheetName] || sheetName;
}

/**
 * Human-readable description of an entry for confirmation/edit messages —
 * includes the key (staff name / PRO name / doctor+procedure) for
 * per-key sheets so it's obvious WHICH record was saved/updated.
 */
function entryDescription_(sheetName, data) {
  const base = friendlyDeptName_(sheetName);
  if (sheetName === 'HR' && data['Staff Name']) return `${base} entry for ${data['Staff Name']}`;
  if (sheetName === 'Marketing' && data['PRO Name']) return `${base} entry for ${data['PRO Name']}`;
  if (sheetName === 'OT' && data['Doctor Name']) return `${base} entry (${data['Doctor Name']}${data['Procedure Type'] ? ' — ' + data['Procedure Type'] : ''})`;
  return `${base} entry`;
}

/**
 * The "department" a sheet's department-wide (key=[]) single daily entry
 * belongs to, for access-control purposes.
 */
function sheetDepartment_(sheetName) {
  if (sheetName === 'CounsellingSummary') return 'Counselling';
  return sheetName;
}

/**
 * Whether `user` may edit `rowObj` (today's row in `sheetName`) directly:
 *  - Must be today's date and not yet locked.
 *  - For department-wide single-entry sheets (UPSERT_KEYS key=[]): anyone
 *    in that department, or Admin/Doctor, can edit — it's one shared daily
 *    report, not personal data.
 *  - For everything else (per-key or append-only sheets): only the person
 *    who submitted that specific row (or Admin/Doctor).
 */
function canEditRow_(sheetName, rowObj, user) {
  if (dateToStr_(rowObj['Date']) !== todayStr_()) return false;
  if (rowObj['Edit Status'] && rowObj['Edit Status'] !== 'Editable') return false;
  if (user.role === 'Admin' || user.role === 'Doctor') return true;
  const keys = UPSERT_KEYS[sheetName];
  if (keys && keys.length === 0) {
    return normDept_(user.department) === normDept_(sheetDepartment_(sheetName));
  }
  return rowObj['Submitted By'] === user.name;
}

/** Case/whitespace-insensitive department comparison — Staff!Department is
 * free text and easy to enter slightly differently (e.g. "reception ",
 * "Front Office") from the department names this app expects. */
function normDept_(d) {
  // Also normalise 'Call Centre' → 'CallCentre' (sheet vs display name)
  return String(d || '').trim().toLowerCase().replace(/\s+/g, '');
}

// ---------- SAME-DAY EDIT ----------

/**
 * Returns today's editable rows this user is allowed to act on for a given
 * sheet — see canEditRow_ for the access rules (department-wide for
 * single-daily sheets, own-rows-only otherwise).
 */
function getMyTodayEntries(sheetName, user) {
  const { rows } = getSheetData_(sheetName);
  return rows.filter(r => canEditRow_(sheetName, r, user));
}

/**
 * Per-session debug helper — any logged-in user can call this for their OWN
 * session to see exactly why a row does/doesn't show in "My entries today".
 * Returns the user object as the server received it, plus, for every row
 * in the sheet, each individual canEditRow_ sub-check. Triggered via a
 * "Why is this empty?" link shown alongside the empty state.
 */
function debugMyEntries(sheetName, user) {
  const { rows } = getSheetData_(sheetName);
  const today = todayStr_();
  const keys = UPSERT_KEYS[sheetName];
  return {
    build: BUILD_VERSION,
    today,
    sheetName,
    user: { name: user.name, department: user.department, role: user.role, loginId: user.loginId },
    upsertKeys: keys,
    rowCount: rows.length,
    rows: rows.map(r => {
      const dateParsed = dateToStr_(r['Date']);
      return {
        _row: r._row,
        date_raw: (r['Date'] instanceof Date) ? r['Date'].toISOString() : String(r['Date']),
        date_type: Object.prototype.toString.call(r['Date']),
        date_parsed: dateParsed,
        date_matches_today: dateParsed === today,
        edit_status: r['Edit Status'],
        edit_status_ok: !(r['Edit Status'] && r['Edit Status'] !== 'Editable'),
        submitted_by: r['Submitted By'],
        submitted_by_matches_user: r['Submitted By'] === user.name,
        dept_compare: normDept_(user.department) + ' vs ' + normDept_(sheetDepartment_(sheetName)),
        dept_matches: normDept_(user.department) === normDept_(sheetDepartment_(sheetName)),
        can_edit_result: canEditRow_(sheetName, r, user)
      };
    })
  };
}

/**
 * For department-wide single-daily sheets (UPSERT_KEYS key=[]), returns
 * today's already-submitted row (if any) so the form can be pre-filled —
 * showing staff exactly what's currently on file, ready to correct and
 * re-save, instead of a blank form that invites a duplicate.
 * Returns null for per-key or append-only sheets (Counselling, OT,
 * Marketing, HR, Store, Purchase), where "My entries today" below the form
 * is the right place to review/edit today's records instead.
 */
function getTodayEntryForForm(sheetName, user) {
  const keys = UPSERT_KEYS[sheetName];
  if (!keys || keys.length > 0) return null;
  const { rows } = getSheetData_(sheetName);
  const today = todayStr_();
  return rows.find(r => dateToStr_(r['Date']) === today) || null;
}

/**
 * Direct edit — allowed only per canEditRow_ (same-day, not locked, and
 * either the row's owner or — for department-wide single-daily sheets —
 * anyone in that department, plus Admin/Doctor always).
 */
function updateMyEntry(sheetName, rowNum, data, user) {
  const sheet = SS.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const rowObj = {};
  headers.forEach((h, i) => rowObj[h] = rowValues[i]);

  if (!canEditRow_(sheetName, rowObj, user)) {
    return { success: false, message: 'This entry can no longer be edited directly. Please submit a correction request instead.' };
  }

  headers.forEach((h, i) => {
    if (data[h] !== undefined && !['Timestamp', 'Date', 'Submitted By', 'Edit Status'].includes(h)) {
      sheet.getRange(rowNum, i + 1).setValue(data[h]);
    }
  });
  sheet.getRange(rowNum, headers.indexOf('Timestamp') + 1).setValue(nowStr_());
  if (headers.indexOf('Submitted By') !== -1) sheet.getRange(rowNum, headers.indexOf('Submitted By') + 1).setValue(user.name);
  SpreadsheetApp.flush();
  logAudit_(user.name, 'SELF-EDIT', sheetName, rowNum, JSON.stringify(data));
  const niceDate = Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy');
  return { success: true, message: `✓ ${entryDescription_(sheetName, data)} for ${niceDate} updated successfully.` };
}

/**
 * Run this once daily (set up a time-driven trigger at midnight) to lock
 * all of yesterday's-and-earlier rows so they can no longer be self-edited.
 */
function lockPastEntries() {
  const sheetNames = ['Reception','Counselling','CounsellingSummary','OT','Diagnostic','Pharmacy','Optical','Marketing','HR','TPA','TPAClaims','TPASettlements','Operations','Store','Purchase','CallCentre','HRSummary','TPASummary'];
  const today = todayStr_();
  sheetNames.forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dateCol = headers.indexOf('Date') + 1;
    const statusCol = headers.indexOf('Edit Status') + 1;
    if (dateCol === 0 || statusCol === 0) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const dates = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (dateToStr_(dates[i][0]) !== today) {
        sheet.getRange(i + 2, statusCol).setValue('Locked');
      }
    }
  });
}

/**
 * Returns this user's entries from the last N days (default 14) for a given
 * sheet — used for the "Request Correction" view on locked/past entries.
 * For department-wide single-daily sheets, shows the whole department's
 * past entries (anyone in that department may request a correction);
 * for everything else, only the user's own past rows.
 */
function getMyRecentEntries(sheetName, user, days) {
  days = days || 14;
  const { rows } = getSheetData_(sheetName);
  const today = todayStr_();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - days);
  const cutoff = Utilities.formatDate(cutoffDate, TZ, 'yyyy-MM-dd');

  const keys = UPSERT_KEYS[sheetName];
  const isDeptWide = keys && keys.length === 0;
  const isAdminOrDoctor = (user.role === 'Admin' || user.role === 'Doctor');

  return rows.filter(r => {
    const d = dateToStr_(r['Date']);
    if (!d || d > today || d < cutoff) return false;
    if (isAdminOrDoctor) return true;
    if (isDeptWide) return normDept_(user.department) === normDept_(sheetDepartment_(sheetName));
    return r['Submitted By'] === user.name;
  }).sort((a, b) => b._row - a._row);
}

// ---------- CORRECTION REQUESTS (for locked/past entries) ----------

// ============================================================
//  TPA — CLAIM & SETTLEMENT TRACKING (FINAL ARCHITECTURE)
// ============================================================
//
//  Two actions for two people:
//  • TPA/cashless person: submitTPAClaim (when filing to CGHS/ECHS/Delhi Govt/TPA)
//    → if patient upgraded, also fills Co-pay Due + Co-pay Collected
//  • Billing person: submitTPASettlement (when money arrives — either from
//    govt/insurer OR from patient paying co-pay balance)
//  Both reference the same Claim Ref No — that's what links everything.
// ============================================================

/**
 * Log a new claim filed to CGHS/ECHS/Delhi Govt/TPA.
 * Auto-computes Co-pay Pending = Co-pay Due − Co-pay Collected.
 * Claim Ref No is mandatory — validated server-side.
 */
function submitTPAClaim(data, user) {
  try { return submitTPAClaim_(data, user); }
  catch(e) { return { success: false, message: '⚠ Error saving claim: ' + e.message }; }
}
function submitTPAClaim_(data, user) {
  const ref = String(data['Claim Ref No'] || '').trim();
  if (!ref) return { success: false, message: '⚠ Claim Ref No is required — this is what links the claim to its settlement. Use your file number, e.g. CGHS/2026/001.' };

  // Prevent duplicate refs
  const { rows } = getSheetData_('TPAClaims');
  const duplicate = rows.find(r => String(r['Claim Ref No'] || '').trim() === ref);
  if (duplicate) return {
    success: false,
    message: `⚠ Claim Ref "${ref}" already exists (filed on ${dateToStr_(duplicate['Date Filed'])}). Use a different reference number, or use Log Settlement to update an existing claim.`
  };

  const due = Number(data['Co-pay Due (₹)']) || 0;
  const collected = Number(data['Co-pay Collected (₹)']) || 0;
  data['Co-pay Pending (₹)'] = Math.max(0, due - collected);
  return submitEntry('TPAClaims', data, user);
}

/**
 * Log a payment received — either govt/TPA paying the claim, or patient
 * paying a co-pay balance. Claim Ref must match an existing TPAClaims row.
 * Auto-computes: Original Claim Amount (from claim), Total Settled So Far,
 * Shortfall, Days to Settle.
 */
function submitTPASettlement(data, user) {
  try { return submitTPASettlement_(data, user); }
  catch(e) { return { success: false, message: '⚠ Error saving settlement: ' + e.message }; }
}
function submitTPASettlement_(data, user) {
  const ref = String(data['Claim Ref No'] || '').trim();
  if (!ref) return { success: false, message: '⚠ Claim Ref No is required — it must match the number you used when filing the claim.' };

  // Find the original claim
  const { rows: claims } = getSheetData_('TPAClaims');
  const claim = claims.find(r => String(r['Claim Ref No'] || '').trim() === ref);
  if (!claim) return {
    success: false,
    message: `⚠ No claim found with Ref "${ref}". Please check the reference number, or file the claim first using Log Claim.`
  };

  // Pull original claim amount
  const originalAmount = Number(claim['Claim Amount (₹)']) || 0;
  data['Original Claim Amount (₹)'] = originalAmount;

  // Sum all previous settlements for this ref to compute running total
  const { rows: settlements } = getSheetData_('TPASettlements');
  const previousTotal = settlements
    .filter(s => String(s['Claim Ref No'] || '').trim() === ref)
    .reduce((sum, s) => sum + (Number(s['Amount Received (₹)']) || 0), 0);
  const thisAmount = Number(data['Amount Received (₹)']) || 0;
  const newTotal = previousTotal + thisAmount;

  data['Total Settled So Far (₹)'] = newTotal;
  data['Shortfall (₹)'] = Math.max(0, originalAmount - newTotal);

  // Days to settle (surgery date → today)
  if (claim['Surgery Date']) {
    try {
      const surgDate = new Date(claim['Surgery Date']);
      const recvDate = new Date();
      const days = Math.round((recvDate - surgDate) / (1000 * 60 * 60 * 24));
      data['Days to Settle'] = days > 0 ? days : 0;
    } catch (e) { data['Days to Settle'] = ''; }
  }

  // Also update Co-pay Collected / Pending on the claim row if this is a co-pay settlement
  if (data['Settlement Type'] === 'Patient Co-pay') {
    const sheet = SS.getSheetByName('TPAClaims');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const collectedIdx = headers.indexOf('Co-pay Collected (₹)') + 1;
    const pendingIdx = headers.indexOf('Co-pay Pending (₹)') + 1;
    if (collectedIdx > 0 && pendingIdx > 0) {
      const prevCollected = Number(claim['Co-pay Collected (₹)']) || 0;
      const newCollected = prevCollected + thisAmount;
      const due = Number(claim['Co-pay Due (₹)']) || 0;
      sheet.getRange(claim._row, collectedIdx).setValue(newCollected);
      sheet.getRange(claim._row, pendingIdx).setValue(Math.max(0, due - newCollected));
      SpreadsheetApp.flush();
    }
  }

  const result = submitEntry('TPASettlements', data, user);
  if (result.success) {
    // Update the claim status to Settled if shortfall is 0
    if (data['Shortfall (₹)'] === 0) {
      const sheet = SS.getSheetByName('TPAClaims');
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusIdx = headers.indexOf('Claim Status') + 1;
      if (statusIdx > 0) sheet.getRange(claim._row, statusIdx).setValue('Settled');
      SpreadsheetApp.flush();
    }
    result.claim = {
      patient: claim['Patient Name'] || '—',
      procedure: claim['Procedure'] || '—',
      originalAmount, newTotal, shortfall: data['Shortfall (₹)']
    };
  }
  return result;
}

/**
 * TPA Aging Report — Admin/Doctor only.
 * Returns: summary totals, pending claims (oldest first, 30-day flag),
 * settled claims with shortfall, and total co-pay still pending.
 */
/**
 * Debug: shows exactly what the server sees for the pipeline.
 * Used when pipeline is empty to diagnose why.
 */
// ============================================================
//  PIPELINE — dedicated functions (separate from dashboard)
//  Reads directly from the Counselling sheet — the only source
//  of per-patient pipeline data. Doctor/Admin only.
// ============================================================

/**
 * Returns all pipeline entries — scheduled and pending — with patient
 * detail. Called independently of the main dashboard for fast loading.
 */
function getPipelineData(user) {
  try {
    if (user.role !== 'Admin' && user.role !== 'Doctor') {
      return { success: false, message: 'Pipeline is visible to Admin and Doctor only.' };
    }
    const sheet = SS.getSheetByName('Counselling');
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: true,
        scheduled: [], pending: [],
        totalScheduled: 0, totalPending: 0,
        debug: {
          sheetExists: !!sheet,
          rowCount: sheet ? sheet.getLastRow() : 0,
          note: 'No data in Counselling sheet yet. Use Counselling form (Per-Patient mode) to add patients.'
        }
      };
    }

    const { rows } = getSheetData_('Counselling');
    const today = todayStr_();

    // Only rows where Converted = 'Yes' and not marked Surgery Done
    const converted = rows.filter(r => {
      const conv = String(r['Converted'] || '').trim().toLowerCase();
      const status = String(r['Pipeline Status'] || '').trim();
      return conv === 'yes' && status !== 'Surgery Done';
    });

    const scheduled = [];
    const pending   = [];
    const monthMap  = {}; // for KPI cards

    converted.forEach(r => {
      const entry = {
        _row:        r._row,
        patient:     r['Patient Reference'] || '—',
        procedure:   r['Procedure']          || '—',
        category:    r['Category']           || '—',
        counsellor:  r['Counselor Name']     || r['Submitted By'] || '—',
        dateLogged:  dateToStr_(r['Date']),
        status:      r['Pipeline Status'] || 'Pending Date'
      };

      const surgDate = r['Surgery Date Given'];
      if (surgDate) {
        let sd = '';
        try {
          sd = (surgDate instanceof Date)
            ? Utilities.formatDate(surgDate, TZ, 'yyyy-MM-dd')
            : String(surgDate).trim().slice(0, 10);
        } catch(e) {}
        entry.surgeryDate = sd;
        entry.isOverdue = sd && sd < today;

        // Month group
        const monthKey = sd ? sd.slice(0, 7) : 'Unknown'; // 'yyyy-MM'
        const monthLabel = sd ? Utilities.formatDate(
          new Date(sd + 'T12:00:00'), TZ, 'MMM yyyy') : 'No date';
        if (!monthMap[monthKey]) monthMap[monthKey] = { key: monthKey, label: monthLabel, count: 0 };
        monthMap[monthKey].count++;
        scheduled.push(entry);
      } else {
        pending.push(entry);
      }
    });

    // Sort scheduled by surgery date (soonest first)
    scheduled.sort((a, b) => (a.surgeryDate || '').localeCompare(b.surgeryDate || ''));

    // Month KPI cards sorted chronologically
    const monthKPIs = Object.values(monthMap).sort((a, b) => a.key.localeCompare(b.key));

    return {
      success: true,
      scheduled, pending,
      totalScheduled: scheduled.length,
      totalPending: pending.length,
      monthKPIs,
      debug: {
        sheetExists: true,
        rowCount: rows.length,
        convertedCount: converted.length,
        note: rows.length === 0 ? 'Counselling sheet is empty. Submit counselling entries in Per-Patient mode.' :
              converted.length === 0 ? 'No rows with Converted=Yes found. Make sure to set Converted to Yes in the counselling form.' : ''
      }
    };
  } catch(e) {
    Logger.log('getPipelineData ERROR: ' + e.message);
    return { success: false, message: 'Pipeline error: ' + e.message };
  }
}

/**
 * Marks a counselling row as Surgery Done — removes it from the active pipeline.
 * Called when Admin/Doctor clicks "Mark Done" on a patient.
 */
function markPipelineDone(rowNum, user) {
  try {
    if (user.role !== 'Admin' && user.role !== 'Doctor') {
      return { success: false, message: 'Only Admin or Doctor can mark pipeline entries as done.' };
    }
    const sheet = SS.getSheetByName('Counselling');
    if (!sheet) return { success: false, message: 'Counselling sheet not found.' };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('Pipeline Status') + 1;
    if (statusCol === 0) return { success: false, message: 'Pipeline Status column not found. Run setupHospitalSheet().' };
    sheet.getRange(rowNum, statusCol).setValue('Surgery Done');
    SpreadsheetApp.flush();
    logAudit_(user.name, 'PIPELINE_DONE', 'Counselling', rowNum, 'Marked Surgery Done');
    return { success: true, message: 'Patient marked as Surgery Done and removed from active pipeline.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Builds today's daily summary as a WhatsApp-ready text message.
 * Returns the message text + the phone numbers to send to.
 * No API needed — frontend opens wa.me links with pre-filled text.
 */
function getDailyWhatsAppReport(user) {
  try {
    const settings  = getSettings();
    const TZ_       = SS.getSpreadsheetTimeZone();
    const today     = Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd');
    const hospital  = settings['Hospital Name'] || 'EyeForce Hospital';
    const DAYS      = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const MONTHS    = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const now       = new Date();
    const niceDate  = DAYS[now.getDay()] + ', ' + now.getDate() + ' ' + MONTHS[now.getMonth()] + ' ' + now.getFullYear();

    function filterToday(sn) {
      try {
        const { rows } = getSheetData_(sn);
        return rows.filter(r => {
          const d = r['Date'];
          if (!d) return false;
          try { return Utilities.formatDate(new Date(d), TZ_, 'yyyy-MM-dd') === today; }
          catch(e2) { return String(d).slice(0,10) === today; }
        });
      } catch(e2) { return []; }
    }
    function s(n) { return Number(n)||0; }
    function fmt(n) { return '\u20b9' + s(n).toLocaleString('en-IN'); }
    function pct(a,b) { return b>0 ? Math.round((a/b)*100)+'%' : '0%'; }
    function bar(val,total,len) {
      if (!total) return '\u2591'.repeat(len);
      const f = Math.round((val/total)*len);
      return '\u2588'.repeat(f) + '\u2591'.repeat(len-f);
    }

    const rec  = filterToday('Reception');
    const ot   = filterToday('OT');
    const diag = filterToday('Diagnostic');
    const phar = filterToday('Pharmacy');
    const opt  = filterToday('Optical');
    const coun = filterToday('Counselling');
    const call = filterToday('CallCentre');
    const mkt  = filterToday('Marketing');
    const ops  = filterToday('Operations');
    const tpa  = filterToday('TPA');

    const totalOPD  = rec.reduce((a,r)=>a+s(r['Total OPD']),0);
    const newPat    = rec.reduce((a,r)=>a+s(r['New Patients']),0);
    const oldPat    = rec.reduce((a,r)=>a+s(r['Old/Follow-up']),0);
    const focPat    = rec.reduce((a,r)=>a+s(r['FOC Patients']),0);
    const noShows   = rec.reduce((a,r)=>a+s(r['No-Shows Today']),0);
    const cghs      = rec.reduce((a,r)=>a+s(r['CGHS']),0);
    const echs      = rec.reduce((a,r)=>a+s(r['ECHS']),0);
    const delhiGovt = rec.reduce((a,r)=>a+s(r['Delhi Govt']),0);
    const cashPat   = rec.reduce((a,r)=>a+s(r['Cash']),0);
    const otherTPA  = rec.reduce((a,r)=>a+s(r['Other TPA/Insurance']),0);
    const srcSocial = rec.reduce((a,r)=>a+s(r['Source - Social Media']),0);
    const srcRef    = rec.reduce((a,r)=>a+s(r['Source - Referral']),0);
    const srcWalk   = rec.reduce((a,r)=>a+s(r['Source - Walk-in/Other']),0);
    const srcNews   = rec.reduce((a,r)=>a+s(r['Source - Newspaper/Hoarding']),0);
    const opdCash   = rec.reduce((a,r)=>a+s(r['OPD Cash Collected (\u20b9)']),0);
    const otCash    = rec.reduce((a,r)=>a+s(r['OT Cash Collected (\u20b9)']),0);
    const diagCash  = rec.reduce((a,r)=>a+s(r['Diagnostic Cash Collected (\u20b9)']),0);
    const othCash   = rec.reduce((a,r)=>a+s(r['Others Cash Collected (\u20b9)']),0);
    const pharmRev  = phar.reduce((a,r)=>a+s(r['Revenue Collected (\u20b9)']),0);
    const optRev    = opt.reduce((a,r)=>a+s(r['Revenue Collected (\u20b9)']),0);
    const totalCash = opdCash+otCash+diagCash+othCash+pharmRev+optRev;
    const otCashC   = ot.reduce((a,r)=>a+s(r['Cash Cases']),0);
    const otCGHS    = ot.reduce((a,r)=>a+s(r['CGHS Cases']),0);
    const otECHS    = ot.reduce((a,r)=>a+s(r['ECHS Cases']),0);
    const otDelhi   = ot.reduce((a,r)=>a+s(r['Delhi Govt Cases']),0);
    const otTPA     = ot.reduce((a,r)=>a+s(r['TPA Cases']),0);
    const totalSurg = otCashC+otCGHS+otECHS+otDelhi+otTPA;
    const otExp     = ot.reduce((a,r)=>a+s(r['Total Expected Claims (\u20b9)']),0);
    const totalCoun = coun.length;
    const converted = coun.filter(r=>String(r['Converted']||'').toLowerCase()==='yes').length;
    const followUp  = coun.reduce((a,r)=>a+s(r['Follow-up Calls Made']),0);
    let pipelineSch = 0, pipelinePend = 0;
    try {
      const { rows: ac } = getSheetData_('Counselling');
      ac.forEach(r => {
        if (String(r['Converted']||'').toLowerCase()==='yes' && String(r['Pipeline Status']||'')!=='Surgery Done') {
          if (r['Surgery Date Given']) pipelineSch++; else pipelinePend++;
        }
      });
    } catch(e2) {}
    const diagCats = ['Biometry','Topography','OCT','Cataract Work-up','LASIK Work-up','Green Laser','Fundus Photography','Visual Field/Perimetry','B-Scan/A-Scan','Other Tests'];
    let dPaid=0, dFree=0;
    diagCats.forEach(c=>{ dPaid+=diag.reduce((a,r)=>a+s(r[c+' - Paid']),0); dFree+=diag.reduce((a,r)=>a+s(r[c+' - Free']),0); });
    const inbound  = call.reduce((a,r)=>a+s(r['Inbound Calls Total']),0);
    const answered = call.reduce((a,r)=>a+s(r['Inbound Calls Answered']),0);
    const missed   = call.reduce((a,r)=>a+s(r['Inbound Calls Missed']),0);
    const outbound = call.reduce((a,r)=>a+s(r['Outbound Calls Total']),0);
    const apptBkd  = call.reduce((a,r)=>a+s(r['Appointments Booked (from calls)']),0);
    const surgBkd  = call.reduce((a,r)=>a+s(r['Surgery Bookings Confirmed (from calls)']),0);
    const missedR  = inbound>0 ? Math.round((missed/inbound)*100) : 0;
    const mktLeads = mkt.reduce((a,r)=>a+s(r['Leads Generated']),0);
    const mktConv  = mkt.reduce((a,r)=>a+s(r['Leads Converted to OPD']),0);
    const mktCalls = mkt.reduce((a,r)=>a+s(r['Calls Made']),0);
    const mktVisit = mkt.reduce((a,r)=>a+s(r['Visits Made']),0);
    const avgWait  = ops.length>0 ? Math.round(ops.reduce((a,r)=>a+s(r['Avg Wait Time (mins)']),0)/ops.length) : 0;
    const compRec  = ops.reduce((a,r)=>a+s(r['Complaints Received']),0);
    const compRes  = ops.reduce((a,r)=>a+s(r['Complaints Resolved']),0);
    const revAsked = ops.reduce((a,r)=>a+s(r['Google Reviews Asked']),0);
    const revRec   = ops.reduce((a,r)=>a+s(r['Google Reviews Received']),0);
    const rxTotal  = phar.reduce((a,r)=>a+s(r['Total Prescriptions']),0);
    const optWalk  = opt.reduce((a,r)=>a+s(r['Walk-in Customers']),0);
    const optFr    = opt.reduce((a,r)=>a+s(r['Frames Sold']),0);
    const optLn    = opt.reduce((a,r)=>a+s(r['Lenses Sold']),0);
    let tpaClaims=0, tpaSett=0;
    try {
      const { rows: tc } = getSheetData_('TPAClaims');
      const { rows: ts } = getSheetData_('TPASettlements');
      tpaClaims = tc.filter(r=>String(r['Date']||'').slice(0,10)===today).reduce((a,r)=>a+s(r['Claim Amount (\u20b9)']),0);
      tpaSett   = ts.filter(r=>String(r['Date']||'').slice(0,10)===today).reduce((a,r)=>a+s(r['Amount Received (\u20b9)']),0);
    } catch(e2) {}
    const flags = [];
    if (missedR>20) flags.push('\ud83d\udcf5 Missed call rate high: '+missedR+'%');
    if (noShows>5)  flags.push('\u26a0\ufe0f High no-shows: '+noShows+' patients');
    if (compRec>compRes) flags.push('\ud83d\udea8 '+(compRec-compRes)+' complaint(s) unresolved');
    if (totalCoun>0 && converted<totalCoun/2) flags.push('\ud83d\udcc9 Counselling conversion below 50%');

    const L  = '\n';
    const DIV = '\u2501'.repeat(24)+L;
    let msg = DIV;
    msg += '\ud83d\udccb *DAILY PERFORMANCE REPORT*'+L;
    msg += '\ud83c\udfe5 *'+hospital+'*'+L;
    msg += '\ud83d\udcc5 '+niceDate+L;
    msg += DIV;
    msg += L+'\ud83d\udc41\ufe0f *OPD \u2014 FRONT DESK*'+L;
    msg += '\u250c'+('\u2500'.repeat(26))+'\u2510'+L;
    msg += '\u2502 Total Patients : *'+String(totalOPD).padStart(4)+'*     \u2502'+L;
    msg += '\u2502 Fresh / New    : '+String(newPat).padStart(4)+'      \u2502'+L;
    msg += '\u2502 Follow-up      : '+String(oldPat).padStart(4)+'      \u2502'+L;
    msg += '\u2502 FOC (Free)     : '+String(focPat).padStart(4)+'      \u2502'+L;
    msg += '\u2502 No-Shows       : '+String(noShows).padStart(4)+'      \u2502'+L;
    msg += '\u2514'+('\u2500'.repeat(26))+'\u2518'+L;
    msg += L+'\ud83d\udcb3 *PAYER MIX*'+L;
    msg += '  \ud83d\udfe2 Cash       : '+cashPat+' ('+pct(cashPat,totalOPD)+')'+L;
    msg += '  \ud83d\udd35 CGHS       : '+cghs+' ('+pct(cghs,totalOPD)+')'+L;
    msg += '  \ud83d\udfe3 ECHS       : '+echs+' ('+pct(echs,totalOPD)+')'+L;
    msg += '  \ud83d\udfe0 Delhi Govt : '+delhiGovt+' ('+pct(delhiGovt,totalOPD)+')'+L;
    msg += '  \u26aa Other TPA  : '+otherTPA+' ('+pct(otherTPA,totalOPD)+')'+L;
    if (totalOPD>0) {
      msg += L+'\ud83c\udf10 *PATIENT SOURCE*'+L;
      msg += '  \ud83d\udc68\u200d\u2695\ufe0f Referral    : '+srcRef+' '+bar(srcRef,totalOPD,8)+L;
      msg += '  \ud83d\udeb6 Walk-in     : '+srcWalk+' '+bar(srcWalk,totalOPD,8)+L;
      msg += '  \ud83d\udcf1 Social Media: '+srcSocial+' '+bar(srcSocial,totalOPD,8)+L;
      msg += '  \ud83d\udcf0 Newspaper   : '+srcNews+' '+bar(srcNews,totalOPD,8)+L;
    }
    msg += L+'\ud83d\udd2c *OT \u2014 SURGICAL CASES*'+L;
    msg += '\u250c'+('\u2500'.repeat(26))+'\u2510'+L;
    msg += '\u2502 Total Surgeries : *'+String(totalSurg).padStart(3)+'*      \u2502'+L;
    msg += '\u2502 Cash            : '+String(otCashC).padStart(3)+'       \u2502'+L;
    msg += '\u2502 CGHS            : '+String(otCGHS).padStart(3)+'       \u2502'+L;
    msg += '\u2502 ECHS            : '+String(otECHS).padStart(3)+'       \u2502'+L;
    msg += '\u2502 Delhi Govt      : '+String(otDelhi).padStart(3)+'       \u2502'+L;
    msg += '\u2502 TPA / Insurance : '+String(otTPA).padStart(3)+'       \u2502'+L;
    msg += '\u2514'+('\u2500'.repeat(26))+'\u2518'+L;
    if (otExp>0) msg += '  \ud83d\udcbc Claims to file: *'+fmt(otExp)+'*'+L;
    msg += L+'\ud83d\udcac *COUNSELLING*'+L;
    msg += '  \ud83d\udc65 Counselled Today : '+totalCoun+L;
    msg += '  \u2705 Converted       : *'+converted+'* ('+pct(converted,totalCoun)+')'+L;
    msg += '  \ud83d\udd04 Follow-up Calls  : '+followUp+L;
    if (totalCoun>0) msg += '  '+bar(converted,totalCoun,15)+' '+pct(converted,totalCoun)+L;
    msg += L+'\ud83d\uddd3\ufe0f *SURGERY PIPELINE*'+L;
    msg += '  \ud83d\udccc Scheduled (date given) : *'+pipelineSch+'*'+L;
    msg += '  \u23f3 Pending (no date yet) : *'+pipelinePend+'*'+L;
    msg += '  \ud83c\udfe5 Total Active Pipeline  : *'+(pipelineSch+pipelinePend)+'*'+L;
    if (dPaid+dFree>0) {
      msg += L+'\ud83d\udd2d *DIAGNOSTIC*'+L;
      msg += '  \ud83d\udcb0 Paid Tests : *'+dPaid+'*'+L;
      msg += '  \ud83c\udd93 Free Tests : '+dFree+L;
      msg += '  \ud83d\udcca Total      : *'+(dPaid+dFree)+'*'+L;
    }
    if (inbound+outbound>0) {
      msg += L+'\ud83d\udcde *CALL CENTRE*'+L;
      msg += '\u250c'+('\u2500'.repeat(26))+'\u2510'+L;
      msg += '\u2502 Inbound    : '+String(inbound).padStart(3)+'              \u2502'+L;
      msg += '\u2502  \u2705 Answered: '+String(answered).padStart(3)+'              \u2502'+L;
      msg += '\u2502  \u274c Missed  : '+String(missed).padStart(3)+' ('+missedR+'%)         \u2502'+L;
      msg += '\u2502 Outbound   : '+String(outbound).padStart(3)+'              \u2502'+L;
      msg += '\u2502 Appts Bkd  : '+String(apptBkd).padStart(3)+'              \u2502'+L;
      msg += '\u2502 Surg Bkd   : '+String(surgBkd).padStart(3)+'              \u2502'+L;
      msg += '\u2514'+('\u2500'.repeat(26))+'\u2518'+L;
    }
    if (mktLeads+mktCalls>0) {
      msg += L+'\ud83d\udce3 *MARKETING / PRO*'+L;
      msg += '  \ud83d\udcde Calls Made     : '+mktCalls+L;
      msg += '  \ud83d\ude97 Visits Made    : '+mktVisit+L;
      msg += '  \ud83c\udfaf Leads Generated: *'+mktLeads+'*'+L;
      msg += '  \u2705 Leads to OPD  : '+mktConv+L;
    }
    if (rxTotal>0) {
      msg += L+'\ud83d\udc8a *PHARMACY*'+L;
      msg += '  \ud83d\udccb Total Rx : '+rxTotal+L;
      msg += '  \ud83d\udcb0 Revenue  : *'+fmt(pharmRev)+'*'+L;
    }
    if (optWalk+optFr>0) {
      msg += L+'\ud83d\udc53 *OPTICAL SHOP*'+L;
      msg += '  \ud83d\udc64 Walk-ins  : '+optWalk+L;
      msg += '  \ud83d\udd76\ufe0f Frames    : '+optFr+L;
      msg += '  \ud83d\udd0d Lenses    : '+optLn+L;
      msg += '  \ud83d\udcb0 Revenue   : *'+fmt(optRev)+'*'+L;
    }
    if (tpaClaims+tpaSett>0) {
      msg += L+'\ud83c\udfe6 *TPA / ACCOUNTS*'+L;
      if (tpaClaims>0) msg += '  \ud83d\udce4 Claims Filed Today   : *'+fmt(tpaClaims)+'*'+L;
      if (tpaSett>0)   msg += '  \ud83d\udce5 Settlements Received : *'+fmt(tpaSett)+'*'+L;
    }
    if (ops.length>0) {
      msg += L+'\u2699\ufe0f *OPERATIONS*'+L;
      if (avgWait>0)   msg += '  \u23f1\ufe0f Avg Wait Time : '+avgWait+' mins'+L;
      if (compRec>0)   msg += '  \ud83d\udce2 Complaints    : '+compRec+' received, '+compRes+' resolved'+L;
      if (revAsked>0)  msg += '  \u2b50 Google Reviews: '+revRec+' received ('+revAsked+' asked)'+L;
    }
    msg += L+DIV;
    msg += '\ud83d\udcb0 *REVENUE SUMMARY \u2014 TODAY*'+L;
    msg += DIV;
    msg += '  OPD Fees      : '+fmt(opdCash)+L;
    msg += '  OT / Surgery  : '+fmt(otCash)+L;
    msg += '  Diagnostic    : '+fmt(diagCash)+L;
    msg += '  Pharmacy      : '+fmt(pharmRev)+L;
    msg += '  Optical       : '+fmt(optRev)+L;
    if (othCash>0) msg += '  Others        : '+fmt(othCash)+L;
    msg += '  ─────────────────────────'+L;
    msg += '  *TOTAL COLLECTED : '+fmt(totalCash)+'*'+L;
    if (tpaSett>0) msg += '  TPA Received     : '+fmt(tpaSett)+L;
    msg += '  *GRAND TOTAL     : '+fmt(totalCash+tpaSett)+'*'+L;
    if (flags.length>0) {
      msg += L+'\ud83d\udea8 *ALERTS & FLAGS*'+L;
      flags.forEach(f => { msg += '  '+f+L; });
    }

    // ── Departments that have NOT submitted today's report ─────────────────
    const deptStatus = [
      { name: 'Reception',    rows: rec,  label: 'Reception / Front Desk' },
      { name: 'OT',           rows: ot,   label: 'OT / Operation Theatre' },
      { name: 'Counselling',  rows: coun, label: 'Counselling' },
      { name: 'Diagnostic',   rows: diag, label: 'Diagnostic' },
      { name: 'Pharmacy',     rows: phar, label: 'Pharmacy' },
      { name: 'Optical',      rows: opt,  label: 'Optical Shop' },
      { name: 'CallCentre',   rows: call, label: 'Call Centre' },
      { name: 'Marketing',    rows: mkt,  label: 'Marketing / PRO' },
      { name: 'Operations',   rows: ops,  label: 'Operations' },
      { name: 'TPA',          rows: tpa,  label: 'TPA / Accounts' },
    ];
    const notUpdated = deptStatus.filter(d => d.rows.length === 0).map(d => d.label);
    const updated    = deptStatus.filter(d => d.rows.length > 0).map(d => d.label);

    msg += L+'\ud83d\udcca *REPORT STATUS*'+L;
    if (notUpdated.length === 0) {
      msg += '  \u2705 All departments have submitted today\'s data'+L;
    } else {
      msg += '  \u2705 Submitted ('+updated.length+'): '+updated.join(', ')+L;
      msg += L+'  \u274c *NOT YET UPDATED ('+notUpdated.length+'):*'+L;
      notUpdated.forEach(d => { msg += '  \u26a0\ufe0f '+d+' \u2014 data not entered in system'+L; });
      msg += L+'  _Please ensure above departments submit their daily report._'+L;
    }

    msg += L+DIV;
    msg += '\u2705 _Report by EyeForce Solutions_'+L;
    msg += '_Practice & Business Intelligence_'+L;
    msg += DIV;

    // ── When should Admin send this report? ────────────────────────────────
    // Best time: after all departments have submitted for the day — typically
    // 7 PM to 8 PM. Avoid sending before 5 PM as data will be incomplete.
    const sendTimeNote = 'Best time to send: after 7 PM when all departments have submitted their data for the day.';

    // Safe phone extraction — handles string, number, or undefined from Settings
    function safePhone(val) {
      if (val === null || val === undefined) return '';
      return String(val).replace(/\D/g, '');
    }
    const adminPhone  = safePhone(settings['WhatsApp - Admin Phone']);
    const doctorPhone = safePhone(settings['WhatsApp - Doctor Phone']);
    return {
      success: true,
      message: msg,
      adminPhone,
      doctorPhone,
      today: niceDate,
      notUpdated,
      updatedCount: updated.length,
      totalDepts: deptStatus.length,
      sendTimeNote
    };
  } catch(e) {
    Logger.log('getDailyWhatsAppReport error: '+e.message);
    return { success:false, message:'Error: '+e.message };
  }
}

function getPipelineDebug(user) {
  const counsellingMode = (getSettings()['Counselling Mode (Per-Patient / Daily Summary)'] || 'Per-Patient');
  const { headers, rows } = getSheetData_('Counselling');
  const { rows: summaryRows } = getSheetData_('CounsellingSummary');
  const convertedRows = rows.filter(r => String(r['Converted'] || '').toLowerCase() === 'yes');
  const withDate = convertedRows.filter(r => r['Surgery Date Given']);

  return {
    counsellingMode,
    counsellingSheet: {
      exists: headers.length > 0,
      headers: headers,
      totalRows: rows.length,
      convertedCount: convertedRows.length,
      withDateCount: withDate.length,
      firstFewRows: rows.slice(0, 5).map(r => ({
        date: dateToStr_(r['Date']),
        submittedBy: r['Submitted By'],
        converted: r['Converted'],
        surgeryDate: r['Surgery Date Given'] ? String(r['Surgery Date Given']) : '(empty)',
        procedure: r['Procedure']
      }))
    },
    summarySheetRows: summaryRows.length
  };
}

function getTPAAgingReport(flagDays, user) {
  if (user.role !== 'Admin' && user.role !== 'Doctor') {
    return { success: false, message: 'TPA Aging Report is Admin/Doctor only.' };
  }
  flagDays = flagDays || 30;
  const today = new Date();

  let claims = [], settlements = [];
  try { claims      = getSheetData_('TPAClaims').rows; }    catch(e) {}
  try { settlements = getSheetData_('TPASettlements').rows; } catch(e) {}

  if (claims.length === 0 && settlements.length === 0) {
    return {
      success: true, flagDays,
      summary: { pendingClaimsCount:0, pendingClaimsAmount:0, flaggedOlderThan30:0,
                 flaggedAmount:0, shortfallCount:0, totalCopayPending:0 },
      pendingClaims: [], shortfallClaims: [],
      emptyState: 'No TPA claims or settlements have been logged yet. Use TPA → Log Claim to start tracking.'
    };
  }

  // Build per-ref settlement totals
  const settledAmounts = {};
  const settledStatus  = {};
  settlements.forEach(s => {
    const ref = String(s['Claim Ref No'] || '').trim();
    if (!ref) return;
    settledAmounts[ref] = (settledAmounts[ref] || 0) + (Number(s['Amount Received (₹)']) || 0);
    settledStatus[ref]  = 'partial'; // at least one settlement exists
  });

  // Pending claims = not fully settled
  const pendingClaims = [];
  const shortfallClaims = [];
  let totalCopayPending = 0;
  let totalClaimsAmount = 0;
  let flaggedCount = 0, flaggedAmount = 0;

  claims.forEach(c => {
    const ref = String(c['Claim Ref No'] || '').trim();
    const claimAmount = Number(c['Claim Amount (₹)']) || 0;
    const settled = settledAmounts[ref] || 0;
    const shortfall = Math.max(0, claimAmount - settled);
    const copayPending = Number(c['Co-pay Pending (₹)']) || 0;
    totalCopayPending += copayPending;

    const filedDate = c['Date Filed'] ? new Date(c['Date Filed']) : null;
    const daysOld = filedDate ? Math.round((today - filedDate) / (1000 * 60 * 60 * 24)) : null;
    const isFlagged = daysOld !== null && daysOld > flagDays;

    if (shortfall > 0) {
      totalClaimsAmount += shortfall;
      if (isFlagged) { flaggedCount++; flaggedAmount += shortfall; }
      pendingClaims.push({
        ref, payer: c['Payer Type'] + (c['Insurer / Payer Name'] ? ' — ' + c['Insurer / Payer Name'] : ''),
        patient: c['Patient Name'] || '—',
        procedure: c['Procedure'] || '—',
        surgeryDate: c['Surgery Date'] || '—',
        dateFiled: dateToStr_(c['Date Filed']),
        claimAmount, settled, outstanding: shortfall,
        daysOld, isFlagged, status: c['Claim Status'] || 'Filed',
        copayPending
      });
      // Also flag as shortfall if partially settled
      if (settled > 0) {
        shortfallClaims.push({ ref, patient: c['Patient Name'] || '—', procedure: c['Procedure'] || '—',
          payer: c['Payer Type'] || '—', claimAmount, settled, shortfall,
          status: String(c['Claim Status'] || '') });
      }
    }
  });
  pendingClaims.sort((a, b) => (b.daysOld || 0) - (a.daysOld || 0));

  return {
    success: true, flagDays,
    summary: {
      pendingClaimsCount: pendingClaims.length,
      pendingClaimsAmount: totalClaimsAmount,
      flaggedOlderThan30: flaggedCount,
      flaggedAmount,
      shortfallCount: shortfallClaims.length,
      totalCopayPending
    },
    pendingClaims,
    shortfallClaims
  };
}

// ============================================================
//  TEMPORARY PAST-DATE ACCESS WINDOWS
//  Admin grants a staff member permission to submit data for
//  a specific past date. Staff fills the form themselves.
//  Admin is only the gate-opener, not the data-filler.
// ============================================================

/**
 * Admin grants a past-date window to a staff member.
 * Returns error if a window for that staff+date already exists and is Open.
 */
function getStaffList(user) {
  if (user.role !== 'Admin') return [];
  const { rows } = getSheetData_('Staff');
  return rows
    .filter(r => r['Status'] === 'Active')
    .map(r => ({ name: r['Staff Name'], department: r['Department'], role: r['Role'] }));
}

function getStaffList(user) {
  if (user.role !== 'Admin') return [];
  const { rows } = getSheetData_('Staff');
  return rows.map(r => ({
    _row: r._row,
    name: r['Staff Name'],
    department: r['Department'],
    role: r['Role'],
    loginId: r['Login ID'],
    status: r['Status']
    // PIN deliberately excluded from client response for security
  }));
}

/**
 * Toggle a staff member Active ↔ Inactive.
 */
function toggleStaffStatus(rowNum, newStatus, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Admin only.' };
  const sheet = SS.getSheetByName('Staff');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  if (statusCol === 0) return { success: false, message: 'Status column not found.' };
  sheet.getRange(rowNum, statusCol).setValue(newStatus);
  SpreadsheetApp.flush();
  logAudit_(user.name, 'STAFF_STATUS', 'Staff', rowNum, newStatus);
  return { success: true, message: 'Status updated to ' + newStatus + '.' };
}

/**
 * Update a staff member's Login ID or PIN.
 */
function updateStaffCredentials(rowNum, loginId, pin, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Admin only.' };
  if (!loginId || loginId.trim().length < 3) return { success: false, message: 'Login ID must be at least 3 characters.' };
  if (pin && (pin.length < 4 || !/^\d+$/.test(pin))) return { success: false, message: 'PIN must be 4+ digits.' };
  const sheet = SS.getSheetByName('Staff');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Check login ID is unique (excluding current row)
  const { rows } = getSheetData_('Staff');
  const conflict = rows.find(r => r['Login ID'] === loginId.trim() && r._row !== rowNum);
  if (conflict) return { success: false, message: 'Login ID "' + loginId + '" is already used by ' + conflict['Staff Name'] + '.' };
  const loginCol = headers.indexOf('Login ID') + 1;
  const pinCol   = headers.indexOf('PIN') + 1;
  if (loginCol > 0) sheet.getRange(rowNum, loginCol).setValue(loginId.trim());
  if (pin && pinCol > 0) sheet.getRange(rowNum, pinCol).setValue(pin.trim());
  SpreadsheetApp.flush();
  logAudit_(user.name, 'STAFF_CREDS', 'Staff', rowNum, 'Login ID updated');
  return { success: true, message: 'Credentials updated for row ' + rowNum + '.' };
}

/**
 * Add a new staff member.
 */
function addStaffMember(name, department, role, loginId, pin, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Admin only.' };
  if (!name || !loginId || !pin) return { success: false, message: 'Name, Login ID and PIN are all required.' };
  if (!/^\d{4,}$/.test(pin)) return { success: false, message: 'PIN must be 4+ digits.' };
  const { rows } = getSheetData_('Staff');
  if (rows.find(r => r['Login ID'] === loginId.trim())) {
    return { success: false, message: 'Login ID "' + loginId + '" is already in use.' };
  }
  const sheet = SS.getSheetByName('Staff');
  sheet.appendRow([name.trim(), department || '', role || 'Staff', loginId.trim(), pin.trim(), 'Active']);
  SpreadsheetApp.flush();
  return { success: true, message: '✓ ' + name + ' added successfully as ' + (role||'Staff') + ' in ' + (department||'—') + '.' };
}

function grantTempAccess(staffName, department, allowedDate, reason, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Only Admin can grant past-date access.' };
  if (!staffName || !allowedDate) return { success: false, message: 'Staff name and date are required.' };
  if (allowedDate >= todayStr_()) return { success: false, message: 'Allowed date must be in the past.' };

  const sheet = SS.getSheetByName('TempAccess');
  if (!sheet) return { success: false, message: 'TempAccess sheet not found — run setupHospitalSheet().' };

  const { rows } = getSheetData_('TempAccess');
  const existing = rows.find(r =>
    r['Staff Name'] === staffName &&
    dateToStr_(r['Allowed Date']) === allowedDate &&
    r['Status'] === 'Open'
  );
  if (existing) return {
    success: false,
    message: `A window for ${staffName} on ${allowedDate} is already open. Close it first before granting a new one.`
  };

  sheet.appendRow([
    staffName, department, allowedDate, reason || '',
    user.name, nowStr_(), 'Open'
  ]);
  SpreadsheetApp.flush();
  return { success: true, message: `✓ Window opened for ${staffName} to submit data for ${allowedDate}. They will see the date option when they next open their form.` };
}

/**
 * Returns all open windows for the current user (for staff form).
 * Only returns 'Open' status windows.
 */
function getMyTempAccess(user) {
  const { rows } = getSheetData_('TempAccess');
  return rows
    .filter(r => r['Staff Name'] === user.name && r['Status'] === 'Open')
    .map(r => ({
      allowedDate: dateToStr_(r['Allowed Date']),
      department: r['Department'] || '',
      reason: r['Reason / Note'] || '',
      _row: r._row
    }))
    .filter(r => r.allowedDate && r.allowedDate < todayStr_()); // must be a past date
}

/**
 * Admin cancels an open window (before staff uses it).
 */
function cancelTempAccess(rowNum, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Only Admin can cancel access windows.' };
  const sheet = SS.getSheetByName('TempAccess');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  if (statusCol > 0) sheet.getRange(rowNum, statusCol).setValue('Cancelled');
  SpreadsheetApp.flush();
  return { success: true, message: 'Window cancelled.' };
}

/**
 * Returns all TempAccess windows (Open/Used/Cancelled) for Admin management UI.
 */
function getAllTempAccess(user) {
  if (user.role !== 'Admin') return { success: false, message: 'Admin only.' };
  const { rows } = getSheetData_('TempAccess');
  return {
    success: true,
    windows: rows.map(r => ({
      _row: r._row,
      staffName: r['Staff Name'],
      department: r['Department'],
      allowedDate: dateToStr_(r['Allowed Date']),
      reason: r['Reason / Note'],
      grantedBy: r['Granted By'],
      grantedOn: r['Granted On'],
      status: r['Status']
    })).reverse() // newest first
  };
}

function submitCorrectionRequest(sheetName, rowNum, fieldName, oldValue, newValue, reason, user) {
  const sheet = SS.getSheetByName('CorrectionRequests');
  sheet.appendRow([nowStr_(), sheetName, rowNum, fieldName, oldValue, newValue, reason, user.name, 'Pending', '', '']);
  notifyCorrectionRequest_(sheetName, fieldName, oldValue, newValue, reason, user);
  return { success: true, message: 'Correction request sent to Admin for approval.' };
}

/**
 * Emails everyone listed in Settings > Notification Emails whenever a
 * correction request comes in, so Admin/Doctor are alerted immediately.
 */
function notifyCorrectionRequest_(sheetName, fieldName, oldValue, newValue, reason, user) {
  try {
    const settings = getSettings();
    const emails = String(settings['Notification Emails'] || '')
      .split(',').map(e => e.trim()).filter(Boolean);
    if (emails.length === 0) return;
    const hospital = settings['Hospital Name'] || 'Hospital';
    const subject = `[${hospital}] Correction Request — ${sheetName}`;
    const body =
      `A correction request has been submitted and needs review.\n\n` +
      `Hospital: ${hospital}\n` +
      `Sheet: ${sheetName}\n` +
      `Field: ${fieldName}\n` +
      `Old Value: ${oldValue}\n` +
      `New Value: ${newValue}\n` +
      `Reason: ${reason}\n` +
      `Requested By: ${user.name} (${user.role})\n` +
      `Time: ${nowStr_()}\n\n` +
      `Please open the EyeForce dashboard → Corrections to approve or reject.`;
    emails.forEach(email => {
      try { MailApp.sendEmail(email, subject, body); } catch (e) { /* ignore individual failures */ }
    });
  } catch (e) { /* notification failures should never block the request */ }
}

/**
 * Returns the count of pending correction requests — used to show a
 * notification badge on the Corrections nav item for Admin/Doctor.
 */
function getPendingCorrectionsCount(user) {
  if (user.role !== 'Admin' && user.role !== 'Doctor') return 0;
  const { rows } = getSheetData_('CorrectionRequests');
  return rows.filter(r => r['Status'] === 'Pending').length;
}

function getPendingCorrections(user) {
  if (user.role !== 'Admin' && user.role !== 'Doctor') return [];
  const { rows } = getSheetData_('CorrectionRequests');
  return rows.filter(r => r['Status'] === 'Pending');
}

function reviewCorrection(requestRow, approve, user) {
  if (user.role !== 'Admin' && user.role !== 'Doctor') {
    return { success: false, message: 'Not authorized.' };
  }
  const crSheet = SS.getSheetByName('CorrectionRequests');
  const headers = crSheet.getRange(1, 1, 1, crSheet.getLastColumn()).getValues()[0];
  const reqValues = crSheet.getRange(requestRow, 1, 1, headers.length).getValues()[0];
  const req = {};
  headers.forEach((h, i) => req[h] = reqValues[i]);

  if (approve) {
    const targetSheet = SS.getSheetByName(req['Department Sheet']);
    const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const colIdx = targetHeaders.indexOf(req['Field Name']) + 1;
    if (colIdx > 0) {
      targetSheet.getRange(Number(req['Row Number']), colIdx).setValue(req['New Value']);
    }
    logAudit_(user.name, 'CORRECTION-APPROVED', req['Department Sheet'], req['Row Number'],
      `${req['Field Name']}: '${req['Old Value']}' -> '${req['New Value']}' (requested by ${req['Requested By']})`);
  }

  crSheet.getRange(requestRow, headers.indexOf('Status') + 1).setValue(approve ? 'Approved' : 'Rejected');
  crSheet.getRange(requestRow, headers.indexOf('Reviewed By') + 1).setValue(user.name);
  crSheet.getRange(requestRow, headers.indexOf('Reviewed On') + 1).setValue(nowStr_());
  return { success: true, message: approve ? 'Correction applied.' : 'Correction rejected.' };
}

// ---------- AUDIT LOG ----------

function logAudit_(userName, action, sheetName, rowNum, details) {
  const sheet = SS.getSheetByName('AuditLog');
  sheet.appendRow([nowStr_(), userName, action, sheetName, rowNum, details]);
}

// ---------- COUNSELLING / PIPELINE ----------

/**
 * Returns this counsellor's converted patients that have a Surgery Date Given,
 * for the "My Pipeline" follow-up view.
 */
function getMyPipeline(user) {
  const { rows } = getSheetData_('Counselling');
  return rows.filter(r =>
    r['Counselor Name'] === user.name &&
    String(r['Converted']).toLowerCase() === 'yes' &&
    r['Surgery Date Given']
  ).map(r => ({
    _row: r._row,
    patient: r['Patient Reference'],
    procedure: r['Procedure'],
    category: r['Category'],
    surgeryDate: r['Surgery Date Given'],
    status: r['Pipeline Status'] || 'Scheduled'
  }));
}

function updatePipelineStatus(rowNum, newStatus, user) {
  const sheet = SS.getSheetByName('Counselling');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const ownerCol = headers.indexOf('Counselor Name') + 1;
  const owner = sheet.getRange(rowNum, ownerCol).getValue();
  if (owner !== user.name && user.role !== 'Admin' && user.role !== 'Doctor') {
    return { success: false, message: 'Not authorized.' };
  }
  const statusCol = headers.indexOf('Pipeline Status') + 1;
  sheet.getRange(rowNum, statusCol).setValue(newStatus);
  logAudit_(user.name, 'PIPELINE-UPDATE', 'Counselling', rowNum, `Status -> ${newStatus}`);
  return { success: true, message: 'Pipeline status updated.' };
}

// ---------- SEARCH ----------

/**
 * Searches Patient Reference across Counselling, and Submitted By / key text
 * fields across other sheets, for a global search tab.
 */
function globalSearch(query, user) {
  query = String(query).toLowerCase().trim();
  if (!query) return [];
  const results = [];
  const sheetsToSearch = ['Counselling', 'Reception', 'Marketing', 'OT'];

  sheetsToSearch.forEach(name => {
    const { rows } = getSheetData_(name);
    rows.forEach(r => {
      const haystack = Object.values(r).join(' ').toLowerCase();
      if (haystack.includes(query)) {
        // Revenue fields hidden from non-Admin/Doctor roles
        const safeRow = Object.assign({}, r);
        if (user.role !== 'Admin' && user.role !== 'Doctor') {
          ['Amount Committed (₹)', 'Amount Collected Today (₹)', 'Quote Given (₹)'].forEach(f => delete safeRow[f]);
        }
        results.push({ sheet: name, row: safeRow });
      }
    });
  });
  return results.slice(0, 50); // cap results
}

// ---------- DASHBOARD AGGREGATION ----------

/**
 * range: 'today' | 'week' | 'month' | 'lastMonth' | 'yearToDate'
 * Returns a consolidated object the dashboard front-end can render directly.
 */
function getDashboardData(range, user, customStart, customEnd) {
  try { return getDashboardData_(range, user, customStart, customEnd); }
  catch(e) {
    Logger.log('getDashboardData ERROR: ' + e.message);
    return { success: false, message: '⚠ Dashboard error: ' + e.message + '. Run setupHospitalSheet() if sheets are missing.' };
  }
}
function getDashboardData_(range, user, customStart, customEnd) {
  if (user.role !== 'Admin' && user.role !== 'Doctor') {
    return { success: false, message: 'The dashboard is available to Admin and Doctor only.' };
  }
  const isFinanceVisible = true;
  const isAdmin = (user.role === 'Admin');
  const { startDate, endDate, compareLabel } = getRangeDates_(range, customStart, customEnd);
  const settings = getSettings();
  const counsellingMode = settings['Counselling Mode (Per-Patient / Daily Summary)'] || 'Per-Patient';
  const isDailySummaryMode = (counsellingMode === 'Daily Summary');

  const reception = filterByDate_('Reception', startDate, endDate);
  // In "Daily Summary" mode, the per-patient Counselling tab is not used at
  // all — every counselling metric comes from CounsellingSummary instead,
  // so this is forced empty here to guarantee no double-counting.
  const counselling = isDailySummaryMode ? [] : filterByDate_('Counselling', startDate, endDate);
  const ot = filterByDate_('OT', startDate, endDate);
  const diagnostic = filterByDate_('Diagnostic', startDate, endDate);
  const marketing = filterByDate_('Marketing', startDate, endDate);
  const tpa = filterByDate_('TPA', startDate, endDate);
  const hr = filterByDate_('HR', startDate, endDate);
  const pharmacy = filterByDate_('Pharmacy', startDate, endDate);
  const optical = filterByDate_('Optical', startDate, endDate);

  // ----- Reception summary -----
  const sum = (rows, field) => rows.reduce((a, r) => a + (Number(r[field]) || 0), 0);
  const receptionSummary = {
    totalOPD:    sum(reception, 'Total OPD'),
    newPatients: sum(reception, 'New Patients'),
    oldPatients: sum(reception, 'Old/Follow-up'),
    foc:         sum(reception, 'FOC Patients'),
    cghs:        sum(reception, 'CGHS'),
    echs:        sum(reception, 'ECHS'),
    delhiGovt:   sum(reception, 'Delhi Govt'),
    cash:        sum(reception, 'Cash'),
    otherTPA:    sum(reception, 'Other TPA/Insurance'),
    noShows:     sum(reception, 'No-Shows Today'),
    // Cash collected at reception counter (three separate buckets + others)
    opdCash:     sum(reception, 'OPD Cash Collected (₹)'),
    otCash:      sum(reception, 'OT Cash Collected (₹)'),
    diagCash:    sum(reception, 'Diagnostic Cash Collected (₹)'),
    othersCash:  sum(reception, 'Others Cash Collected (₹)'),
    sourceBreakdown: {
      socialMedia: sum(reception, 'Source - Social Media'),
      referral:    sum(reception, 'Source - Referral'),
      walkin:      sum(reception, 'Source - Walk-in/Other'),
      newspaper:   sum(reception, 'Source - Newspaper/Hoarding'),
      other:       sum(reception, 'Source - Other')
    }
  };

  // ----- Counselling / conversion funnel -----
  let totalCounselled = counselling.length;
  const converted = counselling.filter(r => String(r['Converted']).toLowerCase() === 'yes');
  let totalConverted = converted.length;
  let conversionRate = totalCounselled ? Math.round((totalConverted / totalCounselled) * 100) : 0;

  const dropoutReasons = {};
  counselling.filter(r => String(r['Converted']).toLowerCase() !== 'yes').forEach(r => {
    const reason = r['Dropout Reason'] || 'Unspecified';
    dropoutReasons[reason] = (dropoutReasons[reason] || 0) + 1;
  });

  // Counsellor leaderboard
  const counsellorMap = {};
  counselling.forEach(r => {
    const name = r['Counselor Name'];
    if (!counsellorMap[name]) counsellorMap[name] = { name, counselled: 0, converted: 0, revenue: 0 };
    counsellorMap[name].counselled++;
    if (String(r['Converted']).toLowerCase() === 'yes') {
      counsellorMap[name].converted++;
      counsellorMap[name].revenue += Number(r['Amount Committed (₹)']) || 0;
    }
  });
  const counsellorLeaderboard = Object.values(counsellorMap).map(c => ({
    name: c.name,
    counselled: c.counselled,
    converted: c.converted,
    conversionRate: c.counselled ? Math.round((c.converted / c.counselled) * 100) : 0,
    revenue: isFinanceVisible ? c.revenue : null
  })).sort((a, b) => b.conversionRate - a.conversionRate);

  // Doctor-wise referral -> conversion -> OT cases
  const doctorMap = {};
  counselling.forEach(r => {
    const doc = r['Referring Doctor'];
    if (!doctorMap[doc]) doctorMap[doc] = { name: doc, referred: 0, converted: 0, otCases: 0 };
    doctorMap[doc].referred++;
    if (String(r['Converted']).toLowerCase() === 'yes') doctorMap[doc].converted++;
  });
  ot.forEach(r => {
    const doc = r['Doctor Name'];
    if (!doctorMap[doc]) doctorMap[doc] = { name: doc, referred: 0, converted: 0, otCases: 0 };
    doctorMap[doc].otCases += Number(r['Number of Cases']) || 0;
  });
  const doctorPerformance = Object.values(doctorMap).map(d => ({
    ...d,
    conversionRate: d.referred ? Math.round((d.converted / d.referred) * 100) : 0
  }));

  // ----- Revenue (Admin/Doctor only) -----
  let revenue = null;
  if (isFinanceVisible) {
    revenue = {
      totalCommitted: sum(converted, 'Amount Committed (₹)'),
      totalCollected: sum(counselling, 'Amount Collected Today (₹)'),
      byCategory: {}
    };
    converted.forEach(r => {
      const cat = r['Category'] || 'Unspecified';
      revenue.byCategory[cat] = (revenue.byCategory[cat] || 0) + (Number(r['Amount Committed (₹)']) || 0);
    });
  }

  // ----- Surgery Pipeline (month-wise) -----
  const pipeline = converted.filter(r => r['Surgery Date Given']);
  // ----- PIPELINE — two clear buckets ----------------------------------------
  // Bucket 1: Patients with surgery date → month-wise + procedure-wise
  // Bucket 2: Patients converted but NO surgery date yet → follow-up list
  const pipelineByMonth = {};
  const pipelinePending = {}; // converted, no date → by procedure
  // Pipeline always reads from the per-patient Counselling sheet regardless
  // of whether the hospital uses Daily Summary or Per-Patient counselling mode.
  // isDailySummaryMode only affects KPI stats (conversion rates, leaderboard)
  // — the forward-looking surgery pipeline must always show converted patients.
  const { rows: allCounselling } = getSheetData_('Counselling');
  allCounselling.filter(r => String(r['Converted']).toLowerCase() === 'yes').forEach(r => {
    const proc = r['Procedure'] || 'Other';
    const rev = isFinanceVisible
      ? (Number(r['Estimated Revenue (₹)']) || Number(r['Price Master Rate (₹)']) || 0) : 0;
    if (r['Surgery Date Given']) {
      let d;
      try { d = new Date(r['Surgery Date Given']); } catch(e) { return; }
      const monthKey = Utilities.formatDate(d, TZ, 'MMM yyyy');
      const sortKey = Utilities.formatDate(d, TZ, 'yyyy-MM');
      if (!pipelineByMonth[monthKey]) pipelineByMonth[monthKey] = {
        month: monthKey, sortKey, total: 0, byProcedure: {}, revenue: 0
      };
      pipelineByMonth[monthKey].total++;
      pipelineByMonth[monthKey].byProcedure[proc] = (pipelineByMonth[monthKey].byProcedure[proc] || 0) + 1;
      if (isFinanceVisible) pipelineByMonth[monthKey].revenue += rev;
    } else {
      // No surgery date — pending follow-up
      if (!pipelinePending[proc]) pipelinePending[proc] = { procedure: proc, count: 0, revenue: 0 };
      pipelinePending[proc].count++;
      if (isFinanceVisible) pipelinePending[proc].revenue += rev;
    }
  });
  // Sort months chronologically
  const pipelineByMonthArr = Object.values(pipelineByMonth).sort((a, b) => a.sortKey.localeCompare(b.sortKey));

  // ----- OPD breakup (Fresh / Follow-up / FOC / payer mix) ------------------
  const opdBreakup = {
    fresh:     receptionSummary.newPatients,
    followUp:  receptionSummary.oldPatients,
    foc:       receptionSummary.foc,
    cghs:      receptionSummary.cghs,
    echs:      receptionSummary.echs,
    delhiGovt: receptionSummary.delhiGovt,
    cash:      receptionSummary.cash,
    otherTPA:  receptionSummary.otherTPA
  };

  // ----- Surgery category KPIs (from OT, grouped by Procedure Type) ----------
  const procedureList = settings['Procedure List'] || [];
  const surgeryCategoryCounts = {};
  procedureList.forEach(p => surgeryCategoryCounts[p] = 0);
  ot.forEach(r => {
    const proc = r['Procedure Type'] || 'Others';
    const cases = (Number(r['Total Cases']) || 0) ||
      (Number(r['Cash Cases']) || 0) + (Number(r['CGHS Cases']) || 0) +
      (Number(r['ECHS Cases']) || 0) + (Number(r['Delhi Govt Cases']) || 0) +
      (Number(r['TPA Cases']) || 0);
    surgeryCategoryCounts[proc] = (surgeryCategoryCounts[proc] || 0) + cases;
  });

  // ----- Diagnostic category KPIs (each test-type column in the Diagnostic sheet) -----
  //
  // ─── REVENUE MODEL ──────────────────────────────────────────────────────────
  //
  // ACTUAL COLLECTED (cash in hand):
  //   OPD fees (staff-entered) + Diagnostic (auto from Price Master)
  //   + OT Cash cases (staff-entered cash collected)
  //   + Pharmacy (staff-entered) + Optical (staff-entered)
  //
  // EXPECTED / RECEIVABLE (operated, payment pending months):
  //   OT CGHS (cases × fixed govt rate from Price Master)
  //   + OT ECHS (cases × fixed govt rate from Price Master)
  //   + OT Private TPA (billed amount entered by staff)
  //
  // PIPELINE PROJECTION (informational only — counselling conversions):
  //   Converted patients × Price Master rate → NOT mixed with any revenue.
  //   Shown separately in "Pipeline" section, clearly labelled "Projection."
  //
  // ─────────────────────────────────────────────────────────────────────────────

  let revenueBreakup = null;
  if (isFinanceVisible) {
    // ── ACTUAL COLLECTED ─────────────────────────────────────────────────────
    // Reception counter handles: OPD fees + OT patient surgery cash + Diagnostic cash
    const opdRev     = sum(reception, 'OPD Cash Collected (₹)');
    const otCash     = sum(reception, 'OT Cash Collected (₹)');
    const diagCash   = sum(reception, 'Diagnostic Cash Collected (₹)');
    const othersCash = sum(reception, 'Others Cash Collected (₹)');
    const pharmacyRev = sum(pharmacy, 'Revenue Collected (₹)');
    const opticalRev  = sum(optical,  'Revenue Collected (₹)');
    // TPA settlements: use detailed (TPASettlements sheet) OR consolidated (TPASummary)
    // depending on which has data — supports both modes transparently.
    const tpaSettlementRows     = filterByDate_('TPASettlements', startDate, endDate);
    const tpaSummaryRows        = filterByDate_('TPASummary', startDate, endDate);
    const settlementsDetailed   = sum(tpaSettlementRows, 'Amount Received (₹)');
    const settlementsConsolidated = sum(tpaSummaryRows, 'Settlements Received Today (₹)')
                                  + sum(tpaSummaryRows, 'Co-pay Collected Today (₹)');
    const settlementsReceived   = settlementsDetailed + settlementsConsolidated;

    const totalCollected = opdRev + otCash + diagCash + othersCash + pharmacyRev + opticalRev + settlementsReceived;

    // ── EXPECTED / RECEIVABLE ────────────────────────────────────────────────
    // Detailed mode: ref-by-ref from TPAClaims − TPASettlements
    // Consolidated mode: running totals from TPASummary latest row
    let claimsPendingAmount = 0;
    let copayPending = 0;
    try {
      const { rows: allClaims } = getSheetData_('TPAClaims');
      const { rows: allSetts }  = getSheetData_('TPASettlements');
      const settledByRef = {};
      allSetts.forEach(s => {
        const ref = String(s['Claim Ref No'] || '').trim();
        if (ref) settledByRef[ref] = (settledByRef[ref] || 0) + (Number(s['Amount Received (₹)']) || 0);
      });
      allClaims.forEach(c => {
        const ref = String(c['Claim Ref No'] || '').trim();
        claimsPendingAmount += Math.max(0, (Number(c['Claim Amount (₹)']) || 0) - (settledByRef[ref] || 0));
        copayPending += Number(c['Co-pay Pending (₹)']) || 0;
      });
    } catch(e) {}
    // Consolidated mode: add TPASummary running totals (both modes can coexist)
    if (tpaSummaryRows.length > 0) {
      const last = tpaSummaryRows[tpaSummaryRows.length - 1];
      claimsPendingAmount += Number(last['Total Expected Pending (₹)']) || 0;
      copayPending        += Number(last['Co-pay Pending (₹)'])         || 0;
    }

    // OT auto-computed by Price Master × cases (useful before claim is filed)
    const otExpectedCGHS  = sum(ot, 'CGHS Expected (₹)');
    const otExpectedECHS  = sum(ot, 'ECHS Expected (₹)');
    const otExpectedDelhi = sum(ot, 'Delhi Govt Expected (₹)');

    revenueBreakup = {
      opd: opdRev, otCash, diagnostic: diagCash, others: othersCash,
      pharmacy: pharmacyRev, optical: opticalRev,
      settlementsReceived, totalCollected,
      otCGHS: otExpectedCGHS, otECHS: otExpectedECHS, otDelhi: otExpectedDelhi,
      claimsPendingAmount, copayPending,
      totalExpected: claimsPendingAmount + copayPending
    };
  }

  // Diagnostic test category counts (for KPI cards)
  // Diagnostic category counts — new paid/free structure.
  // Each test has two columns: '[Test] - Paid' and '[Test] - Free'.
  // The diagnosticCategoryCounts object maps test name → {paid, free, total}.
  const diagCategoryList = settings['Diagnostic Category List'] || [
    'Biometry','Topography','OCT','Cataract Work-up','LASIK Work-up',
    'Green Laser','Fundus Photography','Visual Field/Perimetry','B-Scan/A-Scan','Other Tests'
  ];
  const diagnosticCategoryCounts = {};
  diagCategoryList.forEach(cat => {
    const paid = sum(diagnostic, cat + ' - Paid');
    const free = sum(diagnostic, cat + ' - Free');
    diagnosticCategoryCounts[cat] = { paid, free, total: paid + free };
  });

  // ----- Pipeline by category (all-time totals, across every month) -----
  const pipelineByCategory = {};
  procedureList.forEach(p => pipelineByCategory[p] = { procedure: p, total: 0, revenue: 0 });
  allCounselling.filter(r => String(r['Converted']).toLowerCase() === 'yes' && r['Surgery Date Given']).forEach(r => {
    const proc = r['Procedure'] || 'Others';
    if (!pipelineByCategory[proc]) pipelineByCategory[proc] = { procedure: proc, total: 0, revenue: 0 };
    pipelineByCategory[proc].total++;
    if (isFinanceVisible) {
      // Use auto-computed revenue from Price Master (stored server-side, never shown to staff)
      pipelineByCategory[proc].revenue += Number(r['Estimated Revenue (₹)']) || Number(r['Price Master Rate (₹)']) || 0;
    }
  });

  // ----- Counselling Summary mode (no dedicated counsellor) -----
  // Overrides the counselling-derived figures above using CounsellingSummary
  // instead — the two tabs are mutually exclusive, so nothing is double-counted.
  let categoryBreakdown = null, paymentMix = null;
  if (isDailySummaryMode) {
    const summaryRows = filterByDate_('CounsellingSummary', startDate, endDate);
    categoryBreakdown = {};
    procedureList.forEach(c => categoryBreakdown[c] = { advised: 0, counselled: 0, converted: 0, pending: 0 });
    let totalCounselledSum = 0, totalConvertedSum = 0;
    summaryRows.forEach(r => {
      procedureList.forEach(c => {
        categoryBreakdown[c].advised   += Number(r[c + ' - Advised'])   || 0;
        categoryBreakdown[c].counselled += Number(r[c + ' - Counselled']) || 0;
        categoryBreakdown[c].converted += Number(r[c + ' - Converted']) || 0;
        categoryBreakdown[c].pending   += Number(r[c + ' - Pending'])   || 0;
      });
    });
    procedureList.forEach(c => {
      totalCounselledSum += categoryBreakdown[c].counselled;
      totalConvertedSum  += categoryBreakdown[c].converted;
    });

    totalCounselled = totalCounselledSum;
    totalConverted = totalConvertedSum;
    conversionRate = totalCounselled ? Math.round((totalConverted / totalCounselled) * 100) : 0;

    paymentMix = {
      cash: sum(summaryRows, 'Cash Conversions'),
      cghs: sum(summaryRows, 'CGHS Conversions'),
      echs: sum(summaryRows, 'ECHS Conversions'),
      otherTpa: sum(summaryRows, 'Other TPA Conversions')
    };

    const revenueSurgery = sum(summaryRows, 'Revenue Collected Today (₹)');
    if (isFinanceVisible) {
      revenue = { totalCommitted: null, totalCollected: revenueSurgery, byCategory: null };
      if (revenueBreakup) {
        revenueBreakup.surgery = revenueSurgery;
        revenueBreakup.total = revenueBreakup.opd + revenueBreakup.diagnostic + revenueSurgery + revenueBreakup.optical + revenueBreakup.pharmacy;
      }
    }

    // Pipeline-by-category becomes "currently pending" per category (a snapshot,
    // not month-dated — month-wise pipeline requires per-patient surgery dates,
    // so pipelineByMonth stays empty in this mode).
    procedureList.forEach(c => {
      pipelineByCategory[c].total = categoryBreakdown[c].pending;
      pipelineByCategory[c].revenue = null;
    });
  }

  const marketingSummary = {
    callsMade: sum(marketing, 'Calls Made'),
    visitsMade: sum(marketing, 'Visits Made'),
    leadsGenerated: sum(marketing, 'Leads Generated'),
    leadsConverted: sum(marketing, 'Leads Converted to OPD'),
    conversionRate: sum(marketing, 'Leads Generated') > 0
      ? Math.round((sum(marketing, 'Leads Converted to OPD') / sum(marketing, 'Leads Generated')) * 100) : 0
  };
  // Per-PRO breakdown — who's covering which area and how they're converting
  const proMap = {};
  marketing.forEach(r => {
    const pro = r['PRO Name'] || 'Unspecified';
    if (!proMap[pro]) proMap[pro] = { pro, calls: 0, visits: 0, leads: 0, converted: 0, areas: new Set() };
    proMap[pro].calls += Number(r['Calls Made']) || 0;
    proMap[pro].visits += Number(r['Visits Made']) || 0;
    proMap[pro].leads += Number(r['Leads Generated']) || 0;
    proMap[pro].converted += Number(r['Leads Converted to OPD']) || 0;
    if (r['Area Covered']) proMap[pro].areas.add(r['Area Covered']);
  });
  marketingSummary.byPRO = Object.values(proMap).map(p => ({
    pro: p.pro, calls: p.calls, visits: p.visits, leads: p.leads, converted: p.converted,
    conversionRate: p.leads > 0 ? Math.round((p.converted / p.leads) * 100) : 0,
    areas: Array.from(p.areas).join(', ')
  })).sort((a,b) => b.leads - a.leads);
  // Lead category breakdown
  const leadCatMap = {};
  marketing.forEach(r => {
    const cat = r['Lead Category'] || 'Other';
    if (!leadCatMap[cat]) leadCatMap[cat] = { category: cat, leads: 0, converted: 0 };
    leadCatMap[cat].leads += Number(r['Leads Generated']) || 0;
    leadCatMap[cat].converted += Number(r['Leads Converted to OPD']) || 0;
  });
  marketingSummary.byLeadCategory = Object.values(leadCatMap).sort((a,b) => b.leads - a.leads);

  // ----- TPA health — now using TPAClaims + TPASettlements for accuracy -----
  // The TPA daily form no longer has financial columns (old model removed).
  // Actual TPA figures come directly from the ref-linked claim/settlement sheets.
  let tpaReceived = 0, tpaPendingTotal = 0, tpaCopayPending = 0;
  try {
    const { rows: allClaims_ } = getSheetData_('TPAClaims');
    const { rows: allSetts_ }  = getSheetData_('TPASettlements');
    const settledByRef_ = {};
    allSetts_.forEach(s => {
      const ref = String(s['Claim Ref No'] || '').trim();
      if (ref) settledByRef_[ref] = (settledByRef_[ref] || 0) + (Number(s['Amount Received (₹)']) || 0);
    });
    tpaReceived = allSetts_.reduce((a, s) => a + (Number(s['Amount Received (₹)']) || 0), 0);
    allClaims_.forEach(c => {
      const ref = String(c['Claim Ref No'] || '').trim();
      tpaPendingTotal += Math.max(0, (Number(c['Claim Amount (₹)']) || 0) - (settledByRef_[ref] || 0));
      tpaCopayPending += Number(c['Co-pay Pending (₹)']) || 0;
    });
  } catch(e) {}
  const tpaSummary = {
    received:      isFinanceVisible ? tpaReceived : null,
    pendingAmount: isFinanceVisible ? tpaPendingTotal : null,
    expected:      isFinanceVisible ? tpaPendingTotal : null,
    copayPending:  isFinanceVisible ? tpaCopayPending : null,
    total:         isFinanceVisible ? (tpaReceived + tpaPendingTotal) : null,
    queriesRaised:   sum(tpa, 'Queries Raised'),
    queriesResolved: sum(tpa, 'Queries Resolved'),
    billsMatched:    0,   // field removed from new TPA sheet — kept for backward compat
    billsMismatched: 0
  };

  // ----- Reconciliation flags (pilferage/spillage check) -----
  const flags = [];
  const counsellingTotal = counselling.length;
  if (receptionSummary.totalOPD > 0 && counsellingTotal === 0) {
    flags.push('No counselling entries recorded despite OPD footfall — check for missed logging.');
  }
  const otTotalCases = sum(ot, 'Total Cases');
  if (totalConverted > 0 && otTotalCases === 0) {
    flags.push('Patients converted in Counselling but no OT cases logged for this period — verify surgery completion records.');
  }

  // TPA Aging red flags — claims > 30 days, shortfalls, pending co-pays
  try {
    const allClaims      = getSheetData_('TPAClaims').rows;
    const allSettlements = getSheetData_('TPASettlements').rows;
    const settledRefs    = new Set(allSettlements.map(s => String(s['Claim Reference / File No'] || '').trim()).filter(Boolean));
    const today30 = new Date(); today30.setDate(today30.getDate() - 30);
    const agedClaims = allClaims.filter(c => {
      const ref = String(c['Claim Reference / File No'] || '').trim();
      if (settledRefs.has(ref)) return false;
      const d = c['Date Submitted'] ? new Date(c['Date Submitted']) : null;
      return d && d < today30;
    });
    if (agedClaims.length > 0) {
      flags.push(`⚠ TPA: ${agedClaims.length} claim(s) pending settlement for more than 30 days. Open TPA → Aging Report to review and follow up.`);
    }
    const shortfalls = allSettlements.filter(s => (Number(s['Shortfall (₹)']) || 0) > 0);
    if (shortfalls.length > 0) {
      const totalSF = shortfalls.reduce((a, s) => a + (Number(s['Shortfall (₹)']) || 0), 0);
      flags.push(`⚠ TPA: ${shortfalls.length} settlement(s) received with shortfall — total ₹${totalSF.toLocaleString('en-IN')} not fully paid. Check TPA → Aging Report.`);
    }
    const { rows: allClaims2 } = getSheetData_('TPAClaims');
    const pendingCopay = allClaims2.filter(c => (Number(c['Co-pay Pending (₹)']) || 0) > 0);
    if (pendingCopay.length > 0) {
      const totalPending = pendingCopay.reduce((a, c) => a + (Number(c['Co-pay Pending (₹)']) || 0), 0);
      flags.push(`⚠ TPA: ${pendingCopay.length} upgraded patient(s) with co-pay balance pending — ₹${totalPending.toLocaleString('en-IN')} outstanding.`);
    }
  } catch (e) { /* TPAClaims/TPASettlements sheets may not exist yet on first deploy — ignore */ }

  // ----- Activity heatmap: entries submitted per department, by day of week -----
  const heatmapDepts = ['Reception','Counselling','OT','Diagnostic','Pharmacy','Optical','Marketing','TPA','TPAClaims','TPASettlements','Operations','Store','Purchase','CallCentre','HRSummary','TPASummary'];
  const activityHeatmap = heatmapDepts.map(name => {
    const rows = filterByDate_(name, startDate, endDate);
    const counts = [0,0,0,0,0,0,0];
    rows.forEach(r => {
      if (!r['Date']) return;
      const dow = new Date(r['Date']).getDay();
      counts[dow]++;
    });
    return { dept: name, counts };
  });

  // ----- Operations / Google Reviews -----
  const operations = filterByDate_('Operations', startDate, endDate);
  const cleanlinessOk = operations.filter(r => String(r['Cleanliness Check']).toLowerCase() === 'yes').length;
  const biomedicalDue = operations.filter(r => String(r['Biomedical Servicing Due']).toLowerCase() === 'yes').length;
  const infraIssueNotes = operations.map(r => r['Infrastructure Issue Note']).filter(Boolean);
  const equipBreakdownNotes = operations.map(r => r['Equipment Breakdown Note']).filter(Boolean);
  const stockCheckDone = operations.filter(r => String(r['Stock Check Done']).toLowerCase() === 'yes').length;
  const operationsSummary = {
    daysReported: operations.length,
    avgWaitTime: operations.length ? Math.round(sum(operations, 'Avg Wait Time (mins)') / operations.length) : 0,
    complaintsReceived: sum(operations, 'Complaints Received'),
    complaintsResolved: sum(operations, 'Complaints Resolved'),
    equipmentDowntimeHrs: sum(operations, 'Equipment Downtime (hrs)'),
    equipmentBreakdownNotes: equipBreakdownNotes,
    cleanlinessOkDays: cleanlinessOk,
    biomedicalDueCount: biomedicalDue,
    infrastructureIssues: infraIssueNotes,
    stockCheckDoneDays: stockCheckDone,
    googleReviewsAsked: sum(operations, 'Google Reviews Asked'),
    googleReviewsReceived: sum(operations, 'Google Reviews Received')
  };
  // One-line gist for the Doctor — everything Operations at a glance
  const reviewRate = operationsSummary.googleReviewsAsked > 0
    ? Math.round((operationsSummary.googleReviewsReceived / operationsSummary.googleReviewsAsked) * 100) : 0;
  operationsSummary.gist =
    `Avg wait ${operationsSummary.avgWaitTime} min · ` +
    `Complaints ${operationsSummary.complaintsReceived} received / ${operationsSummary.complaintsResolved} resolved · ` +
    `Cleanliness OK ${cleanlinessOk}/${operations.length} days · ` +
    `Equipment downtime ${operationsSummary.equipmentDowntimeHrs} hrs` +
    (biomedicalDue > 0 ? ` · ⚠ Biomedical servicing due (${biomedicalDue}d)` : '') +
    (infraIssueNotes.length > 0 ? ` · ⚠ ${infraIssueNotes.length} infrastructure issue(s)` : '') +
    ` · Google Reviews ${operationsSummary.googleReviewsReceived}/${operationsSummary.googleReviewsAsked} (${reviewRate}%)`;
  if (biomedicalDue > 0) flags.push(`Operations: biomedical servicing flagged as due on ${biomedicalDue} day(s) this period.`);
  if (infraIssueNotes.length > 0) flags.push(`Operations: ${infraIssueNotes.length} infrastructure issue(s) noted this period — see Operations entries.`);

  // ----- HR consolidated report (attendance, punctuality, grooming, discipline, training, ratings) -----
  const groomingNonCompliant = hr.filter(r => String(r['Grooming Compliance']).toLowerCase() === 'non-compliant');
  const disciplineNotes = hr.filter(r => r['Discipline Note']);
  const grievanceNotes = hr.filter(r => r['Grievance Note']);
  const trainingDone = hr.filter(r => String(r['Training Completed']).toLowerCase() === 'yes').length;
  const trainingPending = hr.filter(r => String(r['Training Completed']).toLowerCase() === 'no').length;
  const ratingsRows = hr.filter(r => r['Behaviour Rating (1-5)'] || r['Performance Score (1-5)']);
  const hrSummary = {
    recordsThisPeriod: hr.length, // = "current report": how many staff-day records were submitted
    present: hr.filter(r => String(r['Attendance']).toLowerCase() === 'present').length,
    absent: hr.filter(r => String(r['Attendance']).toLowerCase() === 'absent').length,
    halfDay: hr.filter(r => String(r['Attendance']).toLowerCase() === 'half-day').length,
    lateArrivalsTotal: sum(hr, 'Late Arrivals'),
    staffWithLateArrival: hr.filter(r => (Number(r['Late Arrivals']) || 0) > 0).length,
    leavesTaken: sum(hr, 'Leaves Taken'),
    groomingNonCompliantCount: groomingNonCompliant.length,
    disciplineIssueCount: disciplineNotes.length,
    grievanceCount: grievanceNotes.length,
    trainingDone,
    trainingPending,
    avgBehaviourRating: ratingsRows.length ? Math.round((sum(ratingsRows, 'Behaviour Rating (1-5)') / ratingsRows.length) * 10) / 10 : null,
    avgPerformanceScore: ratingsRows.length ? Math.round((sum(ratingsRows, 'Performance Score (1-5)') / ratingsRows.length) * 10) / 10 : null
  };
  // By department breakdown — consolidated alongside the daily figures
  const hrByDeptMap = {};
  hr.forEach(r => {
    const d = r['Department'] || 'Unspecified';
    if (!hrByDeptMap[d]) hrByDeptMap[d] = { department: d, present: 0, absent: 0, halfDay: 0, late: 0, groomingIssues: 0, records: 0 };
    hrByDeptMap[d].records++;
    const att = String(r['Attendance']).toLowerCase();
    if (att === 'present') hrByDeptMap[d].present++;
    else if (att === 'absent') hrByDeptMap[d].absent++;
    else if (att === 'half-day') hrByDeptMap[d].halfDay++;
    if ((Number(r['Late Arrivals']) || 0) > 0) hrByDeptMap[d].late++;
    if (String(r['Grooming Compliance']).toLowerCase() === 'non-compliant') hrByDeptMap[d].groomingIssues++;
  });
  hrSummary.byDepartment = Object.values(hrByDeptMap).sort((a,b) => a.department.localeCompare(b.department));
  // "Needs attention" — staff with discipline notes, grievances, or low ratings this period
  const attentionMap = {};
  hr.forEach(r => {
    const name = r['Staff Name'];
    if (!name) return;
    const issues = [];
    if (r['Discipline Note']) issues.push('Discipline: ' + r['Discipline Note']);
    if (r['Grievance Note']) issues.push('Grievance: ' + r['Grievance Note']);
    if (String(r['Grooming Compliance']).toLowerCase() === 'non-compliant') issues.push('Grooming non-compliant');
    if ((Number(r['Behaviour Rating (1-5)']) || 0) > 0 && Number(r['Behaviour Rating (1-5)']) <= 2) issues.push('Low behaviour rating (' + r['Behaviour Rating (1-5)'] + '/5)');
    if ((Number(r['Performance Score (1-5)']) || 0) > 0 && Number(r['Performance Score (1-5)']) <= 2) issues.push('Low performance score (' + r['Performance Score (1-5)'] + '/5)');
    if (issues.length) {
      if (!attentionMap[name]) attentionMap[name] = { name, department: r['Department'] || '', issues: [] };
      attentionMap[name].issues.push(...issues);
    }
  });
  hrSummary.needsAttention = Object.values(attentionMap);
  if (hrSummary.groomingNonCompliantCount > 0) flags.push(`HR: grooming/dressing non-compliance logged ${hrSummary.groomingNonCompliantCount} time(s) this period.`);
  if (hrSummary.absent > 0) flags.push(`HR: ${hrSummary.absent} absence(s) recorded this period.`);

  // ----- Pharmacy & Optical KPI summaries -----
  const pharmacySummary = {
    totalPrescriptions: sum(pharmacy, 'Total Prescriptions'),
    cashPrescriptions: sum(pharmacy, 'Prescriptions - Cash'),
    tpaPrescriptions: sum(pharmacy, 'Prescriptions - TPA/Insurance'),
    stockOutAlerts: sum(pharmacy, 'Stock-Out Alerts')
  };
  const opticalSummary = {
    framesSold: sum(optical, 'Frames Sold'),
    lensesSold: sum(optical, 'Lenses Sold'),
    contactLensSold: sum(optical, 'Contact Lens Sold'),
    walkInCustomers: sum(optical, 'Walk-in Customers')
  };
  opticalSummary.conversionRate = opticalSummary.walkInCustomers > 0
    ? Math.round(((opticalSummary.framesSold + opticalSummary.lensesSold + opticalSummary.contactLensSold) / opticalSummary.walkInCustomers) * 100) : 0;
  if (pharmacySummary.stockOutAlerts > 0) flags.push(`Pharmacy: stock-out alerts raised ${pharmacySummary.stockOutAlerts} time(s) this period — check stock levels.`);

  // ----- Call Centre KPI summary -----
  const callCentreRows = filterByDate_('CallCentre', startDate, endDate);
  const callCentreCategories = [
    'Calls - General Eye Check-up','Calls - Cataract','Calls - LASIK/Refractive',
    'Calls - Glaucoma','Calls - Retina','Calls - Cornea',
    'Calls - Squint/Paediatric','Calls - Other Procedure',
    'Calls - Follow-up / Existing Patient','Calls - Appointment Reminder','Calls - Complaint / Query'
  ];
  const callCentreSummary = {
    inboundTotal:    sum(callCentreRows, 'Inbound Calls Total'),
    inboundAnswered: sum(callCentreRows, 'Inbound Calls Answered'),
    inboundMissed:   sum(callCentreRows, 'Inbound Calls Missed'),
    outboundTotal:   sum(callCentreRows, 'Outbound Calls Total'),
    appointmentsBooked: sum(callCentreRows, 'Appointments Booked (from calls)'),
    surgeryBookings:    sum(callCentreRows, 'Surgery Bookings Confirmed (from calls)'),
    leadsPassedToCounsellor: sum(callCentreRows, 'Leads Passed to Counsellor'),
    byCategory: callCentreCategories.map(c => ({
      label: c.replace('Calls - ', ''),
      count: sum(callCentreRows, c)
    })).filter(c => c.count > 0)
  };
  callCentreSummary.missedCallRate = callCentreSummary.inboundTotal > 0
    ? Math.round((callCentreSummary.inboundMissed / callCentreSummary.inboundTotal) * 100) : 0;
  callCentreSummary.callToAppointmentRate = callCentreSummary.inboundAnswered > 0
    ? Math.round((callCentreSummary.appointmentsBooked / callCentreSummary.inboundAnswered) * 100) : 0;
  if (callCentreSummary.missedCallRate > 20) {
    flags.push(`Call Centre: missed call rate is ${callCentreSummary.missedCallRate}% this period — ${callCentreSummary.inboundMissed} of ${callCentreSummary.inboundTotal} inbound calls unanswered.`);
  }

  // ----- Inventory / Store (current stock = Opening Stock + all-time Received - all-time Issued) -----
  const { rows: invMaster } = getSheetData_('InventoryMaster');
  const { rows: allStore } = getSheetData_('Store');
  const storeInRange = filterByDate_('Store', startDate, endDate);
  const stockMap = {};
  invMaster.forEach(item => {
    stockMap[item['Item Name']] = {
      item: item['Item Name'],
      category: item['Category'],
      unit: item['Unit'],
      reorderLevel: Number(item['Reorder Level']) || 0,
      current: Number(item['Opening Stock']) || 0
    };
  });
  allStore.forEach(r => {
    const name = r['Item Name'];
    if (!stockMap[name]) stockMap[name] = { item: name, category: r['Category'] || 'Others', unit: '', reorderLevel: 0, current: 0 };
    const qty = Number(r['Quantity']) || 0;
    const type = String(r['Transaction Type'] || '');
    if (type.indexOf('Received') === 0) stockMap[name].current += qty;
    else if (type.indexOf('Issued') === 0) stockMap[name].current -= qty;
  });
  const inventoryItems = Object.values(stockMap).map(it => ({
    ...it,
    lowStock: it.reorderLevel > 0 && it.current <= it.reorderLevel
  }));
  const inventoryByCategory = {};
  inventoryItems.forEach(it => {
    if (!inventoryByCategory[it.category]) inventoryByCategory[it.category] = { category: it.category, totalStock: 0, lowStockCount: 0 };
    inventoryByCategory[it.category].totalStock += it.current;
    if (it.lowStock) inventoryByCategory[it.category].lowStockCount++;
  });
  // Items issued during this period, grouped by item and by "issued to"
  const issuedThisPeriod = {};
  const issuedToBreakdown = {};
  storeInRange.filter(r => String(r['Transaction Type'] || '').indexOf('Issued') === 0).forEach(r => {
    const name = r['Item Name'];
    issuedThisPeriod[name] = (issuedThisPeriod[name] || 0) + (Number(r['Quantity']) || 0);
    const to = r['Issued To / Received From'] || 'Unspecified';
    issuedToBreakdown[to] = (issuedToBreakdown[to] || 0) + (Number(r['Quantity']) || 0);
  });
  const inventorySummary = {
    items: inventoryItems,
    byCategory: Object.values(inventoryByCategory),
    lowStockItems: inventoryItems.filter(it => it.lowStock),
    issuedThisPeriod: Object.entries(issuedThisPeriod).map(([item, qty]) => ({ item, qty })).sort((a,b)=>b.qty-a.qty),
    issuedToBreakdown
  };
  // Low-stock flags are shown in the compact Inventory KPI card on the dashboard.
  // Not repeated in the notification banners to avoid cluttering the alerts.
  // (If needed, Admin can see full stock detail in the InventoryMaster sheet.)

  // ----- Purchase / Vendor / Expenses (Admin/Doctor only for amounts) -----
  const purchaseInRange = filterByDate_('Purchase', startDate, endDate);
  const { rows: allPurchase } = getSheetData_('Purchase');
  let purchaseSummary = null;
  if (isFinanceVisible) {
    const totalPurchase = sum(purchaseInRange, 'Amount (₹)');
    const totalPaid = sum(purchaseInRange, 'Amount Paid (₹)');
    const totalPending = sum(purchaseInRange, 'Amount Pending (₹)');

    // Top vendors (lifetime totals, so dues are meaningful regardless of selected range)
    const vendorMap = {};
    allPurchase.forEach(r => {
      const v = r['Vendor Name'] || 'Unspecified';
      if (!vendorMap[v]) vendorMap[v] = { vendor: v, totalAmount: 0, totalPending: 0, items: {} };
      vendorMap[v].totalAmount += Number(r['Amount (₹)']) || 0;
      vendorMap[v].totalPending += Number(r['Amount Pending (₹)']) || 0;
      const item = r['Item/Product'] || 'Other';
      vendorMap[v].items[item] = (vendorMap[v].items[item] || 0) + (Number(r['Amount (₹)']) || 0);
    });
    const topVendors = Object.values(vendorMap).map(v => ({
      vendor: v.vendor,
      totalAmount: v.totalAmount,
      totalPending: v.totalPending,
      topProduct: Object.entries(v.items).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—'
    })).sort((a,b) => b.totalAmount - a.totalAmount).slice(0, 10);

    const totalVendorPendingAllTime = sum(allPurchase, 'Amount Pending (₹)');

    // Expense breakdown by category (this period)
    const expenseByCategory = {};
    purchaseInRange.forEach(r => {
      const cat = r['Category'] || 'Others';
      expenseByCategory[cat] = (expenseByCategory[cat] || 0) + (Number(r['Amount (₹)']) || 0);
    });

    // Month-wise purchase trend (all-time, grouped by month)
    const purchaseByMonth = {};
    allPurchase.forEach(r => {
      if (!r['Date']) return;
      const monthKey = Utilities.formatDate(new Date(r['Date']), TZ, 'MMM yyyy');
      if (!purchaseByMonth[monthKey]) purchaseByMonth[monthKey] = { month: monthKey, total: 0, paid: 0, pending: 0, _sortDate: new Date(r['Date']) };
      purchaseByMonth[monthKey].total += Number(r['Amount (₹)']) || 0;
      purchaseByMonth[monthKey].paid += Number(r['Amount Paid (₹)']) || 0;
      purchaseByMonth[monthKey].pending += Number(r['Amount Pending (₹)']) || 0;
    });

    purchaseSummary = {
      totalPurchase, totalPaid, totalPending,
      totalVendorPendingAllTime,
      topVendors,
      expenseByCategory,
      purchaseByMonth: Object.values(purchaseByMonth).sort((a,b) => a._sortDate - b._sortDate).map(m => ({ month: m.month, total: m.total, paid: m.paid, pending: m.pending }))
    };
  }

  return {
    success: true,
    isAdmin,
    range, compareLabel,
    counsellingMode,
    categoryBreakdown,
    paymentMix,
    reception: receptionSummary,
    opdBreakup,
    counselling: { totalCounselled, totalConverted, conversionRate, dropoutReasons, counsellorLeaderboard },
    doctorPerformance,
    revenue,
    revenueBreakup,
    surgeryCategoryCounts,
    diagnosticCategoryCounts,
    pipelineByMonth: pipelineByMonthArr,
    pipelinePending: Object.values(pipelinePending).sort((a,b) => b.count - a.count),
    pipelineByCategory: Object.values(pipelineByCategory),
    marketing: marketingSummary,
    tpa: tpaSummary,
    callCentre: callCentreSummary,
    operations: operationsSummary,
    hr: hrSummary,
    pharmacy: pharmacySummary,
    optical: opticalSummary,
    inventory: inventorySummary,
    purchase: purchaseSummary,
    activityHeatmap,
    flags
  };
}

function getRangeDates_(range, customStart, customEnd) {
  const now = new Date();
  let start, end, compareLabel;
  end = new Date(now);
  switch (range) {
    case 'today':
      start = new Date(now); compareLabel = 'Today';
      break;
    case 'week':
      start = new Date(now); start.setDate(start.getDate() - 7); compareLabel = 'Last 7 Days';
      break;
    case 'month':
      start = new Date(now.getFullYear(), now.getMonth(), 1); compareLabel = 'This Month';
      break;
    case 'lastMonth':
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end = new Date(now.getFullYear(), now.getMonth(), 0);
      compareLabel = 'Last Month';
      break;
    case 'yearToDate':
      start = new Date(now.getFullYear(), 0, 1); compareLabel = 'Year to Date';
      break;
    case 'custom':
      if (customStart && customEnd) {
        start = new Date(customStart);
        end = new Date(customEnd);
        compareLabel = Utilities.formatDate(start, TZ, 'dd MMM') + ' – ' + Utilities.formatDate(end, TZ, 'dd MMM yyyy');
      } else {
        start = new Date(now); compareLabel = 'Today';
      }
      break;
    default:
      start = new Date(now); compareLabel = 'Today';
  }
  return { startDate: start, endDate: end, compareLabel };
}

function filterByDate_(sheetName, startDate, endDate) {
  const { rows } = getSheetData_(sheetName);
  const startStr = Utilities.formatDate(startDate, TZ, 'yyyy-MM-dd');
  const endStr = Utilities.formatDate(endDate, TZ, 'yyyy-MM-dd');
  return rows.filter(r => {
    const d = dateToStr_(r['Date']);
    return d && d >= startStr && d <= endStr;
  });
}

// ============================================================
//  PERIOD-OVER-PERIOD COMPARISONS
//  (This Month vs Last Month, vs Same Month Last Year,
//   This Quarter vs Last Quarter, vs Same Quarter Last Year,
//   This Year vs Last Year)
// ============================================================

function getComparisonData(user) {
  const isFinanceVisible = (user.role === 'Admin' || user.role === 'Doctor');
  const now = new Date();

  const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const thisMonthEnd   = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const lastMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonthEnd   = new Date(now.getFullYear(), now.getMonth(), 0);
  const sameMonthLYStart = new Date(now.getFullYear() - 1, now.getMonth(), 1);
  const sameMonthLYEnd   = new Date(now.getFullYear() - 1, now.getMonth() + 1, 0);

  const q = Math.floor(now.getMonth() / 3);
  const thisQStart = new Date(now.getFullYear(), q * 3, 1);
  const thisQEnd   = new Date(now.getFullYear(), q * 3 + 3, 0);
  let lastQ = q - 1, lastQYear = now.getFullYear();
  if (lastQ < 0) { lastQ = 3; lastQYear -= 1; }
  const lastQStart = new Date(lastQYear, lastQ * 3, 1);
  const lastQEnd   = new Date(lastQYear, lastQ * 3 + 3, 0);
  const sameQLYStart = new Date(now.getFullYear() - 1, q * 3, 1);
  const sameQLYEnd   = new Date(now.getFullYear() - 1, q * 3 + 3, 0);

  const thisYearStart = new Date(now.getFullYear(), 0, 1);
  const thisYearEnd   = new Date(now.getFullYear(), 11, 31);
  const lastYearStart = new Date(now.getFullYear() - 1, 0, 1);
  const lastYearEnd   = new Date(now.getFullYear() - 1, 11, 31);

  return {
    monthVsLastMonth: {
      current:  getMetricsForPeriod_(thisMonthStart, thisMonthEnd, isFinanceVisible, 'This Month'),
      previous: getMetricsForPeriod_(lastMonthStart, lastMonthEnd, isFinanceVisible, 'Last Month')
    },
    monthVsLastYear: {
      current:  getMetricsForPeriod_(thisMonthStart, thisMonthEnd, isFinanceVisible, 'This Month'),
      previous: getMetricsForPeriod_(sameMonthLYStart, sameMonthLYEnd, isFinanceVisible, 'Same Month Last Year')
    },
    quarterVsLastQuarter: {
      current:  getMetricsForPeriod_(thisQStart, thisQEnd, isFinanceVisible, 'This Quarter'),
      previous: getMetricsForPeriod_(lastQStart, lastQEnd, isFinanceVisible, 'Last Quarter')
    },
    quarterVsLastYear: {
      current:  getMetricsForPeriod_(thisQStart, thisQEnd, isFinanceVisible, 'This Quarter'),
      previous: getMetricsForPeriod_(sameQLYStart, sameQLYEnd, isFinanceVisible, 'Same Quarter Last Year')
    },
    yearVsLastYear: {
      current:  getMetricsForPeriod_(thisYearStart, thisYearEnd, isFinanceVisible, 'This Year'),
      previous: getMetricsForPeriod_(lastYearStart, lastYearEnd, isFinanceVisible, 'Last Year')
    }
  };
}

function getMetricsForPeriod_(startDate, endDate, isFinanceVisible, label) {
  const settings = getSettings();
  const isDailySummaryMode = (settings['Counselling Mode (Per-Patient / Daily Summary)'] || 'Per-Patient') === 'Daily Summary';
  const procedureList = settings['Procedure List'] || [];

  const reception   = filterByDate_('Reception', startDate, endDate);
  const counselling = isDailySummaryMode ? [] : filterByDate_('Counselling', startDate, endDate);
  const ot          = filterByDate_('OT', startDate, endDate);
  const optical     = filterByDate_('Optical', startDate, endDate);
  const pharmacy    = filterByDate_('Pharmacy', startDate, endDate);
  const diagnostic  = filterByDate_('Diagnostic', startDate, endDate);

  const sum = (rows, field) => rows.reduce((a, r) => a + (Number(r[field]) || 0), 0);
  const totalOPD  = sum(reception, 'Total OPD');
  const converted = counselling.filter(r => String(r['Converted']).toLowerCase() === 'yes');
  const surgeryCases = sum(ot, 'Number of Cases');

  let totalCounselled = counselling.length;
  let totalConverted = converted.length;
  let surgeryRevenue = sum(counselling, 'Amount Collected Today (₹)');

  if (isDailySummaryMode) {
    const summaryRows = filterByDate_('CounsellingSummary', startDate, endDate);
    totalCounselled = 0; totalConverted = 0;
    summaryRows.forEach(r => {
      procedureList.forEach(c => {
        totalCounselled += Number(r[c + ' - Counselled']) || 0;
        totalConverted  += Number(r[c + ' - Converted'])  || 0;
      });
    });
    surgeryRevenue = sum(summaryRows, 'Revenue Collected Today (₹)');
  }

  let totalRevenue = null;
  if (isFinanceVisible) {
    totalRevenue = sum(reception, 'OPD Revenue Collected (₹)') + sum(diagnostic, 'Revenue Collected (₹)') +
      surgeryRevenue + sum(optical, 'Revenue Collected (₹)') + sum(pharmacy, 'Revenue Collected (₹)');
  }

  return {
    label,
    totalOPD,
    totalCounselled,
    totalConverted,
    conversionRate: totalCounselled ? Math.round((totalConverted / totalCounselled) * 100) : 0,
    surgeryCases,
    totalRevenue
  };
}

// ============================================================
//  REPORT EXPORTS — PDF (Doctor & Admin) and EXCEL (Admin only)
// ============================================================

/**
 * Builds a clean, printable PDF report for the given range and returns it
 * as a base64 string the front-end can turn into a download.
 */
function generateReportPDF(range, user) {
  const settings = getSettings();
  const d = getDashboardData(range, user);
  const comp = getComparisonData(user);
  const hospital = settings['Hospital Name'] || 'Hospital';
  const isFinanceVisible = (user.role === 'Admin' || user.role === 'Doctor');

  const doc = DocumentApp.create(`__temp_report_${Date.now()}`);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

  body.appendParagraph(hospital).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`Performance Report — ${d.compareLabel}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Generated: ${nowStr_()}  |  Prepared for: ${user.name} (${user.role})`)
    .setFontSize(9).editAsText().setForegroundColor('#6E7480');

  // ---- AI Executive Summary (Admin only, if AI API Key configured) ----
  if (user.role === 'Admin' && settings['AI API Key']) {
    const ai = getAIAnalysis(range, user);
    if (ai.success) {
      body.appendParagraph('Executive Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      ai.text.split('\n\n').filter(p => p.trim()).forEach(p => {
        body.appendParagraph(p.trim());
      });
    }
  }

  // ---- OPD Summary ----
  body.appendParagraph('OPD Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, [
    ['Total OPD', d.reception.totalOPD],
    ['Fresh / New Patients', d.opdBreakup.fresh],
    ['Old / Follow-up', d.opdBreakup.followUp],
    ['FOC Patients', d.opdBreakup.foc],
    ['CGHS', d.opdBreakup.cghs],
    ['ECHS', d.opdBreakup.echs],
    ['Cash', d.opdBreakup.cash],
    ['No-Shows', d.reception.noShows]
  ]);

  // ---- Conversion Funnel ----
  body.appendParagraph('Counselling & Conversion').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, [
    ['Total Counselled', d.counselling.totalCounselled],
    ['Total Converted', d.counselling.totalConverted],
    ['Conversion Rate', d.counselling.conversionRate + '%']
  ]);

  // ---- Surgery / Procedure Categories ----
  body.appendParagraph('Surgeries by Category (OT)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, Object.entries(d.surgeryCategoryCounts).map(([k,v]) => [k, v]));

  // ---- Diagnostic Categories ----
  body.appendParagraph('Diagnostic Volumes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, Object.entries(d.diagnosticCategoryCounts).map(([k,v]) => [k, v]));

  // ---- Revenue Breakup (Admin/Doctor only) ----
  if (d.revenueBreakup) {
    body.appendParagraph('Revenue Collection Breakup (₹)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    appendKVTable_(body, [
      ['OPD', d.revenueBreakup.opd],
      ['Diagnostic', d.revenueBreakup.diagnostic],
      ['Surgery (Counselling Collections)', d.revenueBreakup.surgery],
      ['Optical', d.revenueBreakup.optical],
      ['Pharmacy', d.revenueBreakup.pharmacy],
      ['TOTAL', d.revenueBreakup.total]
    ]);
  }

  // ---- Surgery Pipeline (month-wise) ----
  body.appendParagraph('Surgery Pipeline — Month-wise (Forecast Floor)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (d.pipelineByMonth.length) {
    appendKVTable_(body, d.pipelineByMonth.map(m => [m.month, m.total + (isFinanceVisible ? ` (₹${m.revenue.toLocaleString('en-IN')})` : '')]));
  } else {
    body.appendParagraph('No surgery dates scheduled yet.').setItalic(true);
  }

  // ---- Surgery Pipeline (category-wise) ----
  body.appendParagraph('Surgery Pipeline — Category-wise').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, d.pipelineByCategory.map(c => [c.procedure, c.total + (isFinanceVisible ? ` (₹${c.revenue.toLocaleString('en-IN')})` : '')]));

  // ---- Store / Inventory ----
  body.appendParagraph('Store & Inventory — Current Stock').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKVTable_(body, d.inventory.items.map(it => [it.item + ' (' + it.category + ')', it.current + ' ' + it.unit + (it.lowStock ? '  ⚠ LOW (reorder ' + it.reorderLevel + ')' : '')]));
  if (d.inventory.issuedThisPeriod.length) {
    body.appendParagraph('Items Issued This Period').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    appendKVTable_(body, d.inventory.issuedThisPeriod.map(it => [it.item, it.qty]));
  }

  // ---- Purchase & Expenses (Admin/Doctor only) ----
  if (d.purchase) {
    body.appendParagraph('Purchase & Expenses').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    appendKVTable_(body, [
      ['Total Purchase (₹)', d.purchase.totalPurchase],
      ['Amount Paid (₹)', d.purchase.totalPaid],
      ['Amount Pending — Period (₹)', d.purchase.totalPending],
      ['Total Vendor Dues — All-time (₹)', d.purchase.totalVendorPendingAllTime]
    ]);
    if (d.purchase.topVendors.length) {
      body.appendParagraph('Top Vendors').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      const rows = [['Vendor','Top Product','Total (₹)','Pending (₹)']];
      d.purchase.topVendors.forEach(v => rows.push([v.vendor, v.topProduct, v.totalAmount, v.totalPending]));
      appendTable_(body, rows);
    }
    if (Object.keys(d.purchase.expenseByCategory).length) {
      body.appendParagraph('Expense Breakdown by Category (₹)').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      appendKVTable_(body, Object.entries(d.purchase.expenseByCategory));
    }
  }

  // ---- TPA Summary ----
  body.appendParagraph('TPA / Accounts Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (d.tpa.received !== null) {
    appendKVTable_(body, [
      ['Payment Received (₹)', d.tpa.received],
      ['Payment Pending (₹)', d.tpa.pendingAmount],
      ['Payment Expected (₹)', d.tpa.expected],
      ['Total TPA Business (₹)', d.tpa.total],
      ['Bills Matched', d.tpa.billsMatched],
      ['Bills Mismatched', d.tpa.billsMismatched],
      ['Queries Raised / Resolved', d.tpa.queriesRaised + ' / ' + d.tpa.queriesResolved]
    ]);
  } else {
    body.appendParagraph('Visible to Admin/Doctor only.').setItalic(true);
  }

  // ---- Counsellor Leaderboard ----
  body.appendParagraph('Counsellor Leaderboard').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (d.counselling.counsellorLeaderboard.length) {
    const rows = [['Counsellor','Counselled','Converted','Conv %', isFinanceVisible ? 'Revenue (₹)' : '']];
    d.counselling.counsellorLeaderboard.forEach(c => rows.push([c.name, c.counselled, c.converted, c.conversionRate + '%', isFinanceVisible ? (c.revenue||0) : '']));
    appendTable_(body, rows);
  }

  // ---- Doctor Performance ----
  body.appendParagraph('Doctor-wise Performance').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (d.doctorPerformance.length) {
    const rows = [['Doctor','Referrals','Converted','Conv %','OT Cases']];
    d.doctorPerformance.forEach(doc2 => rows.push([doc2.name, doc2.referred, doc2.converted, doc2.conversionRate + '%', doc2.otCases]));
    appendTable_(body, rows);
  }

  // ---- Period Comparisons ----
  body.appendParagraph('Performance Comparisons').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendComparisonTable_(body, 'This Month vs Last Month', comp.monthVsLastMonth);
  appendComparisonTable_(body, 'This Month vs Same Month Last Year', comp.monthVsLastYear);
  appendComparisonTable_(body, 'This Quarter vs Last Quarter', comp.quarterVsLastQuarter);
  appendComparisonTable_(body, 'This Quarter vs Same Quarter Last Year', comp.quarterVsLastYear);
  appendComparisonTable_(body, 'This Year vs Last Year', comp.yearVsLastYear);

  // ---- Flags ----
  if (d.flags && d.flags.length) {
    body.appendParagraph('Needs Review').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    d.flags.forEach(f => body.appendListItem(f).setGlyphType(DocumentApp.GlyphType.BULLET));
  }

  doc.saveAndClose();

  const docId = doc.getId();
  const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
  const base64 = Utilities.base64Encode(pdfBlob.getBytes());
  DriveApp.getFileById(docId).setTrashed(true);

  return {
    success: true,
    filename: `${hospital.replace(/[^a-zA-Z0-9]/g,'_')}_Report_${d.compareLabel.replace(/\s/g,'_')}.pdf`,
    base64: base64
  };
}

function appendKVTable_(body, pairs) {
  const table = body.appendTable();
  pairs.forEach(([k, v]) => {
    const row = table.appendTableRow();
    row.appendTableCell(String(k));
    row.appendTableCell(String(v === undefined || v === null ? '—' : v));
  });
  formatTable_(table);
}

function appendTable_(body, rows) {
  const table = body.appendTable();
  rows.forEach((r, i) => {
    const row = table.appendTableRow();
    r.forEach(c => {
      const cell = row.appendTableCell(String(c === undefined || c === null ? '' : c));
      if (i === 0) cell.editAsText().setBold(true);
    });
  });
  formatTable_(table);
}

function formatTable_(table) {
  for (let i = 0; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    for (let j = 0; j < row.getNumCells(); j++) {
      row.getCell(j).setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(6).setPaddingRight(6);
    }
  }
}

function appendComparisonTable_(body, title, comp) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  const cur = comp.current, prev = comp.previous;
  const rows = [['Metric', cur.label, prev.label, '% Change']];
  const metrics = [
    ['Total OPD', 'totalOPD'],
    ['Counselled', 'totalCounselled'],
    ['Converted', 'totalConverted'],
    ['Conversion Rate (%)', 'conversionRate'],
    ['Surgery Cases', 'surgeryCases']
  ];
  if (cur.totalRevenue !== null) metrics.push(['Total Revenue (₹)', 'totalRevenue']);
  metrics.forEach(([label, key]) => {
    const c = cur[key], p = prev[key];
    const pct = (p && p !== 0) ? Math.round(((c - p) / p) * 100) + '%' : (c > 0 ? '+100%' : '0%');
    rows.push([label, c, p, pct]);
  });
  appendTable_(body, rows);
}

/**
 * Admin-only Excel export: a multi-sheet workbook with a summary tab and
 * raw daily-entry data for the given range, returned as base64 (.xlsx).
 */
function generateReportExcel(range, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Excel export is available to Admin only.' };

  const settings = getSettings();
  const d = getDashboardData(range, user);
  const { startDate, endDate } = getRangeDates_(range);
  const hospital = settings['Hospital Name'] || 'Hospital';

  const ss = SpreadsheetApp.create(`__temp_export_${Date.now()}`);

  // ---- Summary sheet ----
  const summary = ss.getSheets()[0];
  summary.setName('Summary');
  let r = 1;
  const writeRow = (sheet, row, values) => sheet.getRange(row, 1, 1, values.length).setValues([values]);
  writeRow(summary, r++, [hospital + ' — Performance Summary (' + d.compareLabel + ')']);
  r++;
  writeRow(summary, r++, ['Total OPD', d.reception.totalOPD]);
  writeRow(summary, r++, ['Fresh / New', d.opdBreakup.fresh]);
  writeRow(summary, r++, ['Old / Follow-up', d.opdBreakup.followUp]);
  writeRow(summary, r++, ['FOC', d.opdBreakup.foc]);
  writeRow(summary, r++, ['Counselled', d.counselling.totalCounselled]);
  writeRow(summary, r++, ['Converted', d.counselling.totalConverted]);
  writeRow(summary, r++, ['Conversion Rate (%)', d.counselling.conversionRate]);
  r++;
  writeRow(summary, r++, ['Surgeries by Category']);
  Object.entries(d.surgeryCategoryCounts).forEach(([k,v]) => writeRow(summary, r++, [k, v]));
  r++;
  writeRow(summary, r++, ['Diagnostic Volumes']);
  Object.entries(d.diagnosticCategoryCounts).forEach(([k,v]) => writeRow(summary, r++, [k, v]));
  if (d.revenueBreakup) {
    r++;
    writeRow(summary, r++, ['Revenue Breakup (₹)']);
    writeRow(summary, r++, ['OPD', d.revenueBreakup.opd]);
    writeRow(summary, r++, ['Diagnostic', d.revenueBreakup.diagnostic]);
    writeRow(summary, r++, ['Surgery', d.revenueBreakup.surgery]);
    writeRow(summary, r++, ['Optical', d.revenueBreakup.optical]);
    writeRow(summary, r++, ['Pharmacy', d.revenueBreakup.pharmacy]);
    writeRow(summary, r++, ['TOTAL', d.revenueBreakup.total]);
  }
  r++;
  writeRow(summary, r++, ['Store & Inventory — Current Stock']);
  d.inventory.items.forEach(it => writeRow(summary, r++, [it.item + ' (' + it.category + ')', it.current + ' ' + it.unit + (it.lowStock ? ' — LOW' : '')]));
  if (d.purchase) {
    r++;
    writeRow(summary, r++, ['Purchase & Expenses (₹)']);
    writeRow(summary, r++, ['Total Purchase', d.purchase.totalPurchase]);
    writeRow(summary, r++, ['Amount Paid', d.purchase.totalPaid]);
    writeRow(summary, r++, ['Amount Pending (Period)', d.purchase.totalPending]);
    writeRow(summary, r++, ['Total Vendor Dues (All-time)', d.purchase.totalVendorPendingAllTime]);
    r++;
    writeRow(summary, r++, ['Top Vendors', 'Top Product', 'Total (₹)', 'Pending (₹)']);
    d.purchase.topVendors.forEach(v => writeRow(summary, r++, [v.vendor, v.topProduct, v.totalAmount, v.totalPending]));
  }
  if (d.tpa.received !== null) {
    r++;
    writeRow(summary, r++, ['TPA Summary (₹)']);
    writeRow(summary, r++, ['Received', d.tpa.received]);
    writeRow(summary, r++, ['Pending', d.tpa.pendingAmount]);
    writeRow(summary, r++, ['Expected', d.tpa.expected]);
    writeRow(summary, r++, ['Total Business', d.tpa.total]);
  }
  summary.autoResizeColumns(1, 2);

  // ---- Raw data sheets for the period ----
  const rawSheets = ['Reception','Counselling','OT','Diagnostic','Pharmacy','Optical','Marketing','HR','TPA','TPAClaims','TPASettlements','Operations','Store','Purchase','CallCentre','HRSummary','TPASummary'];
  rawSheets.forEach(name => {
    const filtered = filterByDate_(name, startDate, endDate);
    const { headers } = getSheetData_(name);
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (filtered.length) {
      const values = filtered.map(row => headers.map(h => row[h] !== undefined ? row[h] : ''));
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }
    sheet.autoResizeColumns(1, headers.length);
  });

  SpreadsheetApp.flush();
  const ssId = ss.getId();
  const blob = DriveApp.getFileById(ssId).getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const base64 = Utilities.base64Encode(blob.getBytes());
  DriveApp.getFileById(ssId).setTrashed(true);

  return {
    success: true,
    filename: `${hospital.replace(/[^a-zA-Z0-9]/g,'_')}_Export_${d.compareLabel.replace(/\s/g,'_')}.xlsx`,
    base64: base64
  };
}

// ============================================================
//  AI ANALYSIS (Admin only) — CEO-level narrative commentary
// ============================================================

/**
 * If Settings > "AI API Key" (an Anthropic API key) is filled in, sends a
 * compact summary of this period's numbers + comparisons to Claude and
 * returns a short narrative analysis. Returns null if no key is set.
 */
function getAIAnalysis(range, user) {
  if (user.role !== 'Admin') return { success: false, message: 'AI analysis is available to Admin only.' };
  const settings = getSettings();
  const apiKey = settings['AI API Key'];
  if (!apiKey) {
    return { success: false, message: 'No AI API Key set in Settings. Add an Anthropic API key under Settings > AI API Key to enable this section.' };
  }

  const d = getDashboardData(range, user);
  const comp = getComparisonData(user);

  const summary = {
    period: d.compareLabel,
    opd: d.reception,
    opdBreakup: d.opdBreakup,
    conversion: d.counselling,
    surgeryCategories: d.surgeryCategoryCounts,
    diagnosticVolumes: d.diagnosticCategoryCounts,
    revenueBreakup: d.revenueBreakup,
    pipelineByMonth: d.pipelineByMonth,
    pipelineByCategory: d.pipelineByCategory,
    tpa: d.tpa,
    inventory: { lowStockItems: d.inventory.lowStockItems, issuedThisPeriod: d.inventory.issuedThisPeriod },
    purchase: d.purchase,
    flags: d.flags,
    comparisons: comp
  };

  const prompt =
    'You are a hospital business analyst writing a CEO-level summary for an ophthalmology hospital chain. ' +
    'Based on the JSON data below, write a concise analysis (4-6 short paragraphs) covering: ' +
    '(1) overall performance this period, (2) trend direction vs last month/quarter/year, ' +
    '(3) the strongest and weakest procedure/revenue categories, (4) any pilferage/leakage, low-stock ' +
    'inventory, or operational risks suggested by the flags or numbers, (5) purchase/expense patterns ' +
    'worth noting (top vendors, pending dues, category spend), and (6) 3-4 specific, actionable ' +
    'recommendations. Be direct and specific with numbers. Do not use markdown headers, just plain paragraphs.\n\n' +
    'DATA:\n' + JSON.stringify(summary);

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 1200,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    if (json.content && json.content.length) {
      const text = json.content.map(c => c.text || '').join('\n').trim();
      return { success: true, text: text };
    }
    return { success: false, message: 'AI service did not return a usable response. ' + (json.error ? json.error.message : '') };
  } catch (e) {
    return { success: false, message: 'AI request failed: ' + e.message };
  }
}

// ============================================================
//  DAILY AUTOMATED REPORTS — PDF email (Admin & Doctor) + WhatsApp
//  Set up a daily time-driven trigger on sendDailyReports() (e.g. 9 PM)
//  to have today's PDF report emailed automatically, and (if configured)
//  a short WhatsApp summary sent to Admin.
// ============================================================

function sendDailyReports() {
  const settings = getSettings();
  const hospital = settings['Hospital Name'] || 'Hospital';
  const adminEmails = String(settings['Daily Report Emails - Admin'] || '').split(',').map(e => e.trim()).filter(Boolean);
  const doctorEmails = String(settings['Daily Report Emails - Doctor'] || '').split(',').map(e => e.trim()).filter(Boolean);

  const adminUser = { name: 'Admin (Auto Report)', role: 'Admin', department: 'Admin' };
  const doctorUser = { name: 'Doctor (Auto Report)', role: 'Doctor', department: 'Doctor' };

  if (adminEmails.length) {
    const pdf = generateReportPDF('today', adminUser);
    if (pdf.success) sendPdfEmail_(adminEmails, hospital, pdf);
  }
  if (doctorEmails.length) {
    const pdf = generateReportPDF('today', doctorUser);
    if (pdf.success) sendPdfEmail_(doctorEmails, hospital, pdf);
  }

  sendWhatsAppSummary_(adminUser, settings, hospital);
}

/**
 * Run this on a monthly trigger (Day of month timer — set the day in
 * Settings > "Monthly Report Day (1-28)" purely as documentation; the
 * Apps Script trigger itself controls WHEN this runs). Sends a complete
 * "Last Month" report — including the AI Executive Summary for Admin, if
 * an AI API Key is configured — to whoever is listed in
 * Settings > "Monthly Report Emails - Admin" / "...- Doctor".
 */
/**
 * Run on a time-driven trigger: Apps Script editor → Triggers →
 * + Add Trigger → sendMonthlyReport → Time-driven → Month timer →
 * Day 1 of month → 6am–7am.
 *
 * Always covers the FULL previous calendar month (e.g. if today is
 * 1 June 2026 it covers 1 May–31 May 2026), regardless of which day
 * the trigger fires on — so months with 28/29/30/31 days are always
 * correct without any manual adjustment.
 */
function sendMonthlyReport() {
  const settings = getSettings();
  const hospital = settings['Hospital Name'] || 'Hospital';
  const adminEmails = String(settings['Monthly Report Emails - Admin'] || '').split(',').map(e => e.trim()).filter(Boolean);
  const doctorEmails = String(settings['Monthly Report Emails - Doctor'] || '').split(',').map(e => e.trim()).filter(Boolean);

  // Compute last month's label for the email subject
  const now = new Date();
  const lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonthLabel = Utilities.formatDate(lastMonthDate, TZ, 'MMMM yyyy');

  const adminUser = { name: 'Admin (Auto Report)', role: 'Admin', department: 'Admin' };
  const doctorUser = { name: 'Doctor (Auto Report)', role: 'Doctor', department: 'Doctor' };

  if (adminEmails.length) {
    const pdf = generateReportPDF('lastMonth', adminUser);
    if (pdf.success) sendPdfEmail_(adminEmails, hospital, pdf, true, lastMonthLabel);
  }
  if (doctorEmails.length) {
    const pdf = generateReportPDF('lastMonth', doctorUser);
    if (pdf.success) sendPdfEmail_(doctorEmails, hospital, pdf, true, lastMonthLabel);
  }
}

function sendPdfEmail_(emails, hospital, pdf, isMonthly, monthLabel) {
  const subject = isMonthly
    ? `${hospital} — Monthly Performance Report (${monthLabel || Utilities.formatDate(new Date(), TZ, 'MMMM yyyy')})`
    : `${hospital} — Daily Performance Report (${Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy')})`;
  const body = isMonthly
    ? `Attached is the complete performance report for ${monthLabel} at ${hospital}.\n\nThis is an automated monthly summary generated by the EyeForce Solutions dashboard.`
    : `Attached is today's performance report for ${hospital}.\n\nThis is an automated daily summary generated by the EyeForce Solutions dashboard.`;
  const blob = Utilities.newBlob(Utilities.base64Decode(pdf.base64), 'application/pdf', pdf.filename);
  emails.forEach(email => {
    try { MailApp.sendEmail({ to: email, subject: subject, body: body, attachments: [blob] }); } catch (e) { /* ignore individual failures */ }
  });
}

/**
 * Optional WhatsApp summary — posts a short text summary to a webhook URL
 * configured in Settings > "WhatsApp Webhook URL (Admin)". This is left
 * generic so it can plug into Twilio, Gupshup, Interakt, or any WhatsApp
 * Business API / automation provider that accepts a simple JSON POST.
 * Adjust the payload key ("message") to match your provider if needed.
 * If the setting is blank, this is a no-op.
 */
function sendWhatsAppSummary_(user, settings, hospital) {
  const url = settings['WhatsApp Webhook URL (Admin)'];
  if (!url) return;
  try {
    const d = getDashboardData('today', user);
    const totalSurgeries = Object.values(d.surgeryCategoryCounts).reduce((a, b) => a + b, 0);
    const lines = [
      `*${hospital} — Daily Summary (${Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy')})*`,
      `Total OPD: ${d.reception.totalOPD}`,
      `Counselled: ${d.counselling.totalCounselled} | Converted: ${d.counselling.totalConverted} (${d.counselling.conversionRate}%)`,
      `Surgeries (OT): ${totalSurgeries}`,
      d.revenueBreakup ? `Revenue Today: ₹${d.revenueBreakup.total.toLocaleString('en-IN')}` : '',
      d.flags.length ? `Needs review: ${d.flags.join(' | ')}` : 'No reconciliation flags today.'
    ].filter(Boolean);
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ message: lines.join('\n') }),
      muteHttpExceptions: true
    });
  } catch (e) { /* WhatsApp failures should never block the rest of the flow */ }
}

// ============================================================
//  BULK IMPORT / EXPORT (Admin only)
//  Lets Admin pull a ready-made CSV template for any sheet (e.g. Store,
//  InventoryMaster, or any department's historical data), fill it in
//  Excel/Google Sheets, and paste the data back in as new rows — without
//  retyping anything by hand.
// ============================================================

const BULK_IMPORTABLE_SHEETS = [
  'Reception','Counselling','CounsellingSummary','OT','Diagnostic','Pharmacy','Optical',
  'Marketing','HR','TPA','Operations','Store','Purchase','InventoryMaster'
];

function getBulkImportableSheets(user) {
  if (user.role !== 'Admin') return [];
  return BULK_IMPORTABLE_SHEETS;
}

/**
 * Returns a CSV template (header row + one example row) for the given sheet,
 * as base64, ready for the front-end to turn into a download. Admin opens
 * this in Excel/Google Sheets, fills in rows below the header, saves as CSV,
 * and re-uploads via bulkImportCSV.
 */
function getBulkTemplate(sheetName, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Bulk import/export is available to Admin only.' };
  if (BULK_IMPORTABLE_SHEETS.indexOf(sheetName) === -1) return { success: false, message: 'Unknown sheet.' };

  const sheet = SS.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Build one example row to show the expected format (especially for Date).
  const example = headers.map(h => {
    if (h === 'Timestamp') return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    if (h === 'Date') return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (h === 'Submitted By') return 'Bulk Import';
    if (h === 'Edit Status') return 'Locked';
    if (h === 'Converted' || h === 'Cleanliness Check' || h === 'Biomedical Servicing Due' || h === 'Training Completed' || h === 'Stock Check Done') return 'Yes';
    if (h === 'Transaction Type') return 'Received (Purchase In)';
    if (h === 'Payment Status') return 'Paid';
    return '0';
  });

  const csv = [headers, example].map(row =>
    row.map(v => {
      const s = String(v === undefined || v === null ? '' : v);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',')
  ).join('\r\n');

  return {
    success: true,
    filename: `${sheetName}_Import_Template.csv`,
    base64: Utilities.base64Encode(csv)
  };
}

/**
 * Appends rows from a pasted/uploaded CSV (header row + data rows) to the
 * given sheet. The header row MUST exactly match the sheet's own headers
 * (use getBulkTemplate to get the right format). Any 'Date' values should be
 * yyyy-MM-dd (matching every other entry in the sheet). 'Submitted By' and
 * 'Edit Status' are auto-filled with 'Bulk Import' / 'Locked' if left blank,
 * so historical rows are protected by the same correction-request workflow.
 */
function bulkImportCSV(sheetName, csvText, user) {
  if (user.role !== 'Admin') return { success: false, message: 'Bulk import is available to Admin only.' };
  if (BULK_IMPORTABLE_SHEETS.indexOf(sheetName) === -1) return { success: false, message: 'Unknown sheet.' };

  const sheet = SS.getSheetByName(sheetName);
  const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const parsed = Utilities.parseCsv(csvText);
  if (parsed.length < 2) return { success: false, message: 'No data rows found in the file.' };

  const fileHeaders = parsed[0].map(h => String(h).trim());
  // Validate every file header exists in the sheet (order can differ).
  const missing = fileHeaders.filter(h => sheetHeaders.indexOf(h) === -1);
  if (missing.length) {
    return { success: false, message: 'These columns don\'t match the sheet — please use the downloaded template: ' + missing.join(', ') };
  }

  const dataRows = parsed.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
  const rowsToAppend = dataRows.map(r => {
    const rowObj = {};
    fileHeaders.forEach((h, i) => rowObj[h] = r[i]);
    return sheetHeaders.map(h => {
      let v = rowObj[h];
      if (v === undefined || v === '') {
        if (h === 'Submitted By') return 'Bulk Import';
        if (h === 'Edit Status') return 'Locked';
        if (h === 'Timestamp') return nowStr_();
        return '';
      }
      return v;
    });
  });

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, sheetHeaders.length).setValues(rowsToAppend);
  }
  logAudit_(user.name, 'BULK-IMPORT', sheetName, sheet.getLastRow(), `${rowsToAppend.length} row(s) imported.`);
  return { success: true, message: `${rowsToAppend.length} row(s) imported into ${sheetName}.` };
}
