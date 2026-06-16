/**
 * EyeForce Solutions — Hospital Sheet Setup
 * -------------------------------------------------
 * Run `setupHospitalSheet()` ONCE from the Apps Script editor
 * (select it from the function dropdown, click Run) on a brand-new
 * Google Sheet to create the entire tab structure for a new hospital.
 *
 * To onboard a new hospital:
 *  1. Make a copy of this whole template (Sheet + Apps Script project)
 *  2. Open the copy, run setupHospitalSheet() once
 *  3. Fill in the SETTINGS and PRICE MASTER tabs for that hospital
 *  4. Deploy > New deployment > Web app
 *  5. Share the web app link with that hospital's staff & doctor
 */

function setupHospitalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buildSettingsSheet_(ss);

  createSheetWithHeaders(ss, 'Staff', [
    'Staff Name', 'Department', 'Role', 'Login ID', 'PIN', 'Status'
  ], [
    ['Admin', 'Admin', 'Admin', 'admin', '0000', 'Active'],
    ['Dr. Sharma', 'Doctor', 'Doctor', 'drsharma', '1111', 'Active'],
    ['Reena', 'Reception', 'Staff', 'reena', '1234', 'Active'],
    ['Kamal', 'Counselling', 'Staff', 'kamal', '1234', 'Active']
  ]);

  buildPriceMasterSheet_(ss);

  // InventoryMaster: comprehensive list — Admin sets opening stock and reorder levels.
  // IOLs listed by brand + diopter range (each diopter is a separate inventory item
  // because hospitals stock individual powers). Staff submit Store issues against these.
  buildInventoryMasterSheet_(ss);


  createSheetWithHeaders(ss, 'Reception', [
    'Timestamp', 'Date',
    // Patient counts
    'Total OPD', 'New Patients', 'Old/Follow-up', 'FOC Patients',
    // Payer category breakdown
    'CGHS', 'ECHS', 'Delhi Govt', 'Cash', 'Other TPA/Insurance',
    // Source breakdown
    'Source - Social Media', 'Source - Referral', 'Source - Walk-in/Other',
    'Source - Newspaper/Hoarding', 'Source - Other',
    // Operations
    'No-Shows Today', 'Appointments Booked',
    // Cash collected — Reception counter handles OPD + OT advance + Diagnostic
    'OPD Cash Collected (₹)',          // consultation fees
    'OT Cash Collected (₹)',           // surgery package advance from cash patients
    'Diagnostic Cash Collected (₹)',   // diagnostic tests billed at reception counter
    'Others Cash Collected (₹)',       // any other cash — remarks mandatory
    'Others Remarks',                  // mandatory when Others Cash > 0
    'Submitted By', 'Edit Status'
  ], []);

  createSheetWithHeaders(ss, 'Counselling', [
    'Timestamp', 'Date', 'Counselor Name', 'Patient Reference', 'Referring Doctor',
    'Procedure', 'Category', 'Cash Package',
    'Converted', 'Pipeline Status', 'Surgery Date Given', 'Dropout Reason',
    'Follow-up Calls Made',
    // Backend-computed fields (auto-filled by submitEntry, never shown on form):
    'Price Master Rate (₹)', 'Estimated Revenue (₹)',
    'Submitted By', 'Edit Status'
  ], []);

  // ---- CounsellingSummary: for hospitals with NO dedicated counsellor ----
  // Used INSTEAD OF the per-patient Counselling tab above — controlled by
  // Settings > "Counselling Mode (Per-Patient / Daily Summary)". Only ONE
  // of the two tabs should be filled in for a given hospital, so the
  // dashboard never double-counts. See README for details.
  const procedureCategories = ['Cataract','LASIK','Glaucoma','Retina','Cornea','Squint','Oculoplasty','Anti-VEGF','Others'];
  const summaryHeaders = ['Timestamp', 'Date'];
  procedureCategories.forEach(c => {
    summaryHeaders.push(c + ' - Advised', c + ' - Counselled', c + ' - Converted', c + ' - Pending');
  });
  summaryHeaders.push(
    'Cash Conversions', 'CGHS Conversions', 'ECHS Conversions',
    'Delhi Govt Conversions', 'Other TPA Conversions',
    'Submitted By', 'Edit Status'
  );
  createSheetWithHeaders(ss, 'CounsellingSummary', summaryHeaders, []);

  createSheetWithHeaders(ss, 'OT', [
    'Timestamp', 'Date', 'Doctor Name', 'Procedure Type',
    // OT staff fill case counts only from their OT list — no money fields.
    // Cash revenue is tracked at Reception (OPD Revenue Collected).
    // TPA billed amount is entered by TPA/cashless person in TPAClaims.
    'Cash Cases',                                               // count only
    'CGHS Cases', 'CGHS Upgraded Cases',                       // upgraded = patient chose higher package
    'ECHS Cases', 'ECHS Upgraded Cases',
    'Delhi Govt Cases', 'Delhi Govt Upgraded Cases',
    'TPA Cases', 'TPA Upgraded Cases', 'TPA Insurer',          // insurer name for reference
    // All columns below: auto-computed by backend, never touched by OT staff
    'Total Cases',
    'CGHS Rate (₹)', 'CGHS Expected (₹)',
    'ECHS Rate (₹)', 'ECHS Expected (₹)',
    'Delhi Govt Rate (₹)', 'Delhi Govt Expected (₹)',
    'Total Expected Claims (₹)',
    'Submitted By', 'Edit Status'
  ], []);

  // Diagnostic sheet: paid + free counts per test category.
  // Driven by DIAGNOSTIC_CATS array — to add a new test, add it here and in
  // the 'Diagnostic Category List' setting, then re-run setupHospitalSheet().
  // The form generates fields dynamically from Settings, so it updates
  // immediately after a Settings change (no setup re-run needed for the form).
  const DIAGNOSTIC_CATS = ['Biometry','Topography','OCT','Cataract Work-up','LASIK Work-up',
    'Green Laser','Fundus Photography','Visual Field/Perimetry','B-Scan/A-Scan','Other Tests'];
  const diagHeaders_ = ['Timestamp', 'Date'];
  DIAGNOSTIC_CATS.forEach(c => { diagHeaders_.push(c + ' - Paid', c + ' - Free'); });
  diagHeaders_.push('Diagnostic Revenue Auto (₹)', 'Submitted By', 'Edit Status');
  createSheetWithHeaders(ss, 'Diagnostic', diagHeaders_, []);

  createSheetWithHeaders(ss, 'Pharmacy', [
    'Timestamp', 'Date', 'Total Prescriptions', 'Prescriptions - Cash', 'Prescriptions - TPA/Insurance',
    'Stock-Out Alerts', 'Revenue Collected (₹)', 'Submitted By', 'Edit Status'
    // Revenue Collected (₹) = actual cash collected at pharmacy counter today.
    // This is staff-entered because it reflects real daily collections across
    // all medicines/items — not derivable from a price master.
  ], []);

  createSheetWithHeaders(ss, 'Optical', [
    'Timestamp', 'Date', 'Frames Sold', 'Lenses Sold', 'Contact Lens Sold', 'Walk-in Customers',
    'Revenue Collected (₹)', 'Submitted By', 'Edit Status'
    // Revenue Collected (₹) = actual cash collected at optical counter today.
    // Optical has variable frame/lens prices so staff enter the actual total.
  ], []);

  createSheetWithHeaders(ss, 'Marketing', [
    'Timestamp', 'Date', 'PRO Name', 'Area Covered', 'Calls Made', 'Visits Made',
    'Visit Details', 'Leads Generated', 'Lead Category', 'Leads Converted to OPD',
    'Submitted By', 'Edit Status'
  ], []);

  createSheetWithHeaders(ss, 'HR', [
    'Timestamp', 'Date', 'Department', 'Staff Name', 'Attendance', 'Leaves Taken',
    'Late Arrivals', 'Discipline Note', 'Grooming Compliance', 'Grievance Note',
    'Training Completed', 'Behaviour Rating (1-5)', 'Performance Score (1-5)',
    'Submitted By', 'Edit Status'
  ], []);

  // HRSummary — used when HR Mode = "Daily Summary" (no dedicated HR staff per person).
  // One row per day — total counts across all staff.
  createSheetWithHeaders(ss, 'HRSummary', [
    'Timestamp', 'Date',
    'Total Staff', 'Present', 'Absent', 'Half-day',
    'Late Arrivals', 'Grooming Non-Compliant', 'Discipline Incidents',
    'Grievances', 'Leaves Taken', 'Training Completed',
    'Notes',
    'Submitted By', 'Edit Status'
  ], []);

  // TPASummary — used when TPA Mode = "Daily Summary".
  // Revenue logic: staff enter aggregate cash received and expected amounts.
  // Dashboard uses these directly, same as detailed mode but aggregate.
  createSheetWithHeaders(ss, 'TPASummary', [
    'Timestamp', 'Date',
    'CGHS Cases Operated', 'CGHS Amount Expected (₹)',
    'ECHS Cases Operated', 'ECHS Amount Expected (₹)',
    'Delhi Govt Cases Operated', 'Delhi Govt Amount Expected (₹)',
    'Other TPA Cases Operated', 'Other TPA Amount Expected (₹)',
    'Settlements Received Today (₹)',
    'Total Expected Pending (₹)',
    'Co-pay Collected Today (₹)',
    'Co-pay Pending (₹)',
    'Queries Raised', 'Queries Resolved',
    'Notes',
    'Submitted By', 'Edit Status'
  ], []);

  // ── TPA Daily Operations — summary only, no claim details here ───────────
  createSheetWithHeaders(ss, 'TPA', [
    'Timestamp', 'Date',
    'CGHS Bills Filed Today', 'ECHS Bills Filed Today',
    'Delhi Govt Bills Filed Today', 'Other TPA Bills Filed Today',
    'Queries Raised', 'Queries Resolved', 'Notes',
    'Submitted By', 'Edit Status'
  ], []);

  // ── TPAClaims — one row per claim filed to CGHS/ECHS/Delhi Govt/TPA ───────
  // TPA/cashless person logs when they send the bill to the payer.
  // If patient upgraded → billing person fills Co-pay fields on this same row.
  // Claim Ref is MANDATORY and UNIQUE — it's what links this claim to its
  // settlement in TPASettlements. 30-day clock starts from Date Filed.
  createSheetWithHeaders(ss, 'TPAClaims', [
    'Timestamp', 'Date Filed',
    'Payer Type',                   // CGHS / ECHS / Delhi Govt / Other TPA
    'Insurer / Payer Name',
    'Claim Ref No',                 // MANDATORY — unique ref, e.g. CGHS/2026/001
    'Patient Name',
    'Surgery Date',
    'Procedure',
    'Claim Amount (₹)',             // fixed govt/insurer rate for this procedure
    'Upgraded?',                    // Yes / No
    'Co-pay Due (₹)',               // if upgraded: amount patient must pay (difference)
    'Co-pay Collected (₹)',         // billing person fills this when patient pays
    'Co-pay Pending (₹)',           // auto: Due − Collected
    'Claim Status',                 // Filed / Queried / Settled
    'Notes',
    'Submitted By', 'Edit Status'
  ], []);

  // ── TPASettlements — one row per payment received ──────────────────────────
  // Used for BOTH: (1) CGHS/ECHS/Delhi Govt/TPA paying the hospital, and
  // (2) patient paying co-pay balance. Settlement Type distinguishes them.
  // Claim Ref MUST match a row in TPAClaims — this is how drill-down works
  // and how Expected converts to Actual in the dashboard.
  // Shortfall and Days to Settle are auto-computed on save — staff never enter.
  createSheetWithHeaders(ss, 'TPASettlements', [
    'Timestamp', 'Date Received',
    'Settlement Type',              // Govt/TPA Payment OR Patient Co-pay
    'Claim Ref No',                 // MANDATORY — must match a TPAClaims row
    'Amount Received (₹)',          // actual amount received today
    'Payment Mode',                 // ECS/NEFT / Cheque / Cash / UPI
    // Auto-computed on save — staff never fill these:
    'Original Claim Amount (₹)',    // pulled from matching TPAClaims row
    'Total Settled So Far (₹)',     // sum of all settlements for this ref
    'Shortfall (₹)',                // Original − Total Settled
    'Days to Settle',               // Date Received − Surgery Date from claim
    'Notes',
    'Submitted By', 'Edit Status'
  ], []);

  createSheetWithHeaders(ss, 'Operations', [
    'Timestamp', 'Date', 'Avg Wait Time (mins)', 'Complaints Received', 'Complaints Resolved',
    'Equipment Breakdown Note', 'Equipment Downtime (hrs)', 'Cleanliness Check',
    'Biomedical Servicing Due', 'Infrastructure Issue Note', 'Stock Check Done',
    'Google Reviews Asked', 'Google Reviews Received',
    'Submitted By', 'Edit Status'
  ], []);

  createSheetWithHeaders(ss, 'CallCentre', [
    'Timestamp', 'Date',
    // Volume
    'Inbound Calls Total', 'Inbound Calls Answered', 'Inbound Calls Missed',
    'Outbound Calls Total',
    // Call purpose breakdown — what did callers enquire about?
    'Calls - General Eye Check-up', 'Calls - Cataract', 'Calls - LASIK/Refractive',
    'Calls - Glaucoma', 'Calls - Retina', 'Calls - Cornea',
    'Calls - Squint/Paediatric', 'Calls - Other Procedure',
    'Calls - Follow-up / Existing Patient', 'Calls - Appointment Reminder',
    'Calls - Complaint / Query',
    // Outcomes
    'Appointments Booked (from calls)', 'Surgery Bookings Confirmed (from calls)',
    'Leads Passed to Counsellor',
    'Submitted By', 'Edit Status'
  ], []);

  // TempAccess: Admin grants a staff member a time-limited window to submit
  // data for a specific past date. Staff sees it on their form as a date picker.
  // Once they submit, Status changes to 'Used' and the window closes.
  createSheetWithHeaders(ss, 'TempAccess', [
    'Staff Name', 'Department', 'Allowed Date', 'Reason / Note',
    'Granted By', 'Granted On', 'Status'  // Status: Open / Used / Cancelled
  ], []);

  createSheetWithHeaders(ss, 'Store', [
    'Timestamp', 'Date', 'Item Name', 'Category', 'Transaction Type', 'Quantity',
    'Issued To / Received From', 'Notes', 'Submitted By', 'Edit Status'
  ], []);

  createSheetWithHeaders(ss, 'Purchase', [
    'Timestamp', 'Date', 'Vendor Name', 'Item/Product', 'Category', 'Quantity',
    'Amount (₹)', 'Amount Paid (₹)', 'Amount Pending (₹)', 'Payment Status',
    'Invoice No', 'Submitted By', 'Edit Status'
  ], []);


  createSheetWithHeaders(ss, 'CorrectionRequests', [
    'Timestamp', 'Department Sheet', 'Row Number', 'Field Name', 'Old Value', 'New Value',
    'Reason', 'Requested By', 'Status', 'Reviewed By', 'Reviewed On'
  ], []);

  createSheetWithHeaders(ss, 'AuditLog', [
    'Timestamp', 'User', 'Action', 'Sheet', 'Row', 'Details'
  ], []);

  // Remove default "Sheet1" if present and empty
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) ss.deleteSheet(sheet1);

  SpreadsheetApp.flush();
  Logger.log('Setup complete. Hospital sheet structure created.');
}

/**
 * SAFE SHEET CREATION / UPGRADE — the most important function for
 * production deployments. Rules:
 *
 *  • Sheet does NOT exist → create it fresh (headers + sample data).
 *  • Sheet ALREADY exists → only add columns that are missing from the
 *    end of the header row. NEVER clear, NEVER delete, NEVER reorder.
 *    Existing data rows are completely untouched.
 *
 * This means re-running setupHospitalSheet() after a code update is
 * always safe. New sheets get created, existing sheets get new columns
 * appended — hospital data is never lost.
 */
function createSheetWithHeaders(ss, name, headers, rows) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    // ── FIRST RUN: sheet doesn't exist → create fresh ───────────────────
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF');
    if (rows && rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }

  // ── SUBSEQUENT RUNS: sheet exists → only add missing columns ────────────
  // Read existing headers (row 1)
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
    : [];
  const existingSet = new Set(existingHeaders.filter(Boolean));

  // Find headers in the new definition that don't exist yet
  const missing = headers.filter(h => h && !existingSet.has(h));
  if (missing.length > 0) {
    const startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    sheet.getRange(1, startCol, 1, missing.length)
      .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF');
    sheet.autoResizeColumns(startCol, missing.length);
    Logger.log('  ' + name + ': added ' + missing.length + ' new column(s): ' + missing.join(', '));
  }
  // Existing data rows and existing columns are left completely untouched.
  return sheet;
}

// ============================================================
//  SETTINGS SHEET — built from SETTINGS_SCHEMA (defined in Code.gs,
//  shared across the Apps Script project) so the raw sheet and the
//  in-app Admin Settings page always describe the exact same fields.
//  Result: a grouped, colour-coded, validated sheet — Setting | Value |
//  Notes — instead of a long flat Key/Value list.
// ============================================================

const DEFAULT_SETTINGS_VALUES = {
  'Hospital Name': 'Sample Eye Hospital',
  'Hospital Code': 'HOSP01',
  'Active Status': 'Active',
  'Logo URL': '',
  'EyeForce Logo URL': '',
  'Departments Enabled': 'Reception,Counselling,OT,Diagnostic,Pharmacy,Optical,Marketing,HR,TPA,Operations,Store,Purchase',
  'Counselling Mode (Per-Patient / Daily Summary)': 'Per-Patient',
  'HR Mode (Per-Person / Daily Summary)': 'Per-Person',
  'TPA Mode (Detailed / Daily Summary)': 'Detailed',
  'Doctor List': 'Dr. Sharma,Dr. Mehta',
  'Procedure List': 'Cataract,LASIK,Glaucoma,Retina,Cornea,Squint,Oculoplasty,Anti-VEGF,Others',
  'Diagnostic Category List': 'Biometry,Topography,OCT,Cataract Work-up,LASIK Work-up,Green Laser,Fundus Photography,Visual Field/Perimetry,B-Scan/A-Scan,Other Tests',
  'TPA Category List': 'Cash,CGHS,ECHS,Delhi Govt,Other TPA',
  'Cashless Category List': 'CGHS,ECHS,Delhi Govt,Other TPA',
  'Cash Package List': 'Basic,Standard,Premium,Super Premium',
  'PRO List': 'PRO 1,PRO 2',
  'Area List': 'Area 1,Area 2,Area 3',
  'Lead Category List': 'Doctor Referral,Corporate Tie-up,Health Camp,Community/NGO,Pharmacy Referral,Other',
  'Issued To List': 'OT,Pharmacy,Optical,Diagnostic,Operations,Other',
  'Transaction Type List': 'Received (Purchase In),Issued (Used/Given Out)',
  'Purchase Category List': 'IOL,Viscoelastic (OVD),Drapes & Disposables,Sutures,Betadine/Antiseptics,Post-Op Kits,Anaesthesia Consumables,Pharmacy Stock,Optical Stock,General/Admin,Equipment,Others',
  'Vendor List': 'Vendor 1,Vendor 2,Vendor 3',
  'Payment Status List': 'Paid,Pending,Partial',
  'Daily Report Emails - Admin': '',
  'Daily Report Emails - Doctor': '',
  'Monthly Report Emails - Admin': '',
  'Monthly Report Emails - Doctor': '',
  'Monthly Report Send Day': '1',
  'WhatsApp Webhook URL (Admin)': '',
  'Notification Emails': '',
  'Google Review Link': '',
  'AI API Key': ''
};

function buildSettingsSheet_(ss) {
  let sheet = ss.getSheetByName('Settings');
  const isFirstRun = !sheet;
  if (!sheet) sheet = ss.insertSheet('Settings');

  // ── SUBSEQUENT RUNS: read existing values before clearing the layout ─────
  // We preserve ALL existing setting values — the admin may have customised
  // Hospital Name, Doctor List, Procedure List etc. and we must not lose them.
  var existingValues = {};
  if (!isFirstRun) {
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        existingData.forEach(function(r) {
          var key = String(r[0]).trim();
          if (key) existingValues[key] = r[1]; // preserve the admin's value
        });
      }
    } catch(e) { /* couldn't read, use defaults */ }
  }

  // Now rebuild the layout (clear and reformat) — but use preserved values
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  try {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), Math.max(sheet.getMaxColumns(), 1)).breakApart();
  } catch (e) {}
  sheet.clear();
  sheet.clearFormats();

  // Header row
  sheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Value (edit this column)', 'Notes / Help']]);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF').setFontSize(11);
  sheet.setFrozenRows(1);

  let row = 2;
  SETTINGS_SCHEMA.forEach(group => {
    sheet.getRange(row, 1, 1, 3).merge();
    sheet.getRange(row, 1).setValue('▸  ' + group.title.toUpperCase());
    sheet.getRange(row, 1, 1, 3)
      .setBackground('#FBEFD6').setFontColor('#7A5D1E')
      .setFontWeight('bold').setFontSize(10.5);
    row++;

    group.fields.forEach(f => {
      // Use existing admin value if available, otherwise use default
      const value = (existingValues[f.key] !== undefined && existingValues[f.key] !== '')
        ? existingValues[f.key]
        : (DEFAULT_SETTINGS_VALUES[f.key] !== undefined ? DEFAULT_SETTINGS_VALUES[f.key] : '');
      let help = f.help || '';
      if (f.type === 'list' || f.type === 'checkboxList') {
        help += (help ? ' ' : '') + 'Comma-separated here — or edit one-per-line in the app\'s Settings page.';
      }
      sheet.getRange(row, 1).setValue(f.key);
      sheet.getRange(row, 2).setValue(value);
      sheet.getRange(row, 3).setValue(help);

      // Highlight the editable Value cell
      sheet.getRange(row, 2).setBackground('#FFFDF5');
      // Notes column: small, muted, wrapped
      sheet.getRange(row, 3).setFontColor('#6E7480').setFontSize(9.5).setWrap(true);
      sheet.getRange(row, 1).setFontWeight('bold').setFontColor('#15233B');

      // Dropdown validation for select-type fields (e.g. Active Status,
      // Counselling Mode, Monthly Report Day)
      if (f.type === 'select' && f.options && f.options.length) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(f.options, true)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(row, 2).setDataValidation(rule);
      }
      // Mask the AI key cell visually with a note (Sheets can't truly mask)
      if (f.type === 'password') {
        sheet.getRange(row, 3).setFontColor('#C0564A');
      }
      row++;
    });
  });

  sheet.setColumnWidth(1, 290);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 420);
  sheet.setRowHeights(2, row - 2, 28);
}

// ============================================================
//  PRICE MASTER SHEET — built as a professional, grouped,
//  colour-coded sheet. Staff never see this. Admin fills in
//  the rates once; backend auto-calculates all revenue.
//  Structure: Item | Category | Payment Type | Rate (₹) | Notes
//  Key used in backend: Item + '|' + Payment Type → Rate (₹)
// ============================================================
function buildPriceMasterSheet_(ss) {
  let sheet = ss.getSheetByName('PriceMaster');
  if (!sheet) sheet = ss.insertSheet('PriceMaster');
  sheet.clear(); sheet.clearFormats();

  const headers = ['Item / Procedure', 'Category', 'Payment Type', 'Rate (₹) — fill in', 'Notes'];
  sheet.getRange(1, 1, 1, 5).setValues([headers])
    .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF').setFontSize(11);
  sheet.setFrozenRows(1);

  // Payment types for surgical procedures
  const CASH_PKGS  = ['Cash - Basic','Cash - Standard','Cash - Premium','Cash - Super Premium'];
  const GOVT_TYPES = ['CGHS','ECHS','Delhi Govt'];
  const ALL_PAY    = [...CASH_PKGS, ...GOVT_TYPES, 'Other TPA'];

  function addProc(rows, name, notes) {
    ALL_PAY.forEach(pt => rows.push([name, 'Procedure', pt, '', notes || '']));
  }
  function addDiag(rows, name, notes) {
    ['Cash', 'CGHS', 'ECHS', 'Delhi Govt', 'Other TPA']
      .forEach(pt => rows.push([name, 'Diagnostic', pt, '', notes || '']));
  }

  const rows = [];

  // ── 1. OPD CONSULTATION ──────────────────────────────────────────────────
  rows.push(['__SECTION__', '1. OPD CONSULTATION FEES']);
  rows.push(['OPD - Fresh Patient',         'OPD','Cash',      '','Per new patient consultation']);
  rows.push(['OPD - Follow-up Patient',     'OPD','Cash',      '','Per follow-up / return visit']);
  rows.push(['OPD - FOC (Free of Cost)',    'OPD','FOC',       '0','Free-of-cost patients — for count tracking only']);
  rows.push(['OPD - Fresh Patient',         'OPD','CGHS',      '','']);
  rows.push(['OPD - Fresh Patient',         'OPD','ECHS',      '','']);
  rows.push(['OPD - Fresh Patient',         'OPD','Delhi Govt','','Same as CGHS in most cases']);
  rows.push(['OPD - Fresh Patient',         'OPD','Other TPA', '','']);

  // ── 2. CATARACT SURGERIES ─────────────────────────────────────────────────
  rows.push(['__SECTION__', '2. CATARACT SURGERIES']);
  addProc(rows, 'Phaco + Monofocal IOL', 'Standard phacoemulsification with monofocal lens');
  addProc(rows, 'Phaco + Monofocal Toric IOL', 'Monofocal with cylinder correction');
  addProc(rows, 'Phaco + EDOF IOL', 'Extended Depth of Focus — intermediate + distance');
  addProc(rows, 'Phaco + EDOF Toric IOL', 'EDOF with cylinder correction');
  addProc(rows, 'Phaco + Trifocal IOL', 'Full range of vision — near/intermediate/distance');
  addProc(rows, 'Phaco + Trifocal Toric IOL', 'Trifocal with cylinder correction');
  addProc(rows, 'Manual SICS + Monofocal IOL', 'Small incision cataract surgery — lower cost option');
  addProc(rows, 'Couching / ECCE (rare)', 'Extracapsular cataract extraction — very basic');
  addProc(rows, 'Secondary IOL Implantation', 'IOL implanted in a previously aphakic eye');
  addProc(rows, 'IOL Exchange / Explant', 'Removal and replacement of existing IOL');

  // ── 3. REFRACTIVE / LASER VISION CORRECTION ───────────────────────────────
  rows.push(['__SECTION__', '3. REFRACTIVE & LASER VISION CORRECTION']);
  addProc(rows, 'LASIK (per eye)', 'Standard microkeratome LASIK');
  addProc(rows, 'Femto-LASIK / Bladeless LASIK (per eye)', 'Femtosecond laser LASIK — premium');
  addProc(rows, 'SMILE (per eye)', 'Small Incision Lenticule Extraction');
  addProc(rows, 'LASEK / PRK (per eye)', 'Photorefractive Keratectomy / LASEK');
  addProc(rows, 'Contoura Vision / Topography-guided (per eye)', 'Topography-guided custom LASIK');
  addProc(rows, 'ICL Implantation (per eye)', 'Implantable Collamer / Phakic IOL');
  addProc(rows, 'ICL Toric (per eye)', 'Phakic IOL with toric correction');
  addProc(rows, 'Refractive Lens Exchange (RLE)', 'Clear lens extraction for refractive purpose');

  // ── 4. GLAUCOMA SURGERIES & LASERS ────────────────────────────────────────
  rows.push(['__SECTION__', '4. GLAUCOMA SURGERIES & LASERS']);
  addProc(rows, 'Trabeculectomy', 'Filtration surgery for glaucoma');
  addProc(rows, 'Ahmed Glaucoma Valve Implantation', 'Glaucoma drainage device — tube shunt');
  addProc(rows, 'Baerveldt / Molteno Implant', 'Other glaucoma drainage device');
  addProc(rows, 'Minimally Invasive Glaucoma Surgery (MIGS)', 'iStent, Hydrus, Kahook etc.');
  addProc(rows, 'Cyclodestruction (Cyclodiode Laser)', 'Cyclophotocoagulation for end-stage glaucoma');
  addProc(rows, 'YAG Laser Iridotomy', 'Peripheral iridotomy for narrow angle / AACG');
  addProc(rows, 'Laser Peripheral Iridoplasty (LPI)', 'Argon laser iridoplasty');
  addProc(rows, 'SLT (Selective Laser Trabeculoplasty)', 'Laser for open-angle glaucoma');
  addProc(rows, 'ALT (Argon Laser Trabeculoplasty)', 'Older trabeculoplasty method');

  // ── 5. RETINA SURGERIES & PROCEDURES ──────────────────────────────────────
  rows.push(['__SECTION__', '5. RETINA SURGERIES & PROCEDURES']);
  addProc(rows, 'Vitrectomy (PPV) — Anterior', 'Anterior vitrectomy');
  addProc(rows, 'Vitrectomy (PPV) — Posterior 23G', '23-gauge posterior vitrectomy');
  addProc(rows, 'Vitrectomy (PPV) — Posterior 25G', '25-gauge micro-incision vitrectomy');
  addProc(rows, 'Vitrectomy (PPV) — Posterior 27G', '27-gauge vitrectomy');
  addProc(rows, 'Scleral Buckle', 'For rhegmatogenous retinal detachment');
  addProc(rows, 'Pneumatic Retinopexy', 'Gas bubble for retinal detachment');
  addProc(rows, 'Intravitreal Anti-VEGF (Avastin)', 'Bevacizumab — per injection');
  addProc(rows, 'Intravitreal Anti-VEGF (Lucentis)', 'Ranibizumab — per injection');
  addProc(rows, 'Intravitreal Anti-VEGF (Eylea)', 'Aflibercept — per injection');
  addProc(rows, 'Intravitreal Anti-VEGF (Beovu)', 'Brolucizumab — per injection');
  addProc(rows, 'Intravitreal Triamcinolone (IVTA)', 'Steroid injection for macular oedema');
  addProc(rows, 'Ozurdex Implant (Dexamethasone)', 'Sustained-release steroid implant');
  addProc(rows, 'Laser Photocoagulation (PRP)', 'Panretinal photocoagulation for DR/PDR');
  addProc(rows, 'Focal / Grid Laser (Macular)', 'Focal/grid laser for macular oedema');
  addProc(rows, 'Cryotherapy', 'Retinal cryotherapy');
  addProc(rows, 'Retinopathy of Prematurity (ROP) Laser', 'Laser for ROP in neonates');
  addProc(rows, 'Silicone Oil Injection', 'Tamponade for retinal detachment');
  addProc(rows, 'Silicone Oil Removal', '');
  addProc(rows, 'Gas Injection (C3F8 / SF6)', 'Intraocular gas tamponade');
  addProc(rows, 'Macular Hole Surgery', 'ILM peel + gas tamponade');
  addProc(rows, 'Epiretinal Membrane (ERM) Peeling', '');
  addProc(rows, 'Subretinal / Sub-ILM Haemorrhage Drainage', '');

  // ── 6. CORNEA ─────────────────────────────────────────────────────────────
  rows.push(['__SECTION__', '6. CORNEA PROCEDURES']);
  addProc(rows, 'Penetrating Keratoplasty (PKP)', 'Full thickness corneal transplant');
  addProc(rows, 'DALK (Deep Anterior Lamellar Keratoplasty)', 'Partial thickness anterior cornea transplant');
  addProc(rows, 'DSEK (Descemet Stripping Endothelial Keratoplasty)', 'Posterior lamellar graft');
  addProc(rows, 'DMEK (Descemet Membrane Endothelial Keratoplasty)', 'Descemet membrane only — premium');
  addProc(rows, 'Pterygium Excision (bare sclera)', '');
  addProc(rows, 'Pterygium Excision + Conjunctival Autograft', 'Gold standard — lower recurrence');
  addProc(rows, 'Amniotic Membrane Transplant (AMT)', 'For persistent epithelial defect / chemical burns');
  addProc(rows, 'Cross-Linking (CXL) — Standard (3 mW)', 'Corneal collagen cross-linking for keratoconus');
  addProc(rows, 'Cross-Linking (CXL) — Accelerated (9 mW)', 'Faster CXL protocol');
  addProc(rows, 'Cross-Linking (CXL) — Transepithelial (TXL)', 'Epithelium-on cross-linking');
  addProc(rows, 'Corneal Tattooing', 'Cosmetic / optical corneal tattoo');
  addProc(rows, 'Intrastromal Corneal Ring Segment (ICRS)', 'Keraring / Intacs for keratoconus');
  addProc(rows, 'Foreign Body Removal (Corneal)', '');

  // ── 7. OCULOPLASTY ────────────────────────────────────────────────────────
  rows.push(['__SECTION__', '7. OCULOPLASTY & ADNEXA']);
  addProc(rows, 'Chalazion Surgery', 'Incision and curettage of chalazion');
  addProc(rows, 'Stye / Hordeolum Incision', '');
  addProc(rows, 'Entropion Correction', 'Inward turning of eyelid');
  addProc(rows, 'Ectropion Correction', 'Outward turning of eyelid');
  addProc(rows, 'Ptosis Correction (Levator Resection)', 'Drooping eyelid correction');
  addProc(rows, 'Ptosis Correction (Brow Suspension)', 'Frontalis sling for severe ptosis');
  addProc(rows, 'Blepharoplasty (Upper Eyelid)', 'Cosmetic / functional eyelid surgery');
  addProc(rows, 'Blepharoplasty (Lower Eyelid)', '');
  addProc(rows, 'DCR (Dacryocystorhinostomy)', 'Surgery for blocked tear duct');
  addProc(rows, 'DCR (Endo-DCR — Endoscopic)', 'Endoscopic approach for blocked tear duct');
  addProc(rows, 'Punctal Plug Insertion', 'For dry eye / punctal stenosis');
  addProc(rows, 'Lacrimal Syringing & Probing', 'For blocked nasolacrimal duct in infants');
  addProc(rows, 'Orbital Decompression', 'For thyroid eye disease exophthalmos');
  addProc(rows, 'Enucleation', 'Removal of eyeball');
  addProc(rows, 'Evisceration', 'Removal of intraocular contents');
  addProc(rows, 'Exenteration', 'Removal of orbital contents — for malignancy');
  addProc(rows, 'Orbital Implant', 'Prosthetic sphere after enucleation/evisceration');
  addProc(rows, 'Botox Injection (Oculoplastic)', 'Per session — blepharospasm / cosmetic');
  addProc(rows, 'Eyelid Mass / Tumour Excision', '');
  addProc(rows, 'Eyelid Laceration Repair', '');
  addProc(rows, 'Canthotomy / Cantholysis', 'Orbital compartment syndrome release');

  // ── 8. SQUINT & PAEDIATRIC ────────────────────────────────────────────────
  rows.push(['__SECTION__', '8. SQUINT & PAEDIATRIC OPHTHALMOLOGY']);
  addProc(rows, 'Squint Surgery — 1 Muscle', 'Recession / resection of one extraocular muscle');
  addProc(rows, 'Squint Surgery — 2 Muscles', 'Most common — two muscle correction');
  addProc(rows, 'Squint Surgery — 3+ Muscles', 'Complex strabismus correction');
  addProc(rows, 'Botulinum Toxin for Squint', 'Chemodenervation of extraocular muscle');
  addProc(rows, 'Probing (Congenital NLD Obstruction)', 'Nasolacrimal duct probing in children');
  addProc(rows, 'Amblyopia Treatment (Patching Protocol)', 'Supervised occlusion therapy');
  addProc(rows, 'Examination Under Anaesthesia (EUA)', 'Paediatric eye examination under GA');
  addProc(rows, 'Glaucoma Surgery — Paediatric (Trabeculotomy)', 'For congenital glaucoma');
  addProc(rows, 'Cyclopentolate Refraction (under cycloplegia)', '');

  // ── 9. YAG LASER PROCEDURES ───────────────────────────────────────────────
  rows.push(['__SECTION__', '9. YAG LASER PROCEDURES']);
  addProc(rows, 'YAG Laser Capsulotomy', 'Posterior capsule opacification (PCO) — after cataract');
  addProc(rows, 'YAG Laser Iridotomy (Separate from Glaucoma)', 'Angle closure / prophylactic');
  addProc(rows, 'YAG Laser Vitreolysis', 'Floaters treatment');
  addProc(rows, 'YAG Laser Membranectomy', 'YAG for anterior vitreous membranes');
  addProc(rows, 'Laser Suture Lysis (Post-Trabeculectomy)', '');

  // ── 10. MINOR PROCEDURES & OPD TREATMENTS ────────────────────────────────
  rows.push(['__SECTION__', '10. MINOR PROCEDURES & OPD TREATMENTS']);
  addProc(rows, 'Gonioscopy', 'Angle assessment — usually OPD procedure');
  addProc(rows, 'Intravitreal Injection (Minor/Procedure Room)', 'If not included in Anti-VEGF price');
  addProc(rows, 'Sub-Tenon Injection', 'Periocular steroid injection');
  addProc(rows, 'Subconjunctival Injection', '');
  addProc(rows, 'Intraocular Pressure Check (Goldmann)', 'Applanation tonometry');
  addProc(rows, 'Corneal Scraping (for culture/sensitivity)', 'Microbial keratitis');
  addProc(rows, 'Suture Removal (Corneal / Cataract)', '');
  addProc(rows, 'Contact Lens Fitting & Prescription', '');
  addProc(rows, 'Bandage Contact Lens Application', '');
  addProc(rows, 'Epilation (Trichiasis)', 'Removal of aberrant eyelashes');
  addProc(rows, 'Abscess Drainage (Periocular)', '');
  addProc(rows, 'Enucleation / Prosthetic Eye Fitting', '');
  addProc(rows, 'Fluorescein Staining Examination', '');
  addProc(rows, 'Dry Eye Punctal Cautery', '');
  addProc(rows, 'IPL (Intense Pulsed Light) — Dry Eye', 'Per session');
  addProc(rows, 'LipiFlow / Thermal Pulsation', 'Meibomian gland treatment');
  addProc(rows, 'Low Vision Assessment & Aids Fitting', '');

  // ── 11. DIAGNOSTIC TESTS ─────────────────────────────────────────────────
  rows.push(['__SECTION__', '11. DIAGNOSTIC TESTS']);
  const DIAGS_FULL = [
    ['Biometry (IOL Master / Lenstar)', ''],
    ['Keratometry / Manual K-reading', ''],
    ['Corneal Topography (Placido)', 'Corneal surface mapping'],
    ['Corneal Tomography (Pentacam / Orbscan)', 'Scheimpflug-based — for LASIK screening and keratoconus'],
    ['Pachymetry (Ultrasound / Optical)', 'Corneal thickness measurement'],
    ['Specular Microscopy (Endothelial Cell Count)', 'Corneal endothelium assessment before cataract surgery'],
    ['OCT — Anterior Segment', 'Cornea / angle imaging'],
    ['OCT — Posterior Segment (Macula/Disc)', 'Retina / optic nerve OCT'],
    ['OCT — Wide-Field (Swept Source)', 'Large retinal area'],
    ['OCT Angiography (OCTA)', 'Non-invasive retinal vasculature imaging'],
    ['Fundus Fluorescein Angiography (FFA)', 'Intravenous dye study for retinal vasculature'],
    ['Indocyanine Green Angiography (ICGA)', 'Choroidal vasculature imaging'],
    ['Fundus Photography (Non-Mydriatic)', 'Colour fundus photo without dilation'],
    ['Fundus Photography (Mydriatic / ETDRS)', 'Standard dilated fundus photo'],
    ['Fundus Photography (Wide-Field / Optos)', 'Ultra-wide field retinal imaging'],
    ['Visual Field Test (Humphrey / Octopus)', 'Standard automated perimetry (SAP)'],
    ['Microperimetry (MP)', 'Macular fixation and sensitivity testing'],
    ['B-Scan Ultrasonography', 'For media opacity / posterior segment'],
    ['A-Scan Ultrasonography', 'Axial length for IOL power calculation'],
    ['UBM (Ultrasound Biomicroscopy)', 'Anterior segment high-freq ultrasound'],
    ['Electroretinography (ERG)', 'Electrophysiology — retinal function'],
    ['Visual Evoked Potential (VEP)', 'Electrophysiology — optic nerve / visual pathway'],
    ['Electrooculography (EOG)', 'RPE function testing'],
    ['Multifocal ERG (mfERG)', 'Localised retinal function mapping'],
    ['Cataract Work-up (complete pre-op)', 'Biometry + specular microscopy + dilated fundus'],
    ['LASIK Work-up (complete)', 'Topography + pachymetry + wavefront + biometry'],
    ['Gonioscopy (OPD procedure)', 'Angle assessment for glaucoma'],
    ['Amsler Grid Test', 'Central visual field / macular test'],
    ['Colour Vision Test (Ishihara)', ''],
    ['Contrast Sensitivity Test', ''],
    ['Tear Film Assessment (Schirmer\'s / TBUT)', 'Dry eye evaluation'],
    ['MG (Meibography)', 'Meibomian gland imaging'],
    ['Corneal Confocal Microscopy', 'In-vivo cellular imaging of cornea'],
    ['Pupillometry', 'Dynamic pupil assessment for IOL selection'],
    ['Wavefront Aberrometry', 'Higher-order aberrations — for premium IOL / LASIK'],
    ['Aberrometry (iTrace / OPD-Scan)', 'Optical path difference scanning'],
    ['Fluorescein Angiography — Anterior Segment', 'Iris / angle vasculature'],
    ['Rose Bengal / Lissamine Green Staining', 'Ocular surface staining'],
    ['Schirmer\'s Test I & II', 'Tear production measurement'],
    ['Exophthalmometry (Hertel)', 'Proptosis measurement'],
    ['Synoptophore / Diplopia Testing', 'Binocular vision assessment'],
    ['Cover-Uncover / Prism Cover Test', 'Squint assessment'],
    ['ERG / VEP (combined)', ''],
    ['Photo Documentation (Slit Lamp / External)', ''],
    ['Intraocular Pressure (Goldman / Air-puff)', ''],
  ];
  DIAGS_FULL.forEach(([name, notes]) => addDiag(rows, name, notes));

  // ── 12. IOL & IMPLANT COST TRACKING ──────────────────────────────────────
  rows.push(['__SECTION__', '12. IOL & IMPLANT COST (admin fills cost price — for internal margin tracking)']);
  const IOLS = [
    // Monofocal IOLs
    ['IOL — Aurolab AUROFLEX Hydrophilic (Indian)', 'Aurolab, Madurai — affordable hydrophilic'],
    ['IOL — Appasamy ReSure Monofocal (Indian)', 'Indian company'],
    ['IOL — Alcon AcrySof IQ (USA)', 'Most widely used monofocal globally'],
    ['IOL — J&J Tecnis 1-Piece ZCB00 (USA)', 'High-quality PMMA / acrylic'],
    ['IOL — B&L enVista (USA)', 'Glistening-free hydrophobic'],
    ['IOL — Hoya iSert / Vivinex 1-Piece (Japan)', 'Blue light filter'],
    ['IOL — Zeiss CT LUCIA (Germany)', 'Hydrophobic monofocal'],
    ['IOL — Rayner C-flex / RayOne (UK)', 'Hydrophilic / hydrophobic'],
    ['IOL — Carl Zeiss Aspira-aA (Germany)', ''],
    // Monofocal Toric IOLs
    ['IOL — Alcon AcrySof IQ Toric (USA)', 'T2–T9 cylinder range'],
    ['IOL — J&J Tecnis Toric II (USA)', ''],
    ['IOL — B&L Trulign Toric (USA)', ''],
    ['IOL — Hoya NANEX Toric (Japan)', ''],
    ['IOL — Zeiss CT TORBI (Germany)', ''],
    ['IOL — Rayner T-flex Toric (UK)', ''],
    // EDOF IOLs
    ['IOL — Alcon AcrySof IQ Vivity EDOF (USA)', 'Non-diffractive EDOF'],
    ['IOL — J&J Tecnis Symfony EDOF (USA)', 'Achromatic diffractive EDOF'],
    ['IOL — J&J Tecnis Synergy (USA)', 'Full-range EDOF/multifocal hybrid'],
    ['IOL — Hoya Vivinex iSert XY1A (Japan)', 'EDOF-like extended range'],
    ['IOL — Zeiss AT LARA EDOF (Germany)', ''],
    ['IOL — BVI Mini Well Ready (Belgium)', 'EDOF spherical aberration-based'],
    ['IOL — Physiol FineVision Micro F (Belgium)', 'Trifocal/EDOF hybrid'],
    // EDOF Toric IOLs
    ['IOL — Alcon AcrySof IQ Vivity Toric EDOF (USA)', ''],
    ['IOL — J&J Tecnis Symfony Toric EDOF (USA)', ''],
    ['IOL — Zeiss AT LARA Toric (Germany)', ''],
    // Trifocal IOLs
    ['IOL — Alcon AcrySof IQ PanOptix (USA)', 'Most popular trifocal globally — 50 cm near'],
    ['IOL — J&J Tecnis Synergy Trifocal (USA)', 'Continuous range — hybrid'],
    ['IOL — Zeiss AT LISA tri 839MP (Germany)', 'European standard trifocal'],
    ['IOL — BVI PhysIOL FineVision POD F (Belgium)', 'Diffractive trifocal'],
    ['IOL — Hanita 1stQ ApoLink Trifocal (Israel)', 'Refractive-diffractive trifocal'],
    ['IOL — VSY Biotechnology AT Lisa tri (Turkey)', ''],
    ['IOL — Medicontur Liberty 677MY (Hungary)', 'Quadrifocal / trifocal'],
    ['IOL — Oculentis / LENSTEC Trifocal (Germany)', ''],
    ['IOL — Rayner RayOne Trifocal (UK)', ''],
    // Trifocal Toric IOLs
    ['IOL — Alcon AcrySof IQ PanOptix Toric (USA)', ''],
    ['IOL — Zeiss AT LISA tri Toric (Germany)', ''],
    ['IOL — BVI PhysIOL FineVision Toric (Belgium)', ''],
    ['IOL — J&J Tecnis Synergy Toric (USA)', ''],
    ['IOL — Hanita 1stQ Trifocal Toric (Israel)', ''],
    // Light Adjustable / Special
    ['IOL — RxSight Light Adjustable Lens (LAL) (USA)', 'Post-op power adjustment with UV light'],
    ['IOL — IC-8 (ApoLink Small Aperture) (USA)', 'Small aperture EDOF for high astigmatism / irregular cornea'],
    ['IOL — Sulcoflex Add-On IOL (Rayner)', 'Supplementary piggyback IOL'],
    // OVD / Viscoelastic
    ['OVD — Healon (J&J) Sodium Hyaluronate 1%', 'Standard cohesive OVD'],
    ['OVD — ProVisc (Alcon) Sodium Hyaluronate 1%', ''],
    ['OVD — Viscoat (Alcon) HA + Chondroitin Sulfate', 'Dispersive OVD — corneal protection'],
    ['OVD — DisCoVisc (Alcon) — Visco-adaptive', 'Dual-function OVD'],
    ['OVD — Amvisc (Bausch & Lomb) Sodium Hyaluronate', ''],
    ['OVD — Appavisc HPMC (Appasamy/India)', 'Hydroxypropyl methylcellulose — affordable'],
    ['OVD — Eyevisc HPMC (Aurolab/India)', 'Indian HPMC viscoelastic'],
    ['OVD — BioVisc (Sodium Hyaluronate) Indian brand', ''],
    // Anti-VEGF (for inventory/cost tracking)
    ['Avastin (Bevacizumab — 1.25 mg vial)', 'Off-label — most affordable IVFA'],
    ['Lucentis (Ranibizumab — 0.5 mg vial)', 'Approved for AMD, DME, RVO'],
    ['Eylea (Aflibercept — 2 mg vial)', 'Approved for AMD, DME, RVO, DR'],
    ['Eylea HD (Aflibercept 8 mg)', 'High dose — less frequent dosing'],
    ['Beovu (Brolucizumab — 6 mg)', 'Novartis — 12-weekly dosing'],
    ['Vabysmo (Faricimab)', 'Dual-target anti-VEGF / Ang-2'],
    ['Ozurdex (Dexamethasone Implant 0.7 mg)', 'Sustained-release steroid implant'],
    ['Iluvien (Fluocinolone Acetonide Implant)', 'Long-acting steroid implant for DME'],
    ['Triamcinolone Acetonide (IVTA — 4 mg/0.1 mL)', 'Off-label steroid for macular oedema'],
    // Phaco consumables
    ['Phaco Handpiece Tip Sleeve (disposable)', 'Per case'],
    ['BSS (Balanced Salt Solution 500 mL)', 'Per bottle — Alcon BSS Plus / generic'],
    ['Trypan Blue 0.06% (VisionBlue / generic)', 'Capsule staining per vial'],
    ['5-Fluorouracil (5-FU) 25 mg/mL', 'Post-trabeculectomy subconjunctival injection'],
    ['Mitomycin-C (MMC 0.2 mg/mL)', 'Anti-fibrotic for filtering surgery'],
    ['Intracameral Antibiotics (Cefuroxime / Moxifloxacin)', 'Per vial / dose'],
    ['Intracameral Adrenaline / Mydriatic', 'Pupil dilation during phaco'],
    ['Capsular Tension Ring (CTR)', 'For zonular weakness during cataract surgery'],
    ['Iris Retractor / Pupil Expansion Ring', 'For small pupil phaco'],
    ['Ahmed Glaucoma Valve FP7 (Acrylate Plate)', 'Per device'],
    ['Baerveldt 250 / 350 mm² Implant', 'Per device'],
    ['iStent (Glaukos) MIGS Device', 'Per device — ab-interno trabecular device'],
    ['Silicone Oil 1000 cs', 'Per vial — retinal tamponade'],
    ['Silicone Oil 5000 cs', 'Per vial — heavy silicone oil'],
    ['C3F8 Gas (Perfluoropropane)', 'Long-acting gas tamponade'],
    ['SF6 Gas (Sulphur Hexafluoride)', 'Medium-acting gas tamponade'],
    ['Perfluorocarbon Liquid (PFCL)', 'Heavy liquid for retinal detachment surgery'],
    ['Retinal Cryotherapy Probe (single-use)', ''],
    ['Vitrectomy Cassette / Pack (25G/27G)', 'Per case disposable'],
  ];
  IOLS.forEach(([name, notes]) => {
    rows.push([name, 'IOL/Implant Cost', 'Cost Price', '', notes + ' — fill cost per unit for margin tracking']);
  });

  // ── BATCH WRITE — single setValues call, then format ──────────────────────
  // Build two parallel arrays: allValues (2D for setValues) and sectionRowNums
  var headerRow = [['Item / Procedure', 'Category', 'Payment Type', 'Rate (₹) — fill in', 'Notes']];
  var allValues = [headerRow[0]];
  var sectionRowNums = [];  // 1-based row numbers that are section headers
  var dataRowNums    = [];  // data rows (for Rate column highlight)

  rows.forEach(function(r, i) {
    var rowNum = i + 2; // +2 because row 1 is the header
    if (r[0] === '__SECTION__') {
      allValues.push(['▸  ' + r[1].toUpperCase(), '', '', '', '']);
      sectionRowNums.push(rowNum);
    } else {
      allValues.push(r);
      dataRowNums.push(rowNum);
    }
  });

  // Write ALL rows in one call
  sheet.getRange(1, 1, allValues.length, 5).setValues(allValues);

  // Format header
  sheet.getRange(1, 1, 1, 5)
    .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF').setFontSize(11);

  // Format section headers (batch by building ranges)
  if (sectionRowNums.length > 0) {
    sectionRowNums.forEach(function(rowNum) {
      sheet.getRange(rowNum, 1, 1, 5)
        .setBackground('#FBEFD6').setFontColor('#7A5D1E').setFontWeight('bold').setFontSize(10);
    });
  }

  // Highlight the Rate column for data rows (do first 200 only to stay fast;
  // rest inherit the default white background which is fine)
  var rateHighlight = Math.min(dataRowNums.length, 200);
  for (var i = 0; i < rateHighlight; i++) {
    sheet.getRange(dataRowNums[i], 4).setBackground('#FFFDF5');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 320);

  sheet.getRange(1, 1).setNote(
    'PRICE MASTER — Admin fills Rate (₹) column once.\n' +
    'Sections 1-11: Patient-facing rates (what hospital charges).\n' +
    'Section 12: Internal cost prices (what hospital pays to supplier) — for margin analysis.\n' +
    'Staff never see this sheet. All revenue auto-calculated from these rates.\n' +
    'Key: "Item|Payment Type" → Rate.'
  );
}

// ============================================================
//  INVENTORY MASTER — comprehensive eye hospital inventory
//  BATCH WRITE VERSION — all rows written in one setValues call.
//
//  IOLs are listed by BRAND + TYPE (not per individual diopter).
//  Admin duplicates rows for each diopter power they stock.
//  A note in column G explains this to admin.
//  This keeps the sheet to ~200 rows (fast) vs ~2000 rows (timeout).
// ============================================================
function buildInventoryMasterSheet_(ss) {
  var sheet = ss.getSheetByName('InventoryMaster');
  if (!sheet) sheet = ss.insertSheet('InventoryMaster');
  sheet.clear(); sheet.clearFormats();

  var SEC = '__S__'; // section marker

  // All rows: [name, category, sub-category, unit, openingStock, reorderLevel, notes]
  // SEC rows: [SEC, section title]
  var items = [
    // Header
    ['Item Name', 'Category', 'Sub-Category', 'Unit', 'Opening Stock', 'Reorder Level', 'Notes / Instructions'],

    [SEC, 'IOL — MONOFOCAL (list each diopter power you stock as a separate row — duplicate this row and change diopter in the name)'],
    ['Aurolab AUROFLEX Monofocal +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'Indian brand — hydrophilic acrylic. Add rows for each power.'],
    ['Alcon AcrySof IQ SN60WF +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'Hydrophobic acrylic. Range +10.0 to +34.0 D in 0.5 steps.'],
    ['J&J Tecnis ZCB00 +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'Hydrophobic acrylic. Range +5.0 to +34.0 D.'],
    ['Hoya iSert PC-60AD +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'Hydrophobic acrylic, blue light filter.'],
    ['Zeiss CT LUCIA 601P +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'German monofocal hydrophobic.'],
    ['Rayner RayOne Hydrophobic +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 1, 'UK — hydrophobic.'],
    ['B&L enVista MX60E +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 1, 'Glistening-free hydrophobic.'],
    ['Appasamy Monofocal +21.0 D', 'IOL', 'Monofocal', 'pcs', 0, 2, 'Indian brand. Various powers.'],

    [SEC, 'IOL — MONOFOCAL TORIC (one row per cylinder power per brand; duplicate for each diopter+cylinder combination you stock)'],
    ['Alcon AcrySof IQ Toric T3 (1.50D cyl) +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, 'T2=0.75D, T3=1.50D, T4=2.25D, T5=2.75D, T6=3.00D, T7=3.50D, T8=4.00D, T9=4.50D cylinder'],
    ['Alcon AcrySof IQ Toric T5 (2.75D cyl) +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, 'Mid-range cylinder correction'],
    ['Alcon AcrySof IQ Toric T7 (3.50D cyl) +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, 'High cylinder correction'],
    ['J&J Tecnis Toric II ZCT150 +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, '1.50D cyl — add rows for other cylinders'],
    ['J&J Tecnis Toric II ZCT225 +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, '2.25D cyl'],
    ['Zeiss CT TORBI +21.0 D (Cyl 1.00D)', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, 'German toric'],
    ['Hoya NANEX Toric +21.0 D', 'IOL', 'Monofocal Toric', 'pcs', 0, 1, ''],

    [SEC, 'IOL — EDOF (Extended Depth of Focus)'],
    ['Alcon AcrySof IQ Vivity DFT015 +21.0 D', 'IOL', 'EDOF', 'pcs', 0, 1, 'Non-diffractive EDOF — best for patients worried about halos'],
    ['J&J Tecnis Symfony ZXR00 +21.0 D', 'IOL', 'EDOF', 'pcs', 0, 1, 'Achromatic diffractive EDOF — excellent distance+intermediate'],
    ['J&J Tecnis Synergy ZFR00V +21.0 D', 'IOL', 'EDOF', 'pcs', 0, 1, 'EDOF/multifocal hybrid — continuous range'],
    ['Zeiss AT LARA 829MP +21.0 D', 'IOL', 'EDOF', 'pcs', 0, 1, 'German EDOF'],
    ['BVI Mini Well Ready +21.0 D', 'IOL', 'EDOF', 'pcs', 0, 1, 'Belgian EDOF — spherical aberration based'],

    [SEC, 'IOL — EDOF TORIC'],
    ['Alcon AcrySof IQ Vivity Toric DFT315 +21.0 D (Cyl 1.00D)', 'IOL', 'EDOF Toric', 'pcs', 0, 1, 'DFT315/415/515 for different cylinders'],
    ['J&J Tecnis Symfony Toric ZXT150 +21.0 D', 'IOL', 'EDOF Toric', 'pcs', 0, 1, '1.50D cyl — add other cylinder rows'],
    ['Zeiss AT LARA Toric +21.0 D', 'IOL', 'EDOF Toric', 'pcs', 0, 1, ''],

    [SEC, 'IOL — TRIFOCAL (add rows for each power you stock)'],
    ['Alcon AcrySof IQ PanOptix TFNT00 +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'Most popular trifocal worldwide — near 40cm'],
    ['Zeiss AT LISA tri 839MP +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'German standard trifocal'],
    ['BVI PhysIOL FineVision POD F +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'Belgian diffractive trifocal'],
    ['J&J Tecnis Synergy ZFR00V (Trifocal mode) +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'Continuous range of vision'],
    ['Hanita 1stQ ApoLink Trifocal +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'Israeli refractive-diffractive'],
    ['Rayner RayOne Trifocal +21.0 D', 'IOL', 'Trifocal', 'pcs', 0, 1, 'UK trifocal'],

    [SEC, 'IOL — TRIFOCAL TORIC'],
    ['Alcon AcrySof IQ PanOptix Toric TFNT30 +21.0 D (Cyl 1.00D)', 'IOL', 'Trifocal Toric', 'pcs', 0, 1, 'TFNT30/40/50/60 for cylinders 1.00/1.50/2.25/3.00D'],
    ['Zeiss AT LISA tri Toric 939M +21.0 D', 'IOL', 'Trifocal Toric', 'pcs', 0, 1, ''],
    ['BVI PhysIOL FineVision Toric POD FT +21.0 D', 'IOL', 'Trifocal Toric', 'pcs', 0, 1, ''],
    ['J&J Tecnis Synergy Toric ZFR300 +21.0 D (Cyl 2.25D)', 'IOL', 'Trifocal Toric', 'pcs', 0, 1, ''],

    [SEC, 'IOL — SPECIAL & OTHER'],
    ['RxSight Light Adjustable Lens (LAL) — power TBD', 'IOL', 'Special', 'pcs', 0, 1, 'Post-op UV adjustment — power set after implantation'],
    ['IC-8 Small Aperture IOL (AcuFocus) — any power', 'IOL', 'Special', 'pcs', 0, 1, 'For irregular cornea / high astigmatism'],
    ['Rayner Sulcoflex Add-On IOL (supplementary)', 'IOL', 'Add-On', 'pcs', 0, 1, 'Piggyback IOL for aphakic eyes'],
    ['Capsular Tension Ring (CTR) — Standard', 'IOL Accessory', 'CTR', 'pcs', 0, 5, 'Zonular support'],
    ['Capsular Tension Ring — Cionni Modified (with hooks)', 'IOL Accessory', 'CTR', 'pcs', 0, 2, 'Severe zonular dialysis'],
    ['Malyugin Ring (pupil expansion)', 'IOL Accessory', 'Pupil Dilator', 'pcs', 0, 5, ''],
    ['Iris Retractor Set (4-hook)', 'IOL Accessory', 'Pupil Dilator', 'set', 0, 5, ''],
    ['IOL Injector / Cartridge (disposable)', 'IOL Accessory', 'Injector', 'pcs', 0, 20, 'Single-use for foldable IOL insertion'],
    ['Ahmed Glaucoma Valve FP7 (acrylate plate)', 'Glaucoma Implant', 'Drainage Device', 'pcs', 0, 2, 'For refractory glaucoma'],
    ['Baerveldt 350 mm² Implant', 'Glaucoma Implant', 'Drainage Device', 'pcs', 0, 1, ''],
    ['iStent inject W (Glaukos) — MIGS', 'Glaucoma Implant', 'MIGS', 'pcs', 0, 2, 'Ab-interno trabecular bypass'],
    ['XEN Gel Stent (AbbVie/Allergan)', 'Glaucoma Implant', 'MIGS', 'pcs', 0, 2, 'Subconjunctival drainage'],

    [SEC, 'VISCOELASTIC / OVD (Ophthalmic Viscosurgical Devices)'],
    ['Healon (J&J) — Sodium Hyaluronate 1% 0.85 mL', 'OVD', 'Cohesive', 'syringe', 0, 10, 'Standard cohesive OVD — most widely used'],
    ['ProVisc (Alcon) — Sodium Hyaluronate 1% 0.6 mL', 'OVD', 'Cohesive', 'syringe', 0, 10, ''],
    ['Viscoat (Alcon) — HA+Chondroitin Sulfate 0.5 mL', 'OVD', 'Dispersive', 'syringe', 0, 10, 'Corneal protection in complex cases'],
    ['DisCoVisc (Alcon) — Visco-adaptive 0.5 mL', 'OVD', 'Dual-function', 'syringe', 0, 8, 'Acts as both cohesive and dispersive'],
    ['Healon5 (J&J) — HA 2.3% 0.6 mL', 'OVD', 'High-viscosity', 'syringe', 0, 5, 'For small pupil / hard cataracts'],
    ['Amvisc Plus (B&L) — HA 1.6% 0.8 mL', 'OVD', 'Cohesive', 'syringe', 0, 5, ''],
    ['Appavisc HPMC 2% 0.8 mL (Appasamy, India)', 'OVD', 'HPMC', 'syringe', 0, 15, 'Affordable Indian OVD'],
    ['Eyevisc HPMC 2% (Aurolab, India)', 'OVD', 'HPMC', 'syringe', 0, 15, ''],

    [SEC, 'ANTI-VEGF & INTRAVITREAL AGENTS'],
    ['Avastin (Bevacizumab) 100 mg/4 mL vial', 'Anti-VEGF', 'Off-label', 'vial', 0, 2, 'Off-label use — most affordable. Refrigerate 2-8°C.'],
    ['Lucentis (Ranibizumab) 0.5 mg vial', 'Anti-VEGF', 'Approved', 'vial', 0, 2, 'Approved for AMD / DME / RVO'],
    ['Eylea (Aflibercept) 2 mg vial', 'Anti-VEGF', 'Approved', 'vial', 0, 2, 'Approved for AMD / DME / DR / RVO'],
    ['Eylea HD (Aflibercept 8 mg)', 'Anti-VEGF', 'Approved', 'vial', 0, 2, 'High-dose — fewer injections needed'],
    ['Beovu (Brolucizumab) 6 mg vial', 'Anti-VEGF', 'Approved', 'vial', 0, 2, '12-weekly dosing — AMD'],
    ['Vabysmo (Faricimab) 6 mg vial', 'Anti-VEGF', 'Approved', 'vial', 0, 2, 'Dual-target VEGF-A + Ang-2'],
    ['Ozurdex (Dexamethasone 0.7 mg implant)', 'Intravitreal', 'Steroid Implant', 'implant', 0, 2, 'Sustained-release for RVO macular oedema / DME'],
    ['Triamcinolone Acetonide 40 mg/mL (IVTA)', 'Intravitreal', 'Steroid', 'vial', 0, 3, 'Off-label for macular oedema'],
    ['5-Fluorouracil (5-FU) 25 mg/mL', 'Intravitreal', 'Anti-fibrotic', 'vial', 0, 5, 'Post-trabeculectomy bleb management'],
    ['Mitomycin-C (MMC) — ophthalmic grade', 'Intravitreal', 'Anti-fibrotic', 'vial', 0, 5, 'Filtration surgery / pterygium'],

    [SEC, 'PHACO & OT SURGICAL CONSUMABLES'],
    ['BSS Plus (Alcon) 500 mL', 'Phaco Consumable', 'Irrigation', 'bottle', 0, 10, 'Balanced salt solution for irrigation'],
    ['BSS Standard 500 mL (generic)', 'Phaco Consumable', 'Irrigation', 'bottle', 0, 10, ''],
    ['Trypan Blue 0.06% — VisionBlue / generic', 'Phaco Consumable', 'Dye', 'vial', 0, 20, 'Capsule staining for continuous curvilinear capsulorrhexis'],
    ['Phaco Tip Sleeve / Silicone Sleeve (disposable)', 'Phaco Consumable', 'Handpiece', 'pcs', 0, 20, 'Per case'],
    ['Phaco Cassette / Tubing Set (Alcon/AMO)', 'Phaco Consumable', 'Tubing', 'set', 0, 10, 'Single-use tubing/cassette'],
    ['Vitrectomy Cassette 23G', 'Vitrectomy', 'Cassette', 'set', 0, 5, 'Per case — 23G vitrectomy'],
    ['Vitrectomy Cassette 25G', 'Vitrectomy', 'Cassette', 'set', 0, 5, 'Per case — 25G vitrectomy'],
    ['Vitrectomy Cassette 27G', 'Vitrectomy', 'Cassette', 'set', 0, 5, 'Per case — 27G micro-incision'],
    ['Endoilluminator / Light Pipe (single-use)', 'Vitrectomy', 'Illumination', 'pcs', 0, 5, ''],
    ['Perfluorocarbon Liquid (PFCL) — Perfluoron', 'Vitrectomy', 'Heavy Liquid', 'vial', 0, 3, 'Per RD surgery'],
    ['Silicone Oil 1000 centistokes (5 mL)', 'Vitrectomy', 'Tamponade', 'vial', 0, 3, 'Intraocular tamponade for RD'],
    ['Silicone Oil 5000 cs (heavy) (5 mL)', 'Vitrectomy', 'Tamponade', 'vial', 0, 2, 'For inferior RD / PVR'],
    ['C3F8 Gas (Perfluoropropane) — cylinder', 'Vitrectomy', 'Gas Tamponade', 'cylinder', 0, 1, 'Long-acting (approx 8 weeks)'],
    ['SF6 Gas (Sulphur Hexafluoride) — cylinder', 'Vitrectomy', 'Gas Tamponade', 'cylinder', 0, 1, 'Medium-acting (approx 2 weeks)'],
    ['Intracameral Cefuroxime 1 mg/0.1 mL', 'Phaco Consumable', 'Antibiotic', 'vial', 0, 20, 'Post-phaco prophylaxis (ESCRS recommendation)'],
    ['Intracameral Moxifloxacin 0.5%', 'Phaco Consumable', 'Antibiotic', 'vial', 0, 20, 'Alternative antibiotic prophylaxis'],
    ['Xylocaine (Lignocaine) 2% (retrobulbar)', 'Anaesthesia', 'Local Anaesthetic', 'vial', 0, 10, ''],
    ['Bupivacaine 0.5% + Hyaluronidase', 'Anaesthesia', 'Local Anaesthetic', 'vial', 0, 10, 'Peribulbar / sub-Tenon block'],
    ['Keratome Blade 2.2 mm (phaco incision)', 'Blade', 'Incision', 'pcs', 0, 30, ''],
    ['Keratome Blade 2.4 mm', 'Blade', 'Incision', 'pcs', 0, 20, ''],
    ['Keratome Blade 2.8 mm', 'Blade', 'Incision', 'pcs', 0, 10, ''],
    ['Sideport / Paracentesis Blade 1.0–1.2 mm', 'Blade', 'Incision', 'pcs', 0, 30, 'Secondary port'],
    ['Cystotome Needle 27G (capsulorrhexis)', 'Blade', 'Capsulorrhexis', 'pcs', 0, 30, ''],

    [SEC, 'SUTURES'],
    ['Nylon 10-0 Monofilament — Corneal (Ethilon/Alcon)', 'Suture', 'Corneal', 'pcs', 0, 20, 'Standard corneal suture'],
    ['Nylon 9-0 Monofilament', 'Suture', 'Corneal', 'pcs', 0, 10, ''],
    ['Vicryl 8-0 Absorbable (Ethicon)', 'Suture', 'Conjunctival / Muscle', 'pcs', 0, 20, 'Dissolves in 60-90 days'],
    ['Vicryl 6-0 Absorbable', 'Suture', 'Oculoplasty / Lid', 'pcs', 0, 20, 'Lid and oculoplasty procedures'],
    ['Vicryl 7-0 Absorbable', 'Suture', 'Oculoplasty', 'pcs', 0, 10, ''],
    ['Prolene 10-0 (Polypropylene)', 'Suture', 'Scleral Fixation', 'pcs', 0, 10, 'For scleral-fixated IOL, glaucoma drainage'],
    ['Silk 4-0 Black (bridle suture)', 'Suture', 'Traction', 'pcs', 0, 20, 'Traction suture for globe fixation'],
    ['Nylon 4-0 (squint / oculoplasty)', 'Suture', 'Squint / Oculoplasty', 'pcs', 0, 10, ''],
    ['Mersilk 6-0 / 5-0 (skin)', 'Suture', 'Skin', 'pcs', 0, 15, 'Eyelid skin closure'],
    ['Supramid 5-0 (frontalis sling)', 'Suture', 'Ptosis', 'pcs', 0, 5, 'Ptosis correction'],
    ['Gore-Tex CV-8 (glaucoma tube ligature)', 'Suture', 'Glaucoma', 'pcs', 0, 3, 'Ligature for glaucoma drainage tube'],

    [SEC, 'EYE DROPS & MEDICATIONS (OT + WARD STOCK)'],
    ['Tropicamide 0.8% + Phenylephrine 5% (Mydriatic)', 'Eye Drop', 'Mydriatic', 'bottle', 0, 20, 'Pre-op dilation — combine for best effect'],
    ['Cyclopentolate 1% (Cylate / Zylate)', 'Eye Drop', 'Cycloplegic', 'bottle', 0, 10, 'Cycloplegic refraction in children'],
    ['Atropine 1% (long-acting cycloplegic)', 'Eye Drop', 'Cycloplegic', 'bottle', 0, 5, 'For children / uveitis'],
    ['Proparacaine / Proxymetacaine 0.5%', 'Eye Drop', 'Topical Anaesthetic', 'bottle', 0, 10, 'Pre-procedure topical anaesthesia'],
    ['Povidone-Iodine 5% (Betadine Ophthalmic)', 'Eye Drop', 'Antiseptic', 'bottle', 0, 20, 'Pre-op conjunctival antisepsis — gold standard'],
    ['Moxifloxacin 0.5% (Vigamox / Zymox)', 'Eye Drop', 'Antibiotic', 'bottle', 0, 30, 'Post-operative antibiotic prophylaxis'],
    ['Tobramycin 0.3% (Tobrex)', 'Eye Drop', 'Antibiotic', 'bottle', 0, 20, 'Broad-spectrum antibiotic'],
    ['Chloramphenicol 0.5%', 'Eye Drop', 'Antibiotic', 'bottle', 0, 15, 'Affordable broad-spectrum'],
    ['Prednisolone Acetate 1% (Pred Forte)', 'Eye Drop', 'Steroid', 'bottle', 0, 30, 'Post-op anti-inflammatory — gold standard'],
    ['Dexamethasone 0.1%', 'Eye Drop', 'Steroid', 'bottle', 0, 20, 'Post-op steroid'],
    ['Fluorometholone 0.1% (FML)', 'Eye Drop', 'Steroid', 'bottle', 0, 15, 'Milder steroid for long-term use'],
    ['Loteprednol 0.5% (Lotemax)', 'Eye Drop', 'Steroid', 'bottle', 0, 10, 'Lower IOP risk profile'],
    ['Nepafenac 0.1% (Nevanac)', 'Eye Drop', 'NSAID', 'bottle', 0, 20, 'CMO prevention post-cataract surgery'],
    ['Ketorolac 0.5% (Acular)', 'Eye Drop', 'NSAID', 'bottle', 0, 15, 'NSAID — pain and inflammation'],
    ['Bromfenac 0.1% (Yellox)', 'Eye Drop', 'NSAID', 'bottle', 0, 10, 'Once-daily NSAID'],
    ['Timolol 0.5% (Timoptic)', 'Eye Drop', 'Glaucoma — Beta-blocker', 'bottle', 0, 15, 'IOP reduction — twice daily'],
    ['Brinzolamide 1% (Azopt)', 'Eye Drop', 'Glaucoma — CAI', 'bottle', 0, 10, 'Carbonic anhydrase inhibitor'],
    ['Dorzolamide 2% (Trusopt)', 'Eye Drop', 'Glaucoma — CAI', 'bottle', 0, 10, 'CAI — three times daily'],
    ['Brimonidine 0.2% (Alphagan P)', 'Eye Drop', 'Glaucoma — Alpha-agonist', 'bottle', 0, 10, 'IOP reduction and neuroprotection'],
    ['Latanoprost 0.005% (Xalatan)', 'Eye Drop', 'Glaucoma — Prostaglandin', 'bottle', 0, 15, 'Once-daily evening — prostaglandin'],
    ['Bimatoprost 0.01% (Lumigan)', 'Eye Drop', 'Glaucoma — Prostaglandin', 'bottle', 0, 10, 'Prostaglandin analogue'],
    ['Travoprost 0.004% (Travatan)', 'Eye Drop', 'Glaucoma — Prostaglandin', 'bottle', 0, 10, ''],
    ['Tafluprost 0.0015% (Saflutan — preservative-free)', 'Eye Drop', 'Glaucoma — Prostaglandin', 'bottle', 0, 5, 'Preservative-free option'],
    ['Pilocarpine 2% (Pilocar)', 'Eye Drop', 'Miotic', 'bottle', 0, 5, 'Narrow angle / AACG / Pilocarpine test'],
    ['CMC 0.5% (Refresh Tears)', 'Eye Drop', 'Lubricant / Dry Eye', 'bottle', 0, 30, 'Carboxymethylcellulose lubricant'],
    ['Sodium Hyaluronate 0.1–0.2% (Eyemist)', 'Eye Drop', 'Lubricant / Dry Eye', 'bottle', 0, 20, 'HA lubricant for moderate-severe dry eye'],
    ['Hydroxypropyl Guar (Systane Ultra)', 'Eye Drop', 'Lubricant / Dry Eye', 'bottle', 0, 20, 'Gel-forming lubricant'],
    ['Preservative-free Lubricant Unit Doses', 'Eye Drop', 'Lubricant / Dry Eye', 'box', 0, 10, 'For severe dry eye / post-corneal transplant'],
    ['Cyclosporine 0.05% (Restasis)', 'Eye Drop', 'Immunomodulator', 'vial', 0, 10, 'Chronic dry eye — twice daily for 3+ months'],
    ['Cyclosporine 0.1% (Ikervis — preservative-free)', 'Eye Drop', 'Immunomodulator', 'vial', 0, 5, 'Severe keratitis — once nightly'],
    ['Fluorescein Sodium 0.5% / 1% (diagnostic)', 'Eye Drop', 'Diagnostic', 'bottle', 0, 5, 'Corneal staining under cobalt blue light'],
    ['Rose Bengal 1% (staining for dry eye)', 'Eye Drop', 'Diagnostic', 'bottle', 0, 5, 'Devitalised cell staining'],

    [SEC, 'GENERAL OT & HOSPITAL CONSUMABLES'],
    ['Sterile OT Eye Drape (with collection bag)', 'OT Consumable', 'Drape', 'pcs', 0, 30, 'Single-use — adhesive eye drape with fluid collection bag'],
    ['Disposable OT Gown (sterile)', 'OT Consumable', 'Gown', 'pcs', 0, 50, ''],
    ['Sterile Surgical Gloves 6.0', 'OT Consumable', 'Gloves', 'pairs', 0, 30, ''],
    ['Sterile Surgical Gloves 6.5', 'OT Consumable', 'Gloves', 'pairs', 0, 50, 'Most common size'],
    ['Sterile Surgical Gloves 7.0', 'OT Consumable', 'Gloves', 'pairs', 0, 30, ''],
    ['Sterile Surgical Gloves 7.5', 'OT Consumable', 'Gloves', 'pairs', 0, 20, ''],
    ['Examination Gloves (non-sterile — box of 100)', 'OT Consumable', 'Gloves', 'box', 0, 10, ''],
    ['N95 Mask (FFP2)', 'OT Consumable', 'Mask', 'pcs', 0, 50, ''],
    ['3-Ply Surgical Mask', 'OT Consumable', 'Mask', 'pcs', 0, 200, ''],
    ['Bouffant Cap / Head Cover', 'OT Consumable', 'PPE', 'pcs', 0, 100, ''],
    ['Shoe Cover (OT)', 'OT Consumable', 'PPE', 'pairs', 0, 50, ''],
    ['Syringe 1 mL Tuberculin (insulin)', 'Consumable', 'Syringe', 'pcs', 0, 100, ''],
    ['Syringe 2 mL', 'Consumable', 'Syringe', 'pcs', 0, 100, ''],
    ['Syringe 5 mL', 'Consumable', 'Syringe', 'pcs', 0, 50, ''],
    ['Syringe 10 mL', 'Consumable', 'Syringe', 'pcs', 0, 50, ''],
    ['Needle 26G (intravitreal)', 'Consumable', 'Needle', 'pcs', 0, 50, 'For intravitreal injections'],
    ['Needle 27G (phaco / IVI)', 'Consumable', 'Needle', 'pcs', 0, 100, ''],
    ['Needle 30G (fine cannulation)', 'Consumable', 'Needle', 'pcs', 0, 50, ''],
    ['IV Cannula 20G', 'Consumable', 'IV Access', 'pcs', 0, 30, ''],
    ['IV Cannula 22G', 'Consumable', 'IV Access', 'pcs', 0, 30, ''],
    ['IV Cannula 24G (paediatric)', 'Consumable', 'IV Access', 'pcs', 0, 20, ''],
    ['IV Infusion Set (drip set)', 'Consumable', 'IV Access', 'pcs', 0, 20, ''],
    ['Normal Saline 0.9% 500 mL', 'Consumable', 'IV Fluid', 'bottle', 0, 20, ''],
    ['Cotton Balls (sterile, pack)', 'Consumable', 'Wound Care', 'pack', 0, 20, ''],
    ['Gauze Swabs 4×4 cm (sterile)', 'Consumable', 'Wound Care', 'pack', 0, 20, ''],
    ['Micropore / Transpore Tape 1"', 'Consumable', 'Wound Care', 'roll', 0, 20, 'Post-op eye pad fixation'],
    ['Eye Pad (post-op dressing)', 'Consumable', 'Post-Op', 'pcs', 0, 100, ''],
    ['Eye Shield (Fox shield / protective)', 'Consumable', 'Post-Op', 'pcs', 0, 100, ''],
    ['Betadine 10% Solution (skin prep)', 'Consumable', 'Antiseptic', 'bottle', 0, 10, 'Pre-op skin disinfection'],
    ['70% Isopropyl Alcohol Swabs', 'Consumable', 'Antiseptic', 'pack', 0, 20, ''],
    ['Hand Sanitiser 500 mL', 'Consumable', 'Antiseptic', 'bottle', 0, 20, ''],
    ['OT Scrub / Chlorhexidine Liquid Soap', 'Consumable', 'Antiseptic', 'bottle', 0, 10, ''],
    ['Autoclave Indicator Tape', 'Sterilisation', 'Indicator', 'roll', 0, 10, ''],
    ['Sterilisation Pouches (assorted sizes)', 'Sterilisation', 'Packaging', 'pack', 0, 10, ''],
    ['Ultrasound Gel (for B-scan)', 'Diagnostic Supply', 'Gel', 'bottle', 0, 10, ''],
    ['Fluorescein Strips (Haag-Streit)', 'Diagnostic Supply', 'Staining', 'strips', 0, 200, 'Slit lamp staining'],
    ['Lissamine Green Strips', 'Diagnostic Supply', 'Staining', 'strips', 0, 100, 'Dry eye / ocular surface staining'],
    ['Goldman Tonometer Prism (disposable)', 'Diagnostic Supply', 'Tonometry', 'pcs', 0, 50, 'For applanation tonometry'],
    ['Laser Paper Roll (laser machine)', 'Diagnostic Supply', 'Equipment Supplies', 'roll', 0, 10, ''],
    ['Sharps Disposal Container 5L', 'Safety', 'Waste', 'pcs', 0, 5, ''],
    ['Bio-Hazard Yellow Bag', 'Safety', 'Waste', 'pcs', 0, 50, ''],
    ['Bio-Hazard Red Bag', 'Safety', 'Waste', 'pcs', 0, 50, ''],
    ['Cryotherapy Probe (for retina/CRYO)', 'Surgical Equipment', 'Cryo', 'pcs', 0, 2, 'Single-use cryo probe'],

    [SEC, 'OPTICAL SHOP (if applicable)'],
    ['CR-39 Plastic Lenses (various powers)', 'Optical', 'Lenses', 'pairs', 0, 20, 'Standard optical lenses'],
    ['Polycarbonate Lenses', 'Optical', 'Lenses', 'pairs', 0, 10, 'Impact-resistant'],
    ['Anti-Reflective (AR) Coated Lenses', 'Optical', 'Lenses', 'pairs', 0, 10, ''],
    ['Transition / Photochromic Lenses', 'Optical', 'Lenses', 'pairs', 0, 10, ''],
    ['Reading Glasses (+1.0 to +3.5 — display stock)', 'Optical', 'Ready-Made', 'pairs', 0, 20, ''],
    ['Soft Daily Disposable Contact Lens (box)', 'Optical', 'Contact Lens', 'box', 0, 10, 'Various powers — add per power in stock'],
    ['Monthly Contact Lens (box)', 'Optical', 'Contact Lens', 'box', 0, 10, ''],
    ['Contact Lens Multi-Purpose Solution', 'Optical', 'Contact Lens', 'bottle', 0, 10, ''],
    ['Contact Lens Case', 'Optical', 'Contact Lens', 'pcs', 0, 20, ''],
    ['Sunglass Frames (display)', 'Optical', 'Frames', 'pairs', 0, 20, 'Track by SKU for accuracy'],
    ['Trial Frame (adult — reusable)', 'Optical', 'Equipment', 'pcs', 0, 2, ''],
  ];

  // ── BATCH WRITE ───────────────────────────────────────────────────────────
  var allValues = [];
  var secRows = [];  // {rowNum, label}

  items.forEach(function(item, idx) {
    if (item[0] === SEC) {
      secRows.push({ rowNum: idx + 1, label: item[1] });
      allValues.push(['▸  ' + item[1].toUpperCase(), '', '', '', '', '', '']);
    } else if (idx === 0) {
      // Header row (already structured correctly as 7 values)
      allValues.push(item);
    } else {
      allValues.push(item);
    }
  });

  // Single batch write — all rows at once
  sheet.getRange(1, 1, allValues.length, 7).setValues(allValues);

  // Format header
  sheet.getRange(1, 1, 1, 7)
    .setFontWeight('bold').setBackground('#15233B').setFontColor('#FFFFFF').setFontSize(11);

  // Format section rows (batch)
  secRows.forEach(function(s) {
    sheet.getRange(s.rowNum, 1, 1, 7)
      .setBackground('#E8F0FB').setFontColor('#1A3A6B').setFontWeight('bold').setFontSize(9.5);
  });

  // Highlight Opening Stock and Reorder Level columns
  if (allValues.length > 1) {
    sheet.getRange(2, 5, allValues.length - 1, 2).setBackground('#FFFDF5');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 360);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 300);

  sheet.getRange(1, 5).setNote('Opening Stock: enter current stock count when going live.');
  sheet.getRange(1, 6).setNote('Reorder Level: low-stock alert fires when stock reaches this level.');
  sheet.getRange(1, 1).setNote(
    'INVENTORY MASTER — Comprehensive Eye Hospital Inventory.\n' +
    'IOLs listed by Brand + Type. For each IOL model, duplicate the row\n' +
    'for each diopter power you stock (e.g. +18.0 D, +20.0 D, +22.0 D).\n' +
    'Admin fills Opening Stock and Reorder Level columns.\n' +
    'Store department logs daily issues against these items.'
  );
}
