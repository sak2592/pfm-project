/**
 * Personal Financial Management System - Google Apps Script backend
 * -------------------------------------------------------------------
 * Unlike a single JSON-blob approach, this version gives every part of
 * your app its OWN sheet/tab, so you can open the spreadsheet and read
 * things like a normal table:
 *
 *   - BudgetCategories        (id, name, targetPct, priority, classification)
 *   - CategoriesSubcategories (id, type, category, subcategory)  <-- combined
 *                               category + subcategory table, one row per
 *                               subcategory
 *   - Transactions            (id, date, type, category, subcategory,
 *                               classification, amount, payment, bank,
 *                               member, vendor, notes, recurringId, loanId)
 *   - Loans                   (id, name, lender, principal, outstanding,
 *                               rate, tenure, emi, startDate, type,
 *                               autoOutstanding)
 *   - Goals                   (id, name, icon, target, saved, monthly,
 *                               targetDate, priority, description)
 *   - Assets                  (name, category, purchaseValue,
 *                               currentValue, liquid)
 *   - RecurringItems          (id, type, category, subcategory,
 *                               classification, amount, frequency,
 *                               startDate, endDate, payment, bank, member,
 *                               vendor, notes, escalationPct, active)
 *   - SalaryHistory           (date, salary)
 *   - Settings                (key, value)  <-- current salary + last saved
 *
 * SETUP:
 * 1. Create a new Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Save (Ctrl+S / Cmd+S).
 * 5. Click Deploy -> New deployment.
 * 6. Type: "Web app".
 * 7. Execute as: "Me".
 * 8. Who has access: "Anyone".
 * 9. Click Deploy, authorize when prompted, and copy the Web App URL.
 * 10. Paste that URL into the app's Connection Settings and click
 *     "Save", then "Test Connection".
 *
 * The app talks to this script using:
 *   GET  ?action=ping   -> { status: 'ok' }
 *   GET  ?action=load   -> { status: 'ok', ...savedState }
 *   POST { action:'save', state: {...} } -> { status: 'ok' } or
 *                                             { status: 'error', message }
 * All sheets are created automatically the first time you save.
 */

// ===== SHEET NAMES =====
const SHEETS = {
  categories: 'BudgetCategories',
  categoryConfig: 'CategoriesSubcategories',
  transactions: 'Transactions',
  loans: 'Loans',
  goals: 'Goals',
  assets: 'Assets',
  recurringItems: 'RecurringItems',
  salaryHistory: 'SalaryHistory',
  settings: 'Settings',
  monthlyBudgetTargets: 'MonthlyBudgetTargets'
};

// ===== COLUMN DEFINITIONS FOR SIMPLE (1 row = 1 object) TABLES =====
const HEADERS = {
  categories: ['id', 'name', 'targetPct', 'priority', 'classification'],
  transactions: ['id', 'date', 'type', 'category', 'subcategory', 'classification', 'amount', 'payment', 'bank', 'member', 'vendor', 'notes', 'recurringId', 'loanId'],
  loans: ['id', 'name', 'lender', 'principal', 'outstanding', 'rate', 'tenure', 'emi', 'startDate', 'type', 'autoOutstanding'],
  goals: ['id', 'name', 'icon', 'target', 'saved', 'monthly', 'targetDate', 'priority', 'description'],
  assets: ['name', 'category', 'purchaseValue', 'currentValue', 'liquid'],
  recurringItems: ['id', 'type', 'category', 'subcategory', 'classification', 'amount', 'frequency', 'startDate', 'endDate', 'payment', 'bank', 'member', 'vendor', 'notes', 'escalationPct', 'active'],
  salaryHistory: ['date', 'salary'],
  categoryConfig: ['id', 'type', 'category', 'subcategory'],
  // One row per category per month where the target % has been overridden
  // away from that category's default targetPct (in BudgetCategories).
  // Months not listed here just fall back to the default %.
  monthlyBudgetTargets: ['id', 'month', 'categoryId', 'targetPct']
};

// Fields that should always come back as numbers.
const NUMERIC_FIELDS = ['targetPct', 'amount', 'principal', 'outstanding', 'rate', 'tenure', 'emi',
  'target', 'saved', 'monthly', 'purchaseValue', 'currentValue', 'escalationPct', 'salary'];

// Fields that should always come back as booleans.
const BOOLEAN_FIELDS = ['autoOutstanding', 'liquid', 'active'];

// Fields that hold a date/month and should be normalized back to plain
// text (Sheets sometimes auto-converts date-looking strings into real
// Date cells).
const DATE_FIELDS = ['date', 'startDate', 'endDate', 'targetDate', 'month'];

// ==================================================================
// Generic helpers
// ==================================================================

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function toBool_(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') return v.trim().toUpperCase() === 'TRUE';
  return !!v;
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function toDateStr_(v, monthOnly) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), monthOnly ? 'yyyy-MM' : 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

/**
 * Writes an array of plain objects to a sheet as a simple table:
 * header row + one row per object, in the given column order.
 */
function writeTable_(sheetName, headers, objects) {
  const sheet = getSheet_(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!objects || objects.length === 0) return;
  const rows = objects.map(function (obj) {
    return headers.map(function (h) {
      const val = obj[h];
      return val === undefined || val === null ? '' : val;
    });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Reads a simple table sheet back into an array of plain objects,
 * coercing numeric/boolean/date fields based on the field name.
 */
function readTable_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter(function (row) { return row.some(function (c) { return c !== ''; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        let val = row[i];
        if (NUMERIC_FIELDS.indexOf(h) !== -1) {
          val = toNum_(val);
        } else if (BOOLEAN_FIELDS.indexOf(h) !== -1) {
          val = toBool_(val);
        } else if (DATE_FIELDS.indexOf(h) !== -1) {
          val = toDateStr_(val, h === 'targetDate' || h === 'month');
        }
        obj[h] = val;
      });
      return obj;
    });
}

// ==================================================================
// CategoriesSubcategories (combined category + subcategory table)
// ==================================================================
// state.categoryConfig looks like: { id, type, name, subcategories: [...] }
// We flatten this into one row per subcategory: id, type, category, subcategory
// and re-group them back into that shape on load.

function writeCategoryConfig_(categoryConfig) {
  const rows = [];
  (categoryConfig || []).forEach(function (cc) {
    const subs = (cc.subcategories && cc.subcategories.length) ? cc.subcategories : [''];
    subs.forEach(function (sub) {
      rows.push({ id: cc.id, type: cc.type, category: cc.name, subcategory: sub });
    });
  });
  writeTable_(SHEETS.categoryConfig, HEADERS.categoryConfig, rows);
}

function readCategoryConfig_() {
  const rows = readTable_(SHEETS.categoryConfig, HEADERS.categoryConfig);
  const byId = {};
  const order = [];
  rows.forEach(function (r) {
    if (!byId[r.id]) {
      byId[r.id] = { id: r.id, type: r.type, name: r.category, subcategories: [] };
      order.push(r.id);
    }
    if (r.subcategory !== '') {
      byId[r.id].subcategories.push(r.subcategory);
    }
  });
  return order.map(function (id) { return byId[id]; });
}

// ==================================================================
// Settings (current salary + last saved timestamp)
// ==================================================================

function writeSettings_(state) {
  const rows = [
    { key: 'salary', value: state.salary },
    { key: 'lastSaved', value: new Date().toISOString() }
  ];
  writeTable_(SHEETS.settings, ['key', 'value'], rows);
}

function readSettings_() {
  const rows = readTable_(SHEETS.settings, ['key', 'value']);
  const map = {};
  rows.forEach(function (r) { map[r.key] = r.value; });
  return map;
}

// ==================================================================
// HTTP entry points
// ==================================================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'ping') {
      return jsonResponse_({ status: 'ok', message: 'pong' });
    }

    if (action === 'load') {
      const settings = readSettings_();
      const state = {
        salary: toNum_(settings.salary || 0),
        categories: readTable_(SHEETS.categories, HEADERS.categories),
        categoryConfig: readCategoryConfig_(),
        transactions: readTable_(SHEETS.transactions, HEADERS.transactions),
        loans: readTable_(SHEETS.loans, HEADERS.loans),
        goals: readTable_(SHEETS.goals, HEADERS.goals),
        assets: readTable_(SHEETS.assets, HEADERS.assets),
        recurringItems: readTable_(SHEETS.recurringItems, HEADERS.recurringItems),
        salaryHistory: readTable_(SHEETS.salaryHistory, HEADERS.salaryHistory),
        monthlyBudgetTargets: readTable_(SHEETS.monthlyBudgetTargets, HEADERS.monthlyBudgetTargets)
      };
      return jsonResponse_(Object.assign({ status: 'ok' }, state));
    }

    return jsonResponse_({ status: 'error', message: 'Unknown or missing action for GET request.' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ status: 'error', message: 'No POST body received.' });
    }

    let body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse_({ status: 'error', message: 'Request body is not valid JSON.' });
    }

    const action = body.action || '';

    if (action === 'save') {
      const state = body.state;
      if (!state) {
        return jsonResponse_({ status: 'error', message: 'No "state" field found in save request.' });
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        writeSettings_(state);
        writeTable_(SHEETS.categories, HEADERS.categories, state.categories);
        writeCategoryConfig_(state.categoryConfig);
        writeTable_(SHEETS.transactions, HEADERS.transactions, state.transactions);
        writeTable_(SHEETS.loans, HEADERS.loans, state.loans);
        writeTable_(SHEETS.goals, HEADERS.goals, state.goals);
        writeTable_(SHEETS.assets, HEADERS.assets, state.assets);
        writeTable_(SHEETS.recurringItems, HEADERS.recurringItems, state.recurringItems);
        writeTable_(SHEETS.salaryHistory, HEADERS.salaryHistory, state.salaryHistory);
        writeTable_(SHEETS.monthlyBudgetTargets, HEADERS.monthlyBudgetTargets, state.monthlyBudgetTargets);
      } finally {
        lock.releaseLock();
      }
      return jsonResponse_({ status: 'ok' });
    }

    return jsonResponse_({ status: 'error', message: 'Unknown action for POST request.' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
