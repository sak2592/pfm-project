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
 *                               member, vendor, notes, fundName, recurringId, loanId)
 *   - Loans                   (id, name, lender, principal, outstanding,
 *                               rate, tenure, emi, startDate, type,
 *                               autoOutstanding)
 *   - Goals                   (id, name, icon, target, saved, monthly,
 *                               targetDate, priority, description)
 *   - Assets                  (name, category, purchaseValue,
 *                               currentValue, liquid, source, xirr)
 *   - RecurringItems          (id, type, category, subcategory,
 *                               classification, amount, frequency,
 *                               startDate, endDate, payment, bank, member,
 *                               vendor, notes, escalationPct, fundName, active)
 *   - SalaryHistory           (date, salary)
 *   - Settings                (key, value)  <-- current salary + last saved
 *   - BankAccounts            (id, name, openingBalance)
 *   - Transfers               (id, date, fromBank, toBank, amount, notes)
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
  loanPrepayments: 'LoanPrepayments',
  goals: 'Goals',
  assets: 'Assets',
  recurringItems: 'RecurringItems',
  salaryHistory: 'SalaryHistory',
  settings: 'Settings',
  monthlyBudgetTargets: 'MonthlyBudgetTargets',
  activeSips: 'ActiveSIPs',
  bankAccounts: 'BankAccounts',
  transfers: 'Transfers'
};

// ===== COLUMN DEFINITIONS FOR SIMPLE (1 row = 1 object) TABLES =====
const HEADERS = {
  categories: ['id', 'name', 'targetPct', 'priority', 'classification'],
  transactions: ['id', 'date', 'type', 'category', 'subcategory', 'classification', 'amount', 'payment', 'bank', 'member', 'vendor', 'notes', 'fundName', 'recurringId', 'loanId'],
  loans: ['id', 'name', 'lender', 'principal', 'outstanding', 'rate', 'tenure', 'emi', 'startDate', 'type', 'autoOutstanding'],
  // Each loan's `prepayments` array (lump-sum payments logged in the Prepay popup) gets its
  // own sheet, one row per prepayment, keyed by loanId - same flatten/regroup pattern used for
  // CategoriesSubcategories below. Loans themselves have no column for this nested array, so
  // without this table prepayments are silently dropped on every Sheets round-trip.
  loanPrepayments: ['loanId', 'date', 'amount'],
  goals: ['id', 'name', 'icon', 'target', 'saved', 'monthly', 'targetDate', 'priority', 'description'],
  assets: ['name', 'category', 'purchaseValue', 'currentValue', 'liquid', 'source', 'xirr'],
  recurringItems: ['id', 'type', 'category', 'subcategory', 'classification', 'amount', 'frequency', 'startDate', 'endDate', 'payment', 'bank', 'member', 'vendor', 'notes', 'escalationPct', 'fundName', 'active'],
  salaryHistory: ['date', 'salary'],
  categoryConfig: ['id', 'type', 'category', 'subcategory'],
  // One row per category per month where the target % has been overridden
  // away from that category's default targetPct (in BudgetCategories).
  // Months not listed here just fall back to the default %.
  monthlyBudgetTargets: ['id', 'month', 'categoryId', 'targetPct'],
  // Read-only, derived table: one row per active Mutual Fund / Stocks SIP (a RecurringItem
  // with type=Investment, active=true), enriched with whatever Groww holding (Asset with
  // source='Groww') the app's fuzzy fund-name matcher linked it to. Rebuilt from scratch on
  // every save, so editing it directly in the Sheet has no effect — edit the actual Fixed
  // Investment in the app instead.
  activeSips: ['id', 'category', 'subcategory', 'fundName', 'amount', 'frequency', 'startDate',
    'escalationPct', 'matchedHolding', 'matchedInvested', 'matchedCurrent', 'matchedXirr'],
  // Bank accounts (e.g. HDFC Savings, ICICI Savings) with a starting balance from before
  // the user began logging transactions in the app.
  bankAccounts: ['id', 'name', 'openingBalance'],
  // Money moved between the user's own accounts, kept separate from Transactions so it's
  // never double-counted as Income/Expense.
  transfers: ['id', 'date', 'fromBank', 'toBank', 'amount', 'notes']
};

// Fields that should always come back as numbers.
const NUMERIC_FIELDS = ['targetPct', 'amount', 'principal', 'outstanding', 'rate', 'tenure', 'emi',
  'target', 'saved', 'monthly', 'purchaseValue', 'currentValue', 'escalationPct', 'salary',
  'matchedInvested', 'matchedCurrent', 'matchedXirr', 'openingBalance'];

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

// Numeric fields that should stay null (not become 0) when the cell is blank -
// e.g. XIRR isn't always computable for a holding, and 0% is a real, different value.
const NULLABLE_NUMERIC_FIELDS = ['xirr', 'matchedXirr'];

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
 * Like writeTable_, but instead of clearing and rewriting the whole sheet every time,
 * this diffs against what's already there (matched by the "id" column) and:
 *   - APPENDS brand-new records in one batch call
 *   - UPDATES only the specific rows whose data actually changed
 *   - DELETES only the rows for records removed on the client (e.g. a deleted transaction)
 * Everything else on the sheet is left completely untouched.
 *
 * Only usable for tables whose `headers` include an 'id' column with stable, unique values
 * (Transactions, Loans, Goals, RecurringItems, BankAccounts, Transfers, MonthlyBudgetTargets,
 * BudgetCategories). Tables without a reliable id (Assets, SalaryHistory) still use
 * writeTable_ - falls back to it automatically here too, just in case.
 */
function upsertTable_(sheetName, headers, objects) {
  const idIdx = headers.indexOf('id');
  if (idIdx === -1) {
    writeTable_(sheetName, headers, objects);
    return;
  }

  const sheet = getSheet_(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow();
  const existingRows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];

  // id -> { rowNum, values } for everything currently on the sheet
  const existingById = {};
  existingRows.forEach(function (row, i) {
    const id = row[idIdx];
    if (id !== '' && id !== null && id !== undefined) {
      existingById[String(id)] = { rowNum: i + 2, values: row };
    }
  });

  const incomingIds = {};
  const rowsToAppend = [];

  (objects || []).forEach(function (obj) {
    const id = String(obj.id);
    incomingIds[id] = true;
    const rowValues = headers.map(function (h) {
      const v = obj[h];
      return v === undefined || v === null ? '' : v;
    });

    const existing = existingById[id];
    if (!existing) {
      rowsToAppend.push(rowValues);
    } else {
      // Only write to the sheet if something in this row actually changed
      const changed = rowValues.some(function (v, i) { return String(v) !== String(existing.values[i]); });
      if (changed) {
        sheet.getRange(existing.rowNum, 1, 1, headers.length).setValues([rowValues]);
      }
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  // Rows whose id no longer appears in the incoming data were deleted client-side - remove them
  // from the bottom up so earlier row numbers don't shift while we're still deleting.
  const rowNumsToDelete = [];
  Object.keys(existingById).forEach(function (id) {
    if (!incomingIds[id]) rowNumsToDelete.push(existingById[id].rowNum);
  });
  rowNumsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) {
    sheet.deleteRow(rowNum);
  });
}

/**
 * Writes an array of plain objects to a sheet as a simple table:
 * header row + one row per object, in the given column order.
 * Clears and rewrites the whole sheet - only used for tables without a reliable
 * id column (Assets, SalaryHistory) or ones that are always fully regenerated
 * anyway (ActiveSIPs). Everything else uses upsertTable_ instead.
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
        if (NULLABLE_NUMERIC_FIELDS.indexOf(h) !== -1) {
          val = (val === '' || val === null || val === undefined) ? null : toNum_(val);
        } else if (NUMERIC_FIELDS.indexOf(h) !== -1) {
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
// LoanPrepayments (each loan's `prepayments` array, flattened)
// ==================================================================
// A loan looks like { id, name, ..., prepayments: [{date, amount}, ...] }. HEADERS.loans has
// no column for that nested array, so it's stored separately here (one row per prepayment,
// tagged with loanId) and re-attached to the right loan object on load.

function writeLoanPrepayments_(loans) {
  const rows = [];
  (loans || []).forEach(function (loan) {
    (loan.prepayments || []).forEach(function (p) {
      rows.push({ loanId: loan.id, date: p.date, amount: p.amount });
    });
  });
  writeTable_(SHEETS.loanPrepayments, HEADERS.loanPrepayments, rows);
}

function readLoanPrepayments_() {
  const rows = readTable_(SHEETS.loanPrepayments, HEADERS.loanPrepayments);
  const byLoanId = {};
  rows.forEach(function (r) {
    if (!byLoanId[r.loanId]) byLoanId[r.loanId] = [];
    byLoanId[r.loanId].push({ date: r.date, amount: r.amount });
  });
  return byLoanId;
}

// ==================================================================
// Settings (current salary + last saved timestamp)
// ==================================================================

function writeSettings_(state) {
  const rows = [
    { key: 'salary', value: state.salary },
    { key: 'lastSaved', value: new Date().toISOString() },
    // Echoes back the client's own lastModified timestamp (set on every local save) so that on
    // the next load, the client can tell whether this cloud copy is newer or older than whatever
    // it already has locally, and avoid clobbering newer local edits with a stale cloud copy.
    { key: 'lastModified', value: state.lastModified || Date.now() }
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
      const prepaymentsByLoanId = readLoanPrepayments_();
      const loans = readTable_(SHEETS.loans, HEADERS.loans);
      loans.forEach(function (loan) {
        loan.prepayments = prepaymentsByLoanId[loan.id] || [];
      });
      const state = {
        salary: toNum_(settings.salary || 0),
        lastModified: toNum_(settings.lastModified || 0),
        categories: readTable_(SHEETS.categories, HEADERS.categories),
        categoryConfig: readCategoryConfig_(),
        transactions: readTable_(SHEETS.transactions, HEADERS.transactions),
        loans: loans,
        goals: readTable_(SHEETS.goals, HEADERS.goals),
        assets: readTable_(SHEETS.assets, HEADERS.assets),
        recurringItems: readTable_(SHEETS.recurringItems, HEADERS.recurringItems),
        salaryHistory: readTable_(SHEETS.salaryHistory, HEADERS.salaryHistory),
        monthlyBudgetTargets: readTable_(SHEETS.monthlyBudgetTargets, HEADERS.monthlyBudgetTargets),
        activeSips: readTable_(SHEETS.activeSips, HEADERS.activeSips),
        bankAccounts: readTable_(SHEETS.bankAccounts, HEADERS.bankAccounts),
        transfers: readTable_(SHEETS.transfers, HEADERS.transfers)
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
        upsertTable_(SHEETS.categories, HEADERS.categories, state.categories);
        writeCategoryConfig_(state.categoryConfig);
        upsertTable_(SHEETS.transactions, HEADERS.transactions, state.transactions);
        upsertTable_(SHEETS.loans, HEADERS.loans, state.loans);
        writeLoanPrepayments_(state.loans);
        upsertTable_(SHEETS.goals, HEADERS.goals, state.goals);
        writeTable_(SHEETS.assets, HEADERS.assets, state.assets);
        upsertTable_(SHEETS.recurringItems, HEADERS.recurringItems, state.recurringItems);
        writeTable_(SHEETS.salaryHistory, HEADERS.salaryHistory, state.salaryHistory);
        upsertTable_(SHEETS.monthlyBudgetTargets, HEADERS.monthlyBudgetTargets, state.monthlyBudgetTargets);
        writeTable_(SHEETS.activeSips, HEADERS.activeSips, state.activeSips);
        upsertTable_(SHEETS.bankAccounts, HEADERS.bankAccounts, state.bankAccounts);
        upsertTable_(SHEETS.transfers, HEADERS.transfers, state.transfers);
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
