/**
 * CSITA Christ Church — Treasury Ledger (multi-user)
 * Apps Script backend. Deploy as a Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 * (Security is now handled by app-level login + session tokens below,
 * NOT by the "Only myself" Apps Script restriction — that restriction
 * would block every user except you from ever logging in.)
 *
 * Sheets (auto-created on first run if missing):
 *   Ledger, Categories, Budget, Users
 */

const SHEET_LEDGER = 'Ledger';
const SHEET_CATEGORIES = 'Categories';
const SHEET_BUDGET = 'Budget';
const SHEET_USERS = 'Users';
const SHEET_EVENTS = 'Events';
const SHEET_EVENT_ENTRIES = 'EventEntries';

const LEDGER_HEADERS = [
  'ID', 'Date', 'Type', 'Category', 'SubCategory', 'Description',
  'Amount', 'PaymentMode', 'Reference', 'EnteredOn'
];
const CATEGORY_HEADERS = ['Type', 'Category'];
const BUDGET_HEADERS = ['Year', 'Category', 'BudgetedAmount'];
const USER_HEADERS = ['Username', 'PasswordHash', 'Role', 'CreatedOn'];
const EVENT_HEADERS = ['ID', 'Name', 'StartDate', 'EndDate', 'Status', 'CreatedOn'];
const EVENT_ENTRY_HEADERS = [
  'ID', 'EventID', 'Date', 'Type', 'Description', 'Amount',
  'PaymentMode', 'Reference', 'EnteredOn'
];

const SESSION_DURATION_MS = 12 * 60 * 60 * 1000; // 12 hours

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name, headers) {
  const ss = getSS_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function jsonOut_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Password hashing ---------------- */

function hashPassword_(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return raw.map(b => ((b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))).join('');
}

/**
 * Run this ONCE from the Apps Script editor (select seedDefaultAdmin from
 * the function dropdown, click Run) to create the first admin account.
 * Default login: username "admin", password "ChangeMe123!"
 * Log in with these once the frontend is live, then immediately go to the
 * Users tab (as admin) and change/replace this account.
 * Safe to run again later — it skips creation if "admin" already exists.
 */
function seedDefaultAdmin() {
  const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
  const values = s.getDataRange().getValues();
  const exists = values.slice(1).some(r => r[0] === 'admin');
  if (exists) {
    Logger.log('An "admin" user already exists — skipping.');
    return;
  }
  s.appendRow(['admin', hashPassword_('ChangeMe123!'), 'admin', new Date()]);
  Logger.log('Default admin created — username: admin / password: ChangeMe123!. Change this after first login.');
}

/* ---------------- Sessions ---------------- */

function createSession_(username, role) {
  const token = Utilities.getUuid();
  const payload = JSON.stringify({ username: username, role: role, expires: Date.now() + SESSION_DURATION_MS });
  PropertiesService.getScriptProperties().setProperty('SESSION_' + token, payload);
  return token;
}

function getSession_(token) {
  if (!token) return null;
  const raw = PropertiesService.getScriptProperties().getProperty('SESSION_' + token);
  if (!raw) return null;
  const data = JSON.parse(raw);
  if (Date.now() > data.expires) {
    PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
    return null;
  }
  return data;
}

function requireSession_(token) {
  const session = getSession_(token);
  if (!session) throw new Error('Not logged in, or your session expired. Please log in again.');
  return session;
}

function requireAdmin_(token) {
  const session = requireSession_(token);
  if (session.role !== 'admin') throw new Error('Admin access required for this action.');
  return session;
}

/* ---------------- GET (reads) ---------------- */

function doGet(e) {
  try {
    const action = (e.parameter.action || 'all');
    const token = e.parameter.token;

    if (action === 'verifySession') {
      const session = getSession_(token);
      return jsonOut_({ ok: !!session, username: session ? session.username : null, role: session ? session.role : null });
    }

    // every other read requires a valid, logged-in session
    requireSession_(token);

    const ledgerSheet = getOrCreateSheet_(SHEET_LEDGER, LEDGER_HEADERS);
    const categorySheet = getOrCreateSheet_(SHEET_CATEGORIES, CATEGORY_HEADERS);
    const budgetSheet = getOrCreateSheet_(SHEET_BUDGET, BUDGET_HEADERS);

    if (action === 'ledger') {
      return jsonOut_({ ok: true, data: sheetToObjects_(ledgerSheet) });
    }
    if (action === 'categories') {
      return jsonOut_({ ok: true, data: sheetToObjects_(categorySheet) });
    }
    if (action === 'budget') {
      return jsonOut_({ ok: true, data: sheetToObjects_(budgetSheet) });
    }
    if (action === 'listUsers') {
      requireAdmin_(token);
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const users = sheetToObjects_(s).map(u => ({ Username: u.Username, Role: u.Role, CreatedOn: u.CreatedOn }));
      return jsonOut_({ ok: true, data: users });
    }
    if (action === 'events') {
      const s = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
      return jsonOut_({ ok: true, data: sheetToObjects_(s) });
    }
    if (action === 'eventEntries') {
      const s = getOrCreateSheet_(SHEET_EVENT_ENTRIES, EVENT_ENTRY_HEADERS);
      const all = sheetToObjects_(s);
      const entries = e.parameter.eventId ? all.filter(r => r.EventID === e.parameter.eventId) : all;
      return jsonOut_({ ok: true, data: entries });
    }

    // default: everything the dashboard needs in one call
    const eventsSheet = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
    return jsonOut_({
      ok: true,
      ledger: sheetToObjects_(ledgerSheet),
      categories: sheetToObjects_(categorySheet),
      budget: sheetToObjects_(budgetSheet),
      events: sheetToObjects_(eventsSheet)
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

/* ---------------- POST (writes) ---------------- */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') {
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const users = sheetToObjects_(s);
      const user = users.find(u => u.Username === body.username);
      if (!user || user.PasswordHash !== hashPassword_(body.password)) {
        return jsonOut_({ ok: false, error: 'Invalid username or password.' });
      }
      const token = createSession_(user.Username, user.Role);
      return jsonOut_({ ok: true, token: token, username: user.Username, role: user.Role });
    }

    if (action === 'logout') {
      if (body.token) PropertiesService.getScriptProperties().deleteProperty('SESSION_' + body.token);
      return jsonOut_({ ok: true });
    }

    // everything below requires a valid session
    const session = requireSession_(body.token);

    if (action === 'changePassword') {
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === session.username) {
          if (values[i][1] !== hashPassword_(body.currentPassword)) {
            return jsonOut_({ ok: false, error: 'Current password is incorrect.' });
          }
          s.getRange(i + 1, 2).setValue(hashPassword_(body.newPassword));
          return jsonOut_({ ok: true });
        }
      }
      return jsonOut_({ ok: false, error: 'User not found.' });
    }

    if (action === 'addUser') {
      requireAdmin_(body.token);
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const values = s.getDataRange().getValues();
      const exists = values.slice(1).some(r => r[0] === body.username);
      if (exists) return jsonOut_({ ok: false, error: 'That username already exists.' });
      if (!body.username || !body.password) return jsonOut_({ ok: false, error: 'Username and password are required.' });
      s.appendRow([body.username, hashPassword_(body.password), body.role || 'user', new Date()]);
      return jsonOut_({ ok: true });
    }

    if (action === 'deleteUser') {
      requireAdmin_(body.token);
      if (body.username === session.username) {
        return jsonOut_({ ok: false, error: "You can't delete the account you're logged in as." });
      }
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.username) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    if (action === 'resetUserPassword') {
      requireAdmin_(body.token);
      const s = getOrCreateSheet_(SHEET_USERS, USER_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.username) {
          s.getRange(i + 1, 2).setValue(hashPassword_(body.newPassword));
          return jsonOut_({ ok: true });
        }
      }
      return jsonOut_({ ok: false, error: 'User not found.' });
    }

    if (action === 'addEntry') {
      const s = getOrCreateSheet_(SHEET_LEDGER, LEDGER_HEADERS);
      const id = Utilities.getUuid();
      s.appendRow([
        id,
        body.date,
        body.type,
        body.category,
        body.subCategory || '',
        body.description || '',
        Number(body.amount),
        body.paymentMode || '',
        body.reference || '',
        new Date()
      ]);
      return jsonOut_({ ok: true, id: id });
    }

    if (action === 'deleteEntry') {
      const s = getOrCreateSheet_(SHEET_LEDGER, LEDGER_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    if (action === 'updateEntry') {
      const s = getOrCreateSheet_(SHEET_LEDGER, LEDGER_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          const row = i + 1;
          s.getRange(row, 2, 1, 9).setValues([[
            body.date, body.type, body.category, body.subCategory || '',
            body.description || '', Number(body.amount), body.paymentMode || '',
            body.reference || '', values[i][9]
          ]]);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    if (action === 'addCategory') {
      const s = getOrCreateSheet_(SHEET_CATEGORIES, CATEGORY_HEADERS);
      s.appendRow([body.type, body.category]);
      return jsonOut_({ ok: true });
    }

    if (action === 'deleteCategory') {
      const s = getOrCreateSheet_(SHEET_CATEGORIES, CATEGORY_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.type && values[i][1] === body.category) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    if (action === 'setBudget') {
      const s = getOrCreateSheet_(SHEET_BUDGET, BUDGET_HEADERS);
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(body.year) && values[i][1] === body.category) {
          s.getRange(i + 1, 3).setValue(Number(body.amount));
          found = true;
          break;
        }
      }
      if (!found) {
        s.appendRow([body.year, body.category, Number(body.amount)]);
      }
      return jsonOut_({ ok: true });
    }

    /* ---- Events module (fully separate from the main Ledger) ---- */

    if (action === 'addEvent') {
      if (!body.name) return jsonOut_({ ok: false, error: 'Event name is required.' });
      const s = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
      const id = Utilities.getUuid();
      s.appendRow([id, body.name, body.startDate || '', body.endDate || '', 'Open', new Date()]);
      return jsonOut_({ ok: true, id: id });
    }

    if (action === 'setEventStatus') {
      const s = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.getRange(i + 1, 5).setValue(body.status); // 'Open' or 'Closed'
          return jsonOut_({ ok: true });
        }
      }
      return jsonOut_({ ok: false, error: 'Event not found.' });
    }

    if (action === 'deleteEvent') {
      requireAdmin_(body.token);
      const s = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.deleteRow(i + 1);
          break;
        }
      }
      // cascade: remove this event's entries too
      const es = getOrCreateSheet_(SHEET_EVENT_ENTRIES, EVENT_ENTRY_HEADERS);
      const evalues = es.getDataRange().getValues();
      for (let i = evalues.length - 1; i >= 1; i--) {
        if (evalues[i][1] === body.id) {
          es.deleteRow(i + 1);
        }
      }
      return jsonOut_({ ok: true });
    }

    if (action === 'addEventEntry') {
      const s = getOrCreateSheet_(SHEET_EVENT_ENTRIES, EVENT_ENTRY_HEADERS);
      const id = Utilities.getUuid();
      s.appendRow([
        id,
        body.eventId,
        body.date,
        body.type, // 'Income' or 'Expense'
        body.description || '',
        Number(body.amount),
        body.paymentMode || '',
        body.reference || '',
        new Date()
      ]);
      return jsonOut_({ ok: true, id: id });
    }

    if (action === 'deleteEventEntry') {
      const s = getOrCreateSheet_(SHEET_EVENT_ENTRIES, EVENT_ENTRY_HEADERS);
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

/**
 * Run this once manually from the Apps Script editor to seed default
 * categories so the dashboard isn't empty on first load. Safe to skip
 * and add categories from the UI instead.
 */
function seedDefaultCategories() {
  const s = getOrCreateSheet_(SHEET_CATEGORIES, CATEGORY_HEADERS);
  const defaults = [
    ['Income', 'Offerings'],
    ['Income', 'Tithes'],
    ['Income', 'Special Collections'],
    ['Income', 'Donations'],
    ['Income', 'Harvest Festival'],
    ['Income', 'Other Income'],
    ['Expense', 'Utilities'],
    ['Expense', 'Maintenance'],
    ['Expense', 'Sunday School Supplies'],
    ['Expense', 'Charity & Outreach'],
    ['Expense', 'Honorariums'],
    ['Expense', 'Office & Admin'],
    ['Expense', 'Other Expense']
  ];
  defaults.forEach(row => s.appendRow(row));
}
