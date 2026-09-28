/**
 * Crooked Billet Quiz Questions – API backend (Google Apps Script)
 *
 * Setup:
 * 1. In the Sheet: Extensions > Apps Script. Paste this into Code.gs (replace everything).
 * 2. Project Settings (gear) > Script Properties > Add:  ADMIN_PASSWORD = <your password>
 * 3. Deploy > New deployment > Web app.  Execute as: Me.  Who has access: Anyone.
 * 4. Copy the /exec URL into API_URL in index.html.
 *    (After ANY later edit to this script: Deploy > Manage deployments > pencil > Version: New version > Deploy.)
 *
 * Tabs expected:
 *   Questions:  ID | Category | Question | Answer | Extra Information | Points | Author | Status | Round | Date Added | Last Edited
 *   Categories: Category
 *   People:     Name
 */

const QUESTIONS_TAB = 'Questions';
const CATEGORIES_TAB = 'Categories';
const PEOPLE_TAB = 'People';
const SEP = '\u241F';
const STAMP = 'yyyy-MM-dd HH:mm';

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!pw || req.password !== pw) return out_({ ok: false, error: 'Wrong password' });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return out_(handle_(req)); } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() { return out_({ ok: true, message: 'Quiz API running' }); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  switch (req.action) {
    case 'list':           return list_();
    case 'addQuestion':    return addQuestion_(req.values);
    case 'updateQuestion': return updateQuestion_(req.row, req.sig, req.values);
    case 'deleteQuestion': return deleteQuestion_(req.row, req.sig);
    case 'addCategory':    return addCategory_(req.name);
    case 'renameCategory': return renameCategory_(req.from, req.to);
    case 'deleteCategory': return deleteCategory_(req.name);
    default: throw new Error('Unknown action: ' + req.action);
  }
}

/* ---------- helpers ---------- */

function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), STAMP); }

function qSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(QUESTIONS_TAB);
  if (!sh) throw new Error('No tab named "' + QUESTIONS_TAB + '"');
  return sh;
}

function headers_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(s => String(s).trim());
  const idx = {};
  h.forEach((n, i) => idx[n.toLowerCase()] = i);
  if (idx['category'] === undefined) throw new Error('Questions tab needs a "Category" header');
  return { h, idx };
}

function readRows_(sh, h) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, h.length).getDisplayValues();
  const rows = [];
  data.forEach((r, i) => {
    if (r.every(v => String(v).trim() === '')) return;
    const vals = r.map(String);
    rows.push({ row: i + 2, sig: vals.join(SEP), values: vals });
  });
  return rows;
}

function listCol_(tabName, create, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(tabName);
  if (!sh) {
    if (!create) return [];
    sh = ss.insertSheet(tabName);
    sh.getRange(1, 1).setValue(header);
  }
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function catSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CATEGORIES_TAB);
  if (!sh) { sh = ss.insertSheet(CATEGORIES_TAB); sh.getRange(1, 1).setValue('Category'); }
  return sh;
}

function checkRow_(sh, h, row, sig) {
  if (!row || row < 2 || row > sh.getLastRow()) throw new Error('Row no longer exists – refresh and try again');
  const cur = sh.getRange(row, 1, 1, h.length).getDisplayValues()[0].map(String);
  if (cur.join(SEP) !== sig) throw new Error('That row changed in the sheet – refresh and try again');
  return cur;
}

function nextId_(sh, h, idx) {
  if (idx['id'] === undefined) return '';
  let max = 0;
  readRows_(sh, h).forEach(r => {
    const m = String(r.values[idx['id']]).match(/\d+/);
    if (m) max = Math.max(max, parseInt(m[0], 10));
  });
  return 'Q' + String(max + 1).padStart(4, '0');
}

/* ---------- actions ---------- */

function list_() {
  const sh = qSheet_();
  const { h, idx } = headers_(sh);
  const rows = readRows_(sh, h);
  const categories = listCol_(CATEGORIES_TAB, true, 'Category');
  const seenC = new Set(categories);
  const people = listCol_(PEOPLE_TAB, false);
  const seenP = new Set(people);
  rows.forEach(r => {
    const c = r.values[idx['category']].trim();
    if (c && !seenC.has(c)) { seenC.add(c); categories.push(c); }
    if (idx['author'] !== undefined) {
      const a = r.values[idx['author']].trim();
      if (a && !seenP.has(a)) { seenP.add(a); people.push(a); }
    }
  });
  return { ok: true, headers: h, categories, people, rows };
}

function addQuestion_(values) {
  const sh = qSheet_();
  const { h, idx } = headers_(sh);
  const stamp = now_();
  const id = nextId_(sh, h, idx);
  const row = h.map(name => {
    const k = name.toLowerCase();
    if (k === 'id') return id;
    if (k === 'date added' || k === 'last edited') return stamp;
    return (values && values[name] != null) ? String(values[name]) : '';
  });
  sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setValues([row]);
  return { ok: true, id };
}

function updateQuestion_(row, sig, values) {
  const sh = qSheet_();
  const { h } = headers_(sh);
  const cur = checkRow_(sh, h, row, sig);
  const stamp = now_();
  const next = h.map((name, i) => {
    const k = name.toLowerCase();
    if (k === 'id' || k === 'date added') return cur[i];
    if (k === 'last edited') return stamp;
    return (values && values[name] != null) ? String(values[name]) : cur[i];
  });
  sh.getRange(row, 1, 1, h.length).setValues([next]);
  return { ok: true };
}

function deleteQuestion_(row, sig) {
  const sh = qSheet_();
  const { h } = headers_(sh);
  checkRow_(sh, h, row, sig);
  sh.deleteRow(row);
  return { ok: true };
}

function addCategory_(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Category name is empty');
  if (list_().categories.some(c => c.toLowerCase() === name.toLowerCase())) throw new Error('Category already exists');
  catSheet_().appendRow([name]);
  return { ok: true };
}

function renameCategory_(from, to) {
  to = String(to || '').trim();
  if (!to) throw new Error('New name is empty');
  const sh = qSheet_();
  const { h, idx } = headers_(sh);
  const last = sh.getLastRow();
  if (last >= 2) {
    const rng = sh.getRange(2, idx['category'] + 1, last - 1, 1);
    rng.setValues(rng.getDisplayValues().map(r => [String(r[0]).trim() === from ? to : r[0]]));
  }
  const cs = catSheet_();
  const cl = cs.getLastRow();
  let found = false;
  if (cl >= 2) {
    const crng = cs.getRange(2, 1, cl - 1, 1);
    crng.setValues(crng.getDisplayValues().map(r => {
      if (String(r[0]).trim() === from) { found = true; return [to]; }
      return [r[0]];
    }));
  }
  if (!found) cs.appendRow([to]);
  return { ok: true };
}

function deleteCategory_(name) {
  const sh = qSheet_();
  const { h, idx } = headers_(sh);
  if (readRows_(sh, h).some(r => r.values[idx['category']].trim() === name)) {
    throw new Error('Category still has questions – move or delete them first');
  }
  const cs = catSheet_();
  for (let r = cs.getLastRow(); r >= 2; r--) {
    if (String(cs.getRange(r, 1).getDisplayValue()).trim() === name) cs.deleteRow(r);
  }
  return { ok: true };
}
