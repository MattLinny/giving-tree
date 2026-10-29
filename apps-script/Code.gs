/**
 * Christmas Giving Tree: Google Sheet backend.
 *
 * Setup (once):
 *   1. Open the Needs spreadsheet → Extensions → Apps Script. Paste this whole file over Code.gs and save.
 *   2. Pick `setup` in the function menu and click Run. Approve the permissions it asks for.
 *      It adds the private columns to the Needs tab and creates a Pledges tab.
 *   3. Optional: Project Settings → Script properties → add NOTIFY_EMAIL = the address that should hear
 *      about new needs and pledges.
 *   4. Deploy → New deployment → type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 *   5. Copy the Web app URL into CONFIG.apiUrl in index.html.
 *
 * After editing this file later: Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
 * The URL stays the same.
 *
 * Privacy: the web page only ever receives id, need, details, amount and status for tags that are
 * open, claimed or filled. Names, contact details and new submissions never leave the sheet.
 */

const NEEDS = 'Needs';
const PLEDGES = 'Pledges';
const NEED_COLS = ['id', 'need', 'details', 'for', 'amount', 'status', 'submitted_at', 'submitted_by', 'contact', 'relationship'];
const PLEDGE_COLS = ['submitted_at', 'tag', 'need', 'name', 'contact', 'dropoff'];
const HIDDEN = ['new', 'review', 'pending', 'removed', 'hidden'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const needs = ss.getSheetByName(NEEDS) || ss.getSheets()[0].setName(NEEDS);
  ensureHeaders_(needs, NEED_COLS);
  ensureHeaders_(ss.getSheetByName(PLEDGES) || ss.insertSheet(PLEDGES), PLEDGE_COLS);
}

/**
 * One-off: empties the Needs and Pledges tabs (header rows, colours and status dropdowns stay) and fills
 * Needs with gift-card tags. Pick `seedGiftCards` in the function menu and click Run. No new deployment needed.
 */
function seedGiftCards() {
  const cards = [
    // need, details, for, amount, status
    ['Save-On-Foods gift card', 'Groceries for the holidays and the January stretch.', 'Family of 5', 200, 'open'],
    ['Prepaid Visa gift card', 'Flexible help with bills and Christmas basics.', 'Single mom', 100, 'open'],
    ['Chevron gas card', 'Keeps the car running for work and school drop-offs.', 'Job seeker', 75, 'open'],
    ['Thrifty Foods gift card', 'A little extra for fresh food this month.', 'Senior', 50, 'open'],
    ['Prepaid Visa gift card', 'Winter clothing and school supplies for the kids.', 'Family of 4', 250, 'open'],
    ['Walmart grocery card', 'Christmas dinner and pantry staples.', 'Family of 6', 300, 'open'],
    ['Petro-Canada gas card', 'Covers a few weeks of driving to work.', 'Single dad', 100, 'open'],
    ['Prepaid Visa gift card', 'Help with a phone bill during a job search.', 'Young adult', 50, 'open'],
    ['Save-On-Foods gift card', 'Formula, diapers and groceries for a new baby.', 'New parents', 150, 'open'],
    ['Co-op gas card', 'Driving to medical appointments in town.', 'Family of 3', 150, 'claimed'],
    ['Prepaid Visa gift card', 'A month of rent top-up after a reduced-hours stretch.', 'Family of 5', 500, 'open'],
    ['Quality Foods gift card', 'Groceries through the holidays.', 'Family of 4', 400, 'filled'],
    ['Shell gas card', 'Getting to church and to see family.', 'Senior', 60, 'claimed'],
    ['Prepaid Visa gift card', 'Christmas gifts for three kids.', 'Single mom', 200, 'filled'],
  ];
  const needs = sheet_(NEEDS), pledges = sheet_(PLEDGES);
  [needs, pledges].forEach(sh => {
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
  const head = table_(needs).head, now = new Date();
  const rows = cards.map(([need, details, who, amount, status], i) => {
    const rec = { id: i + 1, need, details, for: who, amount, status, submitted_at: now, submitted_by: 'Care team' };
    return head.map(h => (rec[h] === undefined ? '' : rec[h]));
  });
  needs.getRange(2, 1, rows.length, head.length).setValues(rows);
  CacheService.getScriptCache().remove('needs');
}

/* ── Read: the tree's public data ──
   The answer is cached for 30 s so a busy page doesn't re-read the sheet for every visitor. Any edit to the
   sheet (by hand or by a form) clears the cache, so changes still show up on the next refresh. */
function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('needs');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const out = JSON.stringify(publicNeeds_());
  cache.put('needs', out, 30);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function publicNeeds_() {
  const { head, rows } = table_(sheet_(NEEDS));
  const col = k => head.indexOf(k);
  const get = (r, k) => (col(k) < 0 ? '' : String(r[col(k)]).trim());
  const needs = rows
    .map(r => ({ id: get(r, 'id'), need: get(r, 'need'), details: get(r, 'details'), amount: get(r, 'amount'), status: get(r, 'status').toLowerCase() }))
    .filter(n => n.need && !HIDDEN.includes(n.status));
  return { needs };
}

// Simple trigger: runs on every hand edit in the sheet (no setup needed) and drops the cached tree.
function onEdit() { CacheService.getScriptCache().remove('needs'); }

/* ── Write: new needs and pledges ── */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.website) return json_({ ok: true }); // honeypot field: bots fill it, people never see it
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const res = d.action === 'need' ? addNeed_(d) : d.action === 'pledge' ? addPledge_(d) : { ok: false, error: 'unknown_action' };
    if (res.ok) CacheService.getScriptCache().remove('needs');   // a claim shows on the tree straight away
    return json_(res);
  } finally {
    lock.releaseLock();
  }
}

function addNeed_(d) {
  const need = clean_(d.need, 60), name = clean_(d.name, 80), contact = clean_(d.contact, 120);
  if (!need || !name || !contact) return { ok: false, error: 'missing' };
  const sh = sheet_(NEEDS), { head, rows } = table_(sh);
  const ic = head.indexOf('id');
  const id = rows.reduce((m, r) => Math.max(m, parseInt(r[ic], 10) || 0), 0) + 1;
  const rec = {
    id, need, details: clean_(d.details, 400), amount: clean_(d.amount, 12).replace(/[^0-9.]/g, ''),
    status: 'new', submitted_at: new Date(), submitted_by: name, contact, relationship: clean_(d.relationship, 30),
  };
  sh.appendRow(head.map(h => safe_(rec[h])));
  notify_('New Giving Tree need: ' + need,
    need + (rec.details ? '\n\n' + rec.details : '') + (rec.amount ? '\n\nAbout $' + rec.amount : '') +
    '\n\nFrom: ' + name + ' (' + contact + ')\nFor: ' + (rec.relationship || 'not given') +
    '\n\nIt is on the Needs tab as tag ' + id + ' with status "new". Change it to "open" to put it on the tree.');
  return { ok: true, id };
}

function addPledge_(d) {
  const tag = String(parseInt(d.tag, 10) || ''), name = clean_(d.name, 80), contact = clean_(d.contact, 120);
  if (!tag || !name || !contact) return { ok: false, error: 'missing' };
  const sh = sheet_(NEEDS), { head, rows } = table_(sh);
  const ic = head.indexOf('id'), sc = head.indexOf('status'), nc = head.indexOf('need');
  const i = rows.findIndex(r => String(parseInt(r[ic], 10)) === tag);
  if (i < 0) return { ok: false, error: 'not_found' };
  const status = String(rows[i][sc]).trim().toLowerCase();
  if (status && status !== 'open') return { ok: false, error: 'taken' };
  sh.getRange(i + 2, sc + 1).setValue('claimed');
  const p = sheet_(PLEDGES);
  const rec = { submitted_at: new Date(), tag, need: rows[i][nc], name, contact, dropoff: clean_(d.dropoff, 80) };
  p.appendRow(table_(p).head.map(h => safe_(rec[h])));
  notify_('Giving Tree tag ' + tag + ' claimed',
    'Tag ' + tag + ': ' + rec.need + '\n\nClaimed by: ' + name + ' (' + contact + ')' +
    (rec.dropoff ? '\nDrop-off: ' + rec.dropoff : '') + '\n\nMark it "filled" on the Needs tab once the gift arrives.');
  return { ok: true };
}

/* ── Helpers ── */
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function table_(sh) {
  const v = sh.getDataRange().getValues();
  return { head: v.shift().map(h => String(h).trim().toLowerCase()), rows: v };
}

function ensureHeaders_(sh, cols) {
  const width = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim().toLowerCase());
  cols.forEach(c => {
    if (head.includes(c)) return;
    const at = head[0] === '' && head.length === 1 ? 1 : head.length + 1;
    sh.getRange(1, at).setValue(c);
    if (at === 1) head[0] = c; else head.push(c);
  });
  sh.setFrozenRows(1);
}

function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max); }

// Stop text that starts with = + - @ from being treated as a spreadsheet formula.
function safe_(v) {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

// NOTIFY_EMAIL can hold several addresses, separated by commas, semicolons or spaces.
// Email trouble never blocks a submission: the row is already saved by the time this runs.
function notify_(subject, body) {
  const raw = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || '';
  const to = raw.split(/[\s,;]+/).map(a => a.trim()).filter(a => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a));
  if (!to.length) return;
  try { MailApp.sendEmail(to.join(','), subject, body); }
  catch (err) { console.error('Notification email failed: ' + err); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
