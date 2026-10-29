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

/* ── Read: the tree's public data ── */
function doGet() {
  const { head, rows } = table_(sheet_(NEEDS));
  const col = k => head.indexOf(k);
  const get = (r, k) => (col(k) < 0 ? '' : String(r[col(k)]).trim());
  const needs = rows
    .map(r => ({ id: get(r, 'id'), need: get(r, 'need'), details: get(r, 'details'), amount: get(r, 'amount'), status: get(r, 'status').toLowerCase() }))
    .filter(n => n.need && !HIDDEN.includes(n.status));
  return json_({ needs });
}

/* ── Write: new needs and pledges ── */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.website) return json_({ ok: true }); // honeypot field: bots fill it, people never see it
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.action === 'need') return json_(addNeed_(d));
    if (d.action === 'pledge') return json_(addPledge_(d));
    return json_({ ok: false, error: 'unknown_action' });
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
