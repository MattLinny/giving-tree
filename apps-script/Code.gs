/**
 * Christmas Giving Tree: Google Sheet backend.
 *
 * Setup (once):
 *   1. Open the Needs spreadsheet → Extensions → Apps Script. Paste this whole file over Code.gs and save.
 *   2. Pick `setup` in the function menu and click Run. Approve the permissions it asks for.
 *      It adds the private columns to the Needs tab and creates a Pledges tab.
 *   3. Optional: Project Settings → Script properties → add NOTIFY_EMAIL = the address(es) that should hear
 *      about new needs and claims. The first one is also the reply-to on givers' receipts.
 *   Givers get a receipt email when they claim a tag; edit CHURCH and the DROPOFF_ lines below to suit.
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

// Shown in the email someone gets after claiming a tag. Edit to match how drop-off works at your church.
const CHURCH = 'Oceanside Church';
const DROPOFF_WEEKDAY = 'the church office';                        // e.g. 'the church office, 123 Example St'
const DROPOFF_SUNDAY = 'church on Sunday at Nanaimo Christian School';

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
  const tagNo = pad_(id);
  notify_('New need · Tag ' + tagNo + ' · ' + need + (rec.amount ? ' · $' + rec.amount : ''), 'New need to review: Tag ' + tagNo, [
    ['The need', [['Tag', tagNo], ['Need', need], ['Rough cost', rec.amount ? '$' + rec.amount : ''], ['Details', rec.details]]],
    ['Submitted by', [['Name', name], ['Contact', contact], ['For', rec.relationship || 'Not given']]],
    ['To do', null, [
      'Read it over and contact ' + name + ' if anything needs checking.',
      'On the Needs tab, change Tag ' + tagNo + ' from "new" to "open" to hang it on the tree.',
      'Not a fit? Set it to "removed" and it stays off the tree.',
    ]],
  ]);
  return { ok: true, id };
}

function addPledge_(d) {
  const tag = String(parseInt(d.tag, 10) || ''), name = clean_(d.name, 80);
  const email = clean_(d.email, 120), phone = clean_(d.phone, 40);
  const contact = clean_(d.contact, 120) || [email, phone].filter(String).join(' · ');
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

  const cell = k => (head.indexOf(k) < 0 ? '' : String(rows[i][head.indexOf(k)]).trim());
  const amount = cell('amount').replace(/[^0-9.]/g, ''), tagNo = pad_(tag);
  const gift = rec.need + (amount ? ' · $' + amount : '');
  const sunday = /at church/i.test(rec.dropoff);
  const where = sunday ? DROPOFF_SUNDAY : DROPOFF_WEEKDAY;
  const receipt = isEmail_(email);
  if (receipt) receipt_(email, name, tagNo, gift, rec.need, amount, cell('details'), rec.dropoff, where);

  notify_('Claimed · Tag ' + tagNo + ' · ' + gift + (rec.dropoff ? ' · ' + rec.dropoff.split(' · ')[0] : ''), 'Tag ' + tagNo + ' was claimed', [
    ['The gift', [['Tag', tagNo], ['Need', rec.need], ['Amount', amount ? '$' + amount : ''], ['For', cell('for')], ['Details', cell('details')]]],
    ['Claimed by', [['Name', name], ['Email', email], ['Phone', phone], ['Contact', email || phone ? '' : contact],
      ['Receipt', receipt ? 'Sent to ' + email : 'Not sent (no email given)']]],
    ['Drop-off', [['Day', rec.dropoff || 'Not picked'], ['Where', where]]],
    ['To do', null, [
      'Watch for the gift' + (rec.dropoff ? ' on ' + rec.dropoff.split(' · ')[0] : '') + '. It should come in an envelope or bag marked "Tag ' + tagNo + '".',
      'When it arrives, change Tag ' + tagNo + ' from "claimed" to "filled" on the Needs tab.',
      'If it doesn\u2019t come in, contact ' + name + ', or set the tag back to "open" to put it back on the tree.',
    ]],
  ]);
  return { ok: true, receipt };
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

function pad_(n) { return String(n).padStart(2, '0'); }
function isEmail_(a) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a || ''); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// NOTIFY_EMAIL can hold several addresses, separated by commas, semicolons or spaces.
function admins_() {
  const raw = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || '';
  return raw.split(/[\s,;]+/).map(a => a.trim()).filter(isEmail_);
}

// Email trouble never blocks a submission: the row is already saved by the time this runs.
function send_(opts) {
  try { MailApp.sendEmail(Object.assign({ name: CHURCH + ' · Giving Tree' }, opts)); }
  catch (err) { console.error('Email to ' + opts.to + ' failed: ' + err); }
}

/* Email layout: cream card, red headings, one table per section. Plain text goes along for mail apps that want it. */
const INK = '#4A1A10', RED = '#C22A0F', CREAM = '#FFF4E2';
function frame_(inner) {
  return '<div style="background:' + CREAM + ';padding:28px 12px;font-family:Georgia,\'Times New Roman\',serif;color:' + INK + '">' +
    '<div style="max-width:540px;margin:0 auto;background:#FFFBF4;border:2px solid ' + RED + ';border-radius:16px;padding:26px 26px 22px">' +
    '<div style="font:600 11px/1 Menlo,Consolas,monospace;letter-spacing:.16em;text-transform:uppercase;color:' + RED + ';margin-bottom:14px">Christmas Giving Tree</div>' +
    inner + '</div></div>';
}
function heading_(t) {
  return '<div style="margin:22px 0 8px;padding-bottom:6px;border-bottom:2px dashed rgba(194,42,15,.4);font:700 12px/1 Arial,Helvetica,sans-serif;letter-spacing:.12em;text-transform:uppercase;color:' + RED + '">' + esc_(t) + '</div>';
}
function rowsHtml_(rows) {
  return '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:15px;line-height:1.4">' +
    rows.filter(r => r[1]).map(r => '<tr><td style="padding:5px 12px 5px 0;width:96px;vertical-align:top;color:' + RED + ';font:600 12px/1.6 Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.06em">' +
      esc_(r[0]) + '</td><td style="padding:5px 0;vertical-align:top">' + esc_(r[1]) + '</td></tr>').join('') + '</table>';
}
function stepsHtml_(steps) {
  return '<ol style="margin:4px 0 0;padding-left:20px;font-size:15px;line-height:1.5">' + steps.map(s => '<li style="margin:0 0 6px">' + esc_(s) + '</li>').join('') + '</ol>';
}
function sectionsText_(sections) {
  return sections.map(([title, rows, steps]) => title.toUpperCase() + '\n' +
    (rows ? rows.filter(r => r[1]).map(r => r[0] + ': ' + r[1]).join('\n') : steps.map((s, k) => (k + 1) + '. ' + s).join('\n'))).join('\n\n');
}

// Staff email: a headline, headed sections ([title, rows] or [title, null, steps]) and a button to the sheet.
function notify_(subject, headline, sections) {
  const to = admins_();
  if (!to.length) return;
  const url = SpreadsheetApp.getActive().getUrl();
  const html = frame_('<h1 style="margin:0;font:700 24px/1.15 Georgia,\'Times New Roman\',serif;color:' + RED + '">' + esc_(headline) + '</h1>' + sections.map(([title, rows, steps]) => heading_(title) + (rows ? rowsHtml_(rows) : stepsHtml_(steps))).join('') +
    '<p style="margin:22px 0 0"><a href="' + url + '" style="display:inline-block;padding:11px 20px;border-radius:999px;background:' + RED + ';color:' + CREAM + ';text-decoration:none;font:700 13px/1 Arial,Helvetica,sans-serif;letter-spacing:.06em;text-transform:uppercase">Open the sheet</a></p>');
  send_({ to: to.join(','), subject, htmlBody: html, body: headline + '\n\n' + sectionsText_(sections) + '\n\nThe sheet: ' + url });
}

// The giver's receipt: thanks, what they claimed, and exactly what to do next. Replies go to the staff.
function receipt_(to, name, tagNo, gift, need, amount, details, dropoff, where) {
  const first = name.split(/\s+/)[0];
  const day = dropoff ? dropoff.split(' · ')[0] : '';
  const rows = [['Tag', tagNo + ' · ' + gift], ['Drop-off', dropoff || 'Not picked yet'], ['Where', where], ['Details', details]];
  const steps = [
    'Pick up the gift: ' + need + (amount ? ', about $' + amount : '') + '.',
    'Put it in an envelope or bag and write "Tag ' + tagNo + '" on the front.',
    'Bring it to ' + where + (day ? ' on ' + day : '') + '.',
  ];
  const html = frame_(
    '<h1 style="margin:0 0 10px;font:700 26px/1.15 Georgia,\'Times New Roman\',serif;color:' + RED + '">Thank you so much for your generosity.</h1>' +
    '<p style="margin:0;font-size:16px;line-height:1.5">Hi ' + esc_(first) + ', Tag ' + esc_(tagNo) + ' is yours. Here\u2019s everything you need.</p>' +
    heading_('Your tag') + rowsHtml_(rows) + heading_('What to do') + stepsHtml_(steps) +
    '<p style="margin:20px 0 0;font-size:15px;line-height:1.5">Can\u2019t make that day, or have a question? Just reply to this email.</p>' +
    '<p style="margin:14px 0 0;font-size:15px">' + esc_(CHURCH) + '</p>');
  const text = 'Thank you so much for your generosity.\n\nHi ' + first + ', Tag ' + tagNo + ' is yours. Here\u2019s everything you need.\n\n' +
    sectionsText_([['Your tag', rows], ['What to do', null, steps]]) +
    '\n\nCan\u2019t make that day, or have a question? Just reply to this email.\n\n' + CHURCH;
  const replyTo = admins_()[0];
  send_(Object.assign({ to, subject: 'Thank you! Tag ' + tagNo + ' is yours', htmlBody: html, body: text }, replyTo ? { replyTo } : {}));
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
