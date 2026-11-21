/**
 * Becky's Questionable Screening Process™ — back office
 * Lives inside your Google Sheet (Extensions → Apps Script).
 *
 * What it does
 *  - Saves every application as a row on the "Applications" tab.
 *  - Emails you when a new one arrives.
 *  - Lets your private inbox page read applications and save
 *    your status (Shortlisted / Maybe / No) and notes — but only
 *    with your secret inbox key.
 *
 * Setup (once)
 *  1. Paste this whole file in, replacing what's there. Save.
 *  2. Choose "setup" in the function menu at the top and press Run.
 *     Approve the permissions. Your inbox key appears in the log
 *     at the bottom — keep it somewhere safe.
 *  3. Deploy → New deployment → type: Web app.
 *     Execute as: Me.  Who has access: Anyone.  Deploy.
 *     Copy the Web app URL and give it to Claude.
 */

const SHEET_NAME = 'Applications';
const HEADERS = [
  'Received', 'ID', 'Status', 'Name', 'Phone', 'First date idea',
  'Q1 Why dating apps', 'Q2 Single how long', 'Q3 Friends would say', 'Q4 Saturday morning',
  'Q5 Escape room role', 'Q6 When we disagree', 'Q7 Spontaneous plans', 'Q8 Phone',
  'Q9 Showing interest', 'Q10 Sense of humour',
  'Q11 48-hour test', 'Q12 Ridiculous opinion', 'Q13 Idiot story', 'Q14 Learnt about self', 'Q15 Worth the jeans',
  'Availability', 'Timing notes', 'My notes', 'Slots (raw)'
];
const COL = { status: 3, notes: 24 }; // 1-based column numbers

function setup() {
  sheet_();
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('ADMIN_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('ADMIN_KEY', key);
  }
  Logger.log('Your inbox key is: ' + key);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function key_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'); }
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

/** Applications in, and status/notes updates from your inbox. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sh = sheet_();

    // Inbox: update status or notes (needs your key)
    if (d.action === 'update') {
      if (!key_() || d.key !== key_()) return out_({ ok: false, error: 'key' });
      const r = findRow_(sh, d.id);
      if (r < 0) return out_({ ok: false, error: 'not found' });
      if (d.status !== undefined) sh.getRange(r, COL.status).setValue(clip_(d.status, 20));
      if (d.notes !== undefined) sh.getRange(r, COL.notes).setValue(clip_(d.notes, 5000));
      return out_({ ok: true });
    }

    // Spam trap: real people never fill the hidden "website" box
    if (d.website) return out_({ ok: true });
    if (!d.name || !d.phone) return out_({ ok: false, error: 'missing' });

    const mcq = (d.mcq || []).slice(0, 10).map(x => clip_(x, 200));
    while (mcq.length < 10) mcq.push('');
    const written = (d.written || []).slice(0, 5).map(x => clip_(x, 8000));
    while (written.length < 5) written.push('');
    const id = clip_(d.id || Utilities.getUuid(), 40);

    const row = [
      new Date(), id, 'New', clip_(d.name, 80), "'" + clip_(d.phone, 30), clip_(d.firstDate, 1500),
      ...mcq, ...written,
      (d.availability || []).map(x => clip_(x, 60)).join('\n'),
      clip_(d.notes, 600), '', clip_(JSON.stringify(d.slots || {}), 4000)
    ];

    const existing = findRow_(sh, id);
    if (existing > 0) {
      const keep = sh.getRange(existing, 1, 1, HEADERS.length).getValues()[0];
      row[0] = keep[0]; row[2] = keep[2] || 'New'; row[23] = keep[23];
      sh.getRange(existing, 1, 1, HEADERS.length).setValues([row]);
    } else {
      sh.appendRow(row);
      try {
        MailApp.sendEmail({
          to: Session.getEffectiveUser().getEmail(),
          subject: 'New application: ' + clip_(d.name, 60),
          body: clip_(d.name, 80) + ' has applied.\n\nFirst date idea: ' + clip_(d.firstDate, 500) +
                '\n\nAvailable: ' + (d.availability || []).join(', ') + '\n\nOpen your inbox page to read the rest.'
        });
      } catch (err) { /* email is a bonus; the row is already saved */ }
    }
    return out_({ ok: true, id: id });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Your inbox reads applications here (needs your key). */
function doGet(e) {
  const k = e && e.parameter && e.parameter.key;
  if (!key_() || k !== key_()) return out_({ ok: false, error: 'key' });
  const values = sheet_().getDataRange().getValues();
  values.shift();
  const rows = values.filter(r => r[1]).map(r => {
    let slots = {};
    try { slots = JSON.parse(r[24] || '{}'); } catch (err) {}
    return {
      received: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      id: String(r[1]), status: String(r[2] || 'New'), name: String(r[3]), phone: String(r[4]),
      firstDate: String(r[5]), mcq: r.slice(6, 16).map(String), written: r.slice(16, 21).map(String),
      availability: String(r[21] || '').split('\n').filter(String), timing: String(r[22] || ''),
      myNotes: String(r[23] || ''), slots: slots
    };
  });
  return out_({ ok: true, rows: rows });
}
