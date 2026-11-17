/**
 * Lalitha Naturals Payables – Google Sheet backup (free, runs in your own Google account)
 *
 * SETUP (one time, ~3 minutes):
 *  1. Create a new Google Sheet, e.g. "LN Payables Backup".
 *  2. Extensions → Apps Script. Delete what's there and paste this whole file.
 *  3. Change SYNC_KEY below to your own secret word (letters/numbers, 8+ characters).
 *  4. Deploy → New deployment → type "Web app".
 *       Execute as: Me      Who has access: Anyone
 *     Click Deploy, allow access, and copy the Web app URL (ends with /exec).
 *  5. In the app: Settings → Backup & Sync → paste the URL and the same secret word → Save → Back up now.
 *
 * EMAIL ALERTS (free):
 *  6. In Apps Script, pick "setupEmails" in the function list at the top and press Run (allow access once).
 *  7. Everything else is set inside the app: Settings → Email alerts (who gets it, what time,
 *     how many days ahead, which branch, Monday weekly summary). Use "Send test" to check.
 *     The Sheet reads those settings from the latest backup, so keep "Back up automatically" ON.
 *
 * The secret word lives only here (your Google account) and on your phones – never in the app code on GitHub.
 */
const SYNC_KEY = 'CHANGE-ME-to-your-secret';
const KEEP_BACKUPS = 15;          // older backups are deleted automatically
const CHUNK = 45000;              // a Google Sheet cell holds max 50,000 characters

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.key !== SYNC_KEY) return out({ ok: false, error: 'Wrong secret word' });
    if (body.action === 'ping') return out({ ok: true });
    if (body.action === 'testEmail') {
      const d = body.data && body.data.data ? body.data.data : body.data;
      sendDueEmail(d, new Date(), true);
      return out({ ok: true, to: recipients(d) });
    }
    if (body.action !== 'backup') return out({ ok: false, error: 'Unknown action' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const raw = getSheet(ss, '_backups', ['Backup ID', 'Saved at', 'Device', 'Part', 'Data']);
      const id = 'B' + Date.now();
      const json = JSON.stringify(body.data);
      const rows = [];
      for (let i = 0, part = 0; i < json.length; i += CHUNK, part++) {
        rows.push([id, new Date(), body.device || '', part, json.substring(i, i + CHUNK)]);
      }
      raw.getRange(raw.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
      prune(raw);
      writeReadable(ss, body.data);
      return out({ ok: true, id: id, parts: rows.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if ((e.parameter.key || '') !== SYNC_KEY) return out({ ok: false, error: 'Wrong secret word' });
  if (e.parameter.action !== 'latest') return out({ ok: true });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('_backups');
  if (!raw || raw.getLastRow() < 2) return out({ ok: false, error: 'No backup found yet' });
  const values = raw.getRange(2, 1, raw.getLastRow() - 1, 5).getValues();
  const lastId = values[values.length - 1][0];
  const parts = values.filter(r => r[0] === lastId).sort((a, b) => a[3] - b[3]);
  return out({ ok: true, id: lastId, savedAt: parts[0][1], data: JSON.parse(parts.map(r => r[4]).join('')) });
}

// Keeps only the newest KEEP_BACKUPS backups
function prune(raw) {
  const n = raw.getLastRow() - 1;
  if (n < 1) return;
  const ids = raw.getRange(2, 1, n, 1).getValues().map(r => r[0]);
  const unique = [...new Set(ids)];
  if (unique.length <= KEEP_BACKUPS) return;
  const keep = new Set(unique.slice(-KEEP_BACKUPS));
  const firstKeep = ids.findIndex(id => keep.has(id));
  if (firstKeep > 0) raw.deleteRows(2, firstKeep);
}

// Easy-to-read tabs (photos are left out here; they are inside the backup itself)
function writeReadable(ss, d) {
  const vName = {};
  (d.vendors || []).forEach(v => vName[v.id] = v.name);
  const billNo = {};
  (d.bills || []).forEach(b => billNo[b.id] = b.billNo);
  fill(getSheet(ss, 'Vendors', ['Name', 'Contact', 'Phone', 'GSTIN', 'Opening Balance', 'Address']),
    (d.vendors || []).map(v => [v.name, v.contactPerson || '', v.phone || '', v.gstin || '', Number(v.openingBalance || 0), v.address || '']));
  fill(getSheet(ss, 'Bills', ['Vendor', 'Bill No', 'Branch', 'Brands', 'Bill Date', 'Received', 'Due Date', 'Amount', 'Notes', 'Has Attachment']),
    (d.bills || []).map(b => [vName[b.vendorId] || '', b.billNo, b.branch || '', (b.brands || []).join(', '), b.billDate, b.receivedDate, b.dueDate, Number(b.amount), b.notes || '', b.photo ? 'Yes' : '']));
  fill(getSheet(ss, 'Payments', ['Date', 'Vendor', 'Mode', 'Reference', 'Bill', 'Amount', 'Notes']),
    (d.payments || []).map(p => [p.date, vName[p.vendorId] || '', p.mode, p.reference || '', billNo[p.billId] || 'On Account', Number(p.amount), p.notes || '']));
  fill(getSheet(ss, 'Credit Notes', ['Date', 'Vendor', 'Credit Note No', 'Bill', 'Reason', 'Amount']),
    (d.credits || []).map(c => [c.date, vName[c.vendorId] || '', c.creditNo, billNo[c.billId] || 'General', c.reason || '', Number(c.amount)]));
}

function getSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function fill(sh, rows) {
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Email alerts (settings come from the app: Settings → Email alerts) ================= */

// Run ONCE: Apps Script → choose "setupEmails" at the top → Run → allow access.
function setupEmails() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['hourlyEmailCheck', 'dailyDueEmail'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('hourlyEmailCheck').timeBased().everyHours(1).create();
  const latest = latestBackupData();
  if (latest) sendDueEmail(latest.data, latest.savedAt, true);   // one now so you can see it
}
function setupDailyEmail() { setupEmails(); }   // old name still works

// Runs every hour; sends only at the hour chosen in the app, once per day
function hourlyEmailCheck() {
  const latest = latestBackupData();
  if (!latest) return;
  const em = emailCfg(latest.data);
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, 'Asia/Kolkata', 'H'));
  const today = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  if (em.enabled && hour === Number(em.hour) && props.getProperty('lastDaily') !== today) {
    sendDueEmail(latest.data, latest.savedAt, false);
    props.setProperty('lastDaily', today);
  }
  const isMonday = Utilities.formatDate(now, 'Asia/Kolkata', 'u') === '1';
  if (em.weekly && isMonday && hour === Number(em.hour) && props.getProperty('lastWeekly') !== today) {
    sendWeeklyEmail(latest.data, latest.savedAt);
    props.setProperty('lastWeekly', today);
  }
}

function emailCfg(d) {
  const e = (d && d.config && d.config.email) || {};
  return { enabled: !!e.enabled, to: e.to || '', hour: e.hour === undefined ? 8 : e.hour, daysAhead: e.daysAhead || 3,
           branch: e.branch || 'all', weekly: e.weekly !== false, skipEmpty: !!e.skipEmpty };
}
function recipients(d) {
  const list = emailCfg(d).to.split(/[,;\s]+/).map(x => x.trim()).filter(x => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x));
  return list.length ? list.join(',') : Session.getEffectiveUser().getEmail();
}
function bizName(d) { return (d && d.config && d.config.business && d.config.business.name) || 'Lalitha Naturals'; }

function latestBackupData() {
  const raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_backups');
  if (!raw || raw.getLastRow() < 2) return null;
  const values = raw.getRange(2, 1, raw.getLastRow() - 1, 5).getValues();
  const lastId = values[values.length - 1][0];
  const parts = values.filter(r => r[0] === lastId).sort((a, b) => a[3] - b[3]);
  const payload = JSON.parse(parts.map(r => r[4]).join(''));
  return { savedAt: parts[0][1], data: payload.data || payload };
}

// Pure function (no Google services) – works out what is due
function computeDueSummary(d, today, daysAhead, branch) {
  daysAhead = daysAhead === undefined ? 3 : daysAhead;
  const dayNum = s => { const p = String(s).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / 86400000; };
  const t = dayNum(today);
  const vName = {}; (d.vendors || []).forEach(v => vName[v.id] = v.name);
  const paid = {};
  (d.payments || []).concat(d.credits || []).forEach(x => { if (x.billId) paid[x.billId] = (paid[x.billId] || 0) + Number(x.amount); });
  const open = (d.bills || []).filter(b => !branch || branch === 'all' || (b.branch || '') === branch).map(b => ({
    vendor: vName[b.vendorId] || '', vendorId: b.vendorId, billNo: b.billNo, branch: b.branch || '', due: b.dueDate,
    late: t - dayNum(b.dueDate), bal: Math.round((Number(b.amount) - (paid[b.id] || 0)) * 100) / 100
  })).filter(x => x.bal > 0).sort((a, b) => b.late - a.late);
  const sum = arr => Math.round(arr.reduce((s, x) => s + x.bal, 0) * 100) / 100;
  const overdue = open.filter(x => x.late > 0), dueToday = open.filter(x => x.late === 0), soon = open.filter(x => x.late < 0 && x.late >= -daysAhead);
  return { open, overdue, dueToday, soon, totals: { overdue: sum(overdue), today: sum(dueToday), soon: sum(soon), all: sum(open) } };
}

function inr(n) {
  const neg = n < 0; n = Math.abs(n);
  const [i, f] = n.toFixed(2).split('.');
  const last3 = i.slice(-3), rest = i.slice(0, -3);
  return (neg ? '-' : '') + '₹' + (rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' : '') + last3 + '.' + f;
}
const fmtD = iso => iso ? String(iso).split('-').reverse().join('-') : '';

function buildDueEmail(d, today, savedAt, isTest) {
  const em = emailCfg(d);
  const s = computeDueSummary(d, today, em.daysAhead, em.branch);
  const table = (title, rows, color) => rows.length ? `<h3 style="color:${color};margin:18px 0 6px">${title}</h3>
    <table style="border-collapse:collapse;width:100%;font-size:13px">${rows.map(r =>
      `<tr><td style="padding:6px;border-bottom:1px solid #eee">${r.vendor}</td><td style="padding:6px;border-bottom:1px solid #eee">#${r.billNo}${r.branch ? ' · ' + r.branch.split(' ')[0] : ''}</td>
       <td style="padding:6px;border-bottom:1px solid #eee">Due ${fmtD(r.due)}${r.late > 0 ? ' (' + r.late + 'd late)' : ''}</td>
       <td style="padding:6px;border-bottom:1px solid #eee;text-align:right;font-weight:bold">${inr(r.bal)}</td></tr>`).join('')}</table>` : '';
  const nothing = !s.overdue.length && !s.dueToday.length && !s.soon.length;
  const branchTxt = em.branch && em.branch !== 'all' ? ` · ${em.branch}` : '';
  const html = `<div style="font-family:Arial,sans-serif;max-width:640px">
    ${isTest ? '<p style="background:#fef3c7;padding:8px;border-radius:6px">This is a test email from Settings.</p>' : ''}
    <h2 style="color:#16a34a;margin-bottom:4px">${bizName(d)} – bills to pay</h2>
    <p style="color:#64748b;margin-top:0">${fmtD(today)}${branchTxt} · Total outstanding <b>${inr(s.totals.all)}</b></p>
    <p>Overdue: <b style="color:#be123c">${inr(s.totals.overdue)}</b> · Due today: <b>${inr(s.totals.today)}</b> · Next ${em.daysAhead} day${em.daysAhead == 1 ? '' : 's'}: <b>${inr(s.totals.soon)}</b></p>
    ${table('Overdue', s.overdue, '#be123c')}${table('Due today', s.dueToday, '#ea580c')}${table(`Due in the next ${em.daysAhead} day${em.daysAhead == 1 ? '' : 's'}`, s.soon, '#0f172a')}
    ${nothing ? `<p>Nothing due in the next ${em.daysAhead} days. 🎉</p>` : ''}
    <p style="color:#94a3b8;font-size:11px;margin-top:20px">Based on the app backup from ${savedAt}. Change these emails in the app: Settings → Email alerts.</p></div>`;
  const subject = `${isTest ? '[Test] ' : ''}${bizName(d)} bills: ${inr(s.totals.overdue)} overdue, ${inr(s.totals.today)} due today`;
  return { html, subject, nothing };
}

function sendDueEmail(d, savedAt, isTest) {
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const m = buildDueEmail(d, today, savedAt, isTest);
  if (m.nothing && emailCfg(d).skipEmpty && !isTest) return;
  MailApp.sendEmail({ to: recipients(d), subject: m.subject, htmlBody: m.html });
}

function sendWeeklyEmail(d, savedAt) {
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const em = emailCfg(d);
  const s = computeDueSummary(d, today, 7, em.branch);
  const byV = {};
  s.open.forEach(x => { byV[x.vendor] = byV[x.vendor] || { n: 0, amt: 0, late: 0 }; byV[x.vendor].n++; byV[x.vendor].amt += x.bal; if (x.late > 0) byV[x.vendor].late += x.bal; });
  const rows = Object.keys(byV).sort((a, b) => byV[b].amt - byV[a].amt);
  const buckets = [['Not yet due', x => x.late <= 0], ['1–30 days late', x => x.late > 0 && x.late <= 30], ['31–60 days late', x => x.late > 30 && x.late <= 60], ['60+ days late', x => x.late > 60]];
  const html = `<div style="font-family:Arial,sans-serif;max-width:640px">
    <h2 style="color:#16a34a;margin-bottom:4px">${bizName(d)} – weekly payables summary</h2>
    <p style="color:#64748b;margin-top:0">Week of ${fmtD(today)} · Total outstanding <b>${inr(s.totals.all)}</b></p>
    <h3 style="margin:16px 0 6px">Aging</h3><table style="border-collapse:collapse;font-size:13px">${buckets.map(([l, f]) => {
      const r = s.open.filter(f); return `<tr><td style="padding:4px 12px 4px 0">${l} (${r.length})</td><td style="text-align:right;font-weight:bold">${inr(r.reduce((a, x) => a + x.bal, 0))}</td></tr>`; }).join('')}</table>
    <h3 style="margin:16px 0 6px">By vendor</h3><table style="border-collapse:collapse;width:100%;font-size:13px">
    <tr style="color:#64748b"><td style="padding:6px">Vendor</td><td>Bills</td><td style="text-align:right">Overdue</td><td style="text-align:right;padding:6px">Outstanding</td></tr>
    ${rows.map(v => `<tr><td style="padding:6px;border-top:1px solid #eee">${v}</td><td style="border-top:1px solid #eee">${byV[v].n}</td><td style="text-align:right;border-top:1px solid #eee;color:#be123c">${byV[v].late ? inr(byV[v].late) : '-'}</td><td style="text-align:right;padding:6px;border-top:1px solid #eee;font-weight:bold">${inr(byV[v].amt)}</td></tr>`).join('')}</table>
    <p style="color:#94a3b8;font-size:11px;margin-top:20px">Based on the app backup from ${savedAt}.</p></div>`;
  MailApp.sendEmail({ to: recipients(d), subject: `${bizName(d)} weekly payables: ${inr(s.totals.all)} outstanding`, htmlBody: html });
}
