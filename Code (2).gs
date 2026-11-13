/**
 * Fusion Companion — Google Sign-In user log
 * Paste into: Google Sheet → Extensions → Apps Script → Code.gs
 * Deploy → New deployment → Web app · Execute as: Me · Who has access: Anyone
 */

// 1) Paste your Google OAuth Client ID here (same one as in index.html)
const CLIENT_ID = '667012000264-4tsu4kcsm0ldgpcp2al41ol0rqlna8vg.apps.googleusercontent.com';

// 2) These emails are always Manager + Active (you)
const OWNER_EMAILS = ['pablo.duque.pd@gmail.com', 'pablo.duquepd@gmail.com', 'pablo@fusioncrsfl.com'];

// 3) New sign-up alerts go to this email (leave '' to turn off)
const NOTIFY_EMAIL = 'pablo.duquepd@gmail.com';

const TECH_HEADERS = ['Email','Name','Role','Status','First Sign-In','Last Opened','Opens','Device','Photo','Session'];

function doGet() {
  return ContentService.createTextOutput('Fusion Companion login service is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'login') out = login_(req);
    else if (req.action === 'ping') out = ping_(req);
    else if (req.action === 'list') out = list_(req);
    else if (req.action === 'setStatus') out = setStatus_(req);
    else if (req.action === 'score') out = score_(req);
    else if (req.action === 'lesson') out = lesson_(req);
    else out = { ok: false, error: 'bad_action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function techs_() { return sheet_('Techs', TECH_HEADERS); }
function log_(email, name, event, device) {
  sheet_('Log', ['Time','Email','Name','Event','Device']).appendRow([new Date(), email, name, event, device || '']);
}
function findRow_(sh, email) {
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === email) return { row: i + 1, vals: v[i] };
  }
  return null;
}
function isOwner_(email) {
  return OWNER_EMAILS.map(function (x) { return x.toLowerCase(); }).indexOf(email) !== -1;
}

// Ask Google to confirm the sign-in token is real and was issued for our app
function verify_(credential) {
  if (!credential) throw 'missing_token';
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw 'invalid_token';
  const t = JSON.parse(r.getContentText());
  if (t.aud !== CLIENT_ID) throw 'wrong_client_id';
  if (String(t.email_verified) !== 'true') throw 'email_not_verified';
  return t;
}

function login_(req) {
  const t = verify_(req.credential);
  const email = String(t.email).toLowerCase();
  const name = t.name || email;
  const sh = techs_();
  let f = findRow_(sh, email);
  if (!f) {
    const owner = isOwner_(email);
    sh.appendRow([email, name, owner ? 'Manager' : 'Tech', owner ? 'Active' : 'Pending', new Date(), '', 0, req.device || '', t.picture || '', '']);
    f = findRow_(sh, email);
    if (!owner && NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, 'Fusion Companion: ' + name + ' is waiting for approval',
          name + ' (' + email + ') signed in to Fusion Companion and is waiting for approval.\n\n' +
          'Approve them in the app: Manager dashboard → Pending, or change their Status to Active in the Techs sheet.');
      } catch (e) {}
    }
  }
  const status = String(f.vals[3] || 'Pending');
  if (status !== 'Active') {
    log_(email, name, 'Sign-in blocked (' + status + ')', req.device);
    return { ok: false, reason: status === 'Pending' ? 'pending' : 'inactive', name: name, email: email };
  }
  const session = Utilities.getUuid();
  const r = f.row;
  if (!f.vals[4]) sh.getRange(r, 5).setValue(new Date());
  sh.getRange(r, 2).setValue(name);
  sh.getRange(r, 6).setValue(new Date());
  sh.getRange(r, 7).setValue((Number(f.vals[6]) || 0) + 1);
  sh.getRange(r, 8).setValue(req.device || '');
  sh.getRange(r, 9).setValue(t.picture || '');
  sh.getRange(r, 10).setValue(session);
  log_(email, name, 'Sign-in', req.device);
  return { ok: true, email: email, name: name, picture: t.picture || '', role: String(f.vals[2] || 'Tech'), session: session };
}

function auth_(req) {
  const f = findRow_(techs_(), String(req.email || '').toLowerCase());
  if (!f || !req.session || String(f.vals[9]) !== String(req.session) || String(f.vals[3]) !== 'Active') throw 'not_signed_in';
  return f;
}

function ping_(req) {
  const f = auth_(req);
  const sh = techs_();
  sh.getRange(f.row, 6).setValue(new Date());
  sh.getRange(f.row, 7).setValue((Number(f.vals[6]) || 0) + 1);
  if (req.device) sh.getRange(f.row, 8).setValue(req.device);
  return { ok: true, role: String(f.vals[2] || 'Tech') };
}

function scores_() {
  return sheet_('Scores', ['Time','Email','Name','Quiz ID','Quiz','Score','Total','Percent','Passed']);
}

function score_(req) {
  const f = auth_(req);
  const total = Number(req.total) || 0, score = Number(req.score) || 0;
  const pct = total ? Math.round(100 * score / total) : 0;
  scores_().appendRow([req.ts ? new Date(req.ts) : new Date(), f.vals[0], f.vals[1], String(req.quiz || ''), String(req.title || ''), score, total, pct, pct >= 80 ? 'Yes' : 'No']);
  return { ok: true };
}

function lessons_() {
  return sheet_('Lessons', ['Time','Email','Name','Lesson ID','Lesson','Module']);
}

function lesson_(req) {
  const f = auth_(req);
  lessons_().appendRow([req.ts ? new Date(req.ts) : new Date(), f.vals[0], f.vals[1], String(req.lesson || ''), String(req.title || ''), String(req.module || '')]);
  return { ok: true };
}

function list_(req) {
  const f = auth_(req);
  if (String(f.vals[2]) !== 'Manager') throw 'not_manager';
  const v = techs_().getDataRange().getValues().slice(1);
  // training summary from Scores sheet
  const sv = scores_().getDataRange().getValues().slice(1);
  const tr = {};
  sv.forEach(function (r) {
    const em = String(r[1]).toLowerCase(); if (!em) return;
    tr[em] = tr[em] || { passedSet: {}, attempts: 0, lessonSet: {}, exam: null };
    if (String(r[3]) === 'stt_exam') { tr[em].exam = Math.max(tr[em].exam || 0, Number(r[7]) || 0); return; }
    tr[em].attempts++;
    if (String(r[8]) === 'Yes') tr[em].passedSet[String(r[3])] = true;
  });
  lessons_().getDataRange().getValues().slice(1).forEach(function (r) {
    const em = String(r[1]).toLowerCase(); if (!em) return;
    tr[em] = tr[em] || { passedSet: {}, attempts: 0, lessonSet: {}, exam: null };
    tr[em].lessonSet[String(r[3])] = true;
  });
  const recent = sv.slice(-25).reverse().map(function (r) {
    return { time: r[0] ? new Date(r[0]).toISOString() : '', email: String(r[1]), name: String(r[2]), title: String(r[4]),
             score: Number(r[5]) || 0, total: Number(r[6]) || 0, pct: Number(r[7]) || 0, passed: String(r[8]) === 'Yes' };
  });
  return {
    ok: true,
    recent: recent,
    techs: v.filter(function (r) { return r[0]; }).map(function (r) {
      const t = tr[String(r[0]).toLowerCase()];
      return {
        email: String(r[0]), name: String(r[1]), role: String(r[2]), status: String(r[3]),
        first: r[4] ? new Date(r[4]).toISOString() : '', last: r[5] ? new Date(r[5]).toISOString() : '',
        opens: Number(r[6]) || 0, device: String(r[7]), photo: String(r[8]),
        training: t ? { passed: Object.keys(t.passedSet).length, attempts: t.attempts, lessons: Object.keys(t.lessonSet).length, exam: t.exam } : null
      };
    })
  };
}

function setStatus_(req) {
  const f = auth_(req);
  if (String(f.vals[2]) !== 'Manager') throw 'not_manager';
  if (['Active', 'Pending', 'Inactive'].indexOf(req.status) === -1) throw 'bad_status';
  const target = String(req.target || '').toLowerCase();
  if (isOwner_(target) && req.status !== 'Active') throw 'cannot_lock_owner';
  const sh = techs_();
  const t = findRow_(sh, target);
  if (!t) throw 'not_found';
  sh.getRange(t.row, 4).setValue(req.status);
  if (req.status !== 'Active') sh.getRange(t.row, 10).setValue('');
  log_(target, t.vals[1], 'Status → ' + req.status + ' (by ' + req.email + ')', '');
  return { ok: true };
}
