var CAREERPULSE_BACKEND_VERSION = '2026-10-07-family-db-v7';
/* CareerPulse backend (Google Apps Script) — v7
   SCRIPT PROPERTIES: TOKEN, ADZUNA_ID, ADZUNA_KEY, SUPABASE_URL, SUPABASE_SERVICE_KEY
   OPTIONAL: WA_KEY (CallMeBot), ADMIN_EMAILS (comma-separated allow-list for Admin.html — STRONGLY recommended)
   After ANY edit: Deploy > Manage deployments > Edit > New version > Deploy.
   Triggers: run setup() once for Jobs, setupFamilySupabaseSender() once for Family. */

var P = PropertiesService.getScriptProperties();
function g(k, d) { var v = P.getProperty(k); return (v === null || v === '') ? d : v; }
function flag(v) { return !!v && v !== '0' && v !== 'false'; }
function jsonOut_(v) { return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON); }
function normLoc_(s) { return String(s).trim().replace(/^delhi[\s-]*ncr$/i, 'Delhi'); }   // FIX: regex was mangled, never matched
function checkAdzunaKeys() {
  if (!g('ADZUNA_ID', '') || !g('ADZUNA_KEY', '')) throw new Error('Add ADZUNA_ID and ADZUNA_KEY in Project Settings > Script properties.');
}
/* ================= SENDER IDENTITY (shown on every email) =================
   Script properties (or Admin > Sender): SENDER_NAME, REPLY_TO, CONTACT_PHONE, CONTACT_EMAIL, CONTACT_WEBSITE.
   Note: Apps Script always sends FROM the Google account that owns/runs this script; it cannot spoof another address.
   What we CAN set is the display name and Reply-To, so replies go to the address you choose. */
function senderInfo_() {
  var owner = ''; try { owner = Session.getEffectiveUser().getEmail(); } catch (e) {}
  var ok = function (v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : ''; };
  var reply = ok(g('REPLY_TO', '').trim());
  return { name: g('SENDER_NAME', '').trim() || 'CareerPulse', replyTo: reply || owner, phone: g('CONTACT_PHONE', '').trim(),
    email: ok(g('CONTACT_EMAIL', '').trim()) || reply || owner, web: g('CONTACT_WEBSITE', '').trim() };
}
function footerParts_() {
  var i = senderInfo_(), t = [], h = [];
  if (i.phone) { t.push('Phone: ' + i.phone); h.push('📞 ' + familyHtml_(i.phone)); }
  if (i.email) { t.push('Email: ' + i.email); h.push('✉️ <a href="mailto:' + familyHtml_(i.email) + '">' + familyHtml_(i.email) + '</a>'); }
  if (i.web) { t.push('Web: ' + i.web); h.push('🌐 ' + familyHtml_(i.web)); }
  return { info: i, text: t, html: h };
}
function footerText_() {
  var f = footerParts_();
  return '--\nSent by ' + f.info.name + (f.text.length ? '\n' + f.text.join('\n') : '') + '\nJust reply to this email to reach us, or to change or stop these alerts.';
}
function footerHtml_() {
  var f = footerParts_();
  return '<hr style="border:0;border-top:1px solid #dde6ea;margin:18px 0"><p style="font-size:13px;color:#4b616c;line-height:1.6;margin:0">Sent by <b>' + familyHtml_(f.info.name) + '</b>' +
    (f.html.length ? '<br>' + f.html.join(' &nbsp;·&nbsp; ') : '') + '<br>Just reply to this email to reach us, or to change or stop these alerts.</p>';
}
/* One place that sends every email: display name + Reply-To + contact footer. */
function sendMail_(to, subject, html, plain) {
  var i = senderInfo_(), m = { to: to, subject: subject, body: String(plain || subject) + '\n\n' + footerText_(), name: i.name };
  if (i.replyTo) m.replyTo = i.replyTo;
  if (html) m.htmlBody = html + footerHtml_();
  MailApp.sendEmail(m);
}
function shortKey_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s)).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('').slice(0, 16);
}
function trigExists_(fn) { return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === fn; }); }
function dropTriggers_(fns) { ScriptApp.getProjectTriggers().forEach(function (t) { if (fns.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); }); }

/* ================= ROUTING ================= */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || '');
  try {
    switch (a) {
      case 'deployment_check': return jsonOut_({ ok: true, version: CAREERPULSE_BACKEND_VERSION, adminAuth: 'supabase', family: true, jobsDb: true });
      case 'family_ping': return jsonOut_({ ok: true, service: 'family', version: CAREERPULSE_BACKEND_VERSION });
      case 'backend_version': return jsonOut_({ ok: true, version: CAREERPULSE_BACKEND_VERSION });
      case 'family_status': return jsonOut_(familyPublicStatus_());
      case 'family_suggest': return jsonOut_(familySuggestions_(p.kind, p.q));
      case 'page_config': return jsonOut_(pageConfigPublic_(p.page));
      case 'page_manifest': return jsonOut_(pageManifestPublic_(p.page));
      case 'search_results': return jsonOut_(searchResults_(p.query));
      // Jobs page (TOKEN protected) — profile database
      case 'job_profiles_list': return jsonOut_(jobProfilesList_(p.token));
      case 'job_profile_save': return jsonOut_(jobProfileSave_(p.token, p.profile));
      case 'job_profile_delete': return jsonOut_(jobProfileDelete_(p.token, p.id));
      // Admin page (Supabase session protected)
      case 'portal_admin_list': return jsonOut_(portalAdminList_(p.auth));
      case 'admin_job_data': return jsonOut_(adminJobData_(p.auth));            // FIX: sender settings (was shadowed by profiles)
      case 'admin_job_profiles': return jsonOut_(adminJobProfilesDb_(p.auth));
      case 'admin_job_profile_delete': return jsonOut_(adminJobProfileDelete_(p.auth, p.id));
      case 'admin_job_profile_save': return jsonOut_(adminJobProfileSave_(p.auth, p.profile, p.apply));
      case 'admin_job_search': portalRequireAdmin_(p.auth); return jsonOut_({ ok: true, data: searchResults_(p.query) });
      case 'admin_dashboard': return jsonOut_(adminDashboard_(p.auth));
      case 'admin_sender_get': return jsonOut_(adminSenderGet_(p.auth));
      case 'admin_page_config': return jsonOut_(adminPageConfig_(p.auth, p.page));
      case 'admin_page_all': return jsonOut_(adminPageAllSections_(p.auth, p.page));
    }
    if (a || p.token !== undefined) return radarApi_(e);   // legacy Jobs router (status/save/on/off/test)
    return ContentService.createTextOutput('CareerPulse API is running.');
  } catch (err) { return jsonOut_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || '');
  try {
    switch (a) {
      case 'portal_admin_save': return jsonOut_(portalAdminSave_(p.auth, p.portal || p.config || p.data));
      case 'portal_admin_delete': return jsonOut_(portalAdminDelete_(p.auth, p.id));
      case 'portal_admin_toggle': return jsonOut_(portalAdminToggle_(p.auth, p.id, p.enabled));
      case 'portal_admin_test': return jsonOut_(portalAdminTest_(p.auth, p.portal || p.config || p.data));   // FIX: argument order
      case 'admin_page_draft_save': return jsonOut_(adminPageDraftSave_(p.auth, p.page, p.config));
      case 'admin_page_publish': return jsonOut_(adminPagePublish_(p.auth, p.page));
      case 'admin_page_draft_discard': return jsonOut_(adminPageDraftDiscard_(p.auth, p.page));
      case 'admin_builtin_save': return jsonOut_(adminBuiltinSave_(p.auth, p.page, p.config));
      case 'admin_builtin_reset': return jsonOut_(adminBuiltinReset_(p.auth, p.page));
      // FIX: large payloads (profiles, schedule) must be POSTed, not put in a GET URL
      case 'admin_job_profile_save': return jsonOut_(adminJobProfileSave_(p.auth, p.profile, p.apply));
      case 'admin_job_profile_delete': return jsonOut_(adminJobProfileDelete_(p.auth, p.id));
      case 'admin_sender_save': return jsonOut_(adminSenderSave_(p.auth, p.config));
      case 'admin_sender_test': return jsonOut_(adminSenderTest_(p.auth));
      case 'job_profile_save': return jsonOut_(jobProfileSave_(p.token, p.profile));
      case 'job_profile_delete': return jsonOut_(jobProfileDelete_(p.token, p.id));
    }
    if (p.token !== undefined) return radarApi_(e);
    return jsonOut_({ ok: false, error: 'Unknown POST action' });
  } catch (ex) { return jsonOut_({ ok: false, error: String(ex && ex.message || ex) }); }
}

/* ================= JOBS (legacy token API + scheduler) ================= */
function setOn(on) {
  P.setProperty('ENABLED', on ? '1' : '0');
  dropTriggers_(['digest', 'jobsTick']);
  if (!on) return;
  // Recipient-aware scheduler when the Jobs page has saved recipients; legacy daily digest otherwise.
  if (jobsRecipients_().length) ScriptApp.newTrigger('jobsTick').timeBased().everyMinutes(5).create();
  else ScriptApp.newTrigger('digest').timeBased().everyDays(1).atHour(+g('HOUR', '9')).create();
}
function setup() { setOn(true); }
function enable() { setOn(true); }
function disable() { setOn(false); }

function jsonProp_(k, d) { try { var v = JSON.parse(g(k, '')); return v == null ? d : v; } catch (e) { return d; } }
function jobsRecipients_() {
  var r = jsonProp_('RECIPIENTS_JSON', []);
  return Array.isArray(r) ? r.filter(function (x) { return x && x.address && x.enabled !== false; }) : [];
}

function jobsTick() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    if (g('ENABLED', '0') !== '1') return;
    var base = {
      times: jsonProp_('TIMES_JSON', []), repeat: g('REPEAT', 'daily'), repeatDays: jsonProp_('REPEAT_DAYS_JSON', []),
      timezone: g('TIMEZONE', Session.getScriptTimeZone()), days: g('DAYS', '1'),
      locations: g('LOC', '').split('|').filter(String), jobTypes: g('JT', '').split('|').filter(String),
      workModes: jsonProp_('WORK_MODES_JSON', []), shifts: jsonProp_('SHIFTS_JSON', [])
    };
    if (!base.times.length) base.times = [('0' + g('HOUR', '9')).slice(-2) + ':00'];
    var kw = g('KW', '').split('|').filter(String), now = new Date();
    jobsRecipients_().forEach(function (r) {
      try {
        var key = 'JOBS_' + shortKey_(r.channel + ':' + String(r.address).toLowerCase());
        var s = Object.assign({}, base, { repeats: [base.repeat], channel: r.channel === 'email' ? 'email' : 'whatsapp', address: r.address, firstName: r.firstName || '' });
        if (r.channel === 'whatsapp_group') { Logger.log('Skipping WhatsApp group ' + r.address + ' (not supported by CallMeBot).'); return; }
        s.roles = r.mode === 'specific' ? (r.roles || []) : r.mode === 'random' ? [] : kw;
        s.random = r.mode === 'random';
        var due = scheduleDueSlot_(key, s, now);
        if (due) sendReminder_(key, s, false, due);
      } catch (err) { Logger.log('Jobs reminder failed: ' + err.message); }
    });
  } finally { lock.releaseLock(); }
}

function testNow() {
  Logger.log('SETTINGS: KW=' + g('KW', '') + ' | LOC=' + g('LOC', '') + ' | JT=' + g('JT', '') + ' | DAYS=' + g('DAYS', '2') +
    ' | ENABLED=' + g('ENABLED', '0') + ' | ADZUNA set=' + !!(g('ADZUNA_ID') && g('ADZUNA_KEY')));
  var n = digest(true);
  Logger.log('Jobs found: ' + n);
  Logger.log('Errors: ' + (g('LASTERR', '') || 'none'));
  Logger.log('Triggers: ' + (ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join(',') || 'none'));   // FIX: precedence
}

/* Test / legacy digest: one email to EMAIL (or the owner), plus — for tests — every enabled recipient preview is skipped. */
function fetchJobs(ignoreSeen) {
  var seen = ignoreSeen ? [] : jsonProp_('SEEN', []);
  var kws = g('KW', '').split('|').filter(String);
  if (!kws.length) return { jobs: [], errors: ['No keywords saved. Open jobs.html, add keywords and press Save.'] };
  var res = fetchAdzunaJobs_({ roles: kws, locations: g('LOC', '').split('|').filter(String), days: g('DAYS', '2'),
    jobTypes: g('JT', '').split('|').filter(String), workModes: jsonProp_('WORK_MODES_JSON', []), remote: flag(g('REMOTE', '')) }, 100);
  res.jobs = res.jobs.filter(function (j) { return seen.indexOf(j.id) < 0; });
  return res;
}
function digest(force) {
  if (force !== true && g('ENABLED') !== '1') return 0;
  var to = g('EMAIL', '').split(/[\s,;]+/).filter(String)[0] || Session.getEffectiveUser().getEmail();
  var test = (force === true), r = fetchJobs(test), jobs = r.jobs, errors = r.errors, errText = errors.join(' || ');
  P.setProperty('LASTERR', errText);
  P.setProperty('LAST', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM HH:mm'));
  if (!jobs.length) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (test) sendMail_(to, 'Job radar: no new openings (test)', '', 'Test ran OK but found no openings.\n\nKeywords: ' + g('KW', '') + '\nLocations: ' + g('LOC', '') +
      '\nJob types: ' + g('JT', '') + '\nDays back: ' + g('DAYS', '2') + (errText ? '\n\nProblems:\n' + errors.join('\n') : '\n\nNo errors. Try wider filters.'));
    else if (errText && P.getProperty('ERRDAY') !== today) { P.setProperty('ERRDAY', today); sendMail_(to, 'Job radar: problem with your daily search', '', 'Problems:\n' + errors.join('\n')); }
    return 0;
  }
  var html = '<h3>' + jobs.length + ' new openings</h3>' + jobs.slice(0, 50).map(function (j) {
    return '<p><a href="' + familyHtml_(j.u) + '"><b>' + familyHtml_(j.t) + '</b></a><br>' + familyHtml_(j.c) + ' &middot; ' + familyHtml_(j.l) + '</p>';
  }).join('') + (errText ? '<p style="color:#b00">Some searches failed: ' + familyHtml_(errText) + '</p>' : '');
  sendMail_(to, 'Job radar: ' + jobs.length + ' new openings' + (test ? ' (test)' : ''), html, jobs.length + ' new openings:\n' + jobs.slice(0, 20).map(function (j) { return '- ' + j.t + ' - ' + j.c + ' ' + j.u; }).join('\n'));
  if (!test) P.setProperty('SEEN', JSON.stringify(jsonProp_('SEEN', []).concat(jobs.map(function (j) { return j.id; })).slice(-500)));
  return jobs.length;
}

function radarApi_(e) {
  var p = (e && e.parameter) || {}, result;
  if (!g('TOKEN', '') || p.token !== g('TOKEN')) return jsonOut_({ error: 'Wrong token' });
  var lock = LockService.getScriptLock(), locked = false;
  try {
    var action = p.action || 'status', sent;
    if (['status', 'save', 'on', 'off', 'test', 'search_results'].indexOf(action) < 0) throw new Error('Unknown API action');
    if (action === 'search_results') return jsonOut_(searchResults_(p.query));
    if (action !== 'status') { lock.waitLock(10000); locked = true; }
    if (action === 'save') {
      var map = { kw: 'KW', loc: 'LOC', jt: 'JT', days: 'DAYS', hour: 'HOUR', email: 'EMAIL', wa: 'WA', remote: 'REMOTE' }, updates = {};
      Object.keys(map).forEach(function (k) { if (p[k] !== undefined) updates[map[k]] = String(p[k]); });
      if (p.days !== undefined && (!/^\d+$/.test(p.days) || Number(p.days) < 1 || Number(p.days) > 3650)) throw new Error('Invalid posted-within days');
      if (p.hour !== undefined && (!/^\d+$/.test(p.hour) || Number(p.hour) > 23)) throw new Error('Invalid reminder hour');
      var fields = { emails: 'EMAILS_JSON', mobiles: 'MOBILES_JSON', times: 'TIMES_JSON', repeatDays: 'REPEAT_DAYS_JSON', recipients: 'RECIPIENTS_JSON', workModes: 'WORK_MODES_JSON', shifts: 'SHIFTS_JSON', groups: 'GROUPS_JSON' };
      Object.keys(fields).forEach(function (k) {
        if (p[k] === undefined) return;
        var value = JSON.parse(p[k]);
        if (!Array.isArray(value)) throw new Error(k + ' must be an array');
        if (k === 'times' && (!value.length || value.some(function (t) { return typeof t !== 'string' || !/^([01]\d|2[0-3]):[0-5]\d$/.test(t); }))) throw new Error('Invalid reminder times');
        if (k === 'recipients') {
          var dup = Object.create(null);
          value = value.filter(function (r) {
            if (!r || ['email', 'whatsapp', 'whatsapp_group'].indexOf(r.channel) < 0 || typeof r.address !== 'string' || !r.address.trim() || !Array.isArray(r.roles)) throw new Error('Invalid recipient');
            var key = r.channel + ':' + r.address.trim().toLowerCase();
            if (dup[key]) return false; dup[key] = true; return true;
          });
        }
        var json = JSON.stringify(value);
        if (Utilities.newBlob(json).getBytes().length > 8500) throw new Error(k + ' is too large for Script properties');
        updates[fields[k]] = json;
      });
      if (p.repeat !== undefined) { if (['daily', 'weekdays', 'custom'].indexOf(p.repeat) < 0) throw new Error('Invalid repeat'); updates.REPEAT = p.repeat; }
      if (p.timezone !== undefined) { Utilities.formatDate(new Date(), p.timezone, 'HH:mm'); updates.TIMEZONE = p.timezone; }
      if (updates.TIMES_JSON) updates.HOUR = String(Number(JSON.parse(updates.TIMES_JSON)[0].split(':')[0]));
      if (p.enabled !== undefined) { if (['true', 'false', '1', '0'].indexOf(p.enabled) < 0) throw new Error('Invalid enabled value'); updates.ENABLED = flag(p.enabled) ? '1' : '0'; }
      P.setProperties(updates);
      setOn(g('ENABLED', '0') === '1');
    } else if (action === 'on') setOn(true);
    else if (action === 'off') setOn(false);
    else if (action === 'test') sent = digest(true);
    result = { enabled: g('ENABLED', '0') === '1', last: g('LAST', ''), sent: sent, err: g('LASTERR', ''), kw: g('KW', ''), loc: g('LOC', ''), jt: g('JT', ''),
      days: g('DAYS', '1'), hour: g('HOUR', '9'), email: g('EMAIL', ''), wa: g('WA', ''), remote: flag(g('REMOTE', '')) ? '1' : '',
      scheduleVersion: 6,            // FIX: page checks >= 6; backend used to report 1, so "synced" never showed
      settingsVersion: 6, supportsWhatsAppGroups: false };
    if (P.getProperty('REPEAT') !== null) result.repeat = g('REPEAT');
    if (P.getProperty('TIMEZONE') !== null) result.timezone = g('TIMEZONE');
    var stored = { emails: 'EMAILS_JSON', mobiles: 'MOBILES_JSON', times: 'TIMES_JSON', repeatDays: 'REPEAT_DAYS_JSON', recipients: 'RECIPIENTS_JSON', workModes: 'WORK_MODES_JSON', shifts: 'SHIFTS_JSON', groups: 'GROUPS_JSON' };
    Object.keys(stored).forEach(function (k) { var raw = P.getProperty(stored[k]); if (raw !== null) result[k] = JSON.parse(raw); });
    if (result.err) result.warning = result.err;
  } catch (err) { result = { error: 'Script error: ' + err.message }; P.setProperty('LASTERR', result.error); }
  finally { if (locked) lock.releaseLock(); }
  return jsonOut_(result);
}

function diagnose() {
  var k = g('KW', '').split('|').filter(String)[0] || 'developer', l = g('LOC', '').split('|').filter(String).map(normLoc_)[0] || '', lines = [];
  function run(label, what, where, d) {
    var url = 'https://api.adzuna.com/v1/api/jobs/in/search/1?app_id=' + encodeURIComponent(g('ADZUNA_ID')) + '&app_key=' + encodeURIComponent(g('ADZUNA_KEY')) +
      '&results_per_page=1&max_days_old=' + d + '&what=' + encodeURIComponent(what) + (where ? '&where=' + encodeURIComponent(where) : '');
    try { var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); lines.push(label + ': ' + (r.getResponseCode() !== 200 ? 'HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 120) : JSON.parse(r.getContentText()).count + ' matches')); }
    catch (e) { lines.push(label + ': error ' + e.message); }
  }
  run('1. Keyword + location, saved days', k, l, g('DAYS', '2')); run('2. Keyword only, saved days', k, '', g('DAYS', '2'));
  run('3. Keyword only, 30 days', k, '', 30); run('4. Generic "developer", 30 days', 'developer', '', 30);
  lines.push('If 4 fails, your Adzuna keys are the problem. If 3 is 0, the keyword is too specific.');
  var report = lines.join('\n'); Logger.log(report);
  MailApp.sendEmail(g('EMAIL', '').split(/[\s,;]+/).filter(String)[0] || Session.getEffectiveUser().getEmail(), 'Job radar diagnosis', report);
}

/* ================= ADZUNA SHARED SEARCH ================= */
/* cfg: roles[], locations[], days, jobTypes[], workModes[], shifts[], remote, random */
function fetchAdzunaJobs_(cfg, cap) {
  checkAdzunaKeys();
  var roles = (cfg.roles || []).map(String).filter(String).slice(0, 10); if (!roles.length) roles = [''];   // empty role = any job
  var locs = (cfg.locations || []).map(String).filter(String).slice(0, 10); if (!locs.length) locs = [''];
  var days = Math.max(1, Math.min(365, Math.floor(Number(cfg.days)) || 7));
  var wm = cfg.workModes || [], remoteOnly = cfg.remote || (wm.length === 1 && (wm[0] === 'remote' || wm[0] === 'home'));
  var jt = cfg.jobTypes || [], flags = '';
  if (jt.length === 1) flags = jt[0] === 'full' ? '&full_time=1' : jt[0] === 'part' ? '&part_time=1' : jt[0] === 'contract' ? '&contract=1' : '';   // Adzuna ANDs flags, so only filter on a single type
  var pairs = [];
  roles.forEach(function (role) { locs.forEach(function (loc) { pairs.push({ role: role, loc: loc }); }); });
  pairs = pairs.slice(0, 12);
  var reqs = pairs.map(function (pr) {
    var what = String(pr.role).trim(); if (remoteOnly) what += ' remote'; if ((cfg.shifts || []).length === 1 && cfg.shifts[0] === 'night') what += ' night shift';
    var where = normLoc_(pr.loc); if (/^(remote|india)$/i.test(where)) where = '';
    return { url: 'https://api.adzuna.com/v1/api/jobs/in/search/1?app_id=' + encodeURIComponent(g('ADZUNA_ID')) + '&app_key=' + encodeURIComponent(g('ADZUNA_KEY')) +
      '&results_per_page=20&max_days_old=' + days + '&sort_by=date' + flags + (what.trim() ? '&what=' + encodeURIComponent(what.trim()) : '') + (where ? '&where=' + encodeURIComponent(where) : ''), muteHttpExceptions: true };
  });
  var jobs = [], errors = [], seen = {};
  UrlFetchApp.fetchAll(reqs).forEach(function (resp, i) {   // parallel: avoids the 6-minute execution limit
    if (resp.getResponseCode() !== 200) { errors.push('Adzuna HTTP ' + resp.getResponseCode() + ' for ' + (pairs[i].role || 'any role') + ' / ' + (pairs[i].loc || 'any location')); return; }
    try {
      (JSON.parse(resp.getContentText()).results || []).forEach(function (j) {
        if (!j.id || seen[j.id]) return; seen[j.id] = true;
        jobs.push({ id: String(j.id), t: String(j.title || '').replace(/<[^>]+>/g, ''), c: (j.company || {}).display_name || '', l: (j.location || {}).display_name || '', u: j.redirect_url || '' });
      });
    } catch (err) { errors.push('Bad response for ' + pairs[i].role + ': ' + err.message); }
  });
  return { jobs: jobs.slice(0, cap || 100), errors: errors };
}

/* ================= SEARCH RESULTS CHECK ================= */
function searchResults_(raw) {
  var q; try { q = JSON.parse(raw || '{}'); } catch (e) { return { supported: true, portals: [], error: 'Could not read the search request.' }; }
  var names = Array.isArray(q.portals) ? q.portals : [], pairs = (Array.isArray(q.pairs) ? q.pairs : []).slice(0, 12);
  return { supported: true, portals: names.map(function (n) {
    n = String(n); if (n === 'Adzuna') return adzunaCount_(pairs, q);
    return { name: n, status: 'unsupported', count: null, message: 'No public search service for this site. Open its search link to check results.' };
  }) };
}
function adzunaCount_(pairs, q) {
  var name = 'Adzuna';
  if (!g('ADZUNA_ID', '') || !g('ADZUNA_KEY', '')) return { name: name, status: 'error', count: null, message: 'ADZUNA_ID / ADZUNA_KEY missing in Script properties.' };
  if (!pairs.length) return { name: name, status: 'error', count: null, message: 'Add at least one job role first.' };
  var types = Array.isArray(q.types) ? q.types : [], flags = '';
  if (types.length === 1) flags = types[0] === 'full' ? '&full_time=1' : types[0] === 'part' ? '&part_time=1' : types[0] === 'contract' ? '&contract=1' : '';
  var wm = Array.isArray(q.workModes) ? q.workModes : [], remote = wm.length === 1 && (wm[0] === 'home' || wm[0] === 'remote');
  var days = Math.max(1, Math.min(365, Math.floor(Number(q.days)) || 1));
  var reqs = pairs.map(function (pr) {
    var kw = String((pr && pr.keyword) || '').slice(0, 200), loc = normLoc_(String((pr && pr.location) || '').slice(0, 100)), what = (kw + (remote ? ' remote' : '')).trim();
    return { url: 'https://api.adzuna.com/v1/api/jobs/in/search/1?app_id=' + encodeURIComponent(g('ADZUNA_ID')) + '&app_key=' + encodeURIComponent(g('ADZUNA_KEY')) +
      '&results_per_page=1&max_days_old=' + days + flags + (what ? '&what=' + encodeURIComponent(what) : '') + (loc && !/^(india|remote)$/i.test(loc) ? '&where=' + encodeURIComponent(loc) : ''), muteHttpExceptions: true };
  });
  var total = 0, ok = 0, bad = 0, firstErr = '';
  UrlFetchApp.fetchAll(reqs).forEach(function (r) {
    if (r.getResponseCode() === 200) { try { total += Number(JSON.parse(r.getContentText()).count) || 0; ok++; } catch (e) { bad++; } }
    else { bad++; if (!firstErr) firstErr = 'HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 100); }
  });
  if (!ok) return { name: name, status: 'error', count: null, message: 'Adzuna search failed. ' + firstErr };
  return { name: name, status: 'ok', count: total, message: 'Last ' + days + ' day(s), ' + ok + ' search(es); duplicates possible.' + (bad ? ' ' + bad + ' failed.' : '') };
}

/* ================= SHARED REMINDER SENDER (Family + Jobs) ================= */
function scheduleDueSlot_(key, s, now) {
  var tz = s.timezone || 'Asia/Kolkata';
  try { Utilities.formatDate(now, tz, 'HH:mm'); } catch (e) { tz = 'Asia/Kolkata'; }
  var date = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), time = Utilities.formatDate(now, tz, 'HH:mm');
  var isoDow = Number(Utilities.formatDate(now, tz, 'u')), jsDow = isoDow === 7 ? '0' : String(isoDow);
  var repeats = Array.isArray(s.repeats) && s.repeats.length ? s.repeats.map(String) : [String(s.repeat || 'daily')];
  var allowed = repeats.indexOf('daily') >= 0 || (repeats.indexOf('weekdays') >= 0 && isoDow <= 5) || (repeats.indexOf('weekends') >= 0 && isoDow > 5);
  if (repeats.indexOf('custom') >= 0 && (Array.isArray(s.repeatDays) ? s.repeatDays.map(String) : []).indexOf(jsDow) >= 0) allowed = true;
  if (!allowed) return null;
  var b = time.split(':'), nowMin = Number(b[0]) * 60 + Number(b[1]);
  var due = (Array.isArray(s.times) ? s.times : []).filter(function (t) {
    t = String(t); if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(t)) return false;
    var a = t.split(':'), m = Number(a[0]) * 60 + Number(a[1]);
    return nowMin >= m && nowMin < m + 10;     // FIX: 10-minute catch-up window; the 5-min trigger can drift or skip a run
  })[0];
  if (!due) return null;
  var marker = date + '|' + due;
  if (P.getProperty('SENT_' + key) === marker) return null;
  return { marker: marker, time: due };
}

function sendReminder_(key, s, force, dueSlot) {
  var address = String(s.address || '').trim();
  if (!address) throw new Error('Recipient is empty.');
  if (s.channel === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) throw new Error('Invalid email: ' + address);
  var res = fetchAdzunaJobs_(s, 100), seen = jsonProp_('SEEN_' + key, []);
  var jobs = force ? res.jobs : res.jobs.filter(function (j) { return seen.indexOf(j.id) < 0; });   // NEW: do not resend the same jobs every run
  if (s.random) jobs = jobs.sort(function () { return Math.random() - 0.5; });
  var errors = res.errors, roles = (s.roles || []).join(', ') || 'All jobs', locations = (s.locations || []).join(', ') || 'Any location';
  var hi = s.firstName ? 'Hi ' + familyHtml_(s.firstName) + ',' : '';
  if (s.channel === 'email') {
    var html = '<div style="font-family:Arial,sans-serif;max-width:760px"><h2>CareerPulse job reminder</h2>' + (hi ? '<p>' + hi + '</p>' : '') +
      '<p><b>Roles:</b> ' + familyHtml_(roles) + '<br><b>Locations:</b> ' + familyHtml_(locations) + '</p>';
    html += jobs.length ? jobs.slice(0, 25).map(function (j) { return '<p><a href="' + familyHtml_(j.u) + '"><b>' + familyHtml_(j.t) + '</b></a><br>' + familyHtml_(j.c) + ' &middot; ' + familyHtml_(j.l) + '</p>'; }).join('')
      : '<p>No new matching openings right now. Your reminder remains active.</p>';
    if (errors.length) html += '<p style="color:#a33">Some searches failed: ' + familyHtml_(errors.join(' | ')) + '</p>';
    html += '<p style="color:#60747d;font-size:12px">Posted within ' + familyHtml_(String(s.days || 7)) + ' day(s)</p></div>';
    sendMail_(address, 'CareerPulse: ' + jobs.length + ' new job' + (jobs.length === 1 ? '' : 's') + ' for ' + roles, html,
      'CareerPulse found ' + jobs.length + ' new job(s) for ' + roles + ' in ' + locations + '.\n' + jobs.slice(0, 10).map(function (j) { return '- ' + j.t + ' - ' + j.c + ' ' + j.u; }).join('\n'));
  } else if (s.channel === 'whatsapp') {
    var waKey = g('WA_KEY', ''); if (!waKey) throw new Error('WA_KEY is missing; WhatsApp delivery is not configured.');
    var msg = (s.firstName ? 'Hi ' + s.firstName + ', ' : '') + 'CareerPulse: ' + jobs.length + ' new job(s) for ' + roles + '.\n' + jobs.slice(0, 5).map(function (j) { return '- ' + j.t + ' - ' + j.c + '\n' + j.u; }).join('\n') + waFooter_();
    var wa = UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(address.replace(/[^\d+]/g, '')) + '&text=' + encodeURIComponent(msg) + '&apikey=' + encodeURIComponent(waKey), { muteHttpExceptions: true });
    if (wa.getResponseCode() < 200 || wa.getResponseCode() >= 300) throw new Error('WhatsApp HTTP ' + wa.getResponseCode() + ': ' + wa.getContentText().slice(0, 150));
  } else throw new Error('Unsupported delivery channel: ' + s.channel);
  if (!force) {
    if (dueSlot) P.setProperty('SENT_' + key, dueSlot.marker);
    P.setProperty('SEEN_' + key, JSON.stringify(seen.concat(jobs.map(function (j) { return j.id; })).slice(-100)));
  }
  return { ok: true, jobs: jobs.length, errors: errors, recipient: address };
}
function waFooter_() { var i = senderInfo_(); return '\n\n— ' + i.name + (i.phone ? ' · ' + i.phone : '') + (i.email ? ' · ' + i.email : ''); }
function familyHtml_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }

/* ================= FAMILY (Supabase) ================= */
function familyRequireSupabase_() { if (!g('SUPABASE_URL', '') || !g('SUPABASE_SERVICE_KEY', '')) throw new Error('Add SUPABASE_URL and SUPABASE_SERVICE_KEY in Script properties.'); }
function familySupabaseReminders_() {
  familyRequireSupabase_();
  var key = g('SUPABASE_SERVICE_KEY', '');
  var resp = UrlFetchApp.fetch(String(g('SUPABASE_URL', '')).replace(/\/+$/, '') + '/rest/v1/family_reminders?select=id,created_at,enabled,settings&enabled=eq.true&order=created_at.asc',
    { method: 'get', headers: { apikey: key, Authorization: 'Bearer ' + key }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Supabase HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 300));
  var rows = JSON.parse(resp.getContentText() || '[]'); return Array.isArray(rows) ? rows : [];
}
function familyPublicStatus_() {
  var sb = !!g('SUPABASE_URL', '') && !!g('SUPABASE_SERVICE_KEY', ''), az = !!g('ADZUNA_ID', '') && !!g('ADZUNA_KEY', ''), trig = trigExists_('familySupabaseTick'), n = null, ok = false, error = '';
  if (sb) { try { n = familySupabaseReminders_().length; ok = true; } catch (e) { error = e.message; } }
  return { ok: ok && az && trig, supabaseConnected: ok, adzunaConfigured: az, triggerActive: trig, enabledReminders: n, error: error };
}
function setupFamilySupabaseSender() {
  familyRequireSupabase_(); checkAdzunaKeys(); dropTriggers_(['familySupabaseTick']);
  ScriptApp.newTrigger('familySupabaseTick').timeBased().everyMinutes(5).create();
  Logger.log('Family sender enabled (checks every 5 minutes).');
}
function disableFamilySupabaseSender() { dropTriggers_(['familySupabaseTick']); Logger.log('Family sender disabled.'); }
function testFamilySupabaseConnection() { var rows = familySupabaseReminders_(); Logger.log('Enabled Family reminders: ' + rows.length); return rows.length; }
function testFamilyReminderNow() {
  var rows = familySupabaseReminders_(); if (!rows.length) throw new Error('No enabled Family reminders found.');
  var r = sendReminder_('FAM_' + rows[0].id, familySettings_(rows[0]), true); Logger.log(JSON.stringify(r)); return r;
}
function familySettings_(row) { var s = Object.assign({}, row.settings || {}); if (!Array.isArray(s.jobTypes)) s.jobTypes = []; s.firstName = String(s.firstName || '').slice(0, 60); return s; }
function familySupabaseTick() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return;
  try {
    var now = new Date();
    familySupabaseReminders_().forEach(function (row) {
      try { var key = 'FAM_' + row.id, s = familySettings_(row), due = scheduleDueSlot_(key, s, now); if (due) sendReminder_(key, s, false, due); }
      catch (err) { Logger.log('Family reminder ' + row.id + ' failed: ' + err.message); }
    });
  } finally { lock.releaseLock(); }
}

/* ================= ADMIN AUTH + PORTALS ================= */
function portalRequireAdmin_(tok) {
  familyRequireSupabase_(); if (!tok) throw new Error('Admin session is required.');
  var r = UrlFetchApp.fetch(String(g('SUPABASE_URL', '')).replace(/\/+$/, '') + '/auth/v1/user', { headers: { apikey: g('SUPABASE_SERVICE_KEY', ''), Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Admin session is invalid or expired.');
  var u = JSON.parse(r.getContentText() || '{}'); if (!u.id) throw new Error('Admin user could not be verified.');
  // SECURITY FIX: any signed-in Supabase user used to be treated as admin. Enforce an allow-list.
  var allow = g('ADMIN_EMAILS', '').toLowerCase().split(/[\s,;]+/).filter(String);
  if (allow.length && allow.indexOf(String(u.email || '').toLowerCase()) < 0) throw new Error('This account is not an authorised admin.');
  return u;
}
function portalLoad_() { return jsonProp_('PORTALS_JSON', []); }
function portalSaveAll_(a) {
  var j = JSON.stringify(a || []); if (Utilities.newBlob(j).getBytes().length > 8500) throw new Error('Too many sources for Script properties (9 KB limit).');
  P.setProperty('PORTALS_JSON', j);
}
function portalMask_(s) { s = String(s || ''); return s ? '••••••••' + s.slice(-4) : ''; }
function portalPublic_(p) { return { id: p.id, name: p.name, enabled: p.enabled !== false, type: p.type || 'adzuna', country: p.country || 'in', baseUrl: p.baseUrl || '', appId: p.appId || '', apiKeyMasked: portalMask_(p.apiKey),
  queryParam: p.queryParam || 'what', locationParam: p.locationParam || 'where', resultsPath: p.resultsPath || 'results', titlePath: p.titlePath || 'title', locationPath: p.locationPath || 'location.display_name', companyPath: p.companyPath || 'company.display_name', urlPath: p.urlPath || 'redirect_url' }; }
function portalAdminList_(tok) { portalRequireAdmin_(tok); return { ok: true, portals: portalLoad_().map(portalPublic_) }; }
function portalParse_(raw) {
  var p = typeof raw === 'object' ? raw : JSON.parse(String(raw || '{}'));
  var d = { queryParam: 'what', locationParam: 'where', resultsPath: 'results', titlePath: 'title', locationPath: 'location.display_name', companyPath: 'company.display_name', urlPath: 'redirect_url' };
  p.id = String(p.id || Utilities.getUuid()); p.name = String(p.name || '').trim(); p.type = String(p.type || 'adzuna'); p.country = String(p.country || 'in').trim().toLowerCase();
  p.baseUrl = String(p.baseUrl || '').trim(); p.appId = String(p.appId || '').trim(); p.apiKey = String(p.apiKey || '').trim();
  Object.keys(d).forEach(function (k) { p[k] = String(p[k] || d[k]).trim(); });
  p.enabled = p.enabled !== false;
  if (!p.name) throw new Error('Portal name is required.');
  if (p.type === 'generic' && !/^https:\/\//i.test(p.baseUrl)) throw new Error('Generic Base URL must use HTTPS.');
  return p;
}
function portalAdminSave_(tok, raw) {
  portalRequireAdmin_(tok); var n = portalParse_(raw), a = portalLoad_(), hit = false;
  a = a.map(function (o) { if (o.id !== n.id) return o; hit = true; if (!n.apiKey) n.apiKey = o.apiKey || ''; return n; });
  if (!hit) a.push(n); portalSaveAll_(a); return { ok: true, portal: portalPublic_(n) };
}
function portalAdminDelete_(tok, id) { portalRequireAdmin_(tok); portalSaveAll_(portalLoad_().filter(function (p) { return p.id !== String(id || ''); })); return { ok: true }; }
function portalAdminToggle_(tok, id, en) {
  portalRequireAdmin_(tok); var a = portalLoad_(), hit = false;
  a.forEach(function (p) { if (p.id === String(id || '')) { p.enabled = String(en) === 'true'; hit = true; } });
  if (!hit) throw new Error('Portal not found.'); portalSaveAll_(a); return { ok: true };
}
function portalPath_(o, path) { return String(path || '').split('.').reduce(function (v, k) { return v == null ? undefined : v[k]; }, o); }
function portalSearchOne_(p, what, where, limit) {
  limit = limit || 10;
  if ((p.type || 'adzuna') === 'adzuna') {
    var id = p.appId || g('ADZUNA_ID', ''), key = p.apiKey || g('ADZUNA_KEY', ''); if (!id || !key) throw new Error((p.name || 'Adzuna') + ': credentials missing.');
    var r = UrlFetchApp.fetch('https://api.adzuna.com/v1/api/jobs/' + encodeURIComponent(p.country || 'in') + '/search/1?app_id=' + encodeURIComponent(id) + '&app_key=' + encodeURIComponent(key) + '&results_per_page=' + limit + '&sort_by=date' +
      (what ? '&what=' + encodeURIComponent(what) : '') + (where ? '&where=' + encodeURIComponent(where) : ''), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error((p.name || 'Adzuna') + ' HTTP ' + r.getResponseCode());
    return (JSON.parse(r.getContentText() || '{}').results || []).map(function (j) { return { title: String(j.title || '').replace(/<[^>]+>/g, ''), location: (j.location || {}).display_name || '', company: (j.company || {}).display_name || '', url: j.redirect_url || '' }; });
  }
  var u = p.baseUrl, sep = u.indexOf('?') >= 0 ? '&' : '?';
  if (what) { u += sep + encodeURIComponent(p.queryParam || 'what') + '=' + encodeURIComponent(what); sep = '&'; }
  if (where) u += sep + encodeURIComponent(p.locationParam || 'where') + '=' + encodeURIComponent(where);
  var h = {}; if (p.apiKey) h.Authorization = 'Bearer ' + p.apiKey;
  var r2 = UrlFetchApp.fetch(u, { headers: h, muteHttpExceptions: true });
  if (r2.getResponseCode() < 200 || r2.getResponseCode() >= 300) throw new Error(p.name + ' HTTP ' + r2.getResponseCode());
  var arr = portalPath_(JSON.parse(r2.getContentText() || '{}'), p.resultsPath || 'results'); if (!Array.isArray(arr)) throw new Error(p.name + ': results path is not an array.');
  return arr.slice(0, limit).map(function (j) { return { title: String(portalPath_(j, p.titlePath) || ''), location: String(portalPath_(j, p.locationPath) || ''), company: String(portalPath_(j, p.companyPath) || ''), url: String(portalPath_(j, p.urlPath) || '') }; });
}
function portalAdminTest_(tok, raw) {   // FIX: signature is (tok, raw); callers used to pass (tok, id, raw)
  portalRequireAdmin_(tok); var p = portalParse_(raw), old = portalLoad_().filter(function (x) { return x.id === p.id; })[0];
  if (!p.apiKey && old) p.apiKey = old.apiKey || '';
  var a = portalSearchOne_(p, 'data analyst', 'Delhi', 3); return { ok: true, count: a.length, sample: a.slice(0, 3) };
}
function portalEnabled_() {
  var a = portalLoad_().filter(function (p) { return p.enabled !== false; });
  return a.length ? a : [{ id: 'legacy', name: 'Adzuna', type: 'adzuna', country: 'in', appId: g('ADZUNA_ID', ''), apiKey: g('ADZUNA_KEY', '') }];
}
function familySuggestions_(kind, q) {
  kind = String(kind || ''); q = String(q || '').trim();
  if (['role', 'location'].indexOf(kind) < 0 || q.length < 2) return { ok: true, suggestions: [] };
  var c = CacheService.getScriptCache(), ck = 'fs_' + kind + '_' + shortKey_(q.toLowerCase()), z = c.get(ck);
  if (z) return { ok: true, suggestions: JSON.parse(z) };
  var out = [];
  portalEnabled_().slice(0, 3).forEach(function (p) {
    try { portalSearchOne_(p, kind === 'role' ? q : '', kind === 'location' ? q : '', 20).forEach(function (j) {
      var v = kind === 'role' ? j.title : j.location;
      if (v && v.toLowerCase().indexOf(q.toLowerCase()) >= 0 && !out.some(function (x) { return x.toLowerCase() === v.toLowerCase(); })) out.push(v);   // FIX: only suggestions that match what was typed
    }); } catch (e) { Logger.log('Suggestion ' + p.name + ': ' + e.message); }
  });
  out = out.slice(0, 8); c.put(ck, JSON.stringify(out), 600); return { ok: true, suggestions: out };
}

/* ================= PAGE CMS ================= */
function pageName_(page) { page = String(page || '').toLowerCase(); if (['jobs', 'family'].indexOf(page) < 0) throw new Error('Unknown page.'); return page.toUpperCase(); }
function pageConfigLoad_(page) { var a = jsonProp_('PAGE_CONFIG_' + pageName_(page), []); return Array.isArray(a) ? a : []; }
function cleanSections_(raw) {
  var a = typeof raw === 'object' ? raw : JSON.parse(String(raw || '[]')); if (!Array.isArray(a)) throw new Error('Invalid section configuration.');
  a = a.slice(0, 20).map(function (x) { return { id: String(x.id || Utilities.getUuid()), title: String(x.title || '').slice(0, 120), body: String(x.body || '').slice(0, 2000), position: String(x.position) === 'bottom' ? 'bottom' : 'top', enabled: x.enabled !== false }; });
  if (Utilities.newBlob(JSON.stringify(a)).getBytes().length > 8500) throw new Error('Custom content is too large (9 KB Script-property limit). Shorten the text or remove sections.');   // FIX: explicit limit error
  return a;
}
function pageConfigPublic_(page) { return { ok: true, page: String(page || ''), sections: pageConfigLoad_(page).filter(function (x) { return x.enabled !== false; }) }; }
function adminPageConfig_(tok, page) { portalRequireAdmin_(tok); return { ok: true, page: String(page || ''), sections: pageConfigLoad_(page) }; }
function pageBuiltinDefaults_(page) {
  if (String(page) === 'jobs') return [
    { id: 'switches', target: 'secSwitches', title: 'Switches', order: 10 }, { id: 'filters', target: 'secLook', title: 'What to look for', order: 20 },
    { id: 'reminders', target: 'secReminder', title: 'Reminder settings', order: 30 }, { id: 'recipients', target: 'secRecipients', title: 'Jobs for each recipient', order: 40 },
    { id: 'search', target: 'secSearch', title: 'Search now', order: 50 }, { id: 'results', target: 'secResults', title: 'Check search results', order: 60 },
    { id: 'consultancies', target: 'secConsult', title: 'Find job consultancies', order: 70 }, { id: 'connection', target: 'secConn', title: 'Reminder connection', order: 80 }];
  if (String(page) === 'family') return [
    { id: 'delivery', target: 'deliverySection', title: 'Delivery details', order: 10 }, { id: 'roleLocation', target: 'roleLocationSection', title: 'Job roles & locations', order: 20 },
    { id: 'schedule', target: 'scheduleSection', title: 'Reminder schedule', order: 30 }, { id: 'jobAge', target: 'jobAgeSection', title: 'Jobs from & timezone', order: 40 },
    { id: 'work', target: 'workSection', title: 'Work preference & shift', order: 50 }, { id: 'actions', target: 'actionsSection', title: 'Family service actions', order: 60 }];
  return [];
}
function pageBuiltinSections_(page) {
  var saved = {}; jsonProp_('PAGE_BUILTINS_' + pageName_(page), []).forEach(function (x) { saved[x.id] = x; });
  return pageBuiltinDefaults_(page).map(function (d) { var s = saved[d.id] || {}; return { id: d.id, target: d.target, title: String(s.title || d.title), enabled: s.enabled !== false, order: Number(s.order || d.order) }; }).sort(function (a, b) { return a.order - b.order; });
}
function adminBuiltinSave_(tok, page, raw) {
  portalRequireAdmin_(tok); var allowed = {}; pageBuiltinDefaults_(page).forEach(function (d) { allowed[d.id] = d; });
  var a = typeof raw === 'object' ? raw : JSON.parse(String(raw || '[]')); if (!Array.isArray(a)) throw new Error('Invalid configuration.');
  P.setProperty('PAGE_BUILTINS_' + pageName_(page), JSON.stringify(a.filter(function (x) { return allowed[x.id]; }).map(function (x, i) { return { id: String(x.id), title: String(x.title || allowed[x.id].title).slice(0, 120), enabled: x.enabled !== false, order: (i + 1) * 10 }; })));
  return { ok: true, builtins: pageBuiltinSections_(page) };
}
function adminBuiltinReset_(tok, page) { portalRequireAdmin_(tok); P.deleteProperty('PAGE_BUILTINS_' + pageName_(page)); return { ok: true, builtins: pageBuiltinSections_(page) }; }
function pageManifestPublic_(page) { return { ok: true, page: String(page || ''), builtins: pageBuiltinSections_(page), custom: pageConfigLoad_(page).filter(function (x) { return x.enabled !== false; }) }; }
function pageDraftLoad_(page) { var raw = g('PAGE_DRAFT_' + pageName_(page), ''); if (!raw) return null; try { var a = JSON.parse(raw); return Array.isArray(a) ? a : null; } catch (e) { return null; } }
function adminPageDraftSave_(tok, page, raw) { portalRequireAdmin_(tok); var a = cleanSections_(raw); P.setProperty('PAGE_DRAFT_' + pageName_(page), JSON.stringify(a)); return { ok: true, draft: a }; }
function adminPagePublish_(tok, page) {
  portalRequireAdmin_(tok); var a = pageDraftLoad_(page); if (a === null) throw new Error('No saved draft to publish.');
  P.setProperty('PAGE_CONFIG_' + pageName_(page), JSON.stringify(a)); P.deleteProperty('PAGE_DRAFT_' + pageName_(page)); return { ok: true, sections: a };
}
function adminPageDraftDiscard_(tok, page) { portalRequireAdmin_(tok); P.deleteProperty('PAGE_DRAFT_' + pageName_(page)); return { ok: true, sections: pageConfigLoad_(page) }; }
function adminPageAllSections_(tok, page) { portalRequireAdmin_(tok); var d = pageDraftLoad_(page); return { ok: true, page: String(page || ''), builtins: pageBuiltinSections_(page), custom: pageConfigLoad_(page), draft: d, hasDraft: d !== null }; }

function adminDashboard_(tok) {
  portalRequireAdmin_(tok);
  var fam = 0, famErr = ''; try { fam = familySupabaseReminders_().length; } catch (e) { famErr = e.message; }   // FIX: one failing source no longer kills the dashboard
  return { ok: true, status: {
    jobsSender: (trigExists_('digest') || trigExists_('jobsTick')) ? 'running' : 'stopped', familySender: trigExists_('familySupabaseTick') ? 'running' : 'stopped',
    jobSources: portalEnabled_().length, familyReminders: fam, familyError: famErr, lastJobsRun: g('LAST', '') || 'Not yet', lastJobsError: g('LASTERR', ''), familyCheck: 'Every 5 minutes' } };
}
function adminJobData_(tok) {
  portalRequireAdmin_(tok);
  return { ok: true, data: { enabled: g('ENABLED', '0') === '1', keywords: g('KW', '').split('|').filter(String), locations: g('LOC', '').split('|').filter(String), jobTypes: g('JT', '').split('|').filter(String),
    days: g('DAYS', '2'), remote: g('REMOTE', ''), email: g('EMAIL', ''), hour: g('HOUR', '9'), times: jsonProp_('TIMES_JSON', []), recipients: jobsRecipients_().length, lastRun: g('LAST', ''), lastError: g('LASTERR', '') } };
}

/* ================= JOB PROFILES (Supabase table public.job_profiles) ================= */
function jobDbAuth_(token) { if (!g('TOKEN', '') || String(token || '') !== g('TOKEN', '')) throw new Error('Wrong token'); }
function jobDbReq_(method, path, body, prefer) {
  familyRequireSupabase_(); var key = g('SUPABASE_SERVICE_KEY', ''), opt = { method: method, muteHttpExceptions: true, headers: { apikey: key, Authorization: 'Bearer ' + key } };
  if (body !== undefined) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); } if (prefer) opt.headers.Prefer = prefer;
  var r = UrlFetchApp.fetch(String(g('SUPABASE_URL', '')).replace(/\/$/, '') + '/rest/v1/' + path, opt), st = r.getResponseCode(), txt = r.getContentText();
  if (st < 200 || st >= 300) throw new Error('Database ' + st + ': ' + txt.slice(0, 500)); return txt ? JSON.parse(txt) : [];
}
function cleanJobProfile_(raw) {
  var x = typeof raw === 'string' ? JSON.parse(raw || '{}') : (raw || {}), id = String(x.id || Utilities.getUuid()).replace(/[^A-Za-z0-9_.:@-]/g, '').slice(0, 180) || Utilities.getUuid();
  var kind = ['filter', 'quickfill'].indexOf(String(x.kind)) >= 0 ? String(x.kind) : 'filter', name = String(x.name || (kind === 'quickfill' ? 'Saved reminder' : 'Job filters')).trim().slice(0, 160);
  var packed = JSON.stringify(x.data && typeof x.data === 'object' ? x.data : {}); if (packed.length > 90000) throw new Error('Saved Jobs profile is too large.');
  return { id: id, kind: kind, name: name, data: JSON.parse(packed), updated_at: new Date().toISOString() };
}
function jobProfilesList_(token) { jobDbAuth_(token); return { ok: true, profiles: jobDbReq_('get', 'job_profiles?select=id,kind,name,data,updated_at&order=updated_at.desc') }; }
function jobProfileSave_(token, raw) { jobDbAuth_(token); var row = cleanJobProfile_(raw), rows = jobDbReq_('post', 'job_profiles?on_conflict=id', [row], 'resolution=merge-duplicates,return=representation'); return { ok: true, profile: rows[0] || row }; }
function jobProfileDelete_(token, id) { jobDbAuth_(token); id = String(id || ''); if (!id) throw new Error('Missing profile id.'); jobDbReq_('delete', 'job_profiles?id=eq.' + encodeURIComponent(id), undefined, 'return=minimal'); return { ok: true, id: id }; }
function adminJobProfilesDb_(auth) {
  portalRequireAdmin_(auth); var rows = jobDbReq_('get', 'job_profiles?select=id,kind,name,data,updated_at&order=updated_at.desc');
  return { ok: true, data: { profiles: rows, filters: rows.filter(function (x) { return x.kind === 'filter'; }), quickFill: rows.filter(function (x) { return x.kind === 'quickfill'; }) } };
}
function adminJobProfileDelete_(auth, id) { portalRequireAdmin_(auth); id = String(id || ''); if (!id) throw new Error('Missing profile id.'); jobDbReq_('delete', 'job_profiles?id=eq.' + encodeURIComponent(id), undefined, 'return=minimal'); return { ok: true, id: id }; }

/* Admin edit/create of Jobs filters and Quick Fill records. apply='1' also pushes a 'filter' record to the live sender. */
function adminJobProfileSave_(auth, raw, apply) {
  portalRequireAdmin_(auth);
  var row = cleanJobProfile_(raw), d = row.data || {};
  if (row.kind === 'quickfill') {
    var r = (d.recipients || [])[0];
    if (!r || !r.address) throw new Error('A Quick Fill record needs one recipient.');
  }
  var applied = false;
  if (row.kind === 'filter' && String(apply) === '1') { applyFilterToSender_(d); applied = true; }   // validate first: nothing is stored if invalid
  var rows = jobDbReq_('post', 'job_profiles?on_conflict=id', [row], 'resolution=merge-duplicates,return=representation');
  return { ok: true, profile: rows[0] || row, applied: applied };
}
function applyFilterToSender_(d) {
  var times = Array.isArray(d.times) ? d.times : [];
  if (!times.length || times.some(function (t) { return !/^([01]\d|2[0-3]):[0-5]\d$/.test(String(t)); })) throw new Error('Invalid reminder times.');
  var days = String(d.days || '1'); if (!/^\d+$/.test(days) || Number(days) < 1 || Number(days) > 3650) throw new Error('Invalid posted-within days.');
  var repeat = String(d.repeat || 'daily'); if (['daily', 'weekdays', 'custom'].indexOf(repeat) < 0) throw new Error('Invalid repeat.');
  var tz = String(d.timezone || 'Asia/Kolkata'); Utilities.formatDate(new Date(), tz, 'HH:mm');
  var rec = (Array.isArray(d.recipients) ? d.recipients : []).filter(function (r) {
    if (!r || ['email', 'whatsapp', 'whatsapp_group'].indexOf(r.channel) < 0 || typeof r.address !== 'string' || !r.address.trim() || !Array.isArray(r.roles)) throw new Error('Invalid recipient.');
    return true;
  });
  var u = { KW: (d.kw || []).join('|'), LOC: (d.loc || []).join('|'), JT: (d.types || []).join('|'), DAYS: days, HOUR: String(Number(times[0].split(':')[0])), REPEAT: repeat, TIMEZONE: tz,
    ENABLED: d.enabled ? '1' : '0', TIMES_JSON: JSON.stringify(times), REPEAT_DAYS_JSON: JSON.stringify(d.repeatDays || []), RECIPIENTS_JSON: JSON.stringify(rec),
    WORK_MODES_JSON: JSON.stringify(d.workModes || []), SHIFTS_JSON: JSON.stringify(d.shifts || []) };
  Object.keys(u).forEach(function (k) { if (/_JSON$/.test(k) && Utilities.newBlob(u[k]).getBytes().length > 8500) throw new Error(k + ' is too large for Script properties.'); });
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { P.setProperties(u); setOn(u.ENABLED === '1'); } finally { lock.releaseLock(); }
}

var SENDER_KEYS_ = { name: 'SENDER_NAME', replyTo: 'REPLY_TO', phone: 'CONTACT_PHONE', email: 'CONTACT_EMAIL', web: 'CONTACT_WEBSITE' };
function adminSenderGet_(auth) {
  portalRequireAdmin_(auth); var i = senderInfo_(), raw = {};
  Object.keys(SENDER_KEYS_).forEach(function (k) { raw[k] = g(SENDER_KEYS_[k], ''); });
  return { ok: true, data: raw, effective: { fromAccount: (function () { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } })(), name: i.name, replyTo: i.replyTo, phone: i.phone, email: i.email, web: i.web } };
}
function adminSenderSave_(auth, raw) {
  portalRequireAdmin_(auth); var c = typeof raw === 'object' ? raw : JSON.parse(String(raw || '{}')), set = {}, del = [];
  var mail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  Object.keys(SENDER_KEYS_).forEach(function (k) {
    var v = String(c[k] == null ? '' : c[k]).trim().slice(0, 120);
    if ((k === 'replyTo' || k === 'email') && v && !mail.test(v)) throw new Error('Enter a valid email address for ' + (k === 'replyTo' ? 'Reply-To' : 'Contact email') + '.');
    if (k === 'phone' && v && !/^[+0-9()\-\s]{6,25}$/.test(v)) throw new Error('Enter a valid phone number, e.g. +91 98765 43210.');
    if (k === 'web' && v && !/^https?:\/\//i.test(v)) throw new Error('Website must start with http:// or https://');
    if (v) set[SENDER_KEYS_[k]] = v; else del.push(SENDER_KEYS_[k]);
  });
  P.setProperties(set); del.forEach(function (k) { P.deleteProperty(k); });
  return adminSenderGet_(auth);
}
function adminSenderTest_(auth) {
  var u = portalRequireAdmin_(auth); if (!u.email) throw new Error('Your admin account has no email address.');
  sendMail_(u.email, 'CareerPulse sender test', '<h3>Sender test</h3><p>This is how reminder emails will look. Try replying to this message — the reply should reach the Reply-To address you set.</p>', 'Sender test. Try replying to this message — it should reach the Reply-To address you set.');
  return { ok: true, sentTo: u.email };
}

/* Housekeeping: run occasionally (or add a weekly trigger) to drop state of deleted Family reminders. */
function cleanupFamilyState() {
  var live = {}; familySupabaseReminders_().forEach(function (r) { live[shortKeyless_(r.id)] = 1; });
  Object.keys(P.getProperties()).forEach(function (k) { var m = k.match(/^(SENT|SEEN)_FAM_(.+)$/); if (m && !live[m[2]]) P.deleteProperty(k); });
}
function shortKeyless_(id) { return String(id); }
