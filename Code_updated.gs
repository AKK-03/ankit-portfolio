var CAREERPULSE_BACKEND_VERSION = '2026-10-10-idle-formats-v14';
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
  if (i.phone && showFlag_('mailPhone')) { t.push('Phone: ' + i.phone); h.push('📞 ' + familyHtml_(i.phone)); }
  if (i.email && showFlag_('mailEmail')) { t.push('Email: ' + i.email); h.push('✉️ <a href="mailto:' + familyHtml_(i.email) + '">' + familyHtml_(i.email) + '</a>'); }
  if (i.web && showFlag_('mailWeb')) { t.push('Web: ' + i.web); h.push('🌐 ' + familyHtml_(i.web)); }
  return { info: i, name: showFlag_('mailName') ? i.name : 'CareerPulse', text: t, html: h };
}
function footerText_() {
  var f = footerParts_();
  return '--\nSent by ' + f.name + (f.text.length ? '\n' + f.text.join('\n') : '') + (showFlag_('mailReplyLine') ? '\nJust reply to this email to reach us, or to change or stop these alerts.' : '');
}
function footerHtml_() {
  var f = footerParts_();
  return '<hr style="border:0;border-top:1px solid #dde6ea;margin:18px 0"><p style="font-size:13px;color:#4b616c;line-height:1.6;margin:0">Sent by <b>' + familyHtml_(f.name) + '</b>' +
    (f.html.length ? '<br>' + f.html.join(' &nbsp;·&nbsp; ') : '') + (showFlag_('mailReplyLine') ? '<br>Just reply to this email to reach us, or to change or stop these alerts.' : '') + '</p>';
}
/* One place that sends every email: display name + Reply-To + contact footer. */
function sendMail_(to, subject, html, plain, kindOverride, reminderId) {
  if (/^Job radar:/.test(String(subject)) && receiverBlocked_('jobs',to)) return;
  var i = senderInfo_(), m = { to: to, subject: subject, body: String(plain || subject) + '\n\n' + footerText_(), name: showFlag_('mailName') ? i.name : 'CareerPulse' };
  if (i.replyTo && showFlag_('mailReplyTo')) m.replyTo = i.replyTo;   // switched off in Admin > Sender > Privacy: no "Reply to" line is added
  var kind=kindOverride||(/^Job radar:/.test(String(subject))?'jobs':null);
  if(kind){var manage=receiverMailControls_(kind,to,reminderId);m.body+='\n\nManage your alerts: '+manage.text;if(html)html+=manage.html;}
  if(kind){var tpl=mailTemplatePublic_(kind);if(tpl){var subj0=subject;subject=tpl.subject.replace(/\{\{subject\}\}/g,function(){return subj0;});m.subject=subject;var content=String(html||'').replace(manage.html,'');html=tpl.html.replace(/\{\{subject\}\}/g,function(){return subject;}).replace(/\{\{content\}\}/g,function(){return content;}).replace(/\{\{controls\}\}/g,function(){return manage.html;});}}
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
  var p = (e && e.parameter) || {}, a = String(p.action || ''); rememberUrl_();
  try {
    switch (a) {
      case 'deployment_check': var jh = jobsHealth_(); return jsonOut_({ ok: true, version: CAREERPULSE_BACKEND_VERSION, adminAuth: 'supabase', family: true, jobsDb: true, jobsOk: jh.ok, jobsReason: jh.reason });
      case 'family_ping': var fh = familyHealth_(); return jsonOut_({ ok: true, service: 'family', version: CAREERPULSE_BACKEND_VERSION, senderRunning: fh.ok, senderReason: fh.reason });
      case 'backend_version': return jsonOut_({ ok: true, version: CAREERPULSE_BACKEND_VERSION, execUrl: liveUrl_(), capabilities: ['admin_receiver_set','admin_receiver_reminders','admin_receiver_list','admin_mail_template_save','admin_mail_templates','admin_monitor_record','admin_monitor_logs','admin_family_list','admin_family_set_enabled','admin_family_delete','admin_reminder_delete','admin_receiver_clear','monitor_recovery','admin_sessions','admin_session_ping','admin_session_revoke','admin_logout_all','page_status','admin_pages','admin_page_flags_save','admin_page_add','admin_page_remove','admin_privacy_save','admin_wa_keys','admin_wa_key_save','admin_wa_key_delete','admin_wa_test','admin_delivery_log','admin_delivery_clear','admin_idle_save','admin_mail_template_test'] });
      case 'family_status': return jsonOut_(familyPublicStatus_());
      case 'family_suggest': if (!pageFlag_('family').enabled) return jsonOut_({ ok: false, error: 'page_disabled' }); return jsonOut_(familySuggestions_(p.kind, p.q));
      case 'page_status': return jsonOut_(pageStatusPublic_(p.page));
      case 'page_config': return jsonOut_(pageConfigPublic_(p.page));
      case 'page_manifest': return jsonOut_(pageManifestPublic_(p.page));
      case 'search_results': return jsonOut_(searchResults_(p.query));
      case 'receiver_manage': return receiverManagePage_(p);
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
      case 'admin_monitor_logs': return jsonOut_(monitorLogs_(p.auth));
      case 'admin_family_list': return jsonOut_(adminFamilyList_(p.auth));
      case 'admin_sessions': return jsonOut_(adminSessions_(p.auth));
      case 'admin_pages': return jsonOut_(adminPages_(p.auth));
      case 'admin_wa_keys': return jsonOut_(adminWaKeys_(p.auth));
      case 'admin_delivery_log': return jsonOut_(adminDeliveryLog_(p.auth));
      case 'admin_receiver_list': return jsonOut_(receiverList_(p.auth));
      case 'admin_receiver_reminders': return jsonOut_(receiverReminders_(p.auth,p.type));
      case 'admin_mail_templates': return jsonOut_(mailTemplates_(p.auth));
      case 'admin_sender_get': return jsonOut_(adminSenderGet_(p.auth));
      case 'admin_page_config': return jsonOut_(adminPageConfig_(p.auth, p.page));
      case 'admin_page_all': return jsonOut_(adminPageAllSections_(p.auth, p.page));
    }
    if (a || p.token !== undefined) return radarApi_(e);   // legacy Jobs router (status/save/on/off/test)
    return ContentService.createTextOutput('CareerPulse API is running.');
  } catch (err) { return jsonOut_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || ''); rememberUrl_();
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
      case 'admin_monitor_record': return jsonOut_(monitorRecord_(p.auth,p.record));
      case 'admin_family_set_enabled': return jsonOut_(adminFamilySetEnabled_(p.auth,p.id,p.enabled));
      case 'admin_family_delete': return jsonOut_(adminFamilyDelete_(p.auth,p.ids));
      case 'admin_session_ping': return jsonOut_(adminSessionPing_(p.auth,p.label,p.activeAt));
      case 'admin_idle_save': return jsonOut_(adminIdleSave_(p.auth,p.minutes));
      case 'admin_mail_template_test': return jsonOut_(mailTemplateTest_(p.auth,p.type));
      case 'admin_session_revoke': return jsonOut_(adminSessionRevoke_(p.auth,p.id));
      case 'admin_logout_all': return jsonOut_(adminLogoutAll_(p.auth));
      case 'admin_page_flags_save': return jsonOut_(adminPageFlagsSave_(p.auth,p.page,p.enabled,p.sender,p.message));
      case 'admin_page_add': return jsonOut_(adminPageAdd_(p.auth,p.page,p.title));
      case 'admin_privacy_save': return jsonOut_(adminPrivacySave_(p.auth,p.config));
      case 'admin_wa_key_save': return jsonOut_(adminWaKeySave_(p.auth,p.phone,p.key));
      case 'admin_wa_key_delete': return jsonOut_(adminWaKeyDelete_(p.auth,p.phone));
      case 'admin_wa_test': return jsonOut_(adminWaTest_(p.auth,p.phone));
      case 'admin_delivery_clear': return jsonOut_(adminDeliveryClear_(p.auth));
      case 'admin_page_remove': return jsonOut_(adminPageRemove_(p.auth,p.page));
      case 'admin_reminder_delete': return jsonOut_(adminReminderDelete_(p.auth,p.type,p.address,p.id));
      case 'admin_receiver_clear': return jsonOut_(receiverClear_(p.auth,p.type,p.address,p.reminderId));
      case 'admin_receiver_set': return jsonOut_(receiverSet_(p.auth,p.type,p.address,p.status,p.until,p.reminderId));
      case 'admin_mail_template_save': return jsonOut_(mailTemplateSave_(p.auth,p.type,p.subject,p.html));
      case 'receiver_confirm': return receiverConfirm_(p);
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
    if (!pageFlag_('jobs').sender) return;      // switched off in Admin > Sender
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
        if (r.channel === 'whatsapp_group') { deliveryLog_('whatsapp_group', r.address, 'WhatsApp groups are not supported (CallMeBot only sends to single numbers). Add the people as numbers instead.'); return; }
        s.roles = r.mode === 'specific' ? (r.roles || []) : r.mode === 'random' ? [] : kw;
        s.random = r.mode === 'random';
        var due = scheduleDueSlot_(key, s, now);
        if (due) sendReminder_(key, s, false, due);
      } catch (err) { Logger.log('Jobs reminder failed: ' + err.message); deliveryLog_(r.channel, r.address, err.message); }
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
    return '<p><a href="' + safeUrl_(j.u) + '"><b>' + familyHtml_(j.t) + '</b></a><br>' + familyHtml_(j.c) + ' &middot; ' + familyHtml_(j.l) + '</p>';
  }).join('') + (errText ? '<p style="color:#b00">Some searches failed: ' + familyHtml_(errText) + '</p>' : '');
  sendMail_(to, 'Job radar: ' + jobs.length + ' new openings' + (test ? ' (test)' : ''), html, jobs.length + ' new openings:\n' + jobs.slice(0, 20).map(function (j) { return '- ' + j.t + ' - ' + j.c + ' ' + j.u; }).join('\n'));
  if (!test) P.setProperty('SEEN', JSON.stringify(jsonProp_('SEEN', []).concat(jobs.map(function (j) { return j.id; })).slice(-500)));
  return jobs.length;
}

function radarApi_(e) {
  var p = (e && e.parameter) || {}, result;
  if (!g('TOKEN', '') || p.token !== g('TOKEN')) return jsonOut_({ error: 'Wrong token' });
  if (!pageFlag_('jobs').enabled) return jsonOut_({ error: 'page_disabled', message: pageStatusPublic_('jobs').message });
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
      '&results_per_page=20&max_days_old=' + d + '&what=' + encodeURIComponent(what) + (where ? '&where=' + encodeURIComponent(where) : '');
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
  q.maxResults=Math.max(1,Math.min(100,Math.floor(Number(q.maxResults)||25))); var names = Array.isArray(q.portals) ? q.portals : [], pairs = (Array.isArray(q.pairs) ? q.pairs : []).slice(0, 12);
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
      '&results_per_page=' + q.maxResults + '&max_days_old=' + days + flags + (what ? '&what=' + encodeURIComponent(what) : '') + (loc && !/^(india|remote)$/i.test(loc) ? '&where=' + encodeURIComponent(loc) : ''), muteHttpExceptions: true };
  });
  var total = 0, ok = 0, bad = 0, firstErr = '', fetched=[];
  UrlFetchApp.fetchAll(reqs).forEach(function (r) {
    if (r.getResponseCode() === 200) { try { var parsed=JSON.parse(r.getContentText());total += Number(parsed.count) || 0; (parsed.results||[]).forEach(function(j){if(fetched.length<q.maxResults)fetched.push({title:String(j.title||''),company:String((j.company||{}).display_name||''),location:String((j.location||{}).display_name||''),url:String(j.redirect_url||'')});});ok++; } catch (e) { bad++; } }
    else { bad++; if (!firstErr) firstErr = 'HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 100); }
  });
  if (!ok) return { name: name, status: 'error', count: null, message: 'Adzuna search failed. ' + firstErr };
  return { name: name, status: 'ok', count: total, jobs:fetched, message: 'Fetched '+fetched.length+' job records (maximum '+q.maxResults+'); total matches '+total+'. Last ' + days + ' day(s), ' + ok + ' search(es); duplicates possible.' + (bad ? ' ' + bad + ' failed.' : '') };
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
  if (s.channel === 'email' && receiverBlocked_(key.indexOf('FAM_') === 0 ? 'family' : 'jobs',address,key)) return {ok:true,skipped:true,reason:'Receiver paused or disabled'};
  var res = fetchAdzunaJobs_(s, 100), seen = jsonProp_('SEEN_' + key, []);
  var jobs = force ? res.jobs : res.jobs.filter(function (j) { return seen.indexOf(j.id) < 0; });   // NEW: do not resend the same jobs every run
  if (s.random) jobs = jobs.sort(function () { return Math.random() - 0.5; });
  var errors = res.errors, roles = (s.roles || []).join(', ') || 'All jobs', locations = (s.locations || []).join(', ') || 'Any location';
  var hi = s.firstName ? 'Hi ' + familyHtml_(s.firstName) + ',' : '';
  if (s.channel === 'email') {
    var html = reminderHtml_(hi, roles, locations, jobs, errors, s.days, !!mailTemplatePublic_(key.indexOf('FAM_') === 0 ? 'family' : 'jobs'));
    sendMail_(address, 'CareerPulse: ' + jobs.length + ' new job' + (jobs.length === 1 ? '' : 's') + ' for ' + roles, html,
      'CareerPulse found ' + jobs.length + ' new job(s) for ' + roles + ' in ' + locations + '.\n' + jobs.slice(0, 10).map(function (j) { return '- ' + j.t + ' - ' + j.c + ' ' + j.u; }).join('\n'),key.indexOf('FAM_')===0?'family':'jobs',key);
  } else if (s.channel === 'whatsapp') {
    waSend_(address, waMessage_(s.firstName ? 'Hi ' + s.firstName + ',' : '', roles, jobs));
  } else throw new Error('Unsupported delivery channel: ' + s.channel);
  if (!force) {
    if (dueSlot) P.setProperty('SENT_' + key, dueSlot.marker);
    P.setProperty('SEEN_' + key, JSON.stringify(seen.concat(jobs.map(function (j) { return j.id; })).slice(-100)));
  }
  return { ok: true, jobs: jobs.length, errors: errors, recipient: address };
}
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
    if (!pageFlag_('family').sender) return;   // switched off in Admin > Sender
    var now = new Date();
    familySupabaseReminders_().forEach(function (row) {
      try { var key = 'FAM_' + row.id, s = familySettings_(row), due = scheduleDueSlot_(key, s, now); if (due) sendReminder_(key, s, false, due); }
      catch (err) { Logger.log('Family reminder ' + row.id + ' failed: ' + err.message); deliveryLog_((row.settings || {}).channel, (row.settings || {}).address, err.message); }
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
  var c = jwtClaims_(tok), rv = Number(P.getProperty('ADMIN_REVOKE_' + shortKey_(u.id)) || 0);
  if (rv && Number(c.iat || 0) * 1000 < rv) throw new Error('You were signed out of all devices. Please sign in again.');
  u._sid = sessionId_(c);
  var ss = sessionsLoad_(); if (ss[u._sid] && ss[u._sid].x) throw new Error('This device was signed out. Please sign in again.');
  var idle = idleMinutes_(), en = ss[u._sid];
  if (idle > 0 && en && en.a && Date.now() - en.a > idle * 60000 + 120000) throw new Error('You were signed out after ' + idle + ' minutes of inactivity. Please sign in again.');
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
    jobSources: portalEnabled_().length, familyReminders: fam, familyError: famErr, lastJobsRun: g('LAST', '') || 'Not yet', lastJobsError: g('LASTERR', ''), familyCheck: 'Every 5 minutes', pagesOff: Object.keys(PAGE_BUILTIN_).filter(function (k) { return !pageFlag_(k).enabled; }), sendersOff: Object.keys(PAGE_BUILTIN_).filter(function (k) { return !pageFlag_(k).sender; }), activeSessions: activeAdminSessions_() } };
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
  return { ok: true, data: raw, show: showFlags_(), cc: g('DEFAULT_CC', '91'), effective: { fromAccount: (function () { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } })(), name: i.name, replyTo: showFlag_('mailReplyTo') ? i.replyTo : (function () { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } })(), phone: i.phone, email: i.email, web: i.web } };
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

/* Admin monitoring and receiver delivery preferences. Script Properties persist across deployments. */
function receiverKey_(type,address,reminderId){return 'RCV_'+shortKey_(String(type).toLowerCase()+':'+String(address).trim().toLowerCase()+(reminderId?':'+reminderId:''));}
function receiverList_(auth){portalRequireAdmin_(auth);var p=PropertiesService.getScriptProperties().getProperties(),rows=[];Object.keys(p).filter(function(k){return k.indexOf('RCV_')===0;}).forEach(function(k){try{rows.push(JSON.parse(p[k]));}catch(e){}});return {ok:true,rows:rows};}

/* Secure receiver self-service: signed expiring links; confirmation required. */
function receiverSecret_(){var p=PropertiesService.getScriptProperties(),v=p.getProperty('RECEIVER_LINK_SECRET');if(!v){v=Utilities.getUuid()+Utilities.getUuid();p.setProperty('RECEIVER_LINK_SECRET',v);}return v;}
function receiverSignature_(payload){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload,receiverSecret_())).replace(/=+$/,'');}

function familyAllSupabaseReminders_(){
  familyRequireSupabase_();
  var key=g('SUPABASE_SERVICE_KEY','');
  var url=String(g('SUPABASE_URL','')).replace(/\/+$/,'')+'/rest/v1/family_reminders?select=id,created_at,enabled,settings&order=created_at.asc';
  var resp=UrlFetchApp.fetch(url,{method:'get',headers:{apikey:key,Authorization:'Bearer '+key},muteHttpExceptions:true});
  if(resp.getResponseCode()!==200)throw Error('Family reminders HTTP '+resp.getResponseCode()+': '+resp.getContentText().slice(0,200));
  var rows=JSON.parse(resp.getContentText()||'[]');return Array.isArray(rows)?rows:[];
}

/* Admin list of actual reminders; no invented recipients. */
function receiverReminders_(auth,type){
  portalRequireAdmin_(auth);type=String(type||'family').toLowerCase();var rows=[];
  if(type==='family'){
    familyAllSupabaseReminders_().forEach(function(r){var v=r.settings||{};if(v.channel==='email'&&v.address)rows.push({id:'FAM_'+r.id,address:String(v.address).toLowerCase(),label:(v.roles||[]).join(', ')||'Family reminder',locations:(v.locations||[]).join(', '),enabled:r.enabled!==false});});
  }else if(type==='jobs'){
    jobsRecipients_().forEach(function(r){if(r.channel==='email'&&r.address)rows.push({id:'JOBS_'+shortKey_(r.channel+':'+String(r.address).toLowerCase()),address:String(r.address).toLowerCase(),label:(r.roles||[]).join(', ')||'Jobs reminder',locations:'',enabled:true});});
  }else throw Error('Invalid reminder type');
  rows.forEach(function(r){var rec=receiverRec_(type,r.address,r.id)||receiverRec_(type,r.address,'');r.status=receiverEffective_(rec);r.until=rec&&r.status==='paused'?(rec.until||''):'';});
  return {ok:true,rows:rows};
}
/* Templates are admin-only and persist across deployments. {{content}} and {{controls}} are placeholders. */
function mailTemplates_(auth){portalRequireAdmin_(auth);return {ok:true,jobs:mailTemplatePublic_('jobs'),family:mailTemplatePublic_('family')};}

/* =====================================================================
   v9 ADDITIONS
   - Recipient e-mail buttons (Enable / Pause / Disable / Delete) that really work
   - Receiver override logic (per-reminder record wins over the address-wide one)
   - Admin Family list / enable / delete through the service key (bypasses RLS)
   - Health monitor with failure AND recovery events, downtime, server-side checks
   ===================================================================== */
/* The URL the current request arrived on. Saved automatically so e-mail buttons always use the newest deployment
   (no more editing DEFAULT_WEBAPP_URL_ after a new deployment). */
function liveUrl_() { try { var u = ScriptApp.getService().getUrl() || ''; return /\/exec$/.test(u) ? u : ''; } catch (e) { return ''; } }
function rememberUrl_() { try { var u = liveUrl_(); if (u && g('WEBAPP_URL', '') !== u) P.setProperty('WEBAPP_URL', u); } catch (e) {} }
var DEFAULT_WEBAPP_URL_ = 'https://script.google.com/macros/s/AKfycbwx3QFZKcyvzKzXRca_ms_g7l4rj-G45VvmdMq9-qY13BQ2fJGbXGofJeei5pqDtld3/exec';
/* Public /exec URL used inside e-mail links. Optional script property WEBAPP_URL overrides it.
   (ScriptApp.getService().getUrl() can return a /dev URL when called from a trigger, which broke the links.) */
function webAppUrl_() {
  var u = g('WEBAPP_URL', '').trim();
  if (/^https:\/\/script\.google\.com\/.+\/exec$/.test(u)) return u;
  try { u = ScriptApp.getService().getUrl() || ''; } catch (e) { u = ''; }
  return /\/exec$/.test(u) ? u : DEFAULT_WEBAPP_URL_;
}
function safeUrl_(u) { u = String(u || ''); return familyHtml_(/^https?:\/\//i.test(u) ? u : '#'); }

/* ---------- attractive default reminder e-mail ---------- */
function reminderHtml_(hi, roles, locations, jobs, errors, days, bare) {
  var cards = jobs.length ? jobs.slice(0, 25).map(function (j) {
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e1eaee;border-radius:12px;margin:0 0 10px;background:#ffffff"><tr><td style="padding:14px 16px">' +
      '<div style="font-size:16px;font-weight:bold;color:#17323e">' + familyHtml_(j.t) + '</div>' +
      '<div style="font-size:13px;color:#5b707b;margin:3px 0 10px">' + familyHtml_(j.c || 'Company not listed') + ' &middot; ' + familyHtml_(j.l || 'Location not listed') + '</div>' +
      '<a href="' + safeUrl_(j.u) + '" style="display:inline-block;padding:8px 14px;border-radius:8px;background:#087c6d;color:#ffffff;text-decoration:none;font-size:13px;font-weight:bold">View &amp; apply</a></td></tr></table>';
  }).join('') : '<div style="padding:16px;border:1px dashed #b9ccd3;border-radius:12px;color:#4b616c">No new matching openings right now. Your reminder is still active.</div>';
  var intro = '<p style="font-size:13px;color:#4b616c;margin:0 2px 14px">' + (hi ? '<b>' + hi + '</b><br>' : '') + '<b>Roles:</b> ' + familyHtml_(roles) + '<br><b>Locations:</b> ' + familyHtml_(locations) + '</p>';
  var err = errors.length ? '<p style="color:#a33;font-size:12px">Some searches failed: ' + familyHtml_(errors.join(' | ')) + '</p>' : '';
  var foot = '<p style="color:#60747d;font-size:12px;margin:10px 2px">Posted within ' + familyHtml_(String(days || 7)) + ' day(s)</p>';
  if (bare) return intro + cards + err + foot;      // custom format from Admin > Email formats wraps this
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:auto;background:#f3f7fa;padding:18px;border-radius:16px">' +
    '<div style="background:linear-gradient(135deg,#087c6d,#155d75);background-color:#087c6d;color:#ffffff;padding:22px 24px;border-radius:14px">' +
    '<div style="font-size:12px;letter-spacing:2px;font-weight:bold;opacity:.9">CAREERPULSE</div><div style="font-size:22px;font-weight:bold;margin-top:6px">' + (hi ? hi : 'Your job reminder') + '</div>' +
    '<div style="font-size:13px;margin-top:6px;opacity:.92">' + jobs.length + ' new opening' + (jobs.length === 1 ? '' : 's') + ' for you</div></div>' +
    '<p style="font-size:13px;color:#4b616c;margin:14px 2px"><b>Roles:</b> ' + familyHtml_(roles) + '<br><b>Locations:</b> ' + familyHtml_(locations) + '</p>' + cards +
    (errors.length ? '<p style="color:#a33;font-size:12px">Some searches failed: ' + familyHtml_(errors.join(' | ')) + '</p>' : '') +
    '<p style="color:#60747d;font-size:12px;margin:10px 2px">Posted within ' + familyHtml_(String(days || 7)) + ' day(s)</p></div>';
}

/* ---------- receiver overrides ---------- */
function receiverRec_(type, address, id) {
  var raw = P.getProperty(receiverKey_(type, address, id));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function receiverEffective_(rec) {
  if (!rec) return 'active';
  if (rec.status === 'paused' && rec.until && Date.now() >= Date.parse(rec.until)) return 'active';   // pause expired
  return rec.status || 'active';
}
/* Store a preference. 'active' on one reminder writes an explicit record when an address-wide block exists,
   otherwise the Resume button used to do nothing because the wide block still applied. */
function receiverStore_(type, address, status, until, reminderId, source) {
  var key = receiverKey_(type, address, reminderId);
  if (status === 'active') {
    if (reminderId && receiverRec_(type, address, '')) P.setProperty(key, JSON.stringify({ type: type, address: address, reminderId: String(reminderId), status: 'active', until: '', updated: new Date().toISOString(), source: source || 'admin' }));
    else P.deleteProperty(key);
    return;
  }
  P.setProperty(key, JSON.stringify({ type: type, address: address, reminderId: String(reminderId || ''), status: status, until: status === 'paused' ? String(until || '') : '', updated: new Date().toISOString(), source: source || 'admin' }));
}
function receiverBlocked_(type, address, reminderId) {
  var rec = reminderId ? receiverRec_(type, address, reminderId) : null;
  if (!rec) rec = receiverRec_(type, address, '');
  var st = receiverEffective_(rec);
  return st === 'disabled' || st === 'paused';
}
function receiverSet_(auth, type, address, status, until, reminderId) {
  portalRequireAdmin_(auth); type = String(type || '').toLowerCase(); address = String(address || '').trim().toLowerCase(); status = String(status || '');
  if (['jobs', 'family'].indexOf(type) < 0 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address) || ['active', 'paused', 'disabled'].indexOf(status) < 0) throw Error('Invalid type, email or status');
  if (until && !isFinite(Date.parse(until))) throw Error('Invalid pause end time');
  receiverStore_(type, address, status, until, reminderId, 'admin');
  return { ok: true };
}
function receiverClear_(auth, type, address, reminderId) {
  portalRequireAdmin_(auth); P.deleteProperty(receiverKey_(String(type || '').toLowerCase(), String(address || '').trim().toLowerCase(), reminderId)); return { ok: true };
}

/* ---------- delete a reminder (Family row or Jobs recipient) ---------- */
function familyDeleteRows_(ids) {
  ids = (ids || []).map(String).filter(function (x) { return /^[A-Za-z0-9_-]+$/.test(x); });
  if (!ids.length) throw Error('No valid reminder ids.');
  var rows = jobDbReq_('delete', 'family_reminders?id=in.(' + ids.map(encodeURIComponent).join(',') + ')', undefined, 'return=representation');
  ids.forEach(function (id) { P.deleteProperty('SENT_FAM_' + id); P.deleteProperty('SEEN_FAM_' + id); });
  return rows.length;
}
function reminderDelete_(type, address, id) {
  address = String(address || '').trim().toLowerCase(); id = String(id || '');
  if (type === 'family') {
    if (id.indexOf('FAM_') !== 0) throw Error('Missing reminder id.');
    var n = familyDeleteRows_([id.slice(4)]);
    if (!n) throw Error('Reminder was not found (already deleted?).');
    P.deleteProperty(receiverKey_('family', address, id));
    return;
  }
  var list = jsonProp_('RECIPIENTS_JSON', []), before = list.length;
  list = list.filter(function (r) { return !(r && 'JOBS_' + shortKey_(r.channel + ':' + String(r.address).toLowerCase()) === id); });
  P.setProperty('RECIPIENTS_JSON', JSON.stringify(list));
  P.deleteProperty('SENT_' + id); P.deleteProperty('SEEN_' + id);
  // keep it blocked: the Jobs page re-sends its own recipient list on the next "Save & sync"
  receiverStore_('jobs', address, 'disabled', '', id, 'delete');
  if (list.length === before) Logger.log('Jobs recipient ' + address + ' was not in RECIPIENTS_JSON; blocked instead.');
}
function adminReminderDelete_(auth, type, address, id) {
  portalRequireAdmin_(auth); type = String(type || '').toLowerCase();
  if (['jobs', 'family'].indexOf(type) < 0) throw Error('Invalid reminder type');
  reminderDelete_(type, address, id); return { ok: true };
}

/* ---------- e-mail link handling ---------- */
var RECEIVER_ACTIONS_ = [['active', 'Enable', '#0b8a6f'], ['paused', 'Pause', '#c98a12'], ['disabled', 'Disable', '#5b6b76'], ['delete', 'Delete', '#b83232']];
function receiverMailControls_(type, address, reminderId) {
  var base = webAppUrl_(), expiry = Date.now() + 90 * 86400000, addr = String(address).trim().toLowerCase(), rid = String(reminderId || '');
  var links = RECEIVER_ACTIONS_.map(function (a) {
    var payload = [type, addr, a[0], expiry, rid].join('|');
    return { label: a[1], color: a[2], url: base + '?action=receiver_manage&type=' + encodeURIComponent(type) + '&address=' + encodeURIComponent(addr) + '&status=' + a[0] + '&expires=' + expiry + '&reminderId=' + encodeURIComponent(rid) + '&sig=' + encodeURIComponent(receiverSignature_(payload)) };
  });
  var btns = links.map(function (x) { return '<a href="' + x.url + '" style="display:inline-block;margin:4px 6px 4px 0;padding:9px 16px;border-radius:8px;background:' + x.color + ';color:#ffffff;text-decoration:none;font-weight:bold;font-size:13px">' + x.label + '</a>'; }).join('');
  return {
    text: links.map(function (x) { return x.label + ': ' + x.url; }).join('\n'),
    html: '<div style="margin-top:18px;padding:14px;border:1px solid #dde6ea;border-radius:12px;background:#f7fafb"><div style="font-size:13px;font-weight:bold;color:#17323e;margin-bottom:6px">Manage these ' + type + ' reminders</div>' + btns +
      '<div style="font-size:11px;color:#667782;margin-top:6px">Each button asks you to confirm first. Delete removes the reminder permanently.</div></div>'
  };
}
function receiverVerify_(p) {
  var type = String(p.type || ''), address = String(p.address || '').trim().toLowerCase(), status = String(p.status || ''), expires = Number(p.expires);
  if (['jobs', 'family'].indexOf(type) < 0 || ['active', 'paused', 'disabled', 'delete'].indexOf(status) < 0 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address) || !isFinite(expires) || expires < Date.now() || expires > Date.now() + 92 * 86400000) throw Error('Invalid or expired link');
  var payload = [type, address, status, expires, String(p.reminderId || '')].join('|');
  if (receiverSignature_(payload) !== String(p.sig || '')) throw Error('Invalid signature');
  return { type: type, address: address, status: status, reminderId: String(p.reminderId || '') };
}
function receiverApply_(x) {
  if (x.status === 'delete') { reminderDelete_(x.type, x.address, x.reminderId); return 'Your ' + x.type + ' reminder was deleted. You will not receive it again.'; }
  receiverStore_(x.type, x.address, x.status, '', x.reminderId, 'email');
  return { active: 'Your ' + x.type + ' reminders are enabled again.', paused: 'Your ' + x.type + ' reminders are paused. Use the Enable button in any earlier email to resume.', disabled: 'Your ' + x.type + ' reminders are disabled.' }[x.status];
}
/* Called from the confirm page through google.script.run (a plain <form> POST is blocked inside the Apps Script sandbox iframe). Needs the signed link parameters. */
function receiverApplyPublic(p) { return receiverApply_(receiverVerify_(p || {})); }
function receiverPageShell_(inner, script) {
  return HtmlService.createHtmlOutput('<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{margin:0;font:16px system-ui,Segoe UI,Arial,sans-serif;background:#f3f7fa;color:#17323e;display:grid;place-items:center;min-height:100vh;padding:16px;box-sizing:border-box}.box{max-width:460px;width:100%;background:#fff;border-radius:20px;padding:30px 26px;box-shadow:0 14px 40px #17323e22;text-align:center}.ic{font-size:44px}h2{margin:10px 0 6px}p{color:#4b616c;line-height:1.55}button{border:0;border-radius:12px;padding:13px 26px;font-size:16px;font-weight:700;color:#fff;background:#087c6d;cursor:pointer;margin-top:8px}button:disabled{opacity:.6}.bad{background:#b83232}.out{margin-top:14px;font-weight:600}</style></head><body><div class="box">' + inner + '</div>' + (script || '') + '</body></html>').setTitle('CareerPulse reminders');
}
function receiverManagePage_(p) {
  try {
    var x = receiverVerify_(p), words = { active: 'enable', paused: 'pause', disabled: 'disable', 'delete': 'permanently delete' };
    var params = { type: String(p.type || ''), address: String(p.address || ''), status: String(p.status || ''), expires: String(p.expires || ''), sig: String(p.sig || ''), reminderId: String(p.reminderId || '') };
    var json = JSON.stringify(params).replace(/</g, '\\u003c');
    var danger = x.status === 'delete' || x.status === 'disabled';
    return receiverPageShell_('<div class="ic">' + (x.status === 'delete' ? '🗑️' : x.status === 'active' ? '✅' : x.status === 'paused' ? '⏸️' : '🚫') + '</div><h2>Please confirm</h2><p>Do you want to <b>' + words[x.status] + '</b> the <b>' + familyHtml_(x.type) + '</b> reminders for <b>' + familyHtml_(x.address) + '</b>?</p><button id="ok" class="' + (danger ? 'bad' : '') + '">Yes, ' + words[x.status] + '</button><p class="out" id="out"></p><p style="font-size:13px">To cancel, just close this page.</p>',
      '<script>var P=' + json + ';var ok=document.getElementById("ok"),out=document.getElementById("out");ok.onclick=function(){ok.disabled=true;out.textContent="Saving…";google.script.run.withSuccessHandler(function(m){ok.style.display="none";out.style.color="#0b6b55";out.textContent="✓ "+m}).withFailureHandler(function(e){ok.disabled=false;out.style.color="#b83232";out.textContent="Could not save: "+(e&&e.message?e.message:e)}).receiverApplyPublic(P)}</script>');
  } catch (e) { return receiverPageShell_('<div class="ic">⌛</div><h2>Link expired or invalid</h2><p>No changes were made. Please use the buttons in your latest reminder email.</p>'); }
}
/* Kept so links in older e-mails (plain form POST) still work. */
function receiverConfirm_(p) {
  try { var msg = receiverApply_(receiverVerify_(p)); return receiverPageShell_('<div class="ic">✅</div><h2>Preference saved</h2><p>' + familyHtml_(msg) + '</p>'); }
  catch (e) { return receiverPageShell_('<div class="ic">⌛</div><h2>Link expired or invalid</h2><p>No changes were made.</p>'); }
}

/* ---------- admin: Family reminders through the service key (RLS can silently ignore browser updates / deletes) ---------- */
function adminFamilyList_(auth) { portalRequireAdmin_(auth); return { ok: true, rows: familyAllSupabaseReminders_() }; }
function adminFamilySetEnabled_(auth, id, enabled) {
  portalRequireAdmin_(auth); id = String(id || '');
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw Error('Invalid reminder id.');
  var rows = jobDbReq_('patch', 'family_reminders?id=eq.' + encodeURIComponent(id), { enabled: String(enabled) === 'true' }, 'return=representation');
  if (!rows.length) throw Error('Reminder not found.');
  return { ok: true, enabled: rows[0].enabled };
}
function adminFamilyDelete_(auth, ids) {
  portalRequireAdmin_(auth);
  var list = Array.isArray(ids) ? ids : String(ids || '').indexOf('[') === 0 ? JSON.parse(ids) : String(ids || '').split(',');
  return { ok: true, deleted: familyDeleteRows_(list) };
}

/* ---------- health monitor: failures + recoveries ---------- */
var MON_SERVICES_ = ['jobs', 'family', 'backend', 'supabase'];
function jobsHealth_() {
  var on = g('ENABLED', '0') === '1', trig = trigExists_('digest') || trigExists_('jobsTick'), az = !!g('ADZUNA_ID', '') && !!g('ADZUNA_KEY', '');
  if (!pageFlag_('jobs').sender) return { ok: true, reason: 'Jobs sender is paused from the Admin page' };
  if (!on) return { ok: true, reason: 'Jobs reminders are switched off' };
  if (!trig) return { ok: false, reason: 'Jobs reminders are ON but no scheduler trigger exists. Press Save & sync on the Jobs page or run setup().' };
  if (!az) return { ok: false, reason: 'ADZUNA_ID / ADZUNA_KEY are missing in Script properties.' };
  return { ok: true, reason: 'Jobs sender running' };
}
function familyHealth_() {
  if (!pageFlag_('family').sender) return { ok: true, reason: 'Family sender is paused from the Admin page' };
  return trigExists_('familySupabaseTick') ? { ok: true, reason: 'Family sender running' } : { ok: false, reason: 'Family sender trigger is missing. Run setupFamilySupabaseSender().' };
}
function monSave_(rows) {
  var s = JSON.stringify(rows);
  while (rows.length > 1 && Utilities.newBlob(s).getBytes().length > 8500) { rows.pop(); s = JSON.stringify(rows); }   // a Script property holds at most ~9 KB: older code threw here and stopped logging
  P.setProperty('MONITOR_LOGS', s);
}
/* Logs only transitions: first failure, and the recovery that follows (with downtime). Returns true when something was written. */
function monitorLog_(service, status, durationMs, reason, source, atIso) {
  if (MON_SERVICES_.indexOf(service) < 0) return false;
  status = status === 'working' ? 'working' : 'lost';
  var now = Date.now(), at = Date.parse(atIso || '');
  if (!isFinite(at) || at > now + 60000 || at < now - 7 * 86400000) at = now;
  var state = jsonProp_('MONITOR_STATE', {}), prev = state[service] || {}, row = null;
  reason = String(reason || '').slice(0, 160);
  if (status === 'lost' && prev.status !== 'lost') {
    state[service] = { status: 'lost', since: at };
    row = { t: new Date(at).toISOString(), s: service, v: 'lost', d: Number(durationMs) || 0, r: reason || 'No reason provided', o: source || '' };
  } else if (status === 'working' && prev.status === 'lost') {
    var down = Math.max(0, at - (Number(prev.since) || at));
    state[service] = { status: 'working', since: at };
    row = { t: new Date(at).toISOString(), s: service, v: 'working', d: Number(durationMs) || 0, r: reason || 'Connection restored', m: down, f: Number(prev.since) || 0, o: source || '' };
  } else if (status === 'working' && prev.status !== 'working') {
    state[service] = { status: 'working', since: at };   // first ever observation: remember it, nothing to log
  } else return false;
  P.setProperty('MONITOR_STATE', JSON.stringify(state));
  if (row) { var rows = jsonProp_('MONITOR_LOGS', []); rows.unshift(row); monSave_(rows); }
  return !!row;
}
function monitorRecord_(auth, record) {
  portalRequireAdmin_(auth);
  var x = JSON.parse(record || '[]'), list = (Array.isArray(x) ? x : [x]).filter(function (r) { return r && MON_SERVICES_.indexOf(r.service) >= 0; });
  list.sort(function (a, b) { return (Date.parse(a.at) || 0) - (Date.parse(b.at) || 0); });   // replay queued browser events in order
  list.forEach(function (r) { monitorLog_(r.service, r.status, r.durationMs, r.reason, 'admin page', r.at); });
  return { ok: true, recorded: list.length };
}
function monitorLogs_(auth) {
  portalRequireAdmin_(auth);
  var rows = jsonProp_('MONITOR_LOGS', []).map(function (r) {
    return { time: r.t, service: r.s, status: r.v, durationMs: r.d || 0, reason: r.r || '', downMs: r.m || 0, since: r.f ? new Date(Number(r.f)).toISOString() : '', source: r.o || '' };
  });
  var st = jsonProp_('MONITOR_STATE', {}), state = {};
  Object.keys(st).forEach(function (k) { state[k] = { status: st[k].status, since: new Date(Number(st[k].since)).toISOString() }; });
  return { ok: true, rows: rows, state: state };
}
/* Runs on a 5-minute trigger even when nobody has the Admin page open, so a recovery is always logged. Run setupMonitor() once. */
function monitorTick() {
  monitorLog_('backend', 'working', 0, 'Backend responding', 'server');
  var t = Date.now(), ok = false, err = '';
  try { familySupabaseReminders_(); ok = true; } catch (e) { err = e.message; }
  monitorLog_('supabase', ok ? 'working' : 'lost', Date.now() - t, ok ? 'Supabase reachable' : err, 'server');
  var f = familyHealth_(), j = jobsHealth_();
  monitorLog_('family', f.ok ? 'working' : 'lost', 0, f.reason, 'server');
  monitorLog_('jobs', j.ok ? 'working' : 'lost', 0, j.reason, 'server');
}
function setupMonitor() { dropTriggers_(['monitorTick']); ScriptApp.newTrigger('monitorTick').timeBased().everyMinutes(5).create(); Logger.log('Health monitor enabled (every 5 minutes).'); }
function disableMonitor() { dropTriggers_(['monitorTick']); Logger.log('Health monitor disabled.'); }

/* =====================================================================
   v10: ADMIN SESSIONS  (active-session count, per-device sign-out, sign out everywhere)
   Supabase does not expose a session list to the REST API, so each Admin page registers
   itself with a heartbeat. Revocation is enforced here (portalRequireAdmin_) AND by a
   Supabase global sign-out, which also kills refresh tokens on every device.
   ===================================================================== */
var SESSION_ACTIVE_MS_ = 5 * 60 * 1000;
function jwtClaims_(tok) {
  try { var part = String(tok).split('.')[1].replace(/-/g, '+').replace(/_/g, '/'); while (part.length % 4) part += '='; return JSON.parse(Utilities.newBlob(Utilities.base64Decode(part)).getDataAsString()); } catch (e) { return {}; }
}
function sessionId_(c) { return String(c.session_id || ('t' + (c.iat || 0))); }
function sessionsLoad_() { var o = jsonProp_('ADMIN_SESSIONS', {}); return (o && typeof o === 'object' && !Array.isArray(o)) ? o : {}; }
function sessionsSave_(ss) {
  var now = Date.now();
  Object.keys(ss).forEach(function (k) { if (now - (ss[k].s || 0) > 7 * 86400000) delete ss[k]; });
  var s = JSON.stringify(ss);
  while (Utilities.newBlob(s).getBytes().length > 8500) {          // Script property limit ~9 KB: drop the stalest entry
    var oldest = Object.keys(ss).sort(function (a, b) { return (ss[a].s || 0) - (ss[b].s || 0); })[0]; if (!oldest) break;
    delete ss[oldest]; s = JSON.stringify(ss);
  }
  P.setProperty('ADMIN_SESSIONS', s);
}
function activeAdminSessions_() {
  var ss = sessionsLoad_(), now = Date.now();
  return Object.keys(ss).filter(function (k) { return !ss[k].x && now - (ss[k].s || 0) <= SESSION_ACTIVE_MS_; }).length;
}
function adminSessionPing_(auth, label, activeAt) {
  var u = portalRequireAdmin_(auth), ss = sessionsLoad_(), now = Date.now(), e = ss[u._sid] || { f: now };
  e.u = u.id; e.m = String(u.email || '').slice(0, 60); e.l = String(label || 'Unknown device').slice(0, 50); e.s = now; delete e.x; var aa = Number(activeAt); e.a = (isFinite(aa) && aa > 0) ? Math.min(aa, now) : now;   // last real user activity, reported by the page
  ss[u._sid] = e; sessionsSave_(ss);
  return { ok: true, id: shortKey_(u._sid), active: activeAdminSessions_(), idleMinutes: idleMinutes_() };
}
function adminSessions_(auth) {
  var u = portalRequireAdmin_(auth), ss = sessionsLoad_(), now = Date.now();
  var rows = Object.keys(ss).filter(function (k) { return !ss[k].x; }).map(function (k) {
    var e = ss[k];
    return { id: shortKey_(k), email: e.m || '', device: e.l || 'Unknown device', first: new Date(e.f || e.s).toISOString(), last: new Date(e.s).toISOString(), active: now - e.s <= SESSION_ACTIVE_MS_, current: k === u._sid };
  }).sort(function (a, b) { return Date.parse(b.last) - Date.parse(a.last); });
  return { ok: true, idleMinutes: idleMinutes_(), active: rows.filter(function (r) { return r.active; }).length, known: rows.length, windowMinutes: SESSION_ACTIVE_MS_ / 60000, rows: rows.slice(0, 30) };
}
function adminSessionRevoke_(auth, id) {
  var u = portalRequireAdmin_(auth), ss = sessionsLoad_(), hit = Object.keys(ss).filter(function (k) { return shortKey_(k) === String(id || ''); })[0];
  if (!hit) throw Error('Session not found (it may have expired).');
  ss[hit].x = true; ss[hit].s = Date.now(); sessionsSave_(ss);
  return { ok: true, current: hit === u._sid };
}
function adminLogoutAll_(auth) {
  var u = portalRequireAdmin_(auth);
  P.setProperty('ADMIN_REVOKE_' + shortKey_(u.id), String(Date.now()));       // 1) our own enforcement: tokens issued before now are refused
  var ss = sessionsLoad_(), n = 0;
  Object.keys(ss).forEach(function (k) { if (ss[k].u === u.id) { delete ss[k]; n++; } });
  sessionsSave_(ss);
  var ok = false, code = 0;                                                       // 2) Supabase global sign-out: deletes every session and refresh token for this user
  try {
    var r = UrlFetchApp.fetch(String(g('SUPABASE_URL', '')).replace(/\/+$/, '') + '/auth/v1/logout?scope=global', { method: 'post', headers: { apikey: g('SUPABASE_SERVICE_KEY', ''), Authorization: 'Bearer ' + auth }, muteHttpExceptions: true });
    code = r.getResponseCode(); ok = code >= 200 && code < 300;
  } catch (e) { Logger.log('Global sign-out call failed: ' + e.message); }
  return { ok: true, sessionsCleared: n, supabaseSignedOut: ok, supabaseStatus: code };
}

/* =====================================================================
   v11: PAGE SWITCHES
   - "Page live" switch for any page (built-in jobs / family, plus pages you register)
   - "Sender" switch per page (stops the scheduled reminder e-mails of that page)
   ===================================================================== */
var PAGE_BUILTIN_ = { jobs: 'Jobs page', family: 'Family page' };
var PAGE_DEFAULT_MESSAGE_ = 'This page is temporarily unavailable. Please check back soon.';
function pageFlagsLoad_() { var o = jsonProp_('PAGE_FLAGS_JSON', {}); return (o && typeof o === 'object' && !Array.isArray(o)) ? o : {}; }
function pageSlug_(s) { s = String(s || '').toLowerCase().trim(); if (!/^[a-z0-9][a-z0-9_-]{0,29}$/.test(s)) throw new Error('Page id must be 1-30 characters: letters, numbers, - or _.'); return s; }
function pageFlag_(slug) {
  slug = String(slug || '').toLowerCase(); var f = pageFlagsLoad_()[slug] || {};
  return { slug: slug, enabled: f.enabled !== false, sender: f.sender !== false, message: f.message || '', title: f.title || PAGE_BUILTIN_[slug] || slug, custom: !!f.custom, updated: f.updated || '', by: f.by || '' };
}
function pageStatusPublic_(page) {
  var slug = String(page || '').toLowerCase().replace(/[^a-z0-9_-]/g, '').slice(0, 30), f = pageFlag_(slug);
  return { ok: true, page: slug, enabled: f.enabled, message: f.enabled ? '' : (f.message || PAGE_DEFAULT_MESSAGE_), title: f.title };
}
function adminPages_(auth) {
  portalRequireAdmin_(auth); var all = pageFlagsLoad_(), slugs = Object.keys(PAGE_BUILTIN_);
  Object.keys(all).forEach(function (k) { if (all[k].custom && slugs.indexOf(k) < 0) slugs.push(k); });
  return { ok: true, rows: slugs.map(function (k) { var f = pageFlag_(k); return { slug: k, title: f.title, builtin: !!PAGE_BUILTIN_[k], hasSender: !!PAGE_BUILTIN_[k], enabled: f.enabled, sender: f.sender, message: f.message, updated: f.updated, by: f.by }; }) };
}
function pageFlagsStore_(all) {
  var s = JSON.stringify(all); if (Utilities.newBlob(s).getBytes().length > 8500) throw new Error('Too many page settings. Remove unused custom pages first.');
  P.setProperty('PAGE_FLAGS_JSON', s);
}
function adminPageFlagsSave_(auth, page, enabled, sender, message) {
  var u = portalRequireAdmin_(auth), slug = pageSlug_(page), all = pageFlagsLoad_(), f = all[slug] || {};
  if (!PAGE_BUILTIN_[slug] && !f.custom) throw new Error('Unknown page "' + slug + '". Add it first.');
  var yes = function (v) { return v === true || String(v) === 'true'; };
  if (enabled !== undefined && enabled !== '') f.enabled = yes(enabled);
  if (sender !== undefined && sender !== '') { if (!PAGE_BUILTIN_[slug]) throw new Error('Only the Jobs and Family pages have a sender.'); f.sender = yes(sender); }
  if (message !== undefined) f.message = String(message).trim().slice(0, 200);
  f.updated = new Date().toISOString(); f.by = String(u.email || '').slice(0, 40);
  all[slug] = f; pageFlagsStore_(all);
  return adminPages_(auth);
}
function adminPageAdd_(auth, page, title) {
  portalRequireAdmin_(auth); var slug = pageSlug_(page), all = pageFlagsLoad_();
  if (PAGE_BUILTIN_[slug] || all[slug]) throw new Error('A page with this id already exists.');
  all[slug] = { custom: true, enabled: true, title: String(title || slug).trim().slice(0, 50), updated: new Date().toISOString() }; pageFlagsStore_(all);
  return adminPages_(auth);
}
function adminPageRemove_(auth, page) {
  portalRequireAdmin_(auth); var slug = pageSlug_(page), all = pageFlagsLoad_();
  if (!all[slug] || !all[slug].custom) throw new Error('Only custom pages can be removed.');
  delete all[slug]; pageFlagsStore_(all); return adminPages_(auth);
}

/* =====================================================================
   v12: PRIVACY OF SENDER DETAILS + WORKING WHATSAPP / MOBILE DELIVERY
   ===================================================================== */
var SHOW_FLAGS_ = { mailName: 'SHOW_MAIL_NAME', mailPhone: 'SHOW_MAIL_PHONE', mailEmail: 'SHOW_MAIL_EMAIL', mailWeb: 'SHOW_MAIL_WEB', mailReplyTo: 'SHOW_MAIL_REPLYTO', mailReplyLine: 'SHOW_MAIL_REPLYLINE', waName: 'SHOW_WA_NAME', waPhone: 'SHOW_WA_PHONE', waEmail: 'SHOW_WA_EMAIL' };
function showFlag_(k) { return g(SHOW_FLAGS_[k], '1') !== '0'; }
function showFlags_() { var o = {}; Object.keys(SHOW_FLAGS_).forEach(function (k) { o[k] = showFlag_(k); }); return o; }
function adminPrivacySave_(auth, raw) {
  portalRequireAdmin_(auth); var c = typeof raw === 'object' ? raw : JSON.parse(String(raw || '{}')), sh = c.show || {};
  Object.keys(SHOW_FLAGS_).forEach(function (k) { if (sh[k] !== undefined) P.setProperty(SHOW_FLAGS_[k], (sh[k] === true || String(sh[k]) === 'true') ? '1' : '0'); });
  if (c.cc !== undefined) { var cc = String(c.cc).replace(/\D/g, ''); if (cc && (cc.length > 4)) throw new Error('Country code must be 1 to 4 digits, e.g. 91.'); if (cc) P.setProperty('DEFAULT_CC', cc); else P.deleteProperty('DEFAULT_CC'); }
  return { ok: true, show: showFlags_(), cc: g('DEFAULT_CC', '91') };
}

/* ---- delivery problems list (so "it does not work" shows the real reason in Admin) ---- */
function deliveryLog_(channel, address, msg) {
  try {
    var rows = jsonProp_('DELIVERY_LOG', []), now = new Date().toISOString(); channel = String(channel || ''); address = String(address || ''); msg = String(msg || '').slice(0, 200);
    var top = rows[0];
    if (top && top.c === channel && top.a === address && top.m === msg) { top.n = (top.n || 1) + 1; top.t = now; }
    else rows.unshift({ t: now, c: channel, a: address, m: msg, n: 1 });
    rows = rows.slice(0, 15); var s = JSON.stringify(rows);
    while (rows.length > 1 && Utilities.newBlob(s).getBytes().length > 8000) { rows.pop(); s = JSON.stringify(rows); }
    P.setProperty('DELIVERY_LOG', s);
  } catch (e) {}
}
function adminDeliveryLog_(auth) { portalRequireAdmin_(auth); return { ok: true, rows: jsonProp_('DELIVERY_LOG', []).map(function (r) { return { time: r.t, channel: r.c, address: r.a, message: r.m, count: r.n || 1 }; }) }; }
function adminDeliveryClear_(auth) { portalRequireAdmin_(auth); P.deleteProperty('DELIVERY_LOG'); return { ok: true }; }

/* ---- WhatsApp (CallMeBot) ----
   CallMeBot gives EVERY phone number its own API key (the number must first message the CallMeBot contact to get it).
   The old code used one global WA_KEY for all numbers, so it only ever worked for the one number that key belonged to.
   Keys are now stored per number (Admin > Sender > WhatsApp keys); WA_KEY stays as fallback. */
function waPhone_(raw) {
  var d = String(raw || '').replace(/\D/g, '').replace(/^0+/, ''); if (!d) throw new Error('Phone number is empty.');
  var cc = String(g('DEFAULT_CC', '91')).replace(/\D/g, ''); if (d.length === 10 && cc) d = cc + d;     // 9876543210 -> 919876543210
  if (d.length < 11 || d.length > 15) throw new Error('Phone number looks wrong (' + String(raw).slice(0, 20) + '). Use the international format with country code, e.g. +91 98765 43210.');
  return d;
}
function waSend_(phone, text) {
  var d = waPhone_(phone), keys = jsonProp_('WA_KEYS_JSON', {}), key = keys[d] || g('WA_KEY', '');
  if (!key) throw new Error('No CallMeBot API key for +' + d + '. Add it in Admin > Sender > WhatsApp keys.');
  var r = UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent('+' + d) + '&text=' + encodeURIComponent(text) + '&apikey=' + encodeURIComponent(key), { muteHttpExceptions: true });
  var code = r.getResponseCode(), body = String(r.getContentText() || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200);
  // CallMeBot often answers HTTP 200 with an error sentence in the body, so the body must be checked too
  if (code < 200 || code >= 300 || /invalid|incorrect|not (valid|registered|authori[sz]ed)|error|wrong|blocked|unauthori[sz]ed|fail|too many|limit/i.test(body))
    throw new Error('WhatsApp to +' + d + ' failed (HTTP ' + code + '): ' + (body || 'no reply from CallMeBot') + (keys[d] ? '' : ' (used the shared WA_KEY; this number probably needs its own key)'));
  return { ok: true, detail: body, phone: '+' + d };
}
function waFooter_() {
  var i = senderInfo_(), p = [];
  if (showFlag_('waName')) p.push(i.name); if (i.phone && showFlag_('waPhone')) p.push(i.phone); if (i.email && showFlag_('waEmail')) p.push(i.email);
  return p.length ? '\n\n— ' + p.join(' · ') : '';
}
/* WhatsApp links break above ~2000 characters, so only as many jobs as fit are included. */
function waMessage_(hi, roles, jobs) {
  var head = (hi ? hi + ' ' : '') + 'CareerPulse: ' + jobs.length + ' new job(s) for ' + roles + '.', foot = waFooter_(), msg = head, shown = 0;
  for (var k = 0; k < Math.min(jobs.length, 5); k++) {
    var line = '\n- ' + (jobs[k].t || '') + (jobs[k].c ? ' - ' + jobs[k].c : '') + '\n' + (jobs[k].u || '');
    if (encodeURIComponent(msg + line + foot).length > 1500) break; msg += line; shown++;
  }
  if (jobs.length > shown && shown) msg += '\n+' + (jobs.length - shown) + ' more';
  return msg + foot;
}
function adminWaKeys_(auth) {
  portalRequireAdmin_(auth); var keys = jsonProp_('WA_KEYS_JSON', {});
  return { ok: true, cc: g('DEFAULT_CC', '91'), hasFallback: !!g('WA_KEY', ''), rows: Object.keys(keys).map(function (d) { return { phone: '+' + d, keyMasked: portalMask_(keys[d]) }; }) };
}
function adminWaKeySave_(auth, phone, key) {
  portalRequireAdmin_(auth); var d = waPhone_(phone), k = String(key || '').trim();
  if (!/^[A-Za-z0-9_-]{4,40}$/.test(k)) throw new Error('Enter the API key CallMeBot sent to that number (letters and digits only).');
  var keys = jsonProp_('WA_KEYS_JSON', {}); keys[d] = k;
  if (Utilities.newBlob(JSON.stringify(keys)).getBytes().length > 8500) throw new Error('Too many WhatsApp keys stored. Remove unused numbers first.');
  P.setProperty('WA_KEYS_JSON', JSON.stringify(keys)); return adminWaKeys_(auth);
}
function adminWaKeyDelete_(auth, phone) {
  portalRequireAdmin_(auth); var d = waPhone_(phone), keys = jsonProp_('WA_KEYS_JSON', {}); delete keys[d];
  P.setProperty('WA_KEYS_JSON', JSON.stringify(keys)); return adminWaKeys_(auth);
}
function adminWaTest_(auth, phone) {
  portalRequireAdmin_(auth); var r = waSend_(phone, '✅ CareerPulse test message. WhatsApp delivery works.' + waFooter_());
  return { ok: true, phone: r.phone, reply: r.detail };
}

/* =====================================================================
   v14: AUTO SIGN-OUT AFTER INACTIVITY  +  EMAIL FORMAT FIXES
   ===================================================================== */
function idleMinutes_() { return Math.max(0, Number(g('ADMIN_IDLE_MIN', '0')) || 0); }
function adminIdleSave_(auth, minutes) {
  portalRequireAdmin_(auth); var m = Math.round(Number(minutes));
  if (!isFinite(m) || m < 0 || m > 1440) throw new Error('Enter 0 (off) or a number of minutes up to 1440 (24 hours).');
  if (m) P.setProperty('ADMIN_IDLE_MIN', String(m)); else P.deleteProperty('ADMIN_IDLE_MIN');
  return { ok: true, minutes: m };
}

/* ---- e-mail templates: a Script property holds ~9 KB, but templates may be 25 KB, so they are stored in chunks.
   (Before this, saving any custom format larger than ~9 KB failed with "Argument too large".) ---- */
var TPL_CHUNK_ = 2200;
function mailTemplatePublic_(type) {
  try {
    var n = Number(P.getProperty('MAIL_TEMPLATE_' + type + '_N') || 0);
    if (n > 0) { var s = ''; for (var i = 0; i < n; i++) s += P.getProperty('MAIL_TEMPLATE_' + type + '_' + i) || ''; return JSON.parse(s); }
    return JSON.parse(P.getProperty('MAIL_TEMPLATE_' + type) || 'null');   // formats saved by older versions
  } catch (e) { return null; }
}
function mailTemplateClear_(type) {
  var n = Number(P.getProperty('MAIL_TEMPLATE_' + type + '_N') || 0);
  for (var i = 0; i < Math.max(n, 12); i++) P.deleteProperty('MAIL_TEMPLATE_' + type + '_' + i);
  P.deleteProperty('MAIL_TEMPLATE_' + type + '_N'); P.deleteProperty('MAIL_TEMPLATE_' + type);
}
function mailTemplateSave_(auth, type, subject, html) {
  portalRequireAdmin_(auth);
  if (['jobs', 'family'].indexOf(type) < 0) throw Error('Invalid template type');
  subject = String(subject || '').slice(0, 250); html = String(html || '');
  if (html.length > 25000) throw Error('Template too large (max 25,000 characters).');
  if (html && html.indexOf('{{content}}') < 0) throw Error('Include {{content}} in HTML template');
  if (html && html.indexOf('{{controls}}') < 0) throw Error('Include {{controls}} so recipients can manage reminders');
  mailTemplateClear_(type);
  if (html) {
    var s = JSON.stringify({ subject: subject || '{{subject}}', html: html }), n = Math.ceil(s.length / TPL_CHUNK_);
    for (var i = 0; i < n; i++) P.setProperty('MAIL_TEMPLATE_' + type + '_' + i, s.slice(i * TPL_CHUNK_, (i + 1) * TPL_CHUNK_));
    P.setProperty('MAIL_TEMPLATE_' + type + '_N', String(n));
  }
  return { ok: true };
}
/* Sends the saved format (or the default) to the admin's own address with sample content, so it can be checked end to end. */
function mailTemplateTest_(auth, type) {
  var u = portalRequireAdmin_(auth); type = String(type || '');
  if (['jobs', 'family'].indexOf(type) < 0) throw Error('Invalid template type');
  var bare = !!mailTemplatePublic_(type), jobs = [], titles = ['Senior Software Engineer - Platform and Infrastructure', 'Frontend Developer (React, TypeScript)', 'Quality Assurance Lead - Automation and Performance', 'Backend Engineer, Payments and Risk', 'Product Analyst - Growth and Retention', 'Business Analyst (Banking and Fintech domain)', 'Data Engineer - Streaming Pipelines', 'Customer Success Manager - Enterprise Accounts', 'DevOps Engineer - Kubernetes and Cloud Security', 'Operations Manager, Supply Chain and Logistics'],
    cos = ['Northwind Technologies Private Limited', 'Contoso Digital Solutions', 'Fabrikam Global Services', 'Globex Corporation India', 'Initech Software', 'Umbrella Health Sciences'], locs = ['Gurugram', 'Delhi NCR', 'Noida', 'Bengaluru', 'Hyderabad', 'Pune', 'Remote'];
  for (var k = 0; k < 25; k++) jobs.push({ t: titles[k % titles.length], c: cos[k % cos.length], l: locs[k % locs.length], u: 'https://example.com/job/' + (k + 1) });   // large sample: the most one reminder holds
  var html = reminderHtml_('Hi there,', 'Software Engineer, Frontend Developer, Backend Developer, Full Stack Developer, DevOps Engineer, Data Scientist', locs.join(', '), jobs, [], 7, bare);
  sendMail_(u.email, 'CareerPulse ' + type + ' email format test', html, 'This is a test of your ' + type + ' email format with 25 sample jobs.', type, (type === 'family' ? 'FAM_' : 'JOBS_') + 'FORMATTEST');
  return { ok: true, sentTo: u.email, usedCustomFormat: bare };
}
