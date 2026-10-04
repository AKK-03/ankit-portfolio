/* JOB RADAR REMINDERS — Google Apps Script
 SETUP: 1) script.google.com > New project > paste this file.
 2) Project Settings > Script properties, add: TOKEN (your secret), ADZUNA_ID, ADZUNA_KEY (free at developer.adzuna.com).
    Optional WhatsApp via CallMeBot: WA_KEY (message the CallMeBot bot once to get it).
 3) Run setup() once and allow permissions.
 4) Deploy > New deployment > Web app > Execute as: Me, Access: Anyone. Paste the /exec URL + TOKEN into jobs.html.
 Keywords, locations, job types, email, WhatsApp number and send time are all set from jobs.html. After editing this file, use Deploy > Manage deployments > Edit > New version. */
var P = PropertiesService.getScriptProperties();
function g(k, d) { return P.getProperty(k) || d; }
function setup() { setOn(true); }
function enable() { setOn(true); }
function disable() { setOn(false); }
function setOn(on) {
  P.setProperty('ENABLED', on ? '1' : '0');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'digest') ScriptApp.deleteTrigger(t); });
  if (on) ScriptApp.newTrigger('digest').timeBased().everyDays(1).atHour(+g('HOUR', '9')).create();
}
function fetchJobs() {
  var out = [], seen = JSON.parse(g('SEEN', '[]')), jt = g('JT', '').split('|'), flags = '';
  if (jt.indexOf('full') > -1) flags += '&full_time=1';
  if (jt.indexOf('part') > -1) flags += '&part_time=1';
  if (jt.indexOf('contract') > -1) flags += '&contract=1';
  var locs = g('LOC', '') ? g('LOC').split('|') : [''];
  g('KW', '').split('|').filter(String).forEach(function (k) {
    locs.forEach(function (l) {
      var url = 'https://api.adzuna.com/v1/api/jobs/in/search/1?app_id=' + g('ADZUNA_ID') + '&app_key=' + g('ADZUNA_KEY') +
        '&results_per_page=15&max_days_old=' + g('DAYS', '2') + '&sort_by=date' + flags +
        '&what=' + encodeURIComponent(k + (g('REMOTE') ? ' remote' : '')) + (l ? '&where=' + encodeURIComponent(l) : '');
      try {
        var r = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
        (r.results || []).forEach(function (j) {
          if (seen.indexOf(j.id) < 0 && !out.some(function (o) { return o.id === j.id; }))
            out.push({ id: j.id, t: j.title.replace(/<[^>]+>/g, ''), c: (j.company || {}).display_name || '', l: (j.location || {}).display_name || '', u: j.redirect_url });
        });
      } catch (e) {}
    });
  });
  return { jobs: out, seen: seen };
}
function digest(force) {
  if (force !== true && g('ENABLED') !== '1') return 0;
  var to = g('EMAIL') || Session.getEffectiveUser().getEmail(), r = fetchJobs(), jobs = r.jobs;
  P.setProperty('LAST', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM HH:mm'));
  if (!jobs.length) { if (force === true) MailApp.sendEmail(to, 'Job radar: no new openings', 'Nothing new for your current search settings.'); return 0; }
  var html = '<h3>' + jobs.length + ' new openings</h3>' + jobs.map(function (j) {
    return '<p><a href="' + j.u + '"><b>' + j.t + '</b></a><br>' + j.c + ' · ' + j.l + '</p>';
  }).join('');
  MailApp.sendEmail({ to: to, subject: 'Job radar: ' + jobs.length + ' new openings', htmlBody: html });
  if (g('WA') && g('WA_KEY')) {
    var text = jobs.length + ' new openings:\n' + jobs.slice(0, 5).map(function (j) { return '• ' + j.t + ' — ' + j.c; }).join('\n') + '\nFull list is in your email.';
    try { UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(g('WA')) + '&text=' + encodeURIComponent(text) + '&apikey=' + g('WA_KEY'), { muteHttpExceptions: true }); } catch (e) {}
  }
  P.setProperty('SEEN', JSON.stringify(r.seen.concat(jobs.map(function (j) { return j.id; })).slice(-500)));
  return jobs.length;
}
function doGet(e) {
  var p = e.parameter, res;
  if (!g('TOKEN') || p.token !== g('TOKEN')) res = { error: 'Wrong token' };
  else {
    var sent, map = { kw: 'KW', loc: 'LOC', jt: 'JT', days: 'DAYS', hour: 'HOUR', email: 'EMAIL', wa: 'WA', remote: 'REMOTE' };
    if (p.action === 'on') setOn(true);
    else if (p.action === 'off') setOn(false);
    else if (p.action === 'save') {
      Object.keys(map).forEach(function (k) { if (p[k] !== undefined) P.setProperty(map[k], p[k]); });
      if (g('ENABLED') === '1') setOn(true);
    } else if (p.action === 'test') sent = digest(true);
    res = { enabled: g('ENABLED') === '1', last: g('LAST', ''), sent: sent, kw: g('KW', ''), loc: g('LOC', ''), jt: g('JT', ''), days: g('DAYS', '1'), hour: g('HOUR', '9'), email: g('EMAIL', ''), wa: g('WA', ''), remote: g('REMOTE', '') };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
