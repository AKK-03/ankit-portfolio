/** Separate Apps Script project: public create-only page, authenticated administration. */
function doGet(e) {
  var page = e && e.parameter && e.parameter.page === 'admin' ? 'Admin' : 'Family';
  return HtmlService.createHtmlOutputFromFile(page).setTitle(page === 'Admin' ? 'Reminder administration' : 'Create a job reminder');
}
// Run once from the editor. The underscore prevents browser RPC access.
function setupFamilyStorage_() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FAMILY_SHEET_ID')) {
    var book = SpreadsheetApp.create('Family job reminders');
    book.getSheets()[0].setName('Reminders');
    book.getSheets()[0].appendRow(['ID', 'Created', 'Enabled', 'Settings JSON']);
    props.setProperty('FAMILY_SHEET_ID', book.getId());
  }
}
function requireAdmin_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('FAMILY_ADMIN_TOKEN');
  if (!expected || expected.length < 32 || typeof token !== 'string' || token !== expected) throw new Error('Administrator access required.');
}
function sheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('FAMILY_SHEET_ID');
  if (!id) throw new Error('Reminder storage has not been configured.');
  return SpreadsheetApp.openById(id).getSheetByName('Reminders');
}
function unique_(values, max) {
  if (!Array.isArray(values) || values.length > max) throw new Error('Too many or invalid values.');
  var seen = {};
  return values.map(function(v) { if (typeof v !== 'string' || v.length > 150) throw new Error('Invalid value.'); return v.trim(); }).filter(function(v) {
    var key = v.toLowerCase(); if (!v || seen[key]) return false; seen[key] = true; return true;
  });
}
function cleanReminder_(input) {
  if (!input || typeof input !== 'object') throw new Error('Invalid reminder.');
  var channel = input.channel;
  if (['email', 'whatsapp'].indexOf(channel) < 0) throw new Error('Select Email or Mobile.');
  var address = String(input.address || '').trim();
  if (channel === 'email' ? !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(address) : !/^\+[1-9]\d{7,14}$/.test(address)) throw new Error('Enter a valid email or mobile number with country code.');
  if (address.length > 254) throw new Error('Recipient is too long.');
  var mode = input.mode;
  if (['specific', 'all', 'random'].indexOf(mode) < 0) throw new Error('Invalid job selection.');
  var roles = unique_(input.roles || [], 20), locations = unique_(input.locations || [], 20);
  if (mode === 'specific' && !roles.length) throw new Error('Enter at least one job role.');
  var times = unique_(input.times || [], 12);
  if (!times.length || times.some(function(t) { return !/^([01]\d|2[0-3]):[0-5]\d$/.test(t); })) throw new Error('Choose valid reminder times.');
  var repeat = input.repeat;
  if (['daily', 'weekdays', 'custom'].indexOf(repeat) < 0) throw new Error('Invalid repeat selection.');
  var repeatDays = unique_(input.repeatDays || [], 7);
  if (repeatDays.some(function(v) { return !/^[0-6]$/.test(v); }) || repeat === 'custom' && !repeatDays.length) throw new Error('Select custom repeat days.');
  var timezone = input.timezone;
  if (['Asia/Kolkata', 'UTC', 'America/New_York', 'Europe/London', 'Asia/Dubai', 'Asia/Singapore'].indexOf(timezone) < 0) throw new Error('Invalid time zone.');
  var days = Number(input.days);
  if (!Number.isSafeInteger(days) || days < 1 || days > 3650) throw new Error('Posted within must be 1–3650 days.');
  var workModes = unique_(input.workModes || [], 4), shifts = unique_(input.shifts || [], 3);
  if (workModes.some(function(v) { return ['office','home','hybrid','remote'].indexOf(v) < 0; }) || shifts.some(function(v) { return ['morning','evening','night'].indexOf(v) < 0; })) throw new Error('Invalid work mode or shift.');
  // Whitelist only creation fields. IDs, tokens, action, enabled and existing data are ignored.
  return {channel:channel,address:address,mode:mode,roles:roles,locations:locations,times:times.sort(),repeat:repeat,repeatDays:repeatDays,timezone:timezone,days:days,workModes:workModes,shifts:shifts};
}
function createFamilyReminder(input) {
  var clean = cleanReminder_(input), lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = sheet_();
    if (sheet.getLastRow() > 1000) throw new Error('Reminder limit reached. Contact the administrator.');
    // Reject exact duplicates without returning anyone else's data.
    var json = JSON.stringify(clean), values = sheet.getDataRange().getValues();
    if (values.slice(1).some(function(row) { return row[3] === json; })) throw new Error('This reminder is already registered.');
    var id = Utilities.getUuid();
    sheet.appendRow([id,new Date().toISOString(),true,json]);
    return {ok:true,id:id,message:'Reminder preferences registered. Delivery starts when the administrator connects the reminder sender.'};
  } finally { lock.releaseLock(); }
}
function adminListFamilyReminders(token) {
  requireAdmin_(token);
  return sheet_().getDataRange().getValues().slice(1).filter(function(row) { return row[0]; }).map(function(row) {
    return {id:String(row[0]),created:String(row[1]),enabled:row[2] === true,settings:JSON.parse(row[3])};
  });
}
function adminDeleteFamilyReminder(token, id) {
  requireAdmin_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { var sheet = sheet_(), rows = sheet.getDataRange().getValues();
    for (var i=1;i<rows.length;i++) if (String(rows[i][0]) === id) { sheet.deleteRow(i+1); return {ok:true}; }
    throw new Error('Reminder not found.');
  } finally { lock.releaseLock(); }
}
function adminSetFamilyReminderEnabled(token, id, enabled) {
  requireAdmin_(token);
  if (typeof enabled !== 'boolean') throw new Error('Invalid status.');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { var sheet = sheet_(), rows = sheet.getDataRange().getValues();
    for (var i=1;i<rows.length;i++) if (String(rows[i][0]) === id) { sheet.getRange(i+1,3).setValue(enabled); return {ok:true}; }
    throw new Error('Reminder not found.');
  } finally { lock.releaseLock(); }
}
// Private integration hook for a sender running in this project; not browser callable.
function getFamilyRemindersForDelivery_() {
  return sheet_().getDataRange().getValues().slice(1).filter(function(row) { return row[2] === true; }).map(function(row) { return {id:row[0],settings:JSON.parse(row[3])}; });
}
