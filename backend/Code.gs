/**
 * Outreach backend: one Apps Script for the booking page (book/) and the tracker (admin/).
 *
 * Either bind this script to a Google Sheet, or set the Script Property SHEET_ID to the
 * spreadsheet's id. Then run setup() once and deploy as a web app. See backend/SETUP.md.
 *
 * Sheets
 *   contacts      one row per person or opportunity
 *   availability  weekly windows when Leon can meet (weekday, start, end)
 *   settings      key/value
 *   bookings      one row per booked meeting
 *   log           what happened, newest last
 *
 * Public actions (no key):  ping, profile, slots, book, cancel
 * Admin actions (key):      admin.list, admin.save, admin.import, admin.availability, admin.settings
 */

const TZ = 'America/Los_Angeles';

const SHEETS = {
  contacts: ['id', 'token', 'name', 'salutation', 'unit', 'category', 'tier', 'email', 'office', 'modes',
    'stage', 'last_contact', 'next_followup', 'notes', 'source', 'created_at', 'updated_at'],
  availability: ['weekday', 'start', 'end'],
  settings: ['key', 'value'],
  bookings: ['id', 'token', 'contact_id', 'start', 'end', 'mode', 'place', 'meet_link', 'event_id', 'status', 'note', 'created_at'],
  log: ['time', 'contact_id', 'action', 'detail'],
};

const DAY_FIELDS = ['last_contact', 'next_followup'];

const STAGES = ['待研究', '待发', '已发', '已跟进', '已回', '已约', '已聊', '进组', '暂缓', '婉拒', '无回复结案'];

const DEFAULT_SETTINGS = {
  host_name: 'Leon',
  meeting_minutes: '20',
  slot_step_minutes: '30',
  horizon_days: '21',
  min_notice_hours: '24',
  calendar_id: 'primary',
  coffee_places: '',            // one per line, e.g. "Starbucks, Student Center"
  default_modes: 'office,meet,coffee',
};

const DEFAULT_AVAILABILITY = [
  ['Mon', '13:00', '17:00'],
  ['Wed', '13:00', '17:00'],
  ['Fri', '13:00', '17:00'],
];

const MODES = {
  office: 'Your office',
  meet: 'Google Meet',
  coffee: 'Coffee on campus',
};

/* ───────── entry points ───────── */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ping';
  return respond_(route_(action, e.parameter || {}));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {
    return respond_({ ok: false, error: 'bad_json' });
  }
  return respond_(route_(body.action, body));
}

function route_(action, p) {
  try {
    switch (action) {
      case 'ping': return { ok: true, now: iso_(new Date()) };
      case 'profile': return profile_(p.token);
      case 'slots': return slots_(p.token);
      case 'book': return book_(p);
      case 'cancel': return cancel_(p);
      case 'admin.list': requireKey_(p.key); return adminList_();
      case 'admin.save': requireKey_(p.key); return adminSave_(p.contact || {});
      case 'admin.import': requireKey_(p.key); return adminImport_(p.contacts || []);
      case 'admin.availability': requireKey_(p.key); return adminAvailability_(p.rows || []);
      case 'admin.settings': requireKey_(p.key); return adminSettings_(p.settings || {});
      default: return { ok: false, error: 'unknown_action' };
    }
  } catch (err) {
    const msg = String(err && err.message || err);
    if (msg.indexOf('E:') === 0) return { ok: false, error: msg.slice(2) };
    console.error(err);
    return { ok: false, error: 'server_error', detail: msg };
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor. Creates sheets, defaults and an admin key. */
function setup() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    // Plain text everywhere, so Sheets never turns dates, times or tokens into other types.
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    if (sh.getLastRow() === 0) sh.appendRow(SHEETS[name]);
  });
  if (table_('availability').length === 0) DEFAULT_AVAILABILITY.forEach(function (r) { sheet_('availability').appendRow(r); });
  const have = settings_();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    if (!(k in have)) sheet_('settings').appendRow([k, DEFAULT_SETTINGS[k]]);
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('Admin key: ' + props.getProperty('ADMIN_KEY'));
  return { ok: true };
}

/* ───────── public ───────── */

function profile_(token) {
  const c = contactByToken_(token);
  const s = settings_();
  const modes = String(c.modes || s.default_modes).split(',').map(trim_).filter(function (m) { return MODES[m]; });
  const now = new Date();
  const upcoming = table_('bookings').filter(function (b) {
    return b.token === c.token && b.status === 'booked' && new Date(b.end) > now;
  })[0] || null;
  return {
    ok: true,
    profile: {
      salutation: c.salutation || c.name,
      host: s.host_name,
      minutes: Number(s.meeting_minutes),
      modes: modes.map(function (m) { return { id: m, label: MODES[m] }; }),
      office: c.office || '',
      coffeePlaces: String(s.coffee_places || '').split('\n').map(trim_).filter(String),
    },
    booking: upcoming && publicBooking_(upcoming),
  };
}

function slots_(token) {
  contactByToken_(token);
  return { ok: true, tz: TZ, slots: freeSlots_(new Date()) };
}

function book_(p) {
  const c = contactByToken_(p.token);
  const s = settings_();
  if (!MODES[p.mode]) throw new Error('E:bad_mode');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = new Date();
    const already = table_('bookings').some(function (b) { return b.token === c.token && b.status === 'booked' && new Date(b.end) > now; });
    if (already) throw new Error('E:already_booked');
    const start = new Date(p.start);
    if (isNaN(start)) throw new Error('E:bad_time');
    const free = freeSlots_(new Date());
    if (free.indexOf(iso_(start)) < 0) throw new Error('E:slot_taken');
    const end = new Date(start.getTime() + Number(s.meeting_minutes) * 60000);
    const place = p.mode === 'office' ? (c.office || 'Your office')
      : p.mode === 'coffee' ? (String(p.place || '').slice(0, 120) || 'Coffee on campus (Leon will suggest a spot)')
        : 'Google Meet';
    const ev = createEvent_(c, s, start, end, p.mode, place, String(p.note || '').slice(0, 500));
    const b = {
      id: 'b_' + Utilities.getUuid().slice(0, 8), token: c.token, contact_id: c.id,
      start: iso_(start), end: iso_(end), mode: p.mode, place: place, meet_link: ev.meetLink || '',
      event_id: ev.id || '', status: 'booked', note: String(p.note || '').slice(0, 500), created_at: iso_(new Date()),
    };
    appendObj_('bookings', b);
    updateObj_('contacts', c.id, { stage: '已约', last_contact: day_(new Date()), next_followup: day_(start), updated_at: iso_(new Date()) });
    log_(c.id, 'book', b.start + ' ' + b.mode + ' ' + place);
    return { ok: true, booking: publicBooking_(b) };
  } finally {
    lock.releaseLock();
  }
}

function cancel_(p) {
  const c = contactByToken_(p.token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = table_('bookings').filter(function (x) { return x.id === p.id && x.token === c.token && x.status === 'booked'; })[0];
    if (!b) throw new Error('E:not_found');
    deleteEvent_(settings_(), b.event_id);
    updateObj_('bookings', b.id, { status: 'cancelled' });
    updateObj_('contacts', c.id, { stage: '已回', updated_at: iso_(new Date()) });
    log_(c.id, 'cancel', b.start);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function publicBooking_(b) {
  return { id: b.id, start: b.start, end: b.end, mode: b.mode, modeLabel: MODES[b.mode], place: b.place, meetLink: b.meet_link };
}

/* ───────── slots ───────── */

/** Every free start time (ISO, LA offset) inside the weekly windows, after the notice period. */
function freeSlots_(now) {
  const s = settings_();
  const meeting = Number(s.meeting_minutes) * 60000;
  const step = Number(s.slot_step_minutes) * 60000;
  const earliest = now.getTime() + Number(s.min_notice_hours) * 3600000;
  const windows = table_('availability');
  const horizonEnd = new Date(now.getTime() + Number(s.horizon_days) * 86400000);

  const blocked = busy_(s, now, horizonEnd).concat(
    table_('bookings').filter(function (b) { return b.status === 'booked'; })
      .map(function (b) { return [new Date(b.start).getTime(), new Date(b.end).getTime()]; }));

  const out = [];
  // Step from local noon so daylight-saving changes never skip or repeat a date.
  const noon = Utilities.parseDate(Utilities.formatDate(now, TZ, 'yyyy-MM-dd') + ' 12:00', TZ, 'yyyy-MM-dd HH:mm').getTime();
  for (let d = 0; d <= Number(s.horizon_days); d++) {
    const date = new Date(noon + d * 86400000);
    const ymd = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
    const wd = Utilities.formatDate(date, TZ, 'EEE');
    windows.filter(function (w) { return String(w.weekday).slice(0, 3) === wd; }).forEach(function (w) {
      const ws = Utilities.parseDate(ymd + ' ' + hhmm_(w.start), TZ, 'yyyy-MM-dd HH:mm').getTime();
      const we = Utilities.parseDate(ymd + ' ' + hhmm_(w.end), TZ, 'yyyy-MM-dd HH:mm').getTime();
      for (let t = ws; t + meeting <= we; t += step) {
        if (t < earliest) continue;
        const clash = blocked.some(function (x) { return t < x[1] && t + meeting > x[0]; });
        if (!clash) out.push(iso_(new Date(t)));
      }
    });
  }
  return out.filter(function (v, i, a) { return a.indexOf(v) === i; }).sort();
}

/** Busy intervals from Leon's calendar. All-day events are ignored. */
function busy_(s, from, to) {
  const cal = s.calendar_id && s.calendar_id !== 'primary'
    ? CalendarApp.getCalendarById(s.calendar_id) : CalendarApp.getDefaultCalendar();
  if (!cal) return [];
  return cal.getEvents(from, to)
    .filter(function (e) { return !e.isAllDayEvent(); })
    .map(function (e) { return [e.getStartTime().getTime(), e.getEndTime().getTime()]; });
}

/* ───────── calendar ───────── */

function createEvent_(c, s, start, end, mode, place, note) {
  const title = s.host_name + ' / ' + (c.salutation || c.name) + ' · ' + MODES[mode];
  const desc = [
    mode === 'meet' ? 'Google Meet link is attached to this invite.' : 'Place: ' + place,
    note ? 'Note from ' + (c.salutation || c.name) + ': ' + note : '',
    'Need to change the time? Use the same link you booked with.',
  ].filter(String).join('\n');
  const calId = s.calendar_id || 'primary';
  // Calendar advanced service: sends invites and can attach a Meet link.
  if (typeof Calendar !== 'undefined' && Calendar.Events) {
    const resource = {
      summary: title, description: desc, location: mode === 'meet' ? '' : place,
      start: { dateTime: iso_(start), timeZone: TZ }, end: { dateTime: iso_(end), timeZone: TZ },
      attendees: c.email ? [{ email: c.email }] : [],
    };
    if (mode === 'meet') resource.conferenceData = { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
    const ev = Calendar.Events.insert(resource, calId, { conferenceDataVersion: 1, sendUpdates: 'all' });
    return { id: ev.id, meetLink: ev.hangoutLink || '' };
  }
  // Fallback without the advanced service: no Meet link, still sends the invite.
  const cal = calId === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(calId);
  const ev = cal.createEvent(title, start, end, { description: desc, location: place, guests: c.email || '', sendInvites: true });
  return { id: ev.getId(), meetLink: '' };
}

function deleteEvent_(s, eventId) {
  if (!eventId) return;
  const calId = s.calendar_id || 'primary';
  try {
    if (typeof Calendar !== 'undefined' && Calendar.Events) {
      Calendar.Events.remove(calId, eventId, { sendUpdates: 'all' });
    } else {
      const cal = calId === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(calId);
      const ev = cal.getEventById(eventId);
      if (ev) ev.deleteEvent();
    }
  } catch (err) { console.warn('delete event failed', err); }
}

/* ───────── admin ───────── */

function requireKey_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real || String(key || '') !== real) throw new Error('E:unauthorized');
}

function adminList_() {
  return {
    ok: true, stages: STAGES, modes: MODES, tz: TZ,
    contacts: table_('contacts'), bookings: table_('bookings'),
    availability: table_('availability'), settings: settings_(),
    log: table_('log').slice(-200),
  };
}

function adminSave_(input) {
  const now = iso_(new Date());
  const fields = SHEETS.contacts.filter(function (k) { return ['id', 'token', 'created_at', 'updated_at'].indexOf(k) < 0; });
  const patch = {};
  fields.forEach(function (k) { if (k in input) patch[k] = input[k] == null ? '' : String(input[k]); });
  if (patch.stage && STAGES.indexOf(patch.stage) < 0) throw new Error('E:bad_stage');
  if (input.id) {
    const before = table_('contacts').filter(function (x) { return x.id === input.id; })[0];
    if (!before) throw new Error('E:not_found');
    patch.updated_at = now;
    updateObj_('contacts', input.id, patch);
    if (patch.stage && patch.stage !== before.stage) log_(input.id, 'stage', before.stage + ' → ' + patch.stage);
    return { ok: true, contact: Object.assign({}, before, patch) };
  }
  if (!patch.name) throw new Error('E:name_required');
  const c = Object.assign({}, blank_('contacts'), patch, {
    id: 'c_' + Utilities.getUuid().slice(0, 8),
    token: Utilities.getUuid().replace(/-/g, '').slice(0, 20),
    stage: patch.stage || '待研究', created_at: now, updated_at: now,
  });
  appendObj_('contacts', c);
  log_(c.id, 'create', c.name);
  return { ok: true, contact: c };
}

function adminImport_(list) {
  const have = table_('contacts').map(function (c) { return (c.name + '|' + c.unit).toLowerCase(); });
  let added = 0, skipped = 0;
  list.forEach(function (x) {
    const k = (String(x.name || '') + '|' + String(x.unit || '')).toLowerCase();
    if (!x.name || have.indexOf(k) >= 0) { skipped++; return; }
    adminSave_(x); have.push(k); added++;
  });
  return { ok: true, added: added, skipped: skipped };
}

function adminAvailability_(rows) {
  const clean = rows.map(function (r) { return [String(r.weekday).slice(0, 3), hhmm_(r.start), hhmm_(r.end)]; })
    .filter(function (r) { return /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)$/.test(r[0]) && r[1] < r[2]; });
  const sh = sheet_('availability');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (clean.length) sh.getRange(2, 1, clean.length, 3).setValues(clean);
  log_('', 'availability', JSON.stringify(clean));
  return { ok: true, availability: table_('availability') };
}

function adminSettings_(patch) {
  const sh = sheet_('settings');
  const rows = sh.getDataRange().getValues();
  Object.keys(patch).forEach(function (k) {
    if (!(k in DEFAULT_SETTINGS)) return;
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === k) { sh.getRange(i + 1, 2).setValue(String(patch[k])); found = true; }
    }
    if (!found) sh.appendRow([k, String(patch[k])]);
  });
  return { ok: true, settings: settings_() };
}

/* ───────── sheet helpers ───────── */

/** The data spreadsheet: Script Property SHEET_ID if set, otherwise the sheet this script is bound to. */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (!active) throw new Error('E:no_sheet (set Script Property SHEET_ID to the spreadsheet id)');
  return active;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('E:run_setup_first');
  return sh;
}

function table_(name) {
  const values = sheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  return values.slice(1).filter(function (r) { return r.some(function (v) { return v !== ''; }); }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) {
      const v = r[i];
      o[h] = v instanceof Date ? (DAY_FIELDS.indexOf(h) >= 0 ? day_(v) : iso_(v)) : String(v);
    });
    return o;
  });
}

function blank_(name) {
  const o = {};
  SHEETS[name].forEach(function (k) { o[k] = ''; });
  return o;
}

function appendObj_(name, obj) {
  sheet_(name).appendRow(SHEETS[name].map(function (k) { return obj[k] == null ? '' : obj[k]; }));
}

function updateObj_(name, id, patch) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values[0];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      Object.keys(patch).forEach(function (k) {
        const col = head.indexOf(k);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(patch[k]);
      });
      return;
    }
  }
  throw new Error('E:not_found');
}

function settings_() {
  const sh = ss_().getSheetByName('settings');
  const out = Object.assign({}, DEFAULT_SETTINGS);
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) out[r[0]] = String(r[1]); });
  return out;
}

function contactByToken_(token) {
  if (!token || String(token).length < 12) throw new Error('E:bad_link');
  const c = table_('contacts').filter(function (x) { return x.token === token; })[0];
  if (!c) throw new Error('E:bad_link');
  return c;
}

function log_(contactId, action, detail) {
  try { appendObj_('log', { time: iso_(new Date()), contact_id: contactId, action: action, detail: detail }); } catch (err) { }
}

/* ───────── small utils ───────── */

function iso_(d) { return Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function day_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function trim_(s) { return String(s).trim(); }
function hhmm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = String(v).match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(v);
}
