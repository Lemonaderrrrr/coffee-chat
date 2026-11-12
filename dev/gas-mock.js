// Minimal stand-ins for the Apps Script services Code.gs uses, so it can run under Node
// for tests (dev/test.js) and local preview (dev/server.js). Not used in production.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function tzParts(date, tz) {
  const f = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, hourCycle: 'h23', weekday: 'short',
    year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
  });
  const o = {};
  for (const p of f.formatToParts(date)) o[p.type] = p.value;
  return o;
}

function tzOffsetMin(date, tz) {
  const p = tzParts(date, tz);
  const asUtc = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
  return Math.round((asUtc - Math.floor(date.getTime() / 1000) * 1000) / 60000);
}

const Utilities = {
  getUuid: () => crypto.randomUUID(),
  formatDate(date, tz, pattern) {
    const p = tzParts(date, tz);
    const off = tzOffsetMin(date, tz);
    const sign = off < 0 ? '-' : '+';
    const abs = Math.abs(off);
    const xxx = sign + String(Math.floor(abs / 60)).padStart(2, '0') + ':' + String(abs % 60).padStart(2, '0');
    const map = {
      "yyyy-MM-dd'T'HH:mm:ssXXX": `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}${xxx}`,
      'yyyy-MM-dd': `${p.year}-${p.month}-${p.day}`,
      'EEE': p.weekday,
      'HH:mm': `${p.hour}:${p.minute}`,
    };
    if (!(pattern in map)) throw new Error('mock formatDate: unsupported pattern ' + pattern);
    return map[pattern];
  },
  parseDate(str, tz, pattern) {
    if (pattern !== 'yyyy-MM-dd HH:mm') throw new Error('mock parseDate: unsupported pattern ' + pattern);
    const m = str.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
    const guess = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
    let t = guess - tzOffsetMin(new Date(guess), tz) * 60000;
    t = guess - tzOffsetMin(new Date(t), tz) * 60000;
    return new Date(t);
  },
};

class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => this.sheet._set(this.r + i, this.c + j, v))); return this; }
  setValue(v) { this.sheet._set(this.r, this.c, v); return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet._set(this.r + i, this.c + j, ''); return this; }
  setNumberFormat() { return this; }
  getValues() { return this.sheet.rows.slice(this.r - 1, this.r - 1 + this.nr).map(r => r.slice(this.c - 1, this.c - 1 + this.nc)); }
}

class Sheet {
  constructor(name) { this.name = name; this.rows = []; }
  _set(r, c, v) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() { let n = this.rows.length; while (n > 0 && this.rows[n - 1].every(v => v === '')) n--; return n; }
  getMaxRows() { return 1000; }
  getMaxColumns() { return 26; }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, j) => this._set(r, j + 1, v)); }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  getDataRange() {
    const nr = this.getLastRow();
    const nc = Math.max(0, ...this.rows.slice(0, nr).map(r => r.length));
    return { getValues: () => this.rows.slice(0, nr).map(r => Array.from({ length: nc }, (_, j) => (r[j] === undefined ? '' : r[j]))) };
  }
}

function createGas(opts = {}) {
  const busy = opts.busy || [];
  const sheets = {};
  const props = {};
  const events = [];
  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = new Sheet(n)),
  };
  const busyEvents = () => busy.map(([s, e, allDay]) => ({
    isAllDayEvent: () => !!allDay, getStartTime: () => new Date(s), getEndTime: () => new Date(e),
  }));
  const ctx = {
    console, JSON, Date, Math, String, Number, Object, Array, RegExp, isNaN,
    Utilities,
    Logger: { log: () => {} },
    SpreadsheetApp: { getActive: () => (opts.bound === false ? null : ss), openById: id => { if (id !== 'test-sheet-id') throw new Error('no such sheet ' + id); return ss; } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: s => ({ getContent: () => s, setMimeType() { return this; } }),
    },
    CalendarApp: {
      getDefaultCalendar: () => ({ getEvents: (from, to) => busyEvents().filter(e => e.getStartTime() < to && e.getEndTime() > from) }),
      getCalendarById: () => null,
    },
    Calendar: {
      Events: {
        insert(resource, calId, opts) {
          const ev = { ...resource, id: 'ev_' + events.length, calId, opts };
          if (resource.conferenceData) ev.hangoutLink = 'https://meet.google.com/dev-mock-' + events.length;
          events.push(ev);
          return ev;
        },
        remove(calId, id) { const i = events.findIndex(e => e.id === id); if (i >= 0) events[i].removed = true; },
      },
    },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'backend', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  const call = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());
  return { ctx, call, sheets, props, events, busy };
}

module.exports = { createGas, Utilities };
