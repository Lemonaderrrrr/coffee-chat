// node dev/test.js — runs backend/Code.gs against dev/gas-mock.js
const assert = require('assert');
const { createGas, Utilities } = require('./gas-mock');

const TZ = 'America/Los_Angeles';
const local = (d, hm) => Utilities.parseDate(`${d} ${hm}`, TZ, 'yyyy-MM-dd HH:mm');
let passed = 0;
const t = (name, fn) => { fn(); passed++; console.log('  ✓ ' + name); };

// A busy class on the first Monday after next week, 13:00–14:00 local.
const now = new Date();
const days = [...Array(21).keys()].map(i => Utilities.formatDate(new Date(now.getTime() + (i + 2) * 864e5), TZ, 'yyyy-MM-dd'));
const monday = days.find(d => Utilities.formatDate(local(d, '12:00'), TZ, 'EEE') === 'Mon');
const g = createGas({ busy: [[local(monday, '13:00'), local(monday, '14:00')], [local(monday, '00:00'), local(days[20], '00:00'), true]] });
const { ctx, call } = g;

console.log('backend/Code.gs');
ctx.setup();
const KEY = g.props.ADMIN_KEY;

t('setup creates five sheets and an admin key', () => {
  assert.deepStrictEqual(Object.keys(g.sheets).sort(), ['availability', 'bookings', 'contacts', 'log', 'settings']);
  assert.ok(KEY && KEY.length >= 24);
});

t('admin calls need the key', () => {
  assert.strictEqual(call({ action: 'admin.list' }).error, 'unauthorized');
  assert.strictEqual(call({ action: 'admin.list', key: 'wrong' }).error, 'unauthorized');
});

const seed = [
  { name: 'Test Prof A', salutation: 'Professor A', unit: 'ICS', email: 'a@example.edu', office: 'DBH 1000', stage: '待发' },
  { name: 'Test Prof B', salutation: 'Professor B', unit: 'Math', email: 'b@example.edu', modes: 'office,meet' },
];
t('import adds new contacts and skips repeats', () => {
  assert.deepStrictEqual(call({ action: 'admin.import', key: KEY, contacts: seed }), { ok: true, added: 2, skipped: 0 });
  assert.deepStrictEqual(call({ action: 'admin.import', key: KEY, contacts: seed }), { ok: true, added: 0, skipped: 2 });
});

const list = call({ action: 'admin.list', key: KEY });
const [A, B] = list.contacts;
t('each contact gets a unique token', () => {
  assert.ok(A.token.length >= 20 && B.token.length >= 20 && A.token !== B.token);
  assert.strictEqual(A.stage, '待发');
  assert.strictEqual(B.stage, '待研究');
});

t('a bad link is refused', () => {
  assert.strictEqual(call({ action: 'profile', token: 'nope' }).error, 'bad_link');
  assert.strictEqual(call({ action: 'slots', token: 'x'.repeat(20) }).error, 'bad_link');
});

t('profile shows only what the visitor needs', () => {
  const p = call({ action: 'profile', token: B.token });
  assert.deepStrictEqual(p.profile.modes.map(m => m.id), ['office', 'meet']);
  assert.strictEqual(p.profile.salutation, 'Professor B');
  assert.ok(!('email' in p.profile) && !('notes' in p.profile));
});

const slots = call({ action: 'slots', token: A.token }).slots;
t('slots sit inside Mon/Wed/Fri 13:00–17:00, on a 30-minute grid, 24 h ahead', () => {
  assert.ok(slots.length > 10);
  for (const s of slots) {
    const d = new Date(s);
    assert.ok(['Mon', 'Wed', 'Fri'].includes(Utilities.formatDate(d, TZ, 'EEE')), s);
    const hm = Utilities.formatDate(d, TZ, 'HH:mm');
    assert.ok(hm >= '13:00' && hm <= '16:30' && /:(00|30)$/.test(hm), s);
    assert.ok(d.getTime() >= now.getTime() + 24 * 3600e3 - 60e3, s);
  }
});

t('busy calendar time is skipped; all-day events are ignored', () => {
  const mon = slots.filter(s => s.startsWith(monday)).map(s => Utilities.formatDate(new Date(s), TZ, 'HH:mm'));
  assert.ok(!mon.includes('13:00') && !mon.includes('13:30'), mon.join(','));
  assert.ok(mon.includes('14:00') && mon.includes('16:30'), mon.join(','));
});

const pick = slots.find(s => s.startsWith(monday) && s.includes('T14:00'));
let booking;
t('booking with Google Meet sends an invite with a Meet link and marks the contact 已约', () => {
  const r = call({ action: 'book', token: A.token, start: pick, mode: 'meet', note: 'Looking forward' });
  assert.ok(r.ok, JSON.stringify(r));
  booking = r.booking;
  assert.ok(booking.meetLink.startsWith('https://meet.google.com/'));
  const ev = g.events[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(ev.attendees)), [{ email: 'a@example.edu' }]);
  assert.strictEqual(ev.opts.sendUpdates, 'all');
  const c = call({ action: 'admin.list', key: KEY }).contacts.find(x => x.id === A.id);
  assert.strictEqual(c.stage, '已约');
  assert.strictEqual(c.next_followup, monday);
});

t('a booked time disappears for everyone else', () => {
  assert.ok(!call({ action: 'slots', token: B.token }).slots.includes(pick));
  assert.strictEqual(call({ action: 'book', token: B.token, start: pick, mode: 'office' }).error, 'slot_taken');
});

t('the same contact cannot hold two meetings', () => {
  const other = slots.find(s => s !== pick && !s.startsWith(monday));
  assert.strictEqual(call({ action: 'book', token: A.token, start: other, mode: 'office' }).error, 'already_booked');
});

t('revisiting the link shows the existing booking', () => {
  assert.strictEqual(call({ action: 'profile', token: A.token }).booking.start, pick);
});

t('cancelling frees the time and moves the contact back to 已回', () => {
  assert.ok(call({ action: 'cancel', token: A.token, id: booking.id }).ok);
  assert.ok(g.events[0].removed);
  assert.ok(call({ action: 'slots', token: B.token }).slots.includes(pick));
  assert.strictEqual(call({ action: 'admin.list', key: KEY }).contacts.find(x => x.id === A.id).stage, '已回');
});

t('office bookings use the contact\'s room; unknown modes are refused', () => {
  assert.strictEqual(call({ action: 'book', token: A.token, start: pick, mode: 'teleport' }).error, 'bad_mode');
  const r = call({ action: 'book', token: A.token, start: pick, mode: 'office' });
  assert.strictEqual(r.booking.place, 'DBH 1000');
  assert.strictEqual(r.booking.meetLink, '');
});

t('admin.save validates stages and logs stage changes', () => {
  assert.strictEqual(call({ action: 'admin.save', key: KEY, contact: { id: B.id, stage: 'nope' } }).error, 'bad_stage');
  assert.ok(call({ action: 'admin.save', key: KEY, contact: { id: B.id, stage: '已发', next_followup: '2026-12-01' } }).ok);
  const log = call({ action: 'admin.list', key: KEY }).log;
  assert.ok(log.some(l => l.contact_id === B.id && l.detail === '待研究 → 已发'));
});

t('availability can be replaced from the admin page', () => {
  const r = call({ action: 'admin.availability', key: KEY, rows: [{ weekday: 'Tue', start: '9:30', end: '11:00' }, { weekday: 'Bad', start: '1', end: '2' }] });
  assert.deepStrictEqual(r.availability, [{ weekday: 'Tue', start: '09:30', end: '11:00' }]);
  const s2 = call({ action: 'slots', token: B.token }).slots;
  assert.ok(s2.length > 0 && s2.every(s => Utilities.formatDate(new Date(s), TZ, 'EEE') === 'Tue'));
});

t('slots keep the right UTC offset across the March 2027 daylight-saving change', () => {
  call({ action: 'admin.availability', key: KEY, rows: [{ weekday: 'Mon', start: '13:00', end: '14:00' }] });
  const before = ctx.freeSlots_(new Date('2027-03-03T12:00:00-08:00'));
  const after = ctx.freeSlots_(new Date('2027-03-10T12:00:00-08:00'));
  assert.ok(before.includes('2027-03-08T13:00:00-08:00'), before.join(','));
  assert.ok(after.includes('2027-03-15T13:00:00-07:00'), after.join(','));
});

t('an unbound script reads the sheet id from Script Properties', () => {
  const u = createGas({ bound: false });
  assert.throws(() => u.ctx.setup(), /no_sheet/);
  u.props.SHEET_ID = 'test-sheet-id';
  u.ctx.setup();
  assert.ok(u.call({ action: 'admin.list', key: u.props.ADMIN_KEY }).ok);
});

console.log(`\n${passed} passed`);
