// Local preview: serves the repo and runs backend/Code.gs on mock Google services.
//   node dev/server.js                 → http://localhost:8787  (admin key: dev-key)
//   SEED=/path/to/seed.json node dev/server.js   → imports that seed instead of sample contacts
// Data lives in memory and resets on restart. Nothing here talks to Google.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas, Utilities } = require('./gas-mock');

const PORT = +process.env.PORT || 8787;
const ROOT = path.join(__dirname, '..');
const TZ = 'America/Los_Angeles';

// One fake busy block next Monday 14:00–15:00 so the "skip my calendar" rule is visible.
const now = new Date();
const nextMon = [...Array(10).keys()].map(i => new Date(now.getTime() + (i + 2) * 864e5))
  .find(d => Utilities.formatDate(d, TZ, 'EEE') === 'Mon');
const ymd = Utilities.formatDate(nextMon, TZ, 'yyyy-MM-dd');
const g = createGas({ busy: [[Utilities.parseDate(`${ymd} 14:00`, TZ, 'yyyy-MM-dd HH:mm'), Utilities.parseDate(`${ymd} 15:00`, TZ, 'yyyy-MM-dd HH:mm')]] });
g.props.ADMIN_KEY = 'dev-key';
g.ctx.setup();

const seed = process.env.SEED && fs.existsSync(process.env.SEED)
  ? JSON.parse(fs.readFileSync(process.env.SEED, 'utf8'))
  : [
    { name: 'Sample Professor A', salutation: 'Professor A', unit: 'ICS', tier: 'A · 第 1 波', email: 'a@example.edu', office: 'DBH 0000', stage: '待发' },
    { name: 'Sample Professor B', salutation: 'Professor B', unit: 'Math', tier: 'B · 第 2 波', email: 'b@example.edu', modes: 'office,meet' },
  ];
console.log('import:', g.call({ action: 'admin.import', key: 'dev-key', contacts: seed }));

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png' };
const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type' };

http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  if (url.pathname === '/api') {
    if (req.method === 'OPTIONS') { res.writeHead(204, cors); return res.end(); }
    if (req.method === 'GET') {
      const out = g.ctx.doGet({ parameter: Object.fromEntries(url.searchParams) }).getContent();
      res.writeHead(200, { ...cors, 'Content-Type': 'application/json' }); return res.end(out);
    }
    let body = '';
    req.on('data', c => { body += c; });
    req.on('end', () => {
      const out = g.ctx.doPost({ postData: { contents: body } }).getContent();
      res.writeHead(200, { ...cors, 'Content-Type': 'application/json' });
      res.end(out);
    });
    return;
  }
  let p = path.join(ROOT, decodeURIComponent(url.pathname));
  if (!p.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
  if (fs.existsSync(p) && fs.statSync(p).isDirectory()) p = path.join(p, 'index.html');
  if (!fs.existsSync(p)) { res.writeHead(404); return res.end('not found'); }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(PORT, () => {
  const api = `http://localhost:${PORT}/api`;
  console.log(`\nadmin:  http://localhost:${PORT}/admin/?api=${encodeURIComponent(api)}   key: dev-key`);
  const list = g.call({ action: 'admin.list', key: 'dev-key' });
  list.contacts.slice(0, 3).forEach(c => console.log(`book:   http://localhost:${PORT}/book/?c=${c.token}&api=${encodeURIComponent(api)}   (${c.name})`));
});
