// Minimaler statischer HTTPS-Server für die Entwicklung (keine Abhängigkeiten).
// Nutzt die Zertifikate von `npm run certs` (office-addin-dev-certs) aus ~/.office-addin-dev-certs.
// Ohne Zertifikate: Fallback auf HTTP (reicht für den Demo-Modus im Browser, NICHT für Outlook).
import https from 'node:https';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const port = Number(process.env.PORT || 3000);
const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon', '.xml': 'application/xml; charset=utf-8', '.md': 'text/markdown; charset=utf-8',
};

function handler(req, res) {
  const url = new URL(req.url, 'http://x');
  let p = decodeURIComponent(url.pathname);
  if (p === '/') p = '/src/taskpane.html';
  const file = path.normalize(path.join(root, p));
  if (!file.startsWith(root) || file.includes(`${path.sep}node_modules${path.sep}`) || path.basename(file).startsWith('.')) {
    res.writeHead(403).end('Forbidden'); return;
  }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404).end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(data);
  });
}

const key = path.join(certDir, 'localhost.key');
const crt = path.join(certDir, 'localhost.crt');
if (fs.existsSync(key) && fs.existsSync(crt)) {
  https.createServer({ key: fs.readFileSync(key), cert: fs.readFileSync(crt) }, handler).listen(port, () =>
    console.log(`Mail2DevOps läuft auf https://localhost:${port}/src/taskpane.html`));
} else {
  http.createServer(handler).listen(port, () => {
    console.log(`⚠ Keine Dev-Zertifikate gefunden (npm run certs). HTTP-Fallback – nur für den Demo-Modus im Browser:`);
    console.log(`  http://localhost:${port}/src/taskpane.html`);
  });
}
