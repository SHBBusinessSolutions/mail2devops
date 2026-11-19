// Kopiert die auslieferbaren Dateien nach dist/ (für GitHub Pages oder jedes andere statische Hosting).
// Kein Bundler nötig – das Add-in besteht nur aus statischen Dateien.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dist = path.join(root, 'dist');

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(dist, { recursive: true });
for (const item of ['src', 'assets', 'config.json']) {
  fs.cpSync(path.join(root, item), path.join(dist, item), { recursive: true });
}
// Startseite: leitet auf die Installationsseite weiter (Download-Button für das Manifest)
fs.writeFileSync(
  path.join(dist, 'index.html'),
  '<!doctype html><meta charset="utf-8"><meta http-equiv="refresh" content="0; url=src/install.html"><title>Mail2DevOps</title><a href="src/install.html">Mail2DevOps installieren</a>\n'
);
const count = (d) => fs.readdirSync(d, { withFileTypes: true }).reduce((n, e) => n + (e.isDirectory() ? count(path.join(d, e.name)) : 1), 0);
console.log(`dist/ erstellt (${count(dist)} Dateien)`);
