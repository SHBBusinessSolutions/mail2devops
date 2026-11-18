// Erzeugt manifest.prod.xml mit der produktiven Host-URL.
// Aufruf: npm run set-url -- https://mail2devops.contoso.com
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const target = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\/[^/]+/.test(target)) {
  console.error('Bitte eine https-URL angeben, z.B.: npm run set-url -- https://mail2devops.contoso.com');
  process.exit(1);
}
const src = fs.readFileSync(path.join(root, 'manifest.xml'), 'utf8');
const out = src.replaceAll('https://localhost:3000', target);
fs.writeFileSync(path.join(root, 'manifest.prod.xml'), out);
console.log(`manifest.prod.xml geschrieben (${(src.match(/https:\/\/localhost:3000/g) || []).length} URLs → ${target})`);
