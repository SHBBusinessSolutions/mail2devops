// Prüft, ob ein Hosting-Ziel für ein Outlook Add-in taugt.
// Aufruf: npm run check-host -- https://<hosting-url>
// Geprüft: anonym erreichbar (kein Login-Redirect), HTTPS, richtige Content-Types,
// und ob der Task Pane in Outlook eingebettet werden darf (X-Frame-Options / CSP frame-ancestors).
const base = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\//.test(base)) {
  console.error('Bitte die Hosting-URL angeben, z.B.: npm run check-host -- https://xyz.example.com');
  process.exit(1);
}

const files = [
  ['src/taskpane.html', 'text/html', true],
  ['src/commands.html', 'text/html', true],
  ['src/app.js', 'javascript', false],
  ['src/taskpane.css', 'text/css', false],
  ['config.json', 'json', false],
  ['assets/icon-80.png', 'image/png', false],
];

let problems = 0;
const bad = (msg) => { problems++; console.log('   ✗ ' + msg); };

for (const [p, type, framed] of files) {
  const url = `${base}/${p}`;
  let res;
  try {
    res = await fetch(url, { redirect: 'manual' });
  } catch (e) {
    console.log(`✗ ${p}`); bad(`nicht erreichbar: ${e.message}`); continue;
  }
  const ct = res.headers.get('content-type') || '';
  console.log(`${res.status === 200 ? '✓' : '✗'} ${p}  (${res.status}, ${ct || 'kein Content-Type'})`);
  if (res.status >= 300 && res.status < 400) bad(`Weiterleitung nach ${res.headers.get('location')} – vermutlich Login. Hosting muss anonym (public) sein.`);
  else if (res.status !== 200) bad(`HTTP ${res.status}`);
  else if (!ct.toLowerCase().includes(type)) bad(`falscher Content-Type (erwartet ${type})`);
  if (framed && res.status === 200) {
    const xfo = res.headers.get('x-frame-options');
    const csp = res.headers.get('content-security-policy') || '';
    const fa = (csp.match(/frame-ancestors[^;]*/i) || [''])[0];
    if (xfo && /deny|sameorigin/i.test(xfo)) bad(`X-Frame-Options: ${xfo} → Outlook (Web/neu) kann den Task Pane nicht einbetten`);
    if (fa && !/\*|outlook|office|microsoft/i.test(fa)) bad(`CSP ${fa} → Einbetten in Outlook blockiert`);
    if (!xfo && !fa) console.log('   ✓ Einbetten erlaubt (kein X-Frame-Options / frame-ancestors)');
    else if (!problems) console.log(`   ✓ Einbetten erlaubt (${xfo || fa})`);
  }
}

console.log(problems
  ? `\n${problems} Problem(e): Dieses Hosting taugt so nicht für Outlook. Alternative: GitHub Pages (siehe README).`
  : '\nAlles gut – jetzt: npm run set-url -- ' + base);
process.exit(problems ? 1 : 0);
