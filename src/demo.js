// demo.js – Simulierte E-Mail + simuliertes Azure DevOps, wenn die Seite ausserhalb von Outlook geöffnet wird.
// So lässt sich die komplette Oberfläche im Browser testen (https://localhost:3000/src/taskpane.html).

import { AdoError } from './ado.js';

const b64 = (s) => btoa(unescape(encodeURIComponent(s)));
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

const IMG = `<svg xmlns="http://www.w3.org/2000/svg" width="340" height="110"><rect width="340" height="110" rx="8" fill="#fde7e9"/><rect x="0" y="0" width="6" height="110" fill="#c50f1f"/><text x="22" y="42" font-family="Segoe UI,Arial" font-size="17" font-weight="600" fill="#a4262c">Gutschein ungültig (ERR_VOUCHER_409)</text><text x="22" y="72" font-family="Segoe UI,Arial" font-size="13" fill="#605e5c">Screenshot Checkout · Schritt 3 von 4</text></svg>`;

const BODY = `<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style>p.MsoNormal{margin:0cm;font-size:11pt;font-family:"Calibri",sans-serif}</style></head>
<body lang="DE-CH" link="#0563C1"><div class="WordSection1">
<p class="MsoNormal">Hallo Thomas<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Danke für die schnelle Analyse. Wir haben den Fix auf der Testumgebung geprüft – der Checkout funktioniert jetzt, <b>aber</b> beim Einlösen eines Gutscheincodes erscheint weiterhin folgender Fehler:<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal"><img width="340" height="110" src="cid:image001.png@01DB3A2F.4C1E8B60" alt="Screenshot Fehler"><o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Könnt ihr das bis <span style="color:#C00000;font-weight:bold">Freitag</span> anschauen? Die Weihnachtskampagne startet am Montag und der Code <span style="font-family:Consolas">XMAS26</span> wird im Newsletter beworben.<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Freundliche Grüsse<br>Sandra Keller<br>Leiterin E-Commerce | Muster AG<br><a href="https://www.muster.ch">www.muster.ch</a><o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<img width="1" height="1" src="https://tracking.example.com/p.gif">
<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm">
<p class="MsoNormal"><b>Von:</b> Thomas Brunner &lt;thomas.brunner@devpartner.ch&gt; <br><b>Gesendet:</b> Dienstag, 17. November 2026 16:42<br><b>An:</b> Sandra Keller &lt;sandra.keller@muster.ch&gt;<br><b>Cc:</b> Support DevPartner &lt;support@devpartner.ch&gt;<br><b>Betreff:</b> AW: Checkout-Fehler im Webshop #4711<o:p></o:p></p>
</div>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Hallo Sandra<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Wir konnten das Problem reproduzieren. Ursache war ein Timeout beim Payment-Provider, wenn der Warenkorb mehr als 20 Positionen enthält. Hier die Ergebnisse unserer Tests:<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt">
<tr><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt;background:#D9E2F3"><p class="MsoNormal"><b>Umgebung</b></p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt;background:#D9E2F3"><p class="MsoNormal"><b>Ergebnis</b></p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt;background:#D9E2F3"><p class="MsoNormal"><b>Bemerkung</b></p></td></tr>
<tr><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">DEV</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal" style="color:#00B050">OK</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">Fix deployed</p></td></tr>
<tr><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">TEST</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal" style="color:#00B050">OK</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">Bitte fachlich abnehmen</p></td></tr>
<tr><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">PROD</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal" style="color:#C00000">offen</p></td><td style="border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class="MsoNormal">Release am Donnerstag</p></td></tr>
</table>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Kannst du den Fix auf TEST bitte bis morgen abnehmen?<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<p class="MsoNormal">Gruss<br>Thomas<o:p></o:p></p>
<p class="MsoNormal"><o:p>&nbsp;</o:p></p>
<div id="appendonsend"></div>
<hr style="display:inline-block;width:98%" tabindex="-1">
<div id="divRplyFwdMsg" dir="ltr"><font face="Calibri, sans-serif" style="font-size:11pt" color="#000000"><b>Von:</b> Sandra Keller &lt;sandra.keller@muster.ch&gt;<br><b>Gesendet:</b> Montag, 16. November 2026 09:15<br><b>An:</b> Thomas Brunner &lt;thomas.brunner@devpartner.ch&gt;<br><b>Betreff:</b> Checkout-Fehler im Webshop #4711</font><div>&nbsp;</div></div>
<div dir="ltr"><div style="font-family:Aptos,sans-serif;font-size:12pt;color:rgb(0,0,0)">Guten Morgen Thomas</div>
<div style="font-family:Aptos,sans-serif;font-size:12pt"><br></div>
<div style="font-family:Aptos,sans-serif;font-size:12pt">seit dem Update vom Wochenende brechen mehrere Kunden den Checkout ab. Betroffen sind laut unserem Monitoring:</div>
<ul><li>Bestellungen mit mehr als 20 Positionen</li><li>vor allem Zahlungen mit Kreditkarte</li><li>ca. 12 % aller Bestellungen seit Samstag</li></ul>
<div style="font-family:Aptos,sans-serif;font-size:12pt">Bitte mit <b>hoher Priorität</b> anschauen.</div>
<div style="font-family:Aptos,sans-serif;font-size:12pt"><br></div>
<div style="font-family:Aptos,sans-serif;font-size:12pt">Danke und Gruss<br>Sandra</div></div>
</div></body></html>`;

const PDF = b64('%PDF-1.4\n% Demo-Anhang\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF');

export const demoMail = {
  kind: 'demo',
  hasItem: () => true,
  async readCurrentMail() {
    await wait(150);
    return {
      mode: new URLSearchParams(location.search).has('compose') ? 'compose' : 'read',
      itemId: 'demo-item',
      conversationId: 'demo-conversation',
      subject: 'AW: AW: Checkout-Fehler im Webshop #4711',
      from: { name: 'Sandra Keller', email: 'sandra.keller@muster.ch' },
      to: [{ name: 'Thomas Brunner', email: 'thomas.brunner@devpartner.ch' }],
      cc: [{ name: 'Support DevPartner', email: 'support@devpartner.ch' }],
      date: new Date('2026-11-18T08:03:00'),
      attachments: [
        { id: 'att-img', name: 'image001.png', contentType: 'image/svg+xml', size: IMG.length, isInline: true, type: 'file' },
        { id: 'att-pdf', name: 'Testprotokoll_Checkout.pdf', contentType: 'application/pdf', size: 48213, isInline: false, type: 'file' },
      ],
      bodyHtml: BODY,
    };
  },
  canGetAttachments: () => true,
  async getAttachmentContent(id) {
    await wait(80);
    if (id === 'att-img') return { format: 'base64', content: b64(IMG) };
    return { format: 'base64', content: PDF };
  },
  canGetEml: () => true,
  getEmlBase64: async () => b64('From: sandra.keller@muster.ch\r\nSubject: Demo\r\n\r\nDemo'),
  canComposeSelection: () => new URLSearchParams(location.search).has('compose'),
  getComposeSelection: async () => '<div>Markierter Text aus dem Editor (Demo)</div>',
  onItemChanged: () => {},
  userEmail: () => 'sie@muster.ch',
  openUrl: (url) => window.open(url, '_blank', 'noopener'),
};

/* ------------------------------------------------------------ Mock-ADO */

const PROJECTS = [
  { id: 'p1', name: 'Webshop', types: ['Product Backlog Item', 'Task', 'Bug', 'Feature', 'Epic'] },
  { id: 'p2', name: 'Mobile App', types: ['User Story', 'Task', 'Bug', 'Feature', 'Epic'] },
  { id: 'p3', name: 'Intern', types: ['Issue', 'Task', 'Epic'] },
];
let nextId = 4800;
const ITEMS = [
  [4711, 'Checkout schlägt bei grossen Warenkörben fehl', 'Bug', 'Committed', 'Thomas Brunner', 'Webshop'],
  [4702, 'Gutscheincodes im Checkout einlösen', 'Product Backlog Item', 'Approved', 'Sie', 'Webshop'],
  [4698, 'Payment-Provider Timeout erhöhen', 'Task', 'In Progress', 'Thomas Brunner', 'Webshop'],
  [4690, 'Weihnachtskampagne 2026 – Landingpage', 'Feature', 'In Progress', 'Sie', 'Webshop'],
  [4655, 'Newsletter-Anmeldung DSGVO-konform', 'Product Backlog Item', 'Done', 'Sie', 'Webshop'],
  [4640, 'Produktfilter langsam bei >1000 Artikeln', 'Bug', 'New', '', 'Webshop'],
  [3120, 'Push-Benachrichtigungen für Bestellstatus', 'User Story', 'Active', 'Sie', 'Mobile App'],
  [3118, 'Login mit Face ID', 'User Story', 'New', '', 'Mobile App'],
  [3101, 'App stürzt beim Öffnen der Wunschliste ab', 'Bug', 'Active', 'Thomas Brunner', 'Mobile App'],
  [210, 'Onboarding neue Mitarbeitende', 'Issue', 'To Do', 'Sie', 'Intern'],
].map(([id, title, type, state, assignedTo, project]) => ({ id, title, type, state, assignedTo, project, comments: [], attachments: [] }));

export class MockAdoClient {
  constructor(org) {
    this.org = org;
    this.base = org.url;
    this.isMock = true;
  }
  webUrl(project, id) { return `${this.base}/${encodeURIComponent(project)}/_workitems/edit/${id}`; }
  async testConnection() { await wait(200); return { count: PROJECTS.length }; }
  async getProjects() { await wait(250); return PROJECTS.map(({ id, name }) => ({ id, name })); }
  async getWorkItemTypes(project) { await wait(120); return (PROJECTS.find((p) => p.name === project) || PROJECTS[0]).types; }
  async queryWorkItems(project, { mode = 'recent', text = '' } = {}) {
    await wait(250);
    const t = text.toLowerCase();
    return ITEMS.filter((i) => i.project === project && i.state !== 'Removed')
      .filter((i) => (mode === 'mine' ? i.assignedTo === 'Sie' : true))
      .filter((i) => (t ? i.title.toLowerCase().includes(t) : true))
      .map(strip);
  }
  async getWorkItem(id) {
    await wait(150);
    const i = ITEMS.find((x) => x.id === Number(id));
    if (!i) throw new AdoError(`Nicht gefunden: TF401232: Work Item ${id} existiert nicht oder Sie haben keine Berechtigung.`, 404);
    return strip(i);
  }
  async createWorkItem(project, type, { title, descriptionHtml, tags, assignedTo, parentId } = {}) {
    await wait(400);
    const wi = { id: ++nextId, title, type, state: 'New', assignedTo: assignedTo ? 'Sie' : '', project, comments: [], attachments: [], descriptionHtml, tags, parentId };
    ITEMS.unshift(wi);
    console.info('[Demo] Work Item erstellt', wi);
    return strip(wi);
  }
  async addComment(project, id, html) {
    await wait(300);
    const i = ITEMS.find((x) => x.id === Number(id));
    i.comments.push(html);
    console.info(`[Demo] Kommentar an #${id}`, html);
    return { id: i.comments.length };
  }
  async uploadAttachment(fileName, bytes) {
    await wait(150);
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    const type = { png: 'image/svg+xml', svg: 'image/svg+xml', jpg: 'image/jpeg', gif: 'image/gif', pdf: 'application/pdf' }[ext] || 'application/octet-stream';
    return { id: crypto.randomUUID(), url: URL.createObjectURL(new Blob([bytes], { type })) };
  }
  async attachFiles(id, files) {
    await wait(200);
    ITEMS.find((x) => x.id === Number(id)).attachments.push(...files);
    console.info(`[Demo] Anhänge an #${id}`, files);
    return {};
  }
  /** nur Demo: zuletzt gespeicherte Kommentare zum Anzeigen */
  getComments(id) { return (ITEMS.find((x) => x.id === Number(id)) || {}).comments || []; }
  getDescription(id) { return (ITEMS.find((x) => x.id === Number(id)) || {}).descriptionHtml || ''; }
}

function strip({ id, title, type, state, assignedTo, project }) { return { id, title, type, state, assignedTo, project }; }

export const DEMO_ORG = { id: 'demo', name: 'Demo-Organisation (simuliert)', url: 'https://dev.azure.com/demo', pat: 'demo', demo: true, apiVersion: '7.1', managed: true };
