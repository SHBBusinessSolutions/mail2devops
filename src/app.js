// app.js – UI-Steuerung des Task Panes.
import { loadAdminConfig, getOrgs, saveOrg, deleteOrg, getPrefs, savePrefs, normalizeOrgUrl, getLink, setLink } from './config.js';
import { AdoClient, typeColor } from './ado.js';
import {
  splitThread, sanitizeHtml, textOf, headerBlock, noteBlock, snippetBlock, esc, cleanSubject,
  b64ToBytes, safeFileName, personName, fmtDate,
} from './parse.js';

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const show = (el, on = true) => (typeof el === 'string' ? $(el) : el).classList.toggle('hidden', !on);
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const fmtSize = (n) => (n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB');

const S = {
  demo: false,
  mailApi: null,
  makeClient: (org) => new AdoClient(org),
  mail: null,
  messages: [],
  fullHtml: '',
  orgs: [],
  prefs: {},
  org: null,
  client: null,
  projects: [],
  project: null,
  target: 'existing',
  selectedItem: null,
  types: [],
  newType: null,
  content: 'full',
  threadSel: new Set(),
  snippets: [],
  inlineImages: new Map(), // key → {name, contentType, base64}
  uploadCache: new Map(),
  highlight: null,
  busy: false,
};
let snipId = 0;

/* ================================================================ Boot */

function officeReady(ms) {
  return new Promise((resolve) => {
    if (!window.Office || !Office.onReady) return resolve(null);
    let done = false;
    const t = setTimeout(() => { if (!done) { done = true; resolve(null); } }, ms);
    try {
      Office.onReady((info) => { if (done) return; done = true; clearTimeout(t); resolve(info); });
    } catch { clearTimeout(t); resolve(null); }
  });
}

async function boot() {
  const info = await officeReady(4000);
  const inOutlook = !!(info && info.host && window.Office?.context?.mailbox);
  await loadAdminConfig();

  if (inOutlook) {
    const { officeMail } = await import('./mail.js');
    S.mailApi = officeMail;
  } else {
    const demo = await import('./demo.js');
    S.demo = true;
    S.mailApi = demo.demoMail;
    S.makeClient = (org) => (org.demo ? new demo.MockAdoClient(org) : new AdoClient(org));
    S.demoOrg = demo.DEMO_ORG;
    show('#demoBanner');
  }

  if (window.CSS && CSS.highlights && window.Highlight) {
    S.highlight = new Highlight();
    CSS.highlights.set('m2d-snippet', S.highlight);
  }

  S.prefs = getPrefs();
  wireUi();
  applyPrefsToUi();
  reloadOrgs();

  S.mailApi.onItemChanged(() => loadMail().catch(showMailError));
  await loadMail().catch(showMailError);
  await selectOrg(S.prefs.lastOrgId);
}

function reloadOrgs() {
  S.orgs = getOrgs();
  if (S.demo && S.demoOrg) S.orgs.push(S.demoOrg);
}

function applyPrefsToUi() {
  $('#optHeader').checked = S.prefs.includeHeader !== false;
  $('#optOldestFirst').checked = S.prefs.oldestFirst !== false;
  $('#optImages').checked = S.prefs.transferImages !== false;
  $('#newTags').value = S.prefs.defaultTags || '';
  setContent(S.prefs.defaultContent || 'full', { silent: true });
}

function persistPrefs() { savePrefs(S.prefs); }

/* ================================================================ Mail */

function showMailError(e) {
  console.error(e);
  $('#preview').innerHTML = `<div class="empty">${esc(e.message || 'E-Mail konnte nicht gelesen werden.')}</div>`;
}

async function loadMail() {
  if (!S.mailApi.hasItem()) {
    S.mail = null;
    $('#preview').innerHTML = '<div class="empty">Bitte eine E-Mail auswählen.</div>';
    $('#mailSubject').textContent = '';
    updateSubmit();
    return;
  }
  $('#preview').innerHTML = '<div class="loading">E-Mail wird geladen …</div>';
  const mail = await S.mailApi.readCurrentMail();
  S.mail = mail;
  const { messages, fullHtml } = splitThread(mail.bodyHtml, mail);
  S.messages = messages;
  S.fullHtml = fullHtml;
  S.threadSel = new Set(messages.map((m) => m.index));
  S.snippets = [];
  S.inlineImages = new Map();
  if (S.highlight) S.highlight.clear();
  show('#result', false);

  $('#mailSubject').textContent = mail.subject || '(ohne Betreff)';
  $('#msgCount').textContent = messages.length;
  $('#newTitle').value = cleanSubject(mail.subject);
  show('#btnComposeSel', mail.mode === 'compose' && S.mailApi.canComposeSelection());
  show('#optEmlWrap', mail.mode === 'read' && S.mailApi.canGetEml());
  $('#optEml').checked = false;

  renderPreview();
  renderThreadList();
  renderSnippets();
  renderAttachments();
  renderContentHint();
  updateSubmit();
  loadInlineImages();
  applyConversationLink();
}

async function loadInlineImages() {
  const keys = [...new Set($$('#preview img[data-img-key]').map((i) => i.getAttribute('data-img-key')))];
  if (!keys.length || !S.mailApi.canGetAttachments()) return;
  const inline = (S.mail.attachments || []).filter((a) => a.isInline || /^image\//i.test(a.contentType || ''));
  const used = new Set();
  const map = new Map();
  for (const k of keys.filter((k) => k.startsWith('cid:'))) {
    const cid = k.slice(4).toLowerCase();
    const base = cid.split('@')[0];
    const a = inline.find((x) => !used.has(x.id) && (x.name || '').toLowerCase() === base)
      || inline.find((x) => !used.has(x.id) && cid.includes((x.name || '').toLowerCase()));
    if (a) { map.set(k, a); used.add(a.id); }
  }
  for (const k of keys) {
    if (map.has(k)) continue;
    const a = inline.find((x) => !used.has(x.id));
    if (a) { map.set(k, a); used.add(a.id); }
  }
  for (const [k, a] of map) {
    try {
      const c = await S.mailApi.getAttachmentContent(a.id);
      if (c.format !== 'base64') continue;
      const ct = a.contentType || 'image/png';
      S.inlineImages.set(k, { name: a.name || 'bild.png', contentType: ct, base64: c.content });
      $$(`#preview img[data-img-key="${CSS.escape(k)}"]`).forEach((img) => { img.src = `data:${ct};base64,${c.content}`; });
    } catch (e) { console.warn('Inline-Bild', a.name, e); }
  }
}

function renderPreview() {
  const pv = $('#preview');
  if (!S.messages.length) { pv.innerHTML = '<div class="empty">Kein Inhalt</div>'; return; }
  pv.innerHTML = S.messages.map((m) => `
    <section class="msg" data-msg-index="${m.index}">
      <div class="msg-head">${avatar(personName(m.meta.from))}<span class="msg-from">${esc(personName(m.meta.from) || 'Unbekannt')}</span><span class="msg-date">${esc(fmtDate(m.meta.date))}</span></div>
      <div class="msg-body">${m.html}</div>
    </section>`).join('');
  // cid:-Bilder bis zum Laden nicht als defekt anzeigen
  $$('img', pv).forEach((img) => { if (/^cid:/i.test(img.getAttribute('src') || '')) img.removeAttribute('src'); });
  updatePreviewDim();
}

function orderedMessages() {
  const list = [...S.messages];
  return $('#optOldestFirst').checked ? list.reverse() : list;
}

function renderThreadList() {
  $('#threadList').innerHTML = orderedMessages().map((m) => `
    <label class="thread-item">
      <input type="checkbox" data-i="${m.index}" ${S.threadSel.has(m.index) ? 'checked' : ''}>
      <div class="ti-body">
        <div class="ti-head"><b>${esc(personName(m.meta.from) || 'Unbekannt')}</b><span>${esc(fmtDate(m.meta.date))}</span></div>
        <div class="ti-excerpt">${esc(m.excerpt || '(kein Text)')}</div>
      </div>
    </label>`).join('');
}

function renderSnippets() {
  $('#selCount').textContent = S.snippets.length;
  $('#snippetList').innerHTML = S.snippets.length
    ? S.snippets.map((s) => `
      <div class="snippet" data-id="${s.id}">
        <div class="sn-body">
          <div class="sn-meta">${esc([personName(s.meta?.from), fmtDate(s.meta?.date)].filter(Boolean).join(' · ') || 'Ausschnitt')}</div>
          <div class="sn-text">${esc(s.text.slice(0, 400))}</div>
        </div>
        <button type="button" title="Entfernen" data-remove="${s.id}">✕</button>
      </div>`).join('')
    : '';
}

function renderAttachments() {
  const atts = (S.mail?.attachments || []).filter((a) => !a.isInline);
  const can = S.mailApi.canGetAttachments();
  $('#attList').innerHTML = atts.length
    ? `<div class="lbl">Anhänge der E-Mail ans Work Item hängen</div>` + atts.map((a) => `
      <label class="check att-row">
        <input type="checkbox" data-att="${esc(a.id)}" ${can && a.type !== 'cloud' ? '' : 'disabled'}>
        <span>📎 ${esc(a.name)}</span><span class="size">${a.type === 'cloud' ? 'Cloud-Datei' : fmtSize(a.size || 0)}</span>
      </label>`).join('')
    : '';
}

function renderContentHint() {
  const h = {
    full: 'Die ganze E-Mail wird ein Kommentar.',
    thread: S.messages.length > 1 ? 'Jede Nachricht wird ein eigener Kommentar.' : 'Nur eine Nachricht erkannt.',
    selection: S.snippets.length ? 'Ausschnitte werden als Zitate gespeichert.' : 'Text in der Vorschau markieren, dann „Übernehmen“.',
  }[S.content];
  $('#contentHint').innerHTML = h;
}

function setContent(v, { silent } = {}) {
  S.content = v;
  $$('#segContent button').forEach((b) => b.classList.toggle('active', b.dataset.v === v));
  show('#threadList', v === 'thread');
  show('#snippetList', v === 'selection');
  if (!silent) { renderContentHint(); updateSubmit(); }
  updatePreviewDim();
}

function updatePreviewDim() {
  $$('#preview .msg').forEach((s) => s.classList.toggle('dim', S.content === 'thread' && !S.threadSel.has(+s.dataset.msgIndex)));
}

/* ============================================================ Selection */

function closestSection(node) {
  const el = node && (node.nodeType === 1 ? node : node.parentElement);
  return el ? el.closest('section.msg') : null;
}

function selectionInPreview() {
  const sel = getSelection();
  if (!sel || !sel.rangeCount || sel.isCollapsed) return null;
  const r = sel.getRangeAt(0);
  return $('#preview').contains(r.commonAncestorContainer) ? r : null;
}

function updateFab() {
  const fab = $('#selFab');
  const r = selectionInPreview();
  if (!r || !r.toString().trim()) { show(fab, false); return; }
  const rect = r.getBoundingClientRect();
  show(fab, true);
  const fw = fab.offsetWidth || 110;
  fab.style.left = Math.min(Math.max(8, rect.left + rect.width / 2 - fw / 2), innerWidth - fw - 8) + 'px';
  fab.style.top = Math.min(Math.max(8, rect.bottom + 8), innerHeight - 120) + 'px';
}

function captureSelection() {
  const range = selectionInPreview();
  if (!range) return;
  const div = document.createElement('div');
  div.appendChild(range.cloneContents());
  div.querySelectorAll('.msg-head').forEach((e) => e.remove());
  div.querySelectorAll('section.msg, .msg-body').forEach((e) => { while (e.firstChild) e.parentNode.insertBefore(e.firstChild, e); e.remove(); });
  const text = textOf(div);
  if (!text.trim() && !div.querySelector('img')) return;
  const html = sanitizeHtml(div);
  const sec = closestSection(range.startContainer);
  const msg = sec ? S.messages[+sec.dataset.msgIndex] : null;
  const sn = { id: ++snipId, html, text, meta: msg ? msg.meta : null, range: range.cloneRange() };
  S.snippets.push(sn);
  if (S.highlight) S.highlight.add(sn.range);
  getSelection().removeAllRanges();
  show('#selFab', false);
  setContent('selection');
  renderSnippets();
  renderContentHint();
  updateSubmit();
  toast('Ausschnitt übernommen');
}

async function captureComposeSelection() {
  try {
    const html = await S.mailApi.getComposeSelection();
    if (!html) { toast('Im Editor ist nichts markiert'); return; }
    const clean = sanitizeHtml(html);
    const tdoc = new DOMParser().parseFromString(`<body>${clean}</body>`, 'text/html');
    S.snippets.push({ id: ++snipId, html: clean, text: textOf(tdoc.body), meta: S.messages[0]?.meta || null });
    setContent('selection');
    renderSnippets(); renderContentHint(); updateSubmit();
    toast('Ausschnitt übernommen');
  } catch (e) { setStatus(e.message, 'error'); }
}

function removeSnippet(id) {
  const i = S.snippets.findIndex((s) => s.id === id);
  if (i < 0) return;
  const [sn] = S.snippets.splice(i, 1);
  if (S.highlight && sn.range) S.highlight.delete(sn.range);
  renderSnippets(); renderContentHint(); updateSubmit();
}

/* ============================================================ Org/Projekt */

function renderOrgSelect() {
  const sel = $('#selOrg');
  sel.innerHTML = S.orgs.map((o) => `<option value="${esc(o.id)}">${esc(o.name || o.url)}</option>`).join('');
  sel.value = S.org ? S.org.id : '';
  show('#fieldOrg', S.orgs.length > 1);
  show('#setupNeeded', !S.orgs.length);
  show('#cardTarget', !!S.orgs.length);
}

function renderProjects() {
  const sel = $('#selProject');
  sel.innerHTML = S.projects.length
    ? S.projects.map((p) => `<option value="${esc(p.name)}">${esc(p.name)}</option>`).join('')
    : '<option value="">–</option>';
  sel.value = S.project || '';
  sel.disabled = !S.projects.length;
}

async function selectOrg(id) {
  S.org = S.orgs.find((o) => o.id === id) || S.orgs[0] || null;
  S.client = S.org ? S.makeClient(S.org) : null;
  S.projects = []; S.project = null;
  renderOrgSelect(); renderProjects(); renderWiList([]);
  updateSubmit();
  show('#patMissing', !!(S.org && !S.org.pat));
  show('#targetBody', !!(S.org && S.org.pat));
  if (!S.org) return;
  S.prefs.lastOrgId = S.org.id; persistPrefs();
  if (!S.org.pat) return;
  try {
    setStatus('Projekte werden geladen …');
    S.projects = await S.client.getProjects();
    const pref = (S.prefs.lastProject || {})[S.org.id] || S.org.defaultProject;
    S.project = (S.projects.find((p) => p.name === pref) || S.projects[0] || {}).name || null;
    renderProjects();
    setStatus('');
    await onProjectChanged();
  } catch (e) { setStatus(e.message, 'error'); }
}

async function onProjectChanged() {
  if (!S.project) return;
  S.prefs.lastProject = { ...(S.prefs.lastProject || {}), [S.org.id]: S.project };
  persistPrefs();
  loadTypes();
  if (!S.selectedItem) {
    const pending = S.pendingLink;
    S.pendingLink = null;
    if (pending) return pickById(pending.id, { silent: true });
    refreshSearch();
  }
}

async function loadTypes() {
  S.types = []; renderTypes();
  try {
    S.types = await S.client.getWorkItemTypes(S.project);
    const last = (S.prefs.lastType || {})[S.project];
    S.newType = S.types.includes(last) ? last : S.types[0] || null;
  } catch (e) { console.warn(e); S.types = []; }
  renderTypes(); updateSubmit();
}

function renderTypes() {
  $('#typeChips').innerHTML = S.types.length
    ? S.types.map((t) => `<button type="button" class="chip ${t === S.newType ? 'active' : ''}" data-type="${esc(t)}"><span class="dot" style="background:${typeColor(t)}"></span>${esc(t)}</button>`).join('')
    : '<span class="hint">Typen werden geladen …</span>';
}

/* ============================================================ Work Items */

let searchSeq = 0;
async function refreshSearch() {
  if (!S.client || !S.project) return;
  const raw = $('#wiSearch').value.trim();
  const q = raw.replace(/^(?:AB)?#/i, '');
  const seq = ++searchSeq;
  $('#wiList').innerHTML = '<div class="wi-empty">Suche …</div>';
  try {
    let items;
    if (/^\d+$/.test(q)) {
      try { items = [await S.client.getWorkItem(+q)]; } catch (e) { if (e.status === 404) items = []; else throw e; }
    } else {
      items = await S.client.queryWorkItems(S.project, { mode: $('#wiFilter').value, text: q });
    }
    if (seq === searchSeq) renderWiList(items, q);
  } catch (e) {
    if (seq === searchSeq) $('#wiList').innerHTML = `<div class="wi-empty">${esc(e.message)}</div>`;
  }
}

const TYPE_ABBR = { 'product backlog item': 'PBI', 'user story': 'US', task: 'T', bug: 'B', feature: 'F', epic: 'E', issue: 'I', requirement: 'R' };
function typeIcon(type) {
  const abbr = TYPE_ABBR[(type || '').toLowerCase()] || (type || '?').slice(0, 1).toUpperCase();
  return `<span class="wi-type" style="--tc:${typeColor(type)}" title="${esc(type)}">${esc(abbr)}</span>`;
}

const AVATAR_COLORS = ['#0f6cbd', '#8764b8', '#038387', '#ca5010', '#4f6bed', '#c239b3', '#498205', '#986f0b'];
function avatar(name) {
  const n = (name || '?').trim();
  const parts = n.replace(/[<(].*$/, '').split(/[\s,.]+/).filter(Boolean);
  const ini = ((parts[0] || '?')[0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
  let h = 0;
  for (const c of n) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return `<span class="avatar" style="background:${AVATAR_COLORS[h % AVATAR_COLORS.length]}" aria-hidden="true">${esc(ini)}</span>`;
}

function wiMeta(w) {
  return [w.type, w.state, w.assignedTo, w.project && w.project !== S.project ? w.project : ''].filter(Boolean).map(esc).join(' · ');
}

function renderWiList(items, q = '') {
  S.lastItems = items || [];
  if (!items) return;
  $('#wiList').innerHTML = items.length
    ? items.map((w) => `
      <button type="button" class="wi" data-id="${w.id}" role="option">
        ${typeIcon(w.type)}
        <span class="wi-main"><span class="wi-title"><span class="wi-id">${w.id}</span> ${esc(w.title)}</span><span class="wi-meta">${wiMeta(w)}</span></span>
      </button>`).join('')
    : (q || S.client) ? `<div class="wi-empty">${q ? 'Keine Treffer.' : 'Keine Work Items gefunden.'}</div>` : '';
}

function selectItem(w) {
  S.selectedItem = w;
  show('#wiPicker', !w);
  show('#wiSelected', !!w);
  if (w) {
    const url = S.client.webUrl(w.project || S.project, w.id);
    $('#wiSelected').innerHTML = `
      <div class="wi-sel">
        ${typeIcon(w.type)}
        <div class="wi-main"><div class="wi-title"><span class="wi-id">${w.id}</span> ${esc(w.title)}</div><div class="wi-meta">${wiMeta(w)} · <a href="${esc(url)}" target="_blank" rel="noopener" data-open>Öffnen</a></div></div>
        <button type="button" class="icon-btn" id="btnChangeItem" title="Anderes Item wählen" aria-label="Anderes Item wählen">✕</button>
      </div>`;
  }
  updateSubmit();
}

async function pickById(id, { silent } = {}) {
  try {
    const w = await S.client.getWorkItem(id);
    selectItem(w);
  } catch (e) {
    if (!silent) setStatus(e.message, 'error');
    refreshSearch();
  }
}

function applyConversationLink() {
  show('#linkedHint', false);
  const link = S.mail && getLink(S.mail.conversationId);
  if (link) {
    $('#linkedHint').innerHTML = `<span>🔗</span><span class="grow">Dieser Verlauf wurde bereits an <b>#${link.id}</b> ${esc(link.title || '')} angehängt.</span>`;
    show('#linkedHint', true);
    if (S.target === 'existing' && (!S.selectedItem || S.selectedItem.id !== link.id)) {
      S.selectedItem = null;
      if (S.org && S.org.id === link.orgId && S.client && S.project) pickById(link.id, { silent: true });
      else S.pendingLink = link;
    }
    return;
  }
  // Work-Item-ID im Betreff erkennen ("#4711", "AB#4711")
  const m = (S.mail?.subject || '').match(/(?:^|[\s(\[])(?:AB)?#(\d{1,7})\b/i);
  if (m && !S.selectedItem) {
    $('#wiSearch').value = m[1];
    if (S.client && S.project) refreshSearch();
  }
}

/* ============================================================== Submit */

function setTarget(v) {
  S.target = v;
  $$('#segTarget button').forEach((b) => b.classList.toggle('active', b.dataset.v === v));
  show('#paneExisting', v === 'existing');
  show('#paneNew', v === 'new');
  updateSubmit();
}

function validate() {
  if (!S.mail) return 'Keine E-Mail ausgewählt.';
  if (!S.org) return 'Bitte zuerst eine Organisation in den Einstellungen anlegen.';
  if (!S.org.pat) return 'Für diese Organisation fehlt ein PAT.';
  if (S.target === 'existing' && !S.selectedItem) return 'Bitte ein Work Item auswählen.';
  if (S.target === 'new') {
    if (!S.project) return 'Bitte ein Projekt wählen.';
    if (!S.newType) return 'Bitte einen Typ wählen.';
    if (!$('#newTitle').value.trim()) return 'Bitte einen Titel eingeben.';
    const parent = $('#newParent').value.trim().replace(/^#/, '');
    if (parent && !/^\d+$/.test(parent)) return 'Übergeordnete ID muss eine Zahl sein.';
  }
  if (S.content === 'selection' && !S.snippets.length) return 'Noch kein Ausschnitt markiert.';
  if (S.content === 'thread' && !S.threadSel.size) return 'Keine Nachricht des Verlaufs ausgewählt.';
  return '';
}

function updateSubmit() {
  const btn = $('#btnSubmit');
  const err = validate();
  btn.disabled = !!err || S.busy;
  btn.title = err;
  const n = S.content === 'thread' ? S.threadSel.size : 1;
  const what = S.content === 'selection' ? `${S.snippets.length || ''} Ausschnitt${S.snippets.length === 1 ? '' : 'e'}`.trim()
    : S.content === 'thread' ? `${n} Kommentar${n === 1 ? '' : 'e'}` : 'E-Mail';
  if (S.target === 'new') btn.textContent = `${S.newType || 'Item'} erstellen + ${what}`;
  else btn.textContent = S.selectedItem ? `${what} an #${S.selectedItem.id} anhängen` : 'Work Item auswählen';
}

function setStatus(msg, kind = '') {
  const el = $('#status');
  el.textContent = msg || '';
  el.className = 'status ' + kind;
}

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  show(t, true);
  clearTimeout(toast.t);
  toast.t = setTimeout(() => show(t, false), 1800);
}

function mailMeta() {
  const m = S.mail;
  return { from: m.from, to: m.to, cc: m.cc, date: m.date, subject: m.subject };
}

async function buildContentBlocks() {
  const note = noteBlock($('#note').value.trim());
  const withHeader = $('#optHeader').checked;
  const subject = S.mail.subject;
  let blocks;
  if (S.content === 'full') {
    blocks = [note + (withHeader ? headerBlock(mailMeta()) : '') + S.fullHtml];
  } else if (S.content === 'thread') {
    const msgs = orderedMessages().filter((m) => S.threadSel.has(m.index));
    blocks = msgs.map((m, i) =>
      (i === 0 ? note : '') + (withHeader ? headerBlock(m.meta, { subject, title: msgs.length > 1 ? `📧 ${i + 1}/${msgs.length} ·` : '📧' }) : '') + m.html);
  } else {
    blocks = [note + (withHeader ? headerBlock({ subject }, { title: '✂️ Auszug aus E-Mail:' }) : '') + S.snippets.map(snippetBlock).join('')];
  }
  const out = [];
  for (const b of blocks) out.push(await processImages(b));
  return out;
}

async function processImages(html) {
  if (!/<img/i.test(html)) return html;
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  const transfer = $('#optImages').checked;
  for (const img of [...tpl.content.querySelectorAll('img')]) {
    const key = img.getAttribute('data-img-key');
    const src = img.getAttribute('src') || '';
    const label = (t) => img.replaceWith(document.createTextNode(`[${t}${img.alt ? ': ' + img.alt : ''}]`));
    if (!transfer) { label('Bild'); continue; }
    try {
      let file = null, cacheKey = null;
      if (key && S.inlineImages.has(key)) {
        const im = S.inlineImages.get(key);
        cacheKey = key;
        file = { name: im.name, bytes: () => b64ToBytes(im.base64) };
      } else if (src.startsWith('data:image/')) {
        const m = src.match(/^data:(image\/[\w.+-]+);base64,([\s\S]*)$/);
        cacheKey = 'data:' + src.length + ':' + src.slice(-48);
        if (m) file = { name: 'bild.' + (m[1].split('/')[1] || 'png').replace('svg+xml', 'svg').replace('jpeg', 'jpg'), bytes: () => b64ToBytes(m[2]) };
      } else if (/^https?:/i.test(src) && !key) {
        continue; // externe Bilder bleiben verlinkt
      }
      if (!file) { label('Bild'); continue; }
      let url = S.uploadCache.get(cacheKey);
      if (!url) {
        url = (await S.client.uploadAttachment(file.name, file.bytes())).url;
        S.uploadCache.set(cacheKey, url);
      }
      img.setAttribute('src', url);
      img.removeAttribute('data-img-key');
      img.setAttribute('style', 'max-width:100%');
    } catch (e) {
      console.warn(e);
      label('Bild konnte nicht übertragen werden');
    }
  }
  return tpl.innerHTML;
}

async function collectAttachmentFiles() {
  const files = [];
  for (const cb of $$('#attList input[data-att]:checked')) {
    const att = S.mail.attachments.find((a) => a.id === cb.dataset.att);
    if (!att) continue;
    setStatus(`Anhang „${att.name}“ wird gelesen …`);
    const c = await S.mailApi.getAttachmentContent(att.id);
    if (c.format === 'base64') files.push({ name: att.name, bytes: b64ToBytes(c.content) });
    else if (c.format === 'eml') files.push({ name: safeFileName(att.name) + '.eml', bytes: new TextEncoder().encode(c.content) });
    else if (c.format === 'iCalendar') files.push({ name: safeFileName(att.name) + '.ics', bytes: new TextEncoder().encode(c.content) });
  }
  if ($('#optEml').checked && S.mailApi.canGetEml()) {
    setStatus('Original-E-Mail wird exportiert …');
    files.push({ name: safeFileName(S.mail.subject) + '.eml', bytes: b64ToBytes(await S.mailApi.getEmlBase64()) });
  }
  return files;
}

function setBusy(on) {
  S.busy = on;
  $('#btnSubmit').classList.toggle('busy', on);
  updateSubmit();
}

async function submit() {
  if (S.busy) return;
  const err = validate();
  if (err) { setStatus(err, 'error'); return; }
  setBusy(true);
  S.uploadCache = new Map();
  const c = S.client;
  let wi = S.selectedItem;
  let created = false;
  let commentsDone = 0;
  try {
    const contentTarget = S.target === 'new' ? $('#newContentTarget').value : 'comment';
    setStatus('Inhalt wird aufbereitet …');
    const blocks = await buildContentBlocks();

    if (S.target === 'new') {
      const parent = $('#newParent').value.trim().replace(/^#/, '');
      const fields = {
        title: $('#newTitle').value.trim(),
        descriptionHtml: contentTarget !== 'comment' ? blocks.join('<hr>') : '',
        tags: $('#newTags').value.trim(),
        parentId: parent ? +parent : null,
        assignedTo: $('#newAssignMe').checked ? S.mailApi.userEmail() : '',
      };
      setStatus(`${S.newType} wird erstellt …`);
      try {
        wi = await c.createWorkItem(S.project, S.newType, fields);
      } catch (e) {
        if (!fields.assignedTo || e.status !== 400) throw e;
        wi = await c.createWorkItem(S.project, S.newType, { ...fields, assignedTo: '' }); // Identität unbekannt → ohne Zuweisung
        toast('Zuweisung nicht möglich – Item ohne Zuweisung erstellt');
      }
      if (!wi.project) wi.project = S.project;
      created = true;
      S.prefs.lastType = { ...(S.prefs.lastType || {}), [S.project]: S.newType };
      S.prefs.defaultTags = fields.tags;
      persistPrefs();
    }

    const comments = contentTarget === 'description' ? [] : blocks;
    for (let i = 0; i < comments.length; i++) {
      setStatus(comments.length > 1 ? `Kommentar ${i + 1} von ${comments.length} wird gespeichert …` : 'Kommentar wird gespeichert …');
      await c.addComment(wi.project || S.project, wi.id, comments[i]);
      commentsDone++;
    }

    const files = await collectAttachmentFiles();
    if (files.length) {
      const ups = [];
      for (const f of files) {
        setStatus(`„${f.name}“ wird hochgeladen …`);
        const u = await c.uploadAttachment(f.name, f.bytes);
        ups.push({ url: u.url, name: f.name });
      }
      setStatus('Anhänge werden verknüpft …');
      await c.attachFiles(wi.id, ups, 'Aus E-Mail: ' + (S.mail.subject || ''));
    }

    setLink(S.mail.conversationId, { orgId: S.org.id, project: wi.project || S.project, id: wi.id, title: wi.title });
    setStatus('');
    showResult(wi, { created, comments: comments.length, files: files.length, contentTarget });
    if (created) { setTarget('existing'); selectItem(wi); }
  } catch (e) {
    console.error(e);
    if (created) {
      setTarget('existing'); selectItem(wi);
      setStatus(`#${wi.id} wurde erstellt, aber: ${e.message}${commentsDone ? ` (${commentsDone} Kommentar(e) gespeichert)` : ''}. Erneut senden hängt nur noch an.`, 'error');
    } else {
      setStatus(`Fehler: ${e.message}${commentsDone ? ` – ${commentsDone} Kommentar(e) wurden bereits gespeichert.` : ''}`, 'error');
    }
  } finally {
    setBusy(false);
  }
}

function showResult(wi, { created, comments, files, contentTarget }) {
  const url = S.client.webUrl(wi.project || S.project, wi.id);
  const parts = [];
  if (created) parts.push(`${esc(wi.type)} erstellt`);
  if (contentTarget !== 'comment') parts.push('Beschreibung gefüllt');
  if (comments) parts.push(`${comments} Kommentar${comments === 1 ? '' : 'e'}`);
  if (files) parts.push(`${files} Anhang${files === 1 ? '' : 'e'}`);
  let demoHtml = '';
  if (S.client.isMock) {
    const cs = S.client.getComments(wi.id).slice(-Math.max(comments, 0) || undefined);
    const desc = contentTarget !== 'comment' ? S.client.getDescription(wi.id) : '';
    demoHtml = `<div class="lbl">Demo – so erscheint es in Azure DevOps:</div><div class="demo-comments">` +
      (desc ? `<div class="c"><div class="c-head">Beschreibung</div>${desc}</div>` : '') +
      (comments ? cs.slice().reverse().map((h) => `<div class="c"><div class="c-head">Kommentar</div>${h}</div>`).join('') : '') + `</div>`;
  }
  $('#result').innerHTML = `
    <div class="result">
      <h3>✅ Gespeichert in #${wi.id}</h3>
      <div><b>${esc(wi.title)}</b><br><span class="muted">${parts.join(' · ')}</span></div>
      <div><a class="btn primary small" href="${esc(url)}" target="_blank" rel="noopener" data-open>In Azure DevOps öffnen ↗</a></div>
      ${demoHtml}
    </div>`;
  show('#result', true);
  $('#result').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ============================================================ Settings */

function openSettings() {
  show('#viewMain', false); show('#actionbar', false); show('#viewSettings', true);
  $('#pTags').value = S.prefs.defaultTags || '';
  $('#pContent').value = S.prefs.defaultContent || 'full';
  renderOrgList();
  show('#orgForm', false);
  if (!S.orgs.filter((o) => !o.demo).length) editOrg(null);
}

function closeSettings() {
  S.prefs.defaultTags = $('#pTags').value.trim();
  S.prefs.defaultContent = $('#pContent').value;
  $('#newTags').value = S.prefs.defaultTags;
  persistPrefs();
  show('#viewSettings', false); show('#viewMain', true); show('#actionbar', true);
  reloadOrgs();
  S.selectedItem = null; selectItem(null);
  selectOrg(S.org ? S.org.id : S.prefs.lastOrgId);
}

function renderOrgList() {
  $('#orgList').innerHTML = S.orgs.map((o) => `
    <div class="org-item">
      <div class="grow">
        <div><b>${esc(o.name || o.url)}</b>${o.managed && !o.demo ? '<span class="tag">verwaltet</span>' : ''}${o.demo ? '<span class="tag">Demo</span>' : o.pat ? '<span class="tag ok">PAT ✓</span>' : '<span class="tag bad">PAT fehlt</span>'}</div>
        <div class="url">${esc(o.url)}</div>
      </div>
      ${o.demo ? '' : `<button type="button" class="btn link" data-edit="${esc(o.id)}">Bearbeiten</button>`}
      ${o.managed ? '' : `<button type="button" class="btn link" data-del="${esc(o.id)}">Löschen</button>`}
    </div>`).join('');
}

function editOrg(org) {
  show('#orgForm', true);
  $('#orgFormTitle').textContent = org ? `${org.name || org.url} bearbeiten` : 'Neue Organisation';
  $('#fId').value = org ? org.id : '';
  $('#fName').value = org ? org.name || '' : '';
  $('#fUrl').value = org ? org.url : '';
  $('#fUrl').disabled = !!(org && org.managed);
  $('#fPat').value = org ? org.pat || '' : '';
  $('#fSyncPat').checked = !!(org && org.syncPat);
  $('#fApi').value = org ? org.apiVersion || '7.1' : '7.1';
  $('#fProject').innerHTML = `<option value="">Zuletzt verwendetes</option>` + (org && org.defaultProject ? `<option selected>${esc(org.defaultProject)}</option>` : '');
  $('#testResult').textContent = '';
  updatePatLink();
  $('#orgForm').scrollIntoView({ behavior: 'smooth' });
  (org ? $('#fPat') : $('#fUrl')).focus();
}

function updatePatLink() {
  const u = normalizeOrgUrl($('#fUrl').value);
  $('#patLink').href = u ? `${u}/_usersSettings/tokens` : 'https://dev.azure.com';
}

function formOrg() {
  const existing = S.orgs.find((o) => o.id === $('#fId').value);
  const url = normalizeOrgUrl($('#fUrl').value);
  return {
    ...(existing || {}),
    id: $('#fId').value || undefined,
    name: $('#fName').value.trim() || (existing && existing.name) || url.replace(/^https?:\/\/(dev\.azure\.com\/)?/, ''),
    url,
    pat: $('#fPat').value.trim(),
    syncPat: $('#fSyncPat').checked,
    apiVersion: $('#fApi').value,
    defaultProject: $('#fProject').value,
  };
}

async function testOrg() {
  const o = formOrg();
  const out = $('#testResult');
  if (!o.url) { out.textContent = 'Bitte eine URL eingeben.'; out.className = 'status error'; return; }
  out.textContent = 'Teste Verbindung …'; out.className = 'status';
  try {
    const projects = await new AdoClient(o).getProjects();
    out.textContent = `✓ Verbunden – ${projects.length} Projekt(e) gefunden.`; out.className = 'status ok';
    const cur = $('#fProject').value;
    $('#fProject').innerHTML = `<option value="">Zuletzt verwendetes</option>` + projects.map((p) => `<option ${p.name === cur ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
  } catch (e) { out.textContent = e.message; out.className = 'status error'; }
}

async function submitOrg(ev) {
  ev.preventDefault();
  const o = formOrg();
  if (!/^https:\/\//i.test(o.url) && !/^http:\/\/localhost/i.test(o.url)) {
    $('#testResult').textContent = 'Die URL muss mit https:// beginnen.'; $('#testResult').className = 'status error'; return;
  }
  try {
    const id = await saveOrg(o);
    reloadOrgs();
    if (!S.org || S.org.demo || S.org.id === id) S.org = S.orgs.find((x) => x.id === id);
    renderOrgList();
    show('#orgForm', false);
    toast('Organisation gespeichert');
  } catch (e) { $('#testResult').textContent = e.message; $('#testResult').className = 'status error'; }
}

/* =============================================================== Wiring */

function wireUi() {
  $('#btnSettings').addEventListener('click', () => ($('#viewSettings').classList.contains('hidden') ? openSettings() : closeSettings()));
  $('#btnSetup').addEventListener('click', openSettings);
  $('#btnPatFix').addEventListener('click', () => { openSettings(); if (S.org && !S.org.demo) editOrg(S.org); });
  $('#btnBack').addEventListener('click', closeSettings);
  $('#btnAddOrg').addEventListener('click', () => editOrg(null));
  $('#btnCancelOrg').addEventListener('click', () => show('#orgForm', false));
  $('#btnTest').addEventListener('click', testOrg);
  $('#orgForm').addEventListener('submit', submitOrg);
  $('#fUrl').addEventListener('input', updatePatLink);
  $('#orgList').addEventListener('click', async (e) => {
    const ed = e.target.closest('[data-edit]');
    if (ed) return editOrg(S.orgs.find((o) => o.id === ed.dataset.edit));
    const del = e.target.closest('[data-del]');
    if (del) {
      if (del.dataset.confirm !== '1') { del.dataset.confirm = '1'; del.textContent = 'Wirklich löschen?'; return; }
      await deleteOrg(del.dataset.del);
      reloadOrgs(); renderOrgList();
    }
  });

  $('#selOrg').addEventListener('change', (e) => { selectItem(null); selectOrg(e.target.value); });
  $('#selProject').addEventListener('change', (e) => { S.project = e.target.value; selectItem(null); onProjectChanged(); });

  $('#segTarget').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setTarget(b.dataset.v); });
  $('#segContent').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setContent(b.dataset.v); });

  const doSearch = debounce(refreshSearch, 300);
  $('#wiSearch').addEventListener('input', doSearch);
  $('#wiSearch').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const q = $('#wiSearch').value.trim().replace(/^(?:AB)?#/i, '');
      if (/^\d+$/.test(q)) pickById(+q);
      else if (S.lastItems && S.lastItems.length === 1) selectItem(S.lastItems[0]);
    } else if (e.key === 'ArrowDown') { const f = $('#wiList .wi'); if (f) { e.preventDefault(); f.focus(); } }
  });
  $('#wiFilter').addEventListener('change', refreshSearch);
  $('#wiList').addEventListener('click', (e) => {
    const b = e.target.closest('.wi');
    if (b) selectItem((S.lastItems || []).find((w) => w.id === +b.dataset.id));
  });
  $('#wiList').addEventListener('keydown', (e) => {
    const cur = document.activeElement.closest('.wi');
    if (!cur) return;
    if (e.key === 'ArrowDown' && cur.nextElementSibling) { e.preventDefault(); cur.nextElementSibling.focus(); }
    if (e.key === 'ArrowUp') { e.preventDefault(); (cur.previousElementSibling || $('#wiSearch')).focus(); }
  });
  $('#wiSelected').addEventListener('click', (e) => {
    if (e.target.id === 'btnChangeItem') { selectItem(null); $('#wiSearch').focus(); if (!$('#wiList').children.length) refreshSearch(); }
  });

  $('#typeChips').addEventListener('click', (e) => {
    const b = e.target.closest('[data-type]');
    if (b) { S.newType = b.dataset.type; renderTypes(); updateSubmit(); }
  });
  ['#newTitle', '#newParent'].forEach((s) => $(s).addEventListener('input', updateSubmit));

  $('#threadList').addEventListener('change', (e) => {
    const cb = e.target.closest('input[data-i]');
    if (!cb) return;
    const i = +cb.dataset.i;
    if (cb.checked) S.threadSel.add(i); else S.threadSel.delete(i);
    updatePreviewDim(); updateSubmit();
  });
  $('#snippetList').addEventListener('click', (e) => { const b = e.target.closest('[data-remove]'); if (b) removeSnippet(+b.dataset.remove); });
  $('#btnComposeSel').addEventListener('click', captureComposeSelection);

  $('#optOldestFirst').addEventListener('change', (e) => { S.prefs.oldestFirst = e.target.checked; persistPrefs(); renderThreadList(); });
  $('#optHeader').addEventListener('change', (e) => { S.prefs.includeHeader = e.target.checked; persistPrefs(); });
  $('#optImages').addEventListener('change', (e) => { S.prefs.transferImages = e.target.checked; persistPrefs(); });

  const fab = $('#selFab');
  fab.addEventListener('mousedown', (e) => e.preventDefault());
  fab.addEventListener('click', captureSelection);
  const fabUpdate = debounce(updateFab, 120);
  document.addEventListener('selectionchange', fabUpdate);
  $('#preview').addEventListener('scroll', fabUpdate, { passive: true });
  window.addEventListener('scroll', fabUpdate, { passive: true });
  $('#preview').addEventListener('keydown', (e) => { if (e.key === 'Enter' && selectionInPreview()) { e.preventDefault(); captureSelection(); } });

  $('#btnSubmit').addEventListener('click', submit);

  // Links im Outlook-Fenster sauber im Browser öffnen
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a[data-open]');
    if (a && !S.demo) { e.preventDefault(); S.mailApi.openUrl(a.href); }
  });
}

boot().catch((e) => { console.error(e); setStatus('Start fehlgeschlagen: ' + e.message, 'error'); });
