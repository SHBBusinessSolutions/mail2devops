// parse.js – E-Mail-HTML bereinigen, Verlauf in Einzelnachrichten zerlegen, Kommentar-HTML bauen.
// Reine DOM-Logik ohne Office-Abhängigkeit (im Browser testbar).

export const esc = (s) =>
  String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/** "AW: WG: RE: Betreff" → "Betreff" */
export function cleanSubject(s) {
  let t = (s || '').trim();
  const re = /^\s*(re|aw|wg|fw|fwd|sv|vs|antw|tr)\s*(\[\d+\])?\s*:\s*/i;
  while (re.test(t)) t = t.replace(re, '');
  return t.trim();
}

export function b64ToBytes(b64) {
  const bin = atob(String(b64).replace(/\s+/g, ''));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function safeFileName(s, fallback = 'E-Mail') {
  const t = (s || '').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
  return t || fallback;
}

/* ------------------------------------------------------------------ Text */

const BLOCK_RE = /^(P|DIV|TR|LI|H[1-6]|TABLE|BLOCKQUOTE|SECTION|PRE|HR|UL|OL)$/;

/** Text eines Knotens mit Zeilenumbrüchen für Block-Elemente (ohne Layout-Engine). */
export function textOf(node) {
  let out = '';
  (function walk(n) {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) out += c.nodeValue;
      else if (c.nodeType === 1) {
        const tag = c.tagName.toUpperCase();
        if (tag === 'BR') { out += '\n'; continue; }
        if (tag === 'STYLE' || tag === 'SCRIPT') continue;
        if (tag === 'TD' || tag === 'TH') { walk(c); out += '\t'; continue; }
        const block = BLOCK_RE.test(tag);
        if (block && out && !out.endsWith('\n')) out += '\n';
        walk(c);
        if (block && !out.endsWith('\n')) out += '\n';
      }
    }
  })(node);
  return out.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

const stripTags = (s) =>
  s.replace(/<(style|script)[\s\S]*?<\/\1>/gi, '').replace(/<[^>]*>/g, '').replace(/&nbsp;|&#160;/gi, ' ');

/* ---------------------------------------------------------------- Header */

const H = {
  from: /^(?:from|von|de|da|van|od|från|fra)\s*:\s*(.+)$/i,
  date: /^(?:sent|gesendet|date|datum|envoyé|inviato|verzonden|enviado|skickat|sendt)\s*:\s*(.+)$/i,
  to: /^(?:to|an|à|a|aan|para|till|til)\s*:\s*(.+)$/i,
  cc: /^(?:cc|kopie)\s*:\s*(.+)$/i,
  subject: /^(?:subject|betreff|objet|oggetto|onderwerp|asunto|assunto|ämne|emne)\s*:\s*(.+)$/i,
};
const isHeaderLine = (l) => Object.values(H).some((re) => re.test(l.trim()));

export function parsePerson(s) {
  s = (s || '').trim();
  const m = s.match(/^(.*?)\s*[<\[](?:mailto:)?([^<>\[\]\s]+@[^<>\[\]\s]+)[>\]]\s*$/i);
  if (m) return { name: m[1].replace(/^["'\s]+|["'\s]+$/g, '') || m[2], email: m[2] };
  if (/^[^\s@]+@[^\s@]+$/.test(s)) return { name: s, email: s };
  return { name: s, email: '' };
}

export function parseHeaderText(text) {
  const meta = {};
  for (const line of (text || '').split('\n')) {
    const l = line.trim();
    if (!l) continue;
    for (const [k, re] of Object.entries(H)) {
      if (meta[k]) continue;
      const m = l.match(re);
      if (m) { meta[k] = m[1].trim(); break; }
    }
  }
  if (meta.from) meta.from = parsePerson(meta.from);
  if (!meta.from) {
    // Gmail / Apple / Thunderbird: "Am 12.11.2026 um 10:00 schrieb Max <max@x.ch>:" / "On …, Max <max@x> wrote:"
    const one = (text || '').replace(/\s+/g, ' ').trim();
    const m = one.match(/^(?:on|am|le|el|il|op|den)\s+(.+?)\s*(?:schrieb|wrote|a écrit|escribió|ha scritto|schreef|skrev)\s*(.*?)\s*:?\s*$/i);
    if (m) {
      const german = m[2] && m[2].length > 1;
      const who = german ? m[2] : m[1];
      const pm = who.match(/([^,<>]*?)\s*<([^<>\s]+@[^<>\s]+)>/);
      if (pm) {
        meta.from = { name: pm[1].replace(/^["',\s]+|["',\s]+$/g, '') || pm[2], email: pm[2] };
        meta.date = german ? m[1] : m[1].replace(pm[0], '').replace(/[,\s]+$/, '');
      } else {
        meta.from = { name: who.trim(), email: '' };
        if (german) meta.date = m[1];
      }
    }
  }
  return meta;
}

/* -------------------------------------------------------------- Sanitize */

const DROP_SEL =
  'script,style,meta,link,title,head,xml,object,embed,iframe,frame,frameset,form,input,button,select,textarea,svg,math,noscript,template,base,audio,video,canvas,map,area';
const ALLOWED = new Set(
  'a b strong i em u s strike del ins sub sup br hr p div span font blockquote pre code ul ol li table thead tbody tfoot tr td th caption h1 h2 h3 h4 h5 h6 img small dl dt dd'.split(' ')
);
const KEEP_PROPS = new Set(['font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'color', 'background-color', 'background', 'text-align']);
const TABLE_PROPS = /^(border(-top|-bottom|-left|-right)?(-width|-style|-color)?|border-collapse|padding(-top|-bottom|-left|-right)?|vertical-align|width)$/;
const DEFAULT_COLOR = /^(black|windowtext|auto|inherit|initial|#000|#000000|rgb\(\s*0,\s*0,\s*0\s*\)|#212121|#222|#222222|#201f1e|#242424|#323130)$/i;
const WHITE = /^(white|transparent|window|#fff|#ffffff|rgb\(\s*255,\s*255,\s*255\s*\)|none)$/i;

function filterStyle(style, tag) {
  const out = [];
  const isTable = /^(table|td|th)$/.test(tag);
  for (const decl of (style || '').split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    const p = decl.slice(0, i).trim().toLowerCase();
    let v = decl.slice(i + 1).trim().replace(/\s*!important/i, '');
    if (!v || /expression|url\(|javascript:|behavior/i.test(v)) continue;
    if (KEEP_PROPS.has(p)) {
      if (p === 'color' && DEFAULT_COLOR.test(v)) continue;
      if ((p === 'background-color' || p === 'background') && WHITE.test(v)) continue;
      if (p === 'background') { if (/^(#[0-9a-f]{3,8}|rgba?\([^)]*\)|[a-z]+)$/i.test(v)) out.push('background-color:' + v); continue; }
      if (p === 'font-weight' && /^(normal|400)$/i.test(v)) continue;
      if (p === 'font-style' && /^normal$/i.test(v)) continue;
      if (p.startsWith('text-decoration') && /^none$/i.test(v)) continue;
      if (p === 'text-align' && /^(left|start|justify)$/i.test(v)) continue;
      out.push(`${p}:${v}`);
    } else if (isTable && TABLE_PROPS.test(p)) {
      out.push(`${p}:${v.replace(/windowtext/gi, '#8a8886')}`);
    }
  }
  return out.join(';');
}

function unwrap(el) {
  const p = el.parentNode;
  if (!p) return;
  while (el.firstChild) p.insertBefore(el.firstChild, el);
  el.remove();
}

const isBlank = (el) =>
  !el.querySelector('img,table,hr,li') && !(el.textContent || '').replace(/[\s ​﻿]/g, '');

function collapseBlank(root) {
  root.querySelectorAll('span').forEach((s) => { if (isBlank(s) && !s.querySelector('br')) s.remove(); });
  // Mehrere leere Zeilen hintereinander → eine
  for (const el of [...root.querySelectorAll('div')]) {
    if (!el.isConnected || !isBlank(el)) continue;
    let prev = el.previousSibling;
    while (prev && prev.nodeType === 3 && !prev.nodeValue.trim()) prev = prev.previousSibling;
    if (prev && prev.nodeType === 1 && (prev.tagName === 'DIV' && isBlank(prev))) el.remove();
  }
  // Anfang/Ende trimmen
  const trim = (first) => {
    for (;;) {
      const n = first ? root.firstChild : root.lastChild;
      if (!n) break;
      if (n.nodeType === 3 && !n.nodeValue.replace(/[\s ]/g, '')) { n.remove(); continue; }
      if (n.nodeType === 1 && (n.tagName === 'BR' || (n.tagName === 'DIV' && isBlank(n)))) { n.remove(); continue; }
      break;
    }
  };
  trim(true); trim(false);
}

/**
 * Bereinigt Outlook/Word-HTML: entfernt Skripte, MSO-Ballast, Klassen, die meisten Styles.
 * @param {string|Element} input  HTML-String oder (losgelöstes) Element, dessen Kinder bereinigt werden
 * @returns {string} sauberes HTML
 */
export function sanitizeHtml(input) {
  let root;
  if (typeof input === 'string') {
    root = new DOMParser().parseFromString(`<!doctype html><body>${input}</body>`, 'text/html').body;
  } else root = input;
  const d = root.ownerDocument;

  const tw = d.createTreeWalker(root, NodeFilter.SHOW_COMMENT);
  const comments = [];
  while (tw.nextNode()) comments.push(tw.currentNode);
  comments.forEach((c) => c.remove());
  root.querySelectorAll(DROP_SEL).forEach((e) => e.remove());

  for (const el of [...root.querySelectorAll('*')]) {
    if (!el.isConnected && !root.contains(el)) continue;
    const tag = el.tagName.toLowerCase();
    if (tag.includes(':')) { unwrap(el); continue; } // o:p, v:shape, w:… (Word)
    if (!ALLOWED.has(tag)) { unwrap(el); continue; }

    let style = filterStyle(el.getAttribute('style'), tag);
    if (tag === 'font') {
      const c = el.getAttribute('color');
      if (c && !DEFAULT_COLOR.test(c)) style = (style ? style + ';' : '') + 'color:' + c;
    }
    const keep = {};
    if (tag === 'a') {
      const h = (el.getAttribute('href') || '').trim();
      if (/^(https?:|mailto:)/i.test(h)) keep.href = h;
    }
    if (tag === 'img') {
      const src = (el.getAttribute('src') || '').trim();
      if (!/^(https?:|data:image\/|cid:|blob:)/i.test(src)) { el.remove(); continue; }
      const w = parseInt(el.getAttribute('width'), 10), h = parseInt(el.getAttribute('height'), 10);
      if ((w && w <= 2) || (h && h <= 2)) { el.remove(); continue; } // Tracking-Pixel
      keep.src = src;
      for (const a of ['alt', 'width', 'height', 'data-img-key']) if (el.hasAttribute(a)) keep[a] = el.getAttribute(a);
    }
    if (tag === 'td' || tag === 'th') for (const a of ['colspan', 'rowspan']) if (el.hasAttribute(a)) keep[a] = el.getAttribute(a);
    if (tag === 'ol' && el.hasAttribute('start')) keep.start = el.getAttribute('start');
    if (tag === 'table' && el.getAttribute('border') && el.getAttribute('border') !== '0' && !/border/.test(style)) {
      style = (style ? style + ';' : '') + 'border-collapse:collapse';
      keep.border = '1';
    }

    for (const a of [...el.attributes]) el.removeAttribute(a.name);
    for (const [k, v] of Object.entries(keep)) el.setAttribute(k, v);
    if (style) el.setAttribute('style', style);

    if (tag === 'a') {
      if (!keep.href) { unwrap(el); continue; }
      el.setAttribute('target', '_blank');
      el.setAttribute('rel', 'noopener noreferrer');
    } else if (tag === 'font' || tag === 'span') {
      if (!style) { unwrap(el); continue; }
      if (tag === 'font') {
        const span = d.createElement('span');
        span.setAttribute('style', style);
        while (el.firstChild) span.appendChild(el.firstChild);
        el.replaceWith(span);
      }
    } else if (tag === 'p') {
      // Outlook-Absätze haben margin:0 – als <div> bleibt der Abstand kompakt wie in Outlook
      const div = d.createElement('div');
      if (style) div.setAttribute('style', style);
      while (el.firstChild) div.appendChild(el.firstChild);
      el.replaceWith(div);
    }
  }
  collapseBlank(root);
  return root.innerHTML.trim();
}

/* ---------------------------------------------------------- Thread split */

const OUTLOOK_IMG = /attachments\.office\.net|outlook\.office(365)?\.com|outlook\.live\.com|\/owa\//i;
const RE_TAGS = [
  /<div\b[^>]*\bid=["']?appendonsend\b[^>]*>/gi,
  /<div\b[^>]*\bid=["']?divRplyFwdMsg\b[^>]*>/gi,
  /<div\b[^>]*border-top:\s*solid\s+#[0-9a-f]{3,6}\s+1(?:\.0)?pt[^>]*>/gi,
  /<div\b[^>]*class=["'][^"']*\bgmail_quote\b[^"']*["'][^>]*>/gi,
  /<div\b[^>]*class=["'][^"']*\bmoz-cite-prefix\b[^"']*["'][^>]*>/gi,
  /<blockquote\b[^>]*type=["']?cite\b[^>]*>/gi,
];
const RE_HR = /<hr\b[^>]*>/gi;
const RE_ORIG = /-{3,}\s*(?:Original Message|Ursprüngliche Nachricht|Originalnachricht|Message d'origine|Messaggio originale|Oorspronkelijk bericht|Mensaje original)\s*-{3,}/gi;
const RE_FROMWORD = /(?:^|\s)(?:from|von|de|da|van|od)\s*:/i;

function findCuts(html) {
  const idx = [];
  for (const re of RE_TAGS) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(html))) idx.push(m.index);
  }
  let m;
  RE_HR.lastIndex = 0;
  while ((m = RE_HR.exec(html))) {
    const after = stripTags(html.slice(m.index, m.index + 2000)).slice(0, 300);
    if (RE_FROMWORD.test(after)) idx.push(m.index);
  }
  RE_ORIG.lastIndex = 0;
  while ((m = RE_ORIG.exec(html))) {
    // an den Anfang des umschliessenden Absatzes springen
    const from = Math.max(0, m.index - 400);
    const seg = html.slice(from, m.index);
    const k = Math.max(seg.lastIndexOf('<p'), seg.lastIndexOf('<div'));
    idx.push(k >= 0 ? from + k : m.index);
  }
  idx.sort((a, b) => a - b);
  const out = [];
  for (const i of idx) {
    const last = out[out.length - 1];
    if (last !== undefined && !stripTags(html.slice(last, i)).replace(/[\s ]/g, '')) continue; // Cluster zusammenfassen
    out.push(i);
  }
  return out.filter((i) => i > 0);
}

function findHeaderEl(body) {
  const el = body.querySelector('#divRplyFwdMsg, .gmail_attr, .moz-cite-prefix');
  if (el) return el;
  return [...body.querySelectorAll('div[style]')].find((d) => /border-top:\s*solid/i.test(d.getAttribute('style'))) || null;
}

function removeLeading(body, doc) {
  body.querySelectorAll('#appendonsend').forEach((e) => e.remove());
  const hr = body.querySelector('hr');
  if (hr) {
    const r = doc.createRange();
    r.setStart(body, 0);
    r.setEndBefore(hr);
    if (!r.toString().replace(/[\s ]/g, '')) hr.remove();
  }
  // Text "-----Ursprüngliche Nachricht-----" entfernen
  const tw = doc.createTreeWalker(body, NodeFilter.SHOW_TEXT);
  const hits = [];
  while (tw.nextNode()) { RE_ORIG.lastIndex = 0; if (RE_ORIG.test(tw.currentNode.nodeValue)) hits.push(tw.currentNode); }
  hits.forEach((t) => { RE_ORIG.lastIndex = 0; t.nodeValue = t.nodeValue.replace(RE_ORIG, ''); });
}

/** Header-Zeilen (Von:/An:/Betreff:) am Anfang entfernen, wenn kein Header-Element erkannt wurde. */
function removeHeaderLines(body) {
  for (let n = 0; n < 12; n++) {
    let el = body.firstElementChild;
    while (el && el.children.length === 1 && !textOf(el).includes('\n') && el.firstElementChild.tagName === 'DIV') el = el.firstElementChild;
    if (!el) return;
    const lines = textOf(el).split('\n').map((l) => l.trim()).filter(Boolean);
    if (!lines.length) { el.remove(); continue; }
    if (lines.every(isHeaderLine)) { el.remove(); continue; }
    return;
  }
}

/**
 * Zerlegt einen E-Mail-Body in einzelne Nachrichten (neueste zuerst, wie im Body).
 * @returns {{fullHtml:string, messages:Array<{index:number, meta:object, html:string, text:string, excerpt:string}>}}
 */
export function splitThread(bodyHtml, mailMeta = {}) {
  const doc = new DOMParser().parseFromString(bodyHtml || '', 'text/html');
  let attN = 0;
  doc.querySelectorAll('img').forEach((img) => {
    const src = img.getAttribute('src') || '';
    if (/^cid:/i.test(src)) img.setAttribute('data-img-key', 'cid:' + src.slice(4));
    else if (OUTLOOK_IMG.test(src)) img.setAttribute('data-img-key', 'att:' + attN++);
  });
  const html = doc.body ? doc.body.innerHTML : String(bodyHtml || '');
  const fullHtml = sanitizeHtml(html);

  const cuts = findCuts(html);
  const parts = [];
  let prev = 0;
  for (const c of cuts) if (c > prev) { parts.push(html.slice(prev, c)); prev = c; }
  parts.push(html.slice(prev));

  const messages = [];
  let pendingMeta = null;
  parts.forEach((raw, i) => {
    const pdoc = new DOMParser().parseFromString(`<!doctype html><body>${raw}</body>`, 'text/html');
    const body = pdoc.body;
    let meta;
    if (i === 0) {
      meta = { from: mailMeta.from, to: mailMeta.to, cc: mailMeta.cc, date: mailMeta.date, subject: mailMeta.subject };
    } else if (pendingMeta) {
      meta = pendingMeta;
      pendingMeta = null;
      removeLeading(body, pdoc);
    } else {
      removeLeading(body, pdoc);
      const hdr = findHeaderEl(body);
      if (hdr) {
        meta = parseHeaderText(textOf(hdr));
        hdr.remove();
      } else {
        meta = parseHeaderText(textOf(body).slice(0, 1200));
        removeHeaderLines(body);
      }
    }
    const clean = sanitizeHtml(body);
    const tdoc = new DOMParser().parseFromString(`<body>${clean}</body>`, 'text/html');
    const text = textOf(tdoc.body);
    if (!text && !/<img/i.test(clean)) {
      if (i > 0 && meta && meta.from) pendingMeta = meta; // nur Zitat-Kopf → für nächsten Teil merken
      return;
    }
    messages.push({ index: messages.length, meta: meta || {}, html: clean, text, excerpt: text.replace(/\s+/g, ' ').slice(0, 160) });
  });

  if (!messages.length) {
    messages.push({ index: 0, meta: { ...mailMeta }, html: fullHtml, text: '', excerpt: '' });
  }
  return { fullHtml, messages };
}

/* ------------------------------------------------------------- Formatting */

export function personName(p) {
  if (!p) return '';
  if (typeof p === 'string') return p;
  return p.name || p.email || '';
}

function personHtml(p) {
  if (!p) return '';
  if (typeof p === 'string') return esc(p);
  const name = esc(p.name || p.email || '');
  if (p.email && p.email !== p.name) return `${name} &lt;<a href="mailto:${esc(p.email)}">${esc(p.email)}</a>&gt;`;
  return name;
}

function peopleHtml(list) {
  if (!list) return '';
  if (typeof list === 'string') return esc(list);
  const max = 8;
  const s = list.slice(0, max).map(personHtml).join(', ');
  return list.length > max ? `${s} <i>(+${list.length - max} weitere)</i>` : s;
}

export function fmtDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    if (isNaN(d)) return '';
    return d.toLocaleString('de-CH', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  }
  return String(d);
}

/** Kopfblock (Von/An/Datum) für einen Kommentar. */
export function headerBlock(meta = {}, { subject = '', title = '📧' } = {}) {
  const rows = [];
  if (meta.from) rows.push(['Von', personHtml(meta.from)]);
  const to = peopleHtml(meta.to);
  if (to) rows.push(['An', to]);
  const cc = peopleHtml(meta.cc);
  if (cc) rows.push(['Cc', cc]);
  if (meta.date) rows.push(['Datum', esc(fmtDate(meta.date))]);
  const subj = meta.subject || subject || '(ohne Betreff)';
  return (
    `<div style="border-left:3px solid #0078d4;padding:2px 0 2px 10px;margin:0 0 10px 0;color:#605e5c;font-size:12px">` +
    `<div style="font-size:13px;color:#323130"><b>${title} ${esc(subj)}</b></div>` +
    rows.map(([k, v]) => `<div><b>${k}:</b> ${v}</div>`).join('') +
    `</div>`
  );
}

export function noteBlock(note) {
  if (!note) return '';
  return `<div style="margin:0 0 10px 0"><b>💬 Notiz:</b> ${esc(note).replace(/\n/g, '<br>')}</div>`;
}

export function snippetBlock(sn) {
  const who = sn.meta ? [personName(sn.meta.from), fmtDate(sn.meta.date)].filter(Boolean).join(' · ') : '';
  return (
    (who ? `<div style="margin-top:8px;color:#605e5c;font-size:12px">${esc(who)}:</div>` : '') +
    `<blockquote style="border-left:3px solid #c8c6c4;margin:2px 0 8px 0;padding:2px 10px">${sn.html}</blockquote>`
  );
}
