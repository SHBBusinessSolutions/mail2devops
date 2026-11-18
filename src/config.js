// config.js – Organisationen, PATs und Einstellungen speichern.
// Organisationen + Einstellungen: Office roamingSettings (im Postfach, geräteübergreifend).
// PATs: standardmässig nur lokal (localStorage) – optional geräteübergreifend.
// Ausserhalb von Outlook (Demo/Browser) wird alles in localStorage gespeichert.

const ORGS_KEY = 'm2d_orgs';
const PREFS_KEY = 'm2d_prefs';
const PAT_PREFIX = 'm2d_pat_';
const LINKS_KEY = 'm2d_links';

const hasRoaming = () => !!(window.Office && Office.context && Office.context.roamingSettings);

function lsGet(k) {
  try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : null; } catch { return null; }
}
function lsSet(k, v) {
  try { if (v === null || v === undefined) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch { /* ignore */ }
}
function rsGet(k) {
  if (!hasRoaming()) return lsGet(k);
  const v = Office.context.roamingSettings.get(k);
  return v === undefined ? null : v;
}
function rsSet(k, v) {
  if (!hasRoaming()) { lsSet(k, v); return; }
  if (v === null || v === undefined) Office.context.roamingSettings.remove(k);
  else Office.context.roamingSettings.set(k, v);
}
function rsSave() {
  return new Promise((resolve, reject) => {
    if (!hasRoaming()) return resolve();
    Office.context.roamingSettings.saveAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message || 'Einstellungen konnten nicht gespeichert werden')));
  });
}

/** "contoso" → https://dev.azure.com/contoso ; entfernt abschliessende Slashes */
export function normalizeOrgUrl(u) {
  u = (u || '').trim();
  if (!u) return '';
  if (!/^https?:\/\//i.test(u)) {
    if (!u.includes('.') && !u.includes('/')) u = 'https://dev.azure.com/' + u;
    else u = 'https://' + u;
  }
  return u.replace(/\/+$/, '');
}

const slug = (s) => (s || '').toLowerCase().replace(/^https?:\/\//, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');

let adminConfig = { organizations: [], defaults: {} };

/** Lädt config.json (vom Admin vorkonfigurierte Organisationen/Standards). Optional. */
export async function loadAdminConfig() {
  try {
    const r = await fetch(new URL('../config.json', location.href), { cache: 'no-store' });
    if (r.ok) adminConfig = await r.json();
  } catch { /* keine Admin-Konfiguration */ }
  adminConfig.organizations = Array.isArray(adminConfig.organizations) ? adminConfig.organizations : [];
  adminConfig.defaults = adminConfig.defaults || {};
  return adminConfig;
}

/** Alle Organisationen (verwaltete aus config.json + eigene), inkl. PAT. */
export function getOrgs() {
  const user = rsGet(ORGS_KEY) || [];
  const byId = new Map(user.map((o) => [o.id, o]));
  const result = [];
  for (const a of adminConfig.organizations) {
    if (!a || !a.url) continue;
    const id = a.id || 'admin_' + slug(a.url);
    const u = byId.get(id) || {};
    byId.delete(id);
    result.push({ apiVersion: '7.1', ...a, ...u, id, name: a.name || u.name || a.url, url: normalizeOrgUrl(a.url), managed: true });
  }
  for (const u of byId.values()) result.push({ apiVersion: '7.1', ...u, url: normalizeOrgUrl(u.url), managed: false });
  for (const o of result) if (!o.syncPat) o.pat = lsGet(PAT_PREFIX + o.id) || '';
  return result;
}

export async function saveOrg(org) {
  const list = rsGet(ORGS_KEY) || [];
  const id = org.id || 'org_' + slug(org.url) + '_' + Date.now().toString(36);
  const entry = {
    id,
    name: org.name,
    url: normalizeOrgUrl(org.url),
    apiVersion: org.apiVersion || '7.1',
    defaultProject: org.defaultProject || '',
    syncPat: !!org.syncPat,
  };
  if (org.syncPat) { entry.pat = org.pat || ''; lsSet(PAT_PREFIX + id, null); }
  else lsSet(PAT_PREFIX + id, org.pat || null);
  const i = list.findIndex((o) => o.id === id);
  if (i >= 0) list[i] = entry; else list.push(entry);
  rsSet(ORGS_KEY, list);
  await rsSave();
  return id;
}

export async function deleteOrg(id) {
  const list = (rsGet(ORGS_KEY) || []).filter((o) => o.id !== id);
  rsSet(ORGS_KEY, list);
  lsSet(PAT_PREFIX + id, null);
  await rsSave();
}

const DEFAULT_PREFS = { includeHeader: true, oldestFirst: true, transferImages: true, defaultTags: 'E-Mail', defaultContent: 'full', lastProject: {}, lastType: {} };

export function getPrefs() {
  return { ...DEFAULT_PREFS, ...(adminConfig.defaults || {}), ...(rsGet(PREFS_KEY) || {}) };
}

let prefsTimer = null;
export function savePrefs(prefs) {
  rsSet(PREFS_KEY, prefs);
  clearTimeout(prefsTimer);
  prefsTimer = setTimeout(() => rsSave().catch((e) => console.warn(e)), 600);
}

/* Verknüpfung Konversation → Work Item (lokal, max. 300 Einträge) */
export function getLink(conversationId) {
  if (!conversationId) return null;
  return (lsGet(LINKS_KEY) || {})[conversationId] || null;
}
export function setLink(conversationId, link) {
  if (!conversationId) return;
  const all = lsGet(LINKS_KEY) || {};
  all[conversationId] = { ...link, ts: Date.now() };
  const keys = Object.keys(all);
  if (keys.length > 300) keys.sort((a, b) => all[a].ts - all[b].ts).slice(0, keys.length - 300).forEach((k) => delete all[k]);
  lsSet(LINKS_KEY, all);
}
