// ado.js – schlanker Azure DevOps REST-Client (Work Items, Kommentare, Anhänge).
// Auth: Personal Access Token (Basic). Azure DevOps Services erlaubt CORS mit PAT.

import { normalizeOrgUrl } from './config.js';

export class AdoError extends Error {
  constructor(message, status = 0, details = null) {
    super(message);
    this.status = status;
    this.details = details;
  }
}

const FIELDS = ['System.Id', 'System.Title', 'System.WorkItemType', 'System.State', 'System.AssignedTo', 'System.TeamProject'];

const TYPE_COLORS = {
  'product backlog item': '#009CCC', 'user story': '#009CCC', requirement: '#009CCC', issue: '#B4009E',
  task: '#F2CB1D', bug: '#CC293D', feature: '#773B93', epic: '#FF7B00', impediment: '#B4009E', 'test case': '#004B50',
};
export const typeColor = (t) => TYPE_COLORS[(t || '').toLowerCase()] || '#8a8886';

function mapWi(d) {
  const f = d.fields || {};
  const at = f['System.AssignedTo'];
  return {
    id: d.id,
    title: f['System.Title'] || '',
    type: f['System.WorkItemType'] || '',
    state: f['System.State'] || '',
    assignedTo: at ? (typeof at === 'string' ? at : at.displayName) : '',
    project: f['System.TeamProject'] || '',
  };
}

export class AdoClient {
  constructor(org) {
    this.org = org;
    this.base = normalizeOrgUrl(org.url);
    this.api = org.apiVersion || '7.1';
  }

  get commentsApi() {
    return parseFloat(this.api) >= 7.1 ? `${this.api}-preview.4` : `${this.api}-preview.3`;
  }

  webUrl(project, id) {
    return `${this.base}/${encodeURIComponent(project)}/_workitems/edit/${id}`;
  }

  async request(path, { method = 'GET', body, contentType = 'application/json', query = {} } = {}) {
    if (!this.org.pat) throw new AdoError('Kein Personal Access Token hinterlegt – bitte in den Einstellungen ergänzen.', 401);
    const url = new URL(this.base + path);
    const q = { 'api-version': this.api, ...query };
    for (const [k, v] of Object.entries(q)) if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, v);

    const headers = { Authorization: 'Basic ' + btoa(':' + this.org.pat), Accept: 'application/json' };
    let payload;
    if (body !== undefined) {
      headers['Content-Type'] = contentType;
      payload = contentType.includes('json') ? JSON.stringify(body) : body;
    }
    let res;
    try {
      res = await fetch(url.toString(), { method, headers, body: payload });
    } catch {
      throw new AdoError(`Keine Verbindung zu ${this.base} (Netzwerk oder CORS). Stimmt die Organisations-URL?`, 0);
    }
    if (res.status === 401 || res.status === 203) {
      throw new AdoError('Anmeldung fehlgeschlagen – PAT ungültig, abgelaufen oder ohne Zugriff auf diese Organisation.', 401);
    }
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch {
      if (res.ok) throw new AdoError('Unerwartete Antwort (kein JSON) – stimmt die Organisations-URL?', res.status);
    }
    if (!res.ok) {
      const msg = (data && data.message) || `HTTP ${res.status}`;
      const prefix = res.status === 403 ? 'Keine Berechtigung: ' : res.status === 404 ? 'Nicht gefunden: ' : '';
      throw new AdoError(prefix + msg, res.status, data);
    }
    return data;
  }

  testConnection() {
    return this.request('/_apis/projects', { query: { $top: 1 } });
  }

  async getProjects() {
    const d = await this.request('/_apis/projects', { query: { $top: 500, stateFilter: 'wellFormed' } });
    return (d.value || []).map((p) => ({ id: p.id, name: p.name })).sort((a, b) => a.name.localeCompare(b.name, 'de'));
  }

  /** Relevante Typen in sinnvoller Reihenfolge (PBI/User Story, Task, Bug, Feature, Epic) – prozessunabhängig. */
  async getWorkItemTypes(project) {
    const p = encodeURIComponent(project);
    try {
      const cats = await this.request(`/${p}/_apis/wit/workitemtypecategories`);
      const order = ['Microsoft.RequirementCategory', 'Microsoft.TaskCategory', 'Microsoft.BugCategory', 'Microsoft.FeatureCategory', 'Microsoft.EpicCategory'];
      const names = [];
      for (const ref of order) {
        const c = (cats.value || []).find((x) => x.referenceName === ref);
        if (!c) continue;
        const list = c.workItemTypes && c.workItemTypes.length ? c.workItemTypes : c.defaultWorkItemType ? [c.defaultWorkItemType] : [];
        for (const t of list) if (t && t.name && !names.includes(t.name)) names.push(t.name);
      }
      if (names.length) return names;
    } catch (e) {
      if (e.status === 401) throw e;
    }
    const t = await this.request(`/${p}/_apis/wit/workitemtypes`);
    return (t.value || []).filter((x) => !x.isDisabled).map((x) => x.name);
  }

  async queryWorkItems(project, { mode = 'recent', text = '' } = {}) {
    const conds = ['[System.TeamProject] = @project', "[System.State] <> 'Removed'"];
    if (mode === 'mine') conds.push('[System.AssignedTo] = @Me');
    if (mode === 'created') conds.push('[System.CreatedBy] = @Me');
    const t = (text || '').trim();
    if (t) conds.push(`[System.Title] CONTAINS '${t.replace(/'/g, "''")}'`);
    const wiql = `SELECT [System.Id] FROM WorkItems WHERE ${conds.join(' AND ')} ORDER BY [System.ChangedDate] DESC`;
    const res = await this.request(`/${encodeURIComponent(project)}/_apis/wit/wiql`, { method: 'POST', body: { query: wiql }, query: { $top: 50 } });
    const ids = (res.workItems || []).map((w) => w.id).slice(0, 50);
    return ids.length ? this.getWorkItems(ids) : [];
  }

  async getWorkItems(ids) {
    const d = await this.request('/_apis/wit/workitems', { query: { ids: ids.join(','), fields: FIELDS.join(','), errorPolicy: 'omit' } });
    return (d.value || []).filter(Boolean).map(mapWi);
  }

  async getWorkItem(id) {
    return mapWi(await this.request(`/_apis/wit/workitems/${id}`, { query: { fields: FIELDS.join(',') } }));
  }

  async createWorkItem(project, type, { title, descriptionHtml, tags, assignedTo, parentId } = {}) {
    const ops = [{ op: 'add', path: '/fields/System.Title', value: title }];
    if (descriptionHtml) {
      const field = /^bug$/i.test(type) ? 'Microsoft.VSTS.TCM.ReproSteps' : 'System.Description';
      ops.push({ op: 'add', path: `/fields/${field}`, value: descriptionHtml });
    }
    if (tags) ops.push({ op: 'add', path: '/fields/System.Tags', value: tags });
    if (assignedTo) ops.push({ op: 'add', path: '/fields/System.AssignedTo', value: assignedTo });
    if (parentId) {
      ops.push({ op: 'add', path: '/relations/-', value: { rel: 'System.LinkTypes.Hierarchy-Reverse', url: `${this.base}/_apis/wit/workItems/${parentId}` } });
    }
    const d = await this.request(`/${encodeURIComponent(project)}/_apis/wit/workitems/$${encodeURIComponent(type)}`, {
      method: 'POST', contentType: 'application/json-patch+json', body: ops,
    });
    return mapWi(d);
  }

  addComment(project, id, html) {
    return this.request(`/${encodeURIComponent(project)}/_apis/wit/workItems/${id}/comments`, {
      method: 'POST', body: { text: html }, query: { 'api-version': this.commentsApi },
    });
  }

  /** @param {Uint8Array} bytes */
  uploadAttachment(fileName, bytes) {
    return this.request('/_apis/wit/attachments', {
      method: 'POST', contentType: 'application/octet-stream', body: bytes, query: { fileName },
    });
  }

  attachFiles(id, files, comment = '') {
    const ops = files.map((f) => ({
      op: 'add', path: '/relations/-', value: { rel: 'AttachedFile', url: f.url, attributes: { comment: comment || f.name } },
    }));
    return this.request(`/_apis/wit/workitems/${id}`, { method: 'PATCH', contentType: 'application/json-patch+json', body: ops });
  }
}
