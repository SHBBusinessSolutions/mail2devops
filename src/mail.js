// mail.js – Zugriff auf die aktuelle E-Mail über Office.js (Lesen & Verfassen).

function p(invoke) {
  return new Promise((resolve, reject) =>
    invoke((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error((r.error && r.error.message) || 'Office-Fehler'))));
}

export const supports = (v) => {
  try { return Office.context.requirements.isSetSupported('Mailbox', v); } catch { return false; }
};

const person = (x) => (x ? { name: x.displayName || x.emailAddress || '', email: x.emailAddress || '' } : null);
const mapAtt = (a) => ({ id: a.id, name: a.name, contentType: a.contentType, size: a.size, isInline: !!a.isInline, type: a.attachmentType });

export const officeMail = {
  kind: 'office',

  hasItem: () => !!(Office.context.mailbox && Office.context.mailbox.item),

  async readCurrentMail() {
    const item = Office.context.mailbox.item;
    if (!item) throw new Error('Keine E-Mail ausgewählt.');
    const compose = typeof item.subject !== 'string';
    const ctx = { mode: compose ? 'compose' : 'read', itemId: item.itemId || null, conversationId: item.conversationId || null };
    if (!compose) {
      ctx.subject = item.subject || '';
      ctx.from = person(item.from);
      ctx.to = (item.to || []).map(person);
      ctx.cc = (item.cc || []).map(person);
      ctx.date = item.dateTimeCreated ? new Date(item.dateTimeCreated) : null;
      ctx.attachments = (item.attachments || []).map(mapAtt);
    } else {
      ctx.subject = await p((cb) => item.subject.getAsync(cb));
      ctx.to = ((await p((cb) => item.to.getAsync(cb))) || []).map(person);
      ctx.cc = ((await p((cb) => item.cc.getAsync(cb))) || []).map(person);
      ctx.from = person(Office.context.mailbox.userProfile);
      ctx.date = new Date();
      ctx.attachments = supports('1.8') ? ((await p((cb) => item.getAttachmentsAsync(cb))) || []).map(mapAtt) : [];
    }
    ctx.bodyHtml = await p((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
    return ctx;
  },

  canGetAttachments: () => supports('1.8'),
  getAttachmentContent: (id) => p((cb) => Office.context.mailbox.item.getAttachmentContentAsync(id, cb)),

  canGetEml: () => supports('1.14') && typeof Office.context.mailbox.item.getAsFileAsync === 'function',
  getEmlBase64: () => p((cb) => Office.context.mailbox.item.getAsFileAsync(cb)),

  canComposeSelection: () => typeof Office.context.mailbox.item.getSelectedDataAsync === 'function',
  async getComposeSelection() {
    const v = await p((cb) => Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Html, cb));
    return (v && v.data) || '';
  },

  onItemChanged(handler) {
    if (supports('1.5')) Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler);
  },

  userEmail: () => (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || '',
};
