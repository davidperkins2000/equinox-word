/*
 * Eureka IP – Equinox for Word: the task pane (feature 13.7)
 *
 * Talks to the Equinox Apps Script web app (Word.gs, features 13.1–13.6) with POST text/plain JSON
 * { token, action, … }. The contract is in project doc claude/equinox-word-api.md.
 *
 * The web app address and token are typed into Settings once and kept in this Mac's Word storage
 * (localStorage for this add-in). Nothing secret is in this file, so it can be hosted publicly.
 *
 * Documents are edited in the matters' Google Drive folders (Google Drive for desktop), so the document's
 * path tells the web app which matter it is. Nothing is deleted or overwritten in Equinox or Drive.
 */
'use strict';
const PANE_VERSION = 'w0.9.1';

const KEY = { url: 'eqxw.url', token: 'eqxw.token', img: 'eqxw.img', auto: 'eqxw.auto', pending: 'eqxw.pending' };
const store = {
  get: function (k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage blocked: kept for this session only */ } mem[k] = v; }
};
const mem = {};
const val = function (k) { return store.get(k) || mem[k] || ''; };

const $ = function (id) { return document.getElementById(id); };
const esc = function (t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
const state = { doc: { name: '', path: '' }, codes: [], matters: [], code: '', matter: null, docs: null,
                letter: null, templates: null, recipient: '', intent: '' };

// ---------- the web app ----------
async function call(action, data) {
  const url = val(KEY.url).trim(), token = val(KEY.token).trim();
  if (!url || !token) { openSettings('Type the web app address and token first.'); throw new Error('Settings needed'); }
  let res;
  try {
    res = await fetch(url, { method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // a JSON type would trigger a pre-flight Apps Script can't answer
      body: JSON.stringify(Object.assign({ token: token, action: action }, data || {})) });
  } catch (e) {
    throw new Error('Could not reach the Equinox web app. Check the address in Settings and the internet connection.');
  }
  const text = await res.text();
  let out;
  try { out = JSON.parse(text); } catch (e) {
    throw new Error(res.ok
      ? 'The address answered, but not as Equinox. Check it ends in /exec and the deployment\'s access is "Anyone".'
      : 'The web app answered ' + res.status + '.');
  }
  if (out.version) $('ver').textContent = PANE_VERSION + ' · script ' + out.version;
  if (!out.ok) {
    if (out.error === 'not authorised') openSettings('The token was refused. Copy it again from step24_wordSetup\'s log.');
    throw new Error(out.error || 'failed');
  }
  return out;
}

// ---------- small helpers ----------
function say(el, text, kind) { el.className = (el.className.replace(/\b(bad|good)\b/g, '').trim() + ' ' + (kind || '')).trim(); el.textContent = text || ''; }
async function busy(btn, fn) {
  btn.disabled = true; btn.classList.add('busy');
  try { return await fn(); } finally { btn.disabled = false; btn.classList.remove('busy'); }
}
function openLink(url) {
  if (!url) return;
  try { if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; } } catch (e) { /* fall through */ }
  window.open(url, '_blank');
}
function linkHtml(n) { return !n ? '' : n.url ? '<a data-href="' + esc(n.url) + '">' + esc(n.text) + '</a>' : esc(n.text); }
function size(b) { return !b ? '' : b < 1024 ? b + ' B' : b < 1048576 ? Math.round(b / 1024) + ' KB' : (b / 1048576).toFixed(1) + ' MB'; }

// The open document's name and path. On the Mac this is a file path (Google Drive for desktop puts the
// matters' folders under ~/Library/CloudStorage/GoogleDrive-…); a document not yet saved has none.
function docInfo() {
  let u = '';
  try { u = (Office.context.document && Office.context.document.url) || ''; } catch (e) { u = ''; }
  try { if (/^file:/i.test(u)) u = decodeURIComponent(u.replace(/^file:\/\/(localhost)?/i, '')); else if (/%[0-9a-f]{2}/i.test(u)) u = decodeURIComponent(u); } catch (e) { /* keep as is */ }
  const name = u.split(/[\\/]/).filter(Boolean).pop() || '';
  return { name: name, path: u };
}

// ---------- settings ----------
function openSettings(why) {
  $('setUrl').value = val(KEY.url); $('setToken').value = val(KEY.token); $('setImg').checked = val(KEY.img) === '1'; $('setAuto').checked = autoOn();
  say($('settingsMsg'), why || '', why ? 'bad' : '');
  $('settings').hidden = false; $('main').hidden = true;
}
function closeSettings() { $('settings').hidden = true; $('main').hidden = false; }
async function saveSettings() {
  const url = $('setUrl').value.trim(), token = $('setToken').value.trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec\/?$/.test(url)) { say($('settingsMsg'), 'The address should start https://script.google.com/ and end in /exec.', 'bad'); return; }
  if (token.length < 32) { say($('settingsMsg'), 'That token looks too short. Copy the whole line from the log.', 'bad'); return; }
  store.set(KEY.url, url); store.set(KEY.token, token); store.set(KEY.img, $('setImg').checked ? '1' : ''); store.set(KEY.auto, $('setAuto').checked ? '1' : '0');
  say($('settingsMsg'), 'Testing…');
  try {
    const out = await call('ping');
    say($('settingsMsg'), 'Connected to Equinox ' + (out.version || '') + '.', 'good');
    setTimeout(function () { closeSettings(); start(); }, 700);
  } catch (e) { say($('settingsMsg'), e.message, 'bad'); }
}

// ---------- finding the matter ----------
async function start() {
  state.doc = docInfo();
  $('docLine').textContent = state.doc.name ? state.doc.name : 'This document isn\'t saved yet. Save it in the matter\'s Drive folder, or type a matter below.';
  if (!val(KEY.url) || !val(KEY.token)) { openSettings('Welcome. Type the web app address and token to connect.'); return; }
  if (!state.doc.name) {
    // a letter just opened from a template (New letter), or an unsaved document already given a matter
    const p = pendingLetter(), marked = docMatter();
    const code = (p && p.code) || marked;
    if (code) {
      if (p) { state.recipient = p.recipient || ''; state.recipientChosen = !!p.recipient; state.emailTpl = p.email || ''; if (p.subject) setDocSubject(p.subject); if (!state.intent) state.intent = p.html ? 'build' : 'fill'; }
      $('docLine').textContent = p ? 'New letter for ' + code + '. Save it in the matter\'s Drive folder when you\'re ready.' : 'Unsaved document for ' + code + '.';
      showMatches({ codes: [code], matters: [{ code: code, title: '' }], how: 'typed' });
    }
    return;
  }
  say($('msg'), 'Finding the matter…');
  try {
    const out = await call('lookup', { name: state.doc.name, path: state.doc.path });
    const marked = docMatter();   // a letter saved outside the matter's folder still knows its matter
    if (!(out.codes || []).length && marked) showMatches({ codes: [marked], matters: [{ code: marked, title: '' }], how: 'typed' });
    else showMatches(out);
  } catch (e) { say($('msg'), e.message, 'bad'); }
}

async function typed(ev) {
  ev.preventDefault();
  const text = $('codeBox').value.trim();
  if (!text) return;
  say($('msg'), 'Looking up…');
  try { showMatches(await call('lookup', { text: text })); } catch (e) { say($('msg'), e.message, 'bad'); }
}

function showMatches(out) {
  state.codes = out.codes || []; state.matters = out.matters || [];
  const how = out.how && out.how !== 'typed' && out.how !== 'nothing found' ? 'Matched by ' + out.how + '.' : '';
  if (!state.codes.length) {
    say($('msg'), out.how === 'typed' ? 'No matter found for that.' : 'No matter found for this document. Type its code, a reference or an official number.');
    $('matches').innerHTML = ''; return;
  }
  say($('msg'), how);
  if (out.how !== 'typed' && autoOn()) markAutoOpen();
  $('matches').innerHTML = state.matters.map(function (m) {
    return '<button class="pill" data-code="' + esc(m.code) + '" title="' + esc(m.title) + '">' + esc(m.code) + '</button>';
  }).join('');
  select(state.codes[0]);
}

async function select(code) {
  state.code = code; state.matter = null; state.docs = null;
  Array.prototype.forEach.call(document.querySelectorAll('.pill'), function (b) { b.classList.toggle('on', b.getAttribute('data-code') === code); });
  ['fileOut', 'sumOut', 'coOut'].forEach(function (id) { $(id).innerHTML = ''; });
  $('docs').hidden = true;
  const others = state.codes.filter(function (c) { return c !== code; });
  $('fileAllWrap').hidden = !others.length; $('fileAll').checked = false; $('fileAllN').textContent = others.join(', ');
  $('matter').hidden = false; $('matter').innerHTML = '<span class="muted">Loading ' + esc(code) + '…</span>';
  try {
    const out = await call('matter', { code: code, image: val(KEY.img) === '1' });
    if (state.code !== code) return;     // another matter was chosen meanwhile
    state.matter = out; renderMatter(out); $('actions').hidden = false;
    loadLetter(code).catch(function () {});
    runIntent();
  } catch (e) { $('matter').innerHTML = '<span class="bad">' + esc(e.message) + '</span>'; }
}

function renderMatter(out) {
  const f = out.facts || {}, l = out.links || {};
  const h = [];
  h.push('<div class="code">' + esc(out.code) + '</div>');
  if (f.title) h.push('<div class="title">' + esc(f.title) + '</div>');
  if ((f.applicants || []).length) h.push('<div>' + esc(f.applicants.join('; ')) + '</div>');
  if ((f.inventors || []).length) h.push('<div class="muted">' + esc(f.inventorLabel || 'Inventor') + (f.inventors.length > 1 ? 's' : '') + ': ' + esc(f.inventors.join('; ')) + '</div>');
  if (f.status) h.push('<div>Status: ' + esc(f.status) + '</div>');
  if (f.fatal) h.push('<div class="fatal"><b>' + esc(f.fatal.date) + '</b> : ' + esc(f.fatal.text) + '</div>');
  if (f.priority) h.push('<div>Priority ' + [linkHtml(f.priority.num), esc(f.priority.date)].filter(Boolean).join(' · ') + '</div>');
  if (f.filing) h.push('<div>' + [linkHtml(f.filing.num), esc(f.filing.date), f.filing.granted ? 'granted ' + esc(f.filing.granted) : ''].filter(Boolean).join(' · ') + '</div>');
  else if (f.country) h.push('<div>' + esc(f.country) + '</div>');
  if (f.grant) h.push('<div>' + [linkHtml(f.grant.num), 'granted ' + esc(f.grant.date)].filter(Boolean).join(' · ') + '</div>');
  if ((f.firms || []).length) h.push('<div>' + f.firms.map(function (x) { return x.url ? '<a data-href="' + esc(x.url) + '">' + esc(x.name) + '</a>' : esc(x.name); }).join(', ') +
    (f.assocRef ? ' · ref ' + esc(f.assocRef) : '') + '</div>');
  else if (f.assocRef) h.push('<div>Associate ref ' + esc(f.assocRef) + '</div>');
  const links = [['details', 'Equinox'], ['documents', 'Documents'], ['family', 'Family'], ['contacts', 'Contacts'], ['folder', 'Drive folder']]
    .filter(function (x) { return l[x[0]]; }).map(function (x) { return '<a data-href="' + esc(l[x[0]]) + '">' + x[1] + '</a>'; });
  if (links.length) h.push('<div class="links">' + links.join('') + '</div>');
  if (out.image && /^data:image\//.test(out.image)) h.push('<img alt="" src="' + esc(out.image) + '">');
  $('matter').innerHTML = h.join('');
}

// ---------- opening with the document ----------
// Word for Mac can't load an add-in for every document, but a document carrying this mark opens with the pane
// showing (the manifest's TaskpaneId is Office.AutoShowTaskpaneWithDocument). The mark is saved with the
// document the next time it is saved; documents not in a matter are never marked.
function autoOn() { return val(KEY.auto) !== '0'; }
function markAutoOpen() {
  try {
    const st = Office.context.document.settings;
    if (st.get('Office.AutoShowTaskpaneWithDocument') === true) return;
    st.set('Office.AutoShowTaskpaneWithDocument', true);
    st.saveAsync(function () {});
  } catch (e) { /* not supported here: the pane just doesn't open by itself */ }
}

// ---------- 13.5 File to Equinox ----------
// The document exactly as Word has it now (including changes not yet saved), in slices joined into base64
function documentBase64() {
  return new Promise(function (resolve, reject) {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error(r.error && r.error.message || 'Word could not read the document')); return; }
      const file = r.value, n = file.sliceCount, parts = [];
      const next = function (i) {
        file.getSliceAsync(i, function (s) {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(new Error(s.error && s.error.message || 'Word could not read the document')); return; }
          parts[i] = s.value.data;
          if (i + 1 < n) next(i + 1); else { file.closeAsync(); resolve(toBase64(parts)); }
        });
      };
      next(0);
    });
  });
}
function toBase64(parts) {
  let bin = '';
  parts.forEach(function (p) {
    const a = p instanceof Uint8Array ? p : Uint8Array.from(p);
    for (let i = 0; i < a.length; i += 0x8000) bin += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000));
  });
  return btoa(bin);
}

async function fileToEquinox() {
  const out = $('fileOut');
  state.doc = docInfo();
  if (!state.doc.name) { out.innerHTML = '<div class="bad">Save the document in the matter\'s Drive folder first, so it has a name.</div>'; return; }
  const codes = [state.code].concat($('fileAll').checked ? state.codes.filter(function (c) { return c !== state.code; }) : []);
  out.innerHTML = '<div class="muted">Reading the document…</div>';
  await busy($('fileBtn'), async function () {
    try {
      const b64 = await documentBase64();
      out.innerHTML = '<div class="muted">Filing ' + esc(state.doc.name) + ' (' + size(Math.floor(b64.length * 3 / 4)) + ')…</div>';
      const r = await call('file', { codes: codes, name: state.doc.name, path: state.doc.path, base64: b64 });
      out.innerHTML = (r.results || []).map(function (x) {
        const good = x.outcome === 'filed';
        return '<div class="' + (good ? 'good' : '') + '">' + esc(x.code) + ': ' + esc(x.outcome) + (x.docId ? ' (Equinox document ' + esc(x.docId) + ')' : '') + '</div>';
      }).join('') + (r.inDrive ? '' : '<div class="muted">Not matched to a Drive file, so it was filed by name and contents.</div>');
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// ---------- 13.6 Insert summary ----------
async function insertSummary() {
  const out = $('sumOut');
  await busy($('sumBtn'), async function () {
    try {
      const wantImg = $('sumImage').checked;
      const r = await call('summary', { code: state.code, deadline: $('sumDeadline').checked, image: false });
      let img = '';
      if (wantImg) {
        const m = state.matter && state.matter.image ? state.matter : await call('matter', { code: state.code, image: true });
        img = String(m.image || '').replace(/^data:image\/[a-z+]+;base64,/i, '');
      }
      await Word.run(async function (ctx) {
        const range = ctx.document.getSelection().insertHtml(r.html, 'Replace');
        if (img) range.insertParagraph('', 'After').insertInlinePictureFromBase64(img, 'Start');
        await ctx.sync();
      });
      out.innerHTML = '<div class="good">Summary inserted' + (wantImg && !img ? ' (this matter has no picture)' : '') + '.</div>';
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// ---------- 13.4 Open from Equinox ----------
async function showDocs() {
  const box = $('docs');
  if (!box.hidden) { box.hidden = true; return; }
  box.hidden = false;
  if (state.docs && state.docs.code === state.code) { renderDocs(); return; }
  $('docList').innerHTML = '<li class="muted">Loading documents…</li>';
  try { const r = await call('docs', { code: state.code }); state.docs = { code: state.code, list: r.docs || [] }; renderDocs(); }
  catch (e) { $('docList').innerHTML = '<li class="bad">' + esc(e.message) + '</li>'; }
}
// Only documents Word can edit are offered for check-out; PDFs, images, emails and the rest stay in Equinox
const EDITABLE = /\.(docx|docm|doc|dotx|dotm|rtf|odt)$/i;
function editable(d) { return EDITABLE.test(d.filename || '') || (!d.filename && EDITABLE.test(d.title || '')); }
function renderDocs() {
  const q = $('docFilter').value.trim().toLowerCase();
  const all = state.docs.list, words = all.filter(editable);
  const list = words.filter(function (d) { return !q || (d.filename + ' ' + d.title).toLowerCase().indexOf(q) >= 0; });
  const note = '<li class="muted meta">' + words.length + ' Word document' + (words.length === 1 ? '' : 's') +
    (all.length > words.length ? ' (' + (all.length - words.length) + ' other file' + (all.length - words.length === 1 ? '' : 's') + ', such as PDFs, not shown)' : '') + '</li>';
  $('docList').innerHTML = note + (list.length ? list.map(function (d) {
    return '<li data-id="' + esc(d.id) + '"><div class="fn">' + esc(d.filename || d.title) + '</div><div class="meta">' +
      [esc(d.date), size(d.size)].filter(Boolean).join(' · ') + '</div></li>';
  }).join('') : '<li class="muted">No Word documents' + (q ? ' match' : '') + '.</li>');
}

async function checkout(id) {
  const d = (state.docs.list || []).filter(function (x) { return String(x.id) === String(id); })[0];
  if (!d || !editable(d)) return;
  const out = $('coOut');
  $('docs').hidden = true;
  out.innerHTML = '<div class="muted">Saving ' + esc(d.filename || d.title) + ' to the matter\'s Drive folder…</div>';
  try {
    const r = await call('checkout', { code: state.code, docId: d.id });
    const local = localPath(r);
    const h = ['<div class="good">' + (r.saved ? 'Saved to the Drive folder as ' : 'Already in the Drive folder: ') + '<b>' + esc(r.name) + '</b></div>'];
    h.push('<div>Open it from Finder' + (local ? ':<br><code>' + esc(local) + '</code>' : ' in the matter\'s folder') +
      ' (Google Drive may take a minute to sync), so your edits stay in the Drive folder. <a data-href="' + esc(r.url) + '">View in Drive</a></div>');
    if (r.base64 && canCreate()) { h.push('<button id="openCopy" class="small">Open a copy now</button> <span class="muted">(unsaved, for reading)</span>'); state.copy = r.base64; }
    else if (r.note) h.push('<div class="muted">' + esc(r.note) + '</div>');
    out.innerHTML = h.join('');
  } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
}
// Where the checked-out file will be on this Mac, worked out from this document's own path when it sits in
// the same Google Drive (e.g. …/Shared drives/EUREKA/<matter folder path>/<name>)
function localPath(r) {
  const p = state.doc.path, fp = String(r.folderPath || '').split('/').map(function (x) { return x.trim(); }).filter(Boolean).join('/');
  if (!p || !fp || !/^\//.test(p)) return '';
  const i = p.toLowerCase().indexOf('/' + fp.toLowerCase().replace(/^\/+/, ''));
  if (i >= 0) return p.slice(0, i + 1 + fp.replace(/^\/+/, '').length) + '/' + r.name;
  const top = fp.replace(/^\/+/, '').split('/')[0];   // the folder path's first folder, found higher up this path
  const j = top ? p.indexOf('/' + top + '/') : -1;
  return j >= 0 ? p.slice(0, j + 1) + fp.replace(/^\/+/, '') + '/' + r.name : '';
}
function canCreate() { try { return Office.context.requirements.isSetSupported('WordApi', '1.3'); } catch (e) { return false; } }
async function openCopy() {
  try { await Word.run(async function (ctx) { ctx.application.createDocument(state.copy).open(); await ctx.sync(); }); }
  catch (e) { $('coOut').insertAdjacentHTML('beforeend', '<div class="bad">' + esc(e.message) + '</div>'); }
}

// ---------- 13.8 Letters ----------
// Templates are Word files in the "Equinox letter templates" Drive folder. New letter opens one as a new
// document and leaves a note (this add-in's storage, shared by its panes) so the pane in that new window knows
// the matter and recipient and fills the fields straight away. Fill fields works on any open document.
const CODE_RE = /#[A-Z][A-Z0-9_?+\-]*(?:\|[^#\n<>]{0,24})?#/g;
function pendingLetter() {
  try { const p = JSON.parse(store.get(KEY.pending) || 'null'); return p && Date.now() - p.at < 15 * 60000 ? p : null; } catch (e) { return null; }
}
function docMatter() { try { return Office.context.document.settings.get('Equinox.matter') || ''; } catch (e) { return ''; } }
// the email subject kept with the document: the letter template's (from New letter) or your own
function docSubject() { try { return Office.context.document.settings.get('Equinox.subject') || ''; } catch (e) { return ''; } }
function setDocSubject(s) {
  try { const st = Office.context.document.settings; if ((st.get('Equinox.subject') || '') === s) return; if (s) st.set('Equinox.subject', s); else st.remove('Equinox.subject'); st.saveAsync(function () {}); } catch (e) { /* not kept */ }
}
function markDocMatter(code) {
  try { const st = Office.context.document.settings; if (st.get('Equinox.matter') === code) return; st.set('Equinox.matter', code); st.saveAsync(function () {}); } catch (e) { /* not kept */ }
}

async function loadLetter(code) {
  if (!state.templates) {
    try { state.templates = (await call('templates')).templates || []; } catch (e) { state.templates = []; }
    $('tplSel').innerHTML = state.templates.length ? state.templates.map(function (t) {
      return '<option value="' + esc(t.id) + '">' + esc(t.name) + '</option>'; }).join('') : '<option value="">No templates yet</option>';
  }
  const tpl = (state.templates || []).filter(function (t) { return t.id === $('tplSel').value; })[0] || {};
  const r = await call('letterFields', { code: code, recipient: state.recipientChosen ? state.recipient : '', to: tpl.to || 'client', codes: ['DATE'] });
  if (state.code !== code) return;
  state.letter = r; state.recipient = r.recipient || '';
  $('recipSel').innerHTML = (r.recipients || []).map(function (x) {
    return '<option value="' + esc(x.key) + '"' + (x.key === r.recipient ? ' selected' : '') + '>' + esc(x.label) + (x.email ? ' — ' + esc(x.email) : '') + '</option>';
  }).join('') || '<option value="">No contacts found in Equinox</option>';
  // the email's addresses start as remembered for this matter or its associate or client (set in Studio, 13.13),
  // else the recipient's email (To) and their cc in Equinox (Cc)
  const me = (r.recipients || []).filter(function (x) { return x.key === r.recipient; })[0] || {};
  state.mail = me.mail ? { to: me.mail.to.slice(), cc: me.mail.cc.slice(), bcc: me.mail.bcc.slice() } : { to: me.email ? [me.email] : [], cc: me.cc ? [me.cc] : [], bcc: [] };
  renderMail();
  state.stdSubject = r.subject || '';
  $('subjIn').value = docSubject() || state.stdSubject;
  showWarnings(r.warnings);
  return r;
}
// To / Cc / Bcc for Email PDF: what is chosen, as chips, and a dropdown of every contact on the matter with an email
function renderMail() {
  const box = $('mailTo'); if (!box) return;
  const M = state.mail || { to: [], cc: [], bcc: [] };
  const people = [];
  ((state.letter && state.letter.recipients) || []).forEach(function (x) {
    if (x.email) people.push({ a: x.email, l: x.label + ' — ' + x.email });
    if (x.cc && x.cc !== x.email) people.push({ a: x.cc, l: x.label + ' (their cc) — ' + x.cc });
  });
  box.innerHTML = [['to', 'To'], ['cc', 'Cc'], ['bcc', 'Bcc']].map(function (f) {
    const k = f[0], used = M.to.concat(M.cc, M.bcc);
    return '<b>' + f[1] + '</b><div class="chips">' + M[k].map(function (a, i) {
      return '<span class="chip">' + esc(a) + '<button data-k="' + k + '" data-i="' + i + '" title="Remove">×</button></span>'; }).join('') +
      '<select data-k="' + k + '"><option value="">+ add…</option>' + people.filter(function (p) { return used.indexOf(p.a) < 0; })
        .map(function (p) { return '<option value="' + esc(p.a) + '">' + esc(p.l) + '</option>'; }).join('') + '</select></div>';
  }).join('');
}
function showWarnings(w) {
  const box = $('letterWarn');
  if (box) box.innerHTML = (w || []).map(function (x) { return '<div class="codes">' + esc(x) + '</div>'; }).join('');
}
async function lettersReady() {
  if (!state.code) throw new Error('Choose the matter first.');
  if (!state.letter || state.letter.code !== state.code || (state.letter.recipient || '') !== (state.recipient || '')) await loadLetter(state.code);
  return state.letter;
}

async function newLetter() {
  const out = $('letterOut');
  await busy($('letterBtn'), async function () {
    try {
      if (!canCreate()) throw new Error('This version of Word can\'t open a new document from the add-in.');
      const id = $('tplSel').value;
      if (!id) { out.innerHTML = '<div class="bad">No templates in the "Equinox letter templates" Drive folder yet.</div>'; return; }
      const t = (state.templates || []).filter(function (x) { return x.id === id; })[0] || {};
      out.innerHTML = '<div class="muted">Putting together ' + esc(t.name || 'the letter') + '…</div>';
      let base64;
      if (t.kind === 'parts') {   // master + header + content + sign-off, filled by the script
        const r = await call('letterBuild', { code: state.code, recipient: state.recipientChosen ? state.recipient : '', id: id });
        store.set(KEY.pending, JSON.stringify({ code: state.code, recipient: r.recipient, tpl: id, name: r.name, email: r.email || '', subject: r.subject || '', html: r.html,
          images: r.images || {}, warnings: r.warnings || [], at: Date.now() }));
        base64 = r.master;
        if ((r.missing || []).length) out.insertAdjacentHTML('beforeend', '<div class="bad">Not found or not known: ' + esc(r.missing.join(', ')) + '</div>');
      } else {                    // a whole-document template: opened, then its codes filled
        const r = await call('template', { id: id });
        store.set(KEY.pending, JSON.stringify({ code: state.code, recipient: state.recipient, tpl: id, email: t.email || '', at: Date.now() }));
        base64 = r.base64;
      }
      await Word.run(async function (ctx) { ctx.application.createDocument(base64).open(); await ctx.sync(); });
      out.innerHTML = '<div class="good">Opened in a new window.</div><div class="muted">The Equinox pane there writes the letter. If it doesn\'t open by itself, choose Equinox → Fill fields in that window.</div>';
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// The letter's HTML (header, content, sign-off from letterBuild) written into this document where the master says
// #EQX_LETTER# (else at the end), one Word paragraph per HTML paragraph or line. A paragraph's class names its Word
// style (letter-subject → "Letter Subject"); h1–h3, list items and blockquotes use Word's built-in styles. A list
// item is in List Bullet or List Number, a level further in (List Bullet 2, 3) for each list it sits inside, unless
// its class names the style. A table becomes a Word table in its class's table style (default Eureka Table), the
// text only, first row as a header row when it is <th> cells.
const BUILT_IN = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', ul: 'ListBullet', ol: 'ListNumber', blockquote: 'Quote' };
function styleName(cls) {
  const first = String(cls || '').trim().split(/\s+/)[0];
  return first ? first.split('-').map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ') : '';
}
function letterBlocks(html) {
  const root = new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html').body.firstChild;
  const blocks = [];
  const runsOf = function (node, fmt, lines) {
    node.childNodes.forEach(function (n) {
      if (n.nodeType === 3) { const t = n.nodeValue.replace(/\s+/g, ' '); if (t) lines[lines.length - 1].push(Object.assign({ text: t }, fmt)); return; }
      if (n.nodeType !== 1) return;
      const tag = n.tagName.toLowerCase();
      if (tag === 'br') { lines.push([]); return; }
      if (tag === 'ul' || tag === 'ol') return;
      if ((tag === 'p' || tag === 'div') && lines[lines.length - 1].length) lines.push([]);
      if (tag === 'img') { if (n.getAttribute('data-eqx')) lines[lines.length - 1].push({ img: n.getAttribute('data-eqx'), text: '' }); return; }
      const f = Object.assign({}, fmt);
      if (tag === 'strong' || tag === 'b') f.b = true;
      if (tag === 'em' || tag === 'i') f.i = true;
      if (tag === 'u') f.u = true;
      if (tag === 'a' && n.getAttribute('href')) f.href = n.getAttribute('href');
      if (tag === 'span' && n.className) f.cs = styleName(n.className);   // a Word character style (eureka-mark → "Eureka Mark")
      runsOf(n, f, lines);
    });
  };
  const add = function (node, style, builtIn) {
    const lines = [[]];
    runsOf(node, {}, lines);
    lines.forEach(function (runs) {
      if (runs.length) { runs[0].text = runs[0].text.replace(/^\s+/, ''); runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, ''); }
      if (runs.length && !runs[0].img) runs[0].text = runs[0].text.replace(/^\s+/, '');
      blocks.push({ style: style, builtIn: builtIn, runs: runs.filter(function (r) { return r.text || r.img; }) });
    });
  };
  const list = function (el, depth) {
    const t = el.tagName.toLowerCase();
    Array.prototype.forEach.call(el.children, function (li) {
      if (li.tagName.toLowerCase() !== 'li') return;
      add(li, li.className ? styleName(li.className) : (t === 'ol' ? 'List Number' : 'List Bullet') + (depth > 1 ? ' ' + Math.min(3, depth) : ''), BUILT_IN[t]);
      Array.prototype.forEach.call(li.children, function (c) { if (/^(ul|ol)$/i.test(c.tagName)) list(c, depth + 1); });
    });
  };
  const cellText = function (c) {
    const lines = [[]];
    runsOf(c, {}, lines);
    return lines.map(function (runs) { return runs.map(function (r) { return r.text || ''; }).join('').trim(); }).filter(Boolean).join('\n');
  };
  const walk = function (parent) {
    parent.childNodes.forEach(function (n) {
      if (n.nodeType === 3) { if (n.nodeValue.trim()) blocks.push({ style: '', runs: [{ text: n.nodeValue.trim() }] }); return; }
      if (n.nodeType !== 1) return;
      const tag = n.tagName.toLowerCase();
      if (tag === 'ul' || tag === 'ol') { list(n, 1); return; }
      if (tag === 'table') {
        const trs = Array.prototype.slice.call(n.querySelectorAll('tr'));
        const rows = trs.map(function (tr) { return Array.prototype.map.call(tr.children, cellText); });
        const cols = Math.max.apply(null, [1].concat(rows.map(function (r) { return r.length; })));
        rows.forEach(function (r) { while (r.length < cols) r.push(''); });
        if (rows.length) blocks.push({ table: true, style: n.className ? styleName(n.className) : 'Eureka Table', rows: rows,
          head: !!(trs[0] && trs[0].children[0] && trs[0].children[0].tagName.toLowerCase() === 'th') });
        return;
      }
      if (tag === 'div' || tag === 'section') { walk(n); return; }
      add(n, styleName(n.className), n.className ? '' : (BUILT_IN[tag] || ''));
    });
  };
  walk(root);
  return blocks;
}
async function buildLetter() {
  const out = $('letterOut');
  const p = pendingLetter();
  if (!p || !p.html) { out.innerHTML = '<div class="bad">Nothing waiting to be written: use New letter.</div>'; return; }
  store.set(KEY.pending, JSON.stringify(Object.assign({}, p, { html: '' })));   // write it once only
  state.emailTpl = p.email || state.emailTpl;
  out.innerHTML = '<div class="muted">Writing the letter…</div>';
  try {
    const blocks = letterBlocks(p.html);
    const unknown = await Word.run(async function (ctx) {
      const body = ctx.document.body;
      const hit = body.search('#EQX_LETTER#', { matchCase: true }); hit.load('items');
      let styles = null;
      try { if (Office.context.requirements.isSetSupported('WordApi', '1.5')) { styles = ctx.document.getStyles(); styles.load('items/nameLocal'); } } catch (e) { styles = null; }
      await ctx.sync();
      const have = {};
      if (styles) styles.items.forEach(function (s) { have[s.nameLocal] = true; });
      const known = function (n) { return !n || !styles || have[n] || /^(Letter |Eureka )/.test(n) && !/^Eureka Table$/.test(n); };
      const missing = [];
      let anchor = hit.items.length ? hit.items[0].paragraphs.getFirst() : body.paragraphs.getLast();
      const marker = hit.items.length ? anchor : null;
      let after = null;   // the empty paragraph Word keeps after a table: the next paragraph's text goes in it
      blocks.forEach(function (b) {
        if (b.table) {
          const tbl = (after || anchor).insertTable(b.rows.length, b.rows[0].length, 'After', b.rows);
          if (styles && have[b.style]) tbl.style = b.style;
          else { tbl.styleBuiltIn = 'TableGrid'; if (missing.indexOf(b.style) < 0) missing.push(b.style); }
          if (b.head) tbl.headerRowCount = 1;
          after = tbl.insertParagraph('', 'After'); anchor = after;
          return;
        }
        const para = after || anchor.insertParagraph('', 'After');
        after = null;
        if (b.style && known(b.style)) para.style = b.style;
        else if (b.builtIn) para.styleBuiltIn = b.builtIn;
        else { para.styleBuiltIn = 'Normal'; if (b.style && missing.indexOf(b.style) < 0) missing.push(b.style); }
        b.runs.forEach(function (r) {
          if (r.img) {
            const im = (p.images || {})[r.img];
            if (im && im.base64) { const pic = para.insertInlinePictureFromBase64(im.base64, 'End'); pic.lockAspectRatio = true; if (im.width) pic.width = im.width; }
            return;
          }
          const range = para.insertText(r.text, 'End');
          if (r.b) range.font.bold = true;
          if (r.i) range.font.italic = true;
          if (r.u) range.font.underline = 'Single';
          if (r.href) range.hyperlink = r.href;
          if (r.cs && known(r.cs)) range.style = r.cs; else if (r.cs && missing.indexOf(r.cs) < 0) missing.push(r.cs);
        });
        anchor = para;
      });
      if (marker) marker.delete();
      await ctx.sync();
      return missing;
    });
    markDocMatter(state.code);
    showWarnings(p.warnings);
    out.innerHTML = '<div class="good">' + esc(p.name || 'Letter') + ' written for ' + esc(state.code) + '.</div>' +
      (unknown.length ? '<div class="codes bad">No such Word style in the master, so Normal was used: ' + esc(unknown.join(', ')) + '</div>' : '') +
      '<div class="muted">Save it in the matter\'s Drive folder, then Save PDF or Email PDF.</div>';
    // codes in the master itself (letterhead, footer) and any the parts left
    await fillFields(true);
  } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
}

// Every #CODE# in the document (body, headers and footers) replaced by its value. A code with no value
// removes its line when the line holds nothing else but a short label (e.g. "By email: #EMAIL#").
async function fillFields(after) {
  const out = $('letterOut');
  after = after === true;
  await busy($('fillBtn'), async function () {
    try {
      if (!state.code) throw new Error('Choose the matter first.');
      const parts = async function (ctx) {
        const list = [ctx.document.body];
        const secs = ctx.document.sections; secs.load('items'); await ctx.sync();
        secs.items.forEach(function (s) { ['Primary', 'FirstPage'].forEach(function (t) { list.push(s.getHeader(t)); list.push(s.getFooter(t)); }); });
        return list;
      };
      // 1. the codes in the document (body, headers, footers)
      const codes = await Word.run(async function (ctx) {
        const list = await parts(ctx);
        list.forEach(function (x) { x.load('text'); }); await ctx.sync();
        const found = {};
        list.forEach(function (x) { (String(x.text || '').match(CODE_RE) || []).forEach(function (k) { found[k] = true; }); });
        return Object.keys(found);
      });
      if (!codes.length) { if (!after) out.innerHTML = '<div class="muted">No #CODES# in this document.</div>'; return; }
      // 2. their values from Equinox, for the chosen recipient
      const L = await call('letterFields', { code: state.code, recipient: state.recipient, codes: codes.map(function (k) { return k.slice(1, -1); }) });
      const F = L.fields || {}, I = L.images || {};
      // 3. replace them
      const res = await Word.run(async function (ctx) {
        const list = await parts(ctx);
        const unknown = [], blank = [], hits = [];
        codes.forEach(function (k) {
          const key = k.slice(1, -1);
          if (!Object.prototype.hasOwnProperty.call(F, key) && !I[key]) { unknown.push(k); return; }
          list.forEach(function (x) { const r = x.search(k, { matchCase: true }); r.load('items'); hits.push({ k: k, v: String(F[key] || ''), img: I[key], r: r }); });
        });
        await ctx.sync();
        let filled = 0;
        const empties = [];
        hits.forEach(function (h) {
          h.r.items.forEach(function (range) {
            if (h.img) { const pic = range.insertInlinePictureFromBase64(h.img.base64, 'Replace'); pic.lockAspectRatio = true; if (h.img.width) pic.width = h.img.width; filled++; }
            else if (h.v) { range.insertText(h.v, 'Replace'); filled++; }
            else { empties.push(range.paragraphs.getFirst()); range.insertText('', 'Replace'); }
          });
          if (!h.v && !h.img && h.r.items.length && blank.indexOf(h.k) < 0) blank.push(h.k);
        });
        await ctx.sync();
        // lines left holding only a short label (or nothing) go; each such line once
        empties.forEach(function (para) { para.load('text'); });
        await ctx.sync();
        const gone = {};
        empties.forEach(function (para) {
          const rest = String(para.text || '').replace(CODE_RE, '').trim();
          if ((/^[A-Za-z][A-Za-z .]{0,20}:$/.test(rest) || !rest || /^(in the name of|by email|by fax)$/i.test(rest)) && !gone[rest]) { gone[rest] = true; para.delete(); }
        });
        await ctx.sync();
        return { filled: filled, unknown: unknown, blank: blank, total: codes.length };
      });
      markDocMatter(state.code);
      const p = pendingLetter(); if (p && p.code === state.code) { state.emailTpl = p.email || state.emailTpl; store.set(KEY.pending, ''); }
      if (after) {
        if (res.unknown.length) out.insertAdjacentHTML('beforeend', '<div class="codes bad">Not known, left in the text: ' + esc(res.unknown.join(' ')) + '</div>');
        return;
      }
      showWarnings(L.warnings);
      const h = ['<div class="good">Filled ' + res.filled + ' field' + (res.filled === 1 ? '' : 's') + ' for ' + esc(state.code) + '.</div>'];
      if (res.blank.length) h.push('<div class="codes">Nothing in Equinox for ' + esc(res.blank.join(' ')) + ' (left out)</div>');
      if (res.unknown.length) h.push('<div class="codes bad">Not known, left in the text: ' + esc(res.unknown.join(' ')) + '</div>');
      if (!L.live) h.push('<div class="muted">Equinox couldn\'t be reached, so details came from the index only.</div>');
      out.innerHTML = h.join('');
    } catch (e) { if (after) out.insertAdjacentHTML('beforeend', '<div class="bad">' + esc(e.message) + '</div>'); else out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// ---------- 13.9 PDF to Drive and Gmail ----------
function documentPdfBase64() {
  return new Promise(function (resolve, reject) {
    Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4194304 }, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error(r.error && r.error.message || 'Word could not make the PDF')); return; }
      const file = r.value, n = file.sliceCount, parts = [];
      const next = function (i) {
        file.getSliceAsync(i, function (s) {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(new Error(s.error && s.error.message || 'Word could not make the PDF')); return; }
          parts[i] = s.value.data;
          if (i + 1 < n) next(i + 1); else { file.closeAsync(); resolve(toBase64(parts)); }
        });
      };
      next(0);
    });
  });
}
async function savePdf(quiet) {
  const out = $('letterOut');
  if (!state.code) throw new Error('Choose the matter first.');
  state.doc = docInfo();
  if (!quiet) out.innerHTML = '<div class="muted">Making the PDF…</div>';
  const b64 = await documentPdfBase64();
  const r = await call('savePdf', { code: state.code, name: state.doc.name, base64: b64 });
  markDocMatter(state.code);
  state.pdf = r;
  out.innerHTML = '<div class="good">' + (r.saved ? 'Saved ' : 'Already in the Drive folder: ') + '<b>' + esc(r.name) + '</b></div>' +
    '<div><a data-href="' + esc(r.url) + '">View in Drive</a>' + (r.folderPath ? ' <span class="muted">· ' + esc(r.folderPath) + '</span>' : '') + '</div>';
  return r;
}
async function savePdfClick() {
  await busy($('pdfBtn'), async function () { try { await savePdf(false); } catch (e) { $('letterOut').innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; } });
}
async function emailPdf() {
  const out = $('letterOut');
  await busy($('emailBtn'), async function () {
    try {
      await lettersReady();
      out.innerHTML = '<div class="muted">Making the PDF…</div>';
      const pdf = await savePdf(true);
      out.insertAdjacentHTML('beforeend', '<div class="muted">Making the Gmail draft…</div>');
      const M = state.mail || { to: [], cc: [], bcc: [] };
      const tpl = state.emailTpl || (((state.templates || []).filter(function (t) { return t.id === $('tplSel').value; })[0]) || {}).email || '';
      const r = await call('pdfDraft', { code: state.code, fileId: pdf.fileId, recipient: state.recipient, to: M.to, cc: M.cc, bcc: M.bcc, subject: $('subjIn').value.trim(), template: tpl });
      out.innerHTML = '<div class="good">Gmail draft ready' + (r.to ? ' to ' + esc(r.to) : ' (add the recipient in Gmail)') + ', with ' + esc(pdf.name) + ' attached.</div>' +
        '<div><a data-href="' + esc(r.url) + '">Open the draft in Gmail</a></div>' +
        '<div class="muted">Once you send it, it is filed to Equinox with the PDF.</div>';
      openLink(r.url);
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// ---------- ribbon buttons ----------
// Each Equinox-tab button opens this pane with ?do=… ; the action runs once the matter is known.
function runIntent() {
  const what = state.intent; state.intent = '';
  if (!what || !state.code) return;
  const go = { letter: function () { $('letterAct').scrollIntoView(); $('tplSel').focus(); }, build: buildLetter, fill: function () { fillFields(); }, pdf: savePdfClick, email: emailPdf, file: fileToEquinox }[what];
  if (go) setTimeout(go, 50);
}

// ---------- wiring ----------
function wire() {
  $('gear').onclick = function () { $('settings').hidden ? openSettings('') : closeSettings(); };
  $('saveSettings').onclick = saveSettings;
  $('closeSettings').onclick = closeSettings;
  $('showToken').onclick = function () { const t = $('setToken'); t.type = t.type === 'password' ? 'text' : 'password'; this.textContent = t.type === 'password' ? 'Show' : 'Hide'; };
  $('codeForm').onsubmit = typed;
  $('matches').onclick = function (e) { const b = e.target.closest('.pill'); if (b) select(b.getAttribute('data-code')); };
  $('fileBtn').onclick = fileToEquinox;
  $('sumBtn').onclick = insertSummary;
  $('docsBtn').onclick = showDocs;
  $('docFilter').oninput = function () { if (state.docs) renderDocs(); };
  $('docList').onclick = function (e) { const li = e.target.closest('li[data-id]'); if (li) checkout(li.getAttribute('data-id')); };
  $('letterBtn').onclick = newLetter;
  $('fillBtn').onclick = function () { fillFields(); };
  $('pdfBtn').onclick = savePdfClick;
  $('emailBtn').onclick = emailPdf;
  $('subjIn').oninput = function () { const v = this.value.trim(); setDocSubject(v && v !== state.stdSubject ? v : ''); };
  $('subjIn').onchange = function () { if (!this.value.trim()) this.value = state.stdSubject || ''; };
  $('mailTo').onchange = function (e) {
    const k = e.target.getAttribute('data-k'); if (!k || !e.target.value || !state.mail) return;
    state.mail[k].push(e.target.value); renderMail();
  };
  $('mailTo').onclick = function (e) {
    const b = e.target.closest('button[data-k]'); if (!b || !state.mail) return;
    state.mail[b.getAttribute('data-k')].splice(Number(b.getAttribute('data-i')), 1); renderMail();
  };
  $('recipSel').onchange = function () { state.recipient = this.value; state.recipientChosen = true; loadLetter(state.code).catch(function (e) { say($('msg'), e.message, 'bad'); }); };
  $('tplSel').onchange = function () { if (!state.recipientChosen && state.code) loadLetter(state.code).catch(function () {}); };
  document.body.addEventListener('click', function (e) {
    const a = e.target.closest('a[data-href]'); if (a) { e.preventDefault(); openLink(a.getAttribute('data-href')); }
    if (e.target.id === 'openCopy') openCopy();
  });
  $('ver').textContent = PANE_VERSION;
}

if (typeof Office !== 'undefined') {
  Office.onReady(function (info) {
    wire();
    try { state.intent = new URLSearchParams(location.search).get('do') || ''; } catch (e) { state.intent = ''; }
    if (info.host !== Office.HostType.Word) { $('docLine').textContent = 'Open this page from Word (Home → Equinox).'; return; }
    start();
  });
} else {
  document.addEventListener('DOMContentLoaded', function () { wire(); $('docLine').textContent = 'Open this page from Word (Home → Equinox).'; });
}
