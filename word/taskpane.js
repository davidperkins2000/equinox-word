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
const PANE_VERSION = 'w0.3.0';

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
      if (p) { state.recipient = p.recipient || ''; state.emailTpl = p.email || ''; if (!state.intent) state.intent = 'fill'; }
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
const CODE_RE = /#[A-Z0-9_?+\-]+#/g;
function pendingLetter() {
  try { const p = JSON.parse(store.get(KEY.pending) || 'null'); return p && Date.now() - p.at < 15 * 60000 ? p : null; } catch (e) { return null; }
}
function docMatter() { try { return Office.context.document.settings.get('Equinox.matter') || ''; } catch (e) { return ''; } }
function markDocMatter(code) {
  try { const st = Office.context.document.settings; if (st.get('Equinox.matter') === code) return; st.set('Equinox.matter', code); st.saveAsync(function () {}); } catch (e) { /* not kept */ }
}

async function loadLetter(code) {
  if (!state.templates) {
    try { state.templates = (await call('templates')).templates || []; } catch (e) { state.templates = []; }
    $('tplSel').innerHTML = state.templates.length ? state.templates.map(function (t) {
      return '<option value="' + esc(t.id) + '">' + esc(t.name) + '</option>'; }).join('') : '<option value="">No templates yet</option>';
  }
  const r = await call('letterFields', { code: code, recipient: state.recipient });
  if (state.code !== code) return;
  state.letter = r; state.recipient = r.recipient || '';
  $('recipSel').innerHTML = (r.recipients || []).map(function (x) {
    return '<option value="' + esc(x.key) + '"' + (x.key === r.recipient ? ' selected' : '') + '>' + esc(x.label) + (x.email ? ' — ' + esc(x.email) : '') + '</option>';
  }).join('') || '<option value="">No contacts found in Equinox</option>';
  return r;
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
      out.innerHTML = '<div class="muted">Opening ' + esc(t.name || 'the template') + '…</div>';
      const r = await call('template', { id: id });
      store.set(KEY.pending, JSON.stringify({ code: state.code, recipient: state.recipient, tpl: id, email: t.email || '', at: Date.now() }));
      await Word.run(async function (ctx) { ctx.application.createDocument(r.base64).open(); await ctx.sync(); });
      out.innerHTML = '<div class="good">Opened in a new window.</div><div class="muted">If its fields aren\'t filled in a moment, choose Equinox → Fill fields there.</div>';
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
  });
}

// Every #CODE# in the document (body, headers and footers) replaced by its value. A code with no value
// removes its line when the line holds nothing else but a short label (e.g. "By email: #EMAIL#").
async function fillFields() {
  const out = $('letterOut');
  await busy($('fillBtn'), async function () {
    try {
      const L = await lettersReady(), F = L.fields || {};
      const res = await Word.run(async function (ctx) {
        const parts = [ctx.document.body];
        const secs = ctx.document.sections; secs.load('items'); await ctx.sync();
        secs.items.forEach(function (s) { ['Primary', 'FirstPage'].forEach(function (t) { parts.push(s.getHeader(t)); parts.push(s.getFooter(t)); }); });
        parts.forEach(function (p) { p.load('text'); }); await ctx.sync();
        const found = {};
        parts.forEach(function (p) { (String(p.text || '').match(CODE_RE) || []).forEach(function (k) { found[k] = true; }); });
        const codes = Object.keys(found), unknown = [], blank = [], hits = [];
        codes.forEach(function (k) {
          const key = k.slice(1, -1);
          if (!Object.prototype.hasOwnProperty.call(F, key)) { unknown.push(k); return; }
          parts.forEach(function (p) { const r = p.search(k, { matchCase: true }); r.load('items'); hits.push({ k: k, v: String(F[key] || ''), r: r }); });
        });
        await ctx.sync();
        let filled = 0;
        const empties = [];
        hits.forEach(function (h) {
          h.r.items.forEach(function (range) {
            if (h.v) { range.insertText(h.v, 'Replace'); filled++; }
            else { empties.push(range.paragraphs.getFirst()); range.insertText('', 'Replace'); }
          });
          if (!h.v && h.r.items.length && blank.indexOf(h.k) < 0) blank.push(h.k);
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
      const h = ['<div class="good">' + (res.total ? 'Filled ' + res.filled + ' field' + (res.filled === 1 ? '' : 's') + ' for ' + esc(state.code) + '.' : 'No #CODES# in this document.') + '</div>'];
      if (res.blank.length) h.push('<div class="codes">Nothing in Equinox for ' + esc(res.blank.join(' ')) + ' (left out)</div>');
      if (res.unknown.length) h.push('<div class="codes bad">Not known, left in the text: ' + esc(res.unknown.join(' ')) + '</div>');
      if (!L.live) h.push('<div class="muted">Equinox couldn\'t be reached, so addresses came from the index only.</div>');
      out.innerHTML = h.join('');
    } catch (e) { out.innerHTML = '<div class="bad">' + esc(e.message) + '</div>'; }
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
      const rec = (state.letter.recipients || []).filter(function (x) { return x.key === state.recipient; })[0] || {};
      const tpl = state.emailTpl || (((state.templates || []).filter(function (t) { return t.id === $('tplSel').value; })[0]) || {}).email || '';
      const r = await call('pdfDraft', { code: state.code, fileId: pdf.fileId, recipient: state.recipient, to: rec.email || '', template: tpl });
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
  const go = { letter: function () { $('letterAct').scrollIntoView(); $('tplSel').focus(); }, fill: fillFields, pdf: savePdfClick, email: emailPdf, file: fileToEquinox }[what];
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
  $('fillBtn').onclick = fillFields;
  $('pdfBtn').onclick = savePdfClick;
  $('emailBtn').onclick = emailPdf;
  $('recipSel').onchange = function () { state.recipient = this.value; loadLetter(state.code).catch(function (e) { say($('msg'), e.message, 'bad'); }); };
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
