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
const PANE_VERSION = 'w0.1.0';

const KEY = { url: 'eqxw.url', token: 'eqxw.token', img: 'eqxw.img' };
const store = {
  get: function (k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage blocked: kept for this session only */ } mem[k] = v; }
};
const mem = {};
const val = function (k) { return store.get(k) || mem[k] || ''; };

const $ = function (id) { return document.getElementById(id); };
const esc = function (t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
const state = { doc: { name: '', path: '' }, codes: [], matters: [], code: '', matter: null, docs: null };

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
  $('setUrl').value = val(KEY.url); $('setToken').value = val(KEY.token); $('setImg').checked = val(KEY.img) === '1';
  say($('settingsMsg'), why || '', why ? 'bad' : '');
  $('settings').hidden = false; $('main').hidden = true;
}
function closeSettings() { $('settings').hidden = true; $('main').hidden = false; }
async function saveSettings() {
  const url = $('setUrl').value.trim(), token = $('setToken').value.trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec\/?$/.test(url)) { say($('settingsMsg'), 'The address should start https://script.google.com/ and end in /exec.', 'bad'); return; }
  if (token.length < 32) { say($('settingsMsg'), 'That token looks too short. Copy the whole line from the log.', 'bad'); return; }
  store.set(KEY.url, url); store.set(KEY.token, token); store.set(KEY.img, $('setImg').checked ? '1' : '');
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
  if (!state.doc.name) return;
  say($('msg'), 'Finding the matter…');
  try { showMatches(await call('lookup', { name: state.doc.name, path: state.doc.path })); }
  catch (e) { say($('msg'), e.message, 'bad'); }
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
  } catch (e) { $('matter').innerHTML = '<span class="bad">' + esc(e.message) + '</span>'; }
}

function renderMatter(out) {
  const f = out.facts || {}, l = out.links || {};
  const h = [];
  h.push('<div class="code">' + esc(out.code) + '</div>');
  if (f.title) h.push('<div class="title">' + esc(f.title) + '</div>');
  if ((f.applicants || []).length) h.push('<div>' + esc(f.applicants.join('; ')) + '</div>');
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
function renderDocs() {
  const q = $('docFilter').value.trim().toLowerCase();
  const list = state.docs.list.filter(function (d) { return !q || (d.filename + ' ' + d.title).toLowerCase().indexOf(q) >= 0; });
  $('docList').innerHTML = list.length ? list.map(function (d) {
    const word = /\.(docx?|dotx|rtf)$/i.test(d.filename);
    return '<li data-id="' + esc(d.id) + '"><div class="fn">' + esc(d.filename || d.title) + '</div><div class="meta">' +
      [esc(d.date), size(d.size), word ? '' : 'not a Word file'].filter(Boolean).join(' · ') + '</div></li>';
  }).join('') : '<li class="muted">No documents' + (q ? ' match' : '') + '.</li>';
}

async function checkout(id) {
  const d = (state.docs.list || []).filter(function (x) { return String(x.id) === String(id); })[0];
  if (!d) return;
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
  const p = state.doc.path, fp = String(r.folderPath || '');
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
  document.body.addEventListener('click', function (e) {
    const a = e.target.closest('a[data-href]'); if (a) { e.preventDefault(); openLink(a.getAttribute('data-href')); }
    if (e.target.id === 'openCopy') openCopy();
  });
  $('ver').textContent = PANE_VERSION;
}

if (typeof Office !== 'undefined') {
  Office.onReady(function (info) {
    wire();
    if (info.host !== Office.HostType.Word) { $('docLine').textContent = 'Open this page from Word (Home → Equinox).'; return; }
    start();
  });
} else {
  document.addEventListener('DOMContentLoaded', function () { wire(); $('docLine').textContent = 'Open this page from Word (Home → Equinox).'; });
}
