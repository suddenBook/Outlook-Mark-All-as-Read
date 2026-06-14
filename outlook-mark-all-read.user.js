// ==UserScript==
// @name         Outlook - Mark ALL folders & subfolders as read
// @namespace    https://local/outlook-mark-all-read
// @version      1.1
// @description  One-click button to mark every Outlook folder & subfolder as read. Works on consumer (outlook.live.com) and work/school (outlook.office.com / outlook.cloud.microsoft).
// @match        https://outlook.live.com/*
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.cloud.microsoft/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
  'use strict';
  if (window.__owaMarkAllReadLoaded) return;
  window.__owaMarkAllReadLoaded = true;

  // ===================== CONFIG =====================
  const EXCLUDE_NAMES = [];        // folders to leave untouched, e.g. ["Inbox", "Junk Email"]
  const CONCURRENCY   = 6;         // parallel mark requests
  const DEFAULT_LABEL = 'Mark all folders read';

  // ====== 1. Harvest auth token + API origin + anchor mailbox from OWA's own traffic ======
  // Consumer uses  Authorization: MSAuth1.0 usertoken="..."
  // Work/school uses Authorization: Bearer <JWT>  (API on https://outlook.office.com)
  let auth = null, apiBase = null, anchor = null;
  const isToken = (v) => typeof v === 'string' && /^(Bearer\s|MSAuth1\.0\b)/i.test(v);
  const note = (urlStr, headerMap) => {
    // Only trust auth/anchor/base from OWA's own endpoints, so we never pick up a
    // Bearer token minted for a different resource (e.g. Microsoft Graph).
    if (!urlStr || !/\/owa\//i.test(urlStr)) return;
    const a = headerMap['authorization'];
    if (isToken(a)) auth = a;
    if (headerMap['x-anchormailbox']) anchor = headerMap['x-anchormailbox'];
    if (/\/owa\/[^?]*service\.svc/i.test(urlStr)) {
      try { apiBase = new URL(urlStr, location.href).origin; } catch (e) {}
    }
  };
  const toMap = (...sources) => {
    const m = {};
    const add = (k, v) => { if (k != null) m[('' + k).toLowerCase()] = v; };
    sources.forEach(h => {
      if (!h) return;
      if (h instanceof Headers) h.forEach((v, k) => add(k, v));
      else if (Array.isArray(h)) h.forEach(([k, v]) => add(k, v));
      else Object.keys(h).forEach(k => add(k, h[k]));
    });
    return m;
  };

  const origFetch = window.fetch;
  window.fetch = function (input, init) {
    try {
      const url = (input instanceof Request) ? input.url : ('' + input);
      const map = toMap(input instanceof Request ? input.headers : null, init && init.headers);
      note(url, map);
    } catch (e) {}
    return origFetch.apply(this, arguments);
  };
  const xOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (m, u) { try { this.__owaUrl = u; } catch (e) {} return xOpen.apply(this, arguments); };
  const xSet = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    try { note(this.__owaUrl, toMap([[k, v]])); } catch (e) {}
    return xSet.apply(this, arguments);
  };

  const waitForToken = async (ms = 25000) => {
    const t0 = Date.now();
    while (!auth && Date.now() - t0 < ms) await new Promise(r => setTimeout(r, 250));
    return auth;
  };

  // ===================== 2. API helpers =====================
  let seq = 1000;
  const base = () => apiBase
    || (/(^|\.)cloud\.microsoft$/i.test(location.hostname) ? 'https://outlook.office.com' : location.origin);
  const sameOrigin = () => base() === location.origin;
  const url = (action) => `${base()}/owa/service.svc?action=${action}&app=Mail&n=${seq++}`;
  const headers = (action) => {
    const h = { 'action': action, 'authorization': auth, 'content-type': 'application/json; charset=utf-8', 'prefer': 'IdType="ImmutableId"' };
    if (anchor) h['x-anchormailbox'] = anchor;
    return h;
  };

  // POST an action; for cross-origin (work/school) try without cookies first, fall back to with-cookies on auth error.
  async function api(action, body) {
    const modes = sameOrigin() ? ['include'] : ['omit', 'include'];
    let lastErr = 'request failed';
    for (const cred of modes) {
      try {
        const r = await origFetch(url(action), { method: 'POST', credentials: cred, headers: headers(action), body: JSON.stringify(body) });
        if (r.ok) return await r.json();
        lastErr = `HTTP ${r.status} (${action}) @ ${base()}`;
        if (r.status !== 401 && r.status !== 403) break;   // only auth errors are worth retrying with the other mode
      } catch (e) { lastErr = `${e} (${action}) @ ${base()}`; }
    }
    throw new Error(lastErr);
  }

  async function listAllFolders() {
    const all = [];
    let offset = 0;
    for (let page = 0; page < 50; page++) {
      const body = {
        "__type": "FindFolderJsonRequest:#Exchange",
        "Header": { "__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08" },
        "Body": {
          "__type": "FindFolderRequest:#Exchange",
          "FolderShape": {
            "__type": "FolderResponseShape:#Exchange", "BaseShape": "IdOnly",
            "AdditionalProperties": [
              { "__type": "PropertyUri:#Exchange", "FieldURI": "FolderDisplayName" },
              { "__type": "PropertyUri:#Exchange", "FieldURI": "FolderClass" },
              { "__type": "PropertyUri:#Exchange", "FieldURI": "UnreadCount" }
            ]
          },
          "Paging": { "__type": "IndexedPageView:#Exchange", "BasePoint": "Beginning", "Offset": offset, "MaxEntriesReturned": 1000 },
          "ParentFolderIds": [ { "__type": "DistinguishedFolderId:#Exchange", "Id": "msgfolderroot" } ],
          "Traversal": "Deep"
        }
      };
      const root = (await api('FindFolder', body)).Body.ResponseMessages.Items[0].RootFolder;
      const fs = root.Folders || [];
      all.push(...fs);
      if (root.IncludesLastItemInRange || fs.length === 0) break;
      offset = (root.IndexedPagingOffset != null) ? root.IndexedPagingOffset : offset + fs.length;
    }
    return all
      .filter(f => (f.FolderClass || '').startsWith('IPF.Note') || !f.FolderClass)   // mail folders only
      .map(f => ({ name: f.DisplayName, id: f.FolderId.Id, unread: f.UnreadCount || 0 }));
  }

  async function markFolderRead(folder) {
    const body = {
      "__type": "MarkAllItemsAsReadJsonRequest:#Exchange",
      "Header": { "__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08" },
      "Body": {
        "__type": "MarkAllItemsAsReadRequest:#Exchange",
        "ReadFlag": true, "SuppressReadReceipts": false,
        "FolderIds": [ { "__type": "FolderId:#Exchange", "Id": folder.id } ],
        "ItemIdsToExclude": []
      }
    };
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const j = await api('MarkAllItemsAsRead', body);
        if (j.Body.ResponseMessages.Items[0].ResponseClass === 'Success') return true;
      } catch (e) { if (attempt === 1) console.warn('[Mark all read]', folder.name, e.message); }
      await new Promise(r => setTimeout(r, 300));
    }
    return false;
  }

  // ===================== 3. Orchestration =====================
  async function run(btn) {
    const label = (t) => { btn.textContent = t; };
    try {
      btn.disabled = true;
      label('Getting token…');
      if (!await waitForToken()) {
        alert('Could not capture an auth token yet.\nLet your mailbox finish loading, click a folder, then try again.');
        return;
      }
      label('Listing folders…');
      const exclude = EXCLUDE_NAMES.map(s => s.toLowerCase());
      const targets = (await listAllFolders()).filter(f => f.unread > 0 && !exclude.includes((f.name || '').toLowerCase()));
      const totalUnread = targets.reduce((s, f) => s + f.unread, 0);

      if (!targets.length) { alert('No unread messages in any folder.'); return; }
      if (!confirm(`Mark ${targets.length} folder(s) / ${totalUnread} unread message(s) as read?`)) return;

      let done = 0, ok = 0;
      const q = targets.slice();
      const worker = async () => {
        while (q.length) {
          const f = q.shift();
          if (await markFolderRead(f)) ok++;
          done++;
          label(`Marking ${done}/${targets.length}…`);
        }
      };
      await Promise.all(Array.from({ length: CONCURRENCY }, worker));

      const failed = targets.length - ok;
      label(failed ? `Done: ${ok} ok, ${failed} failed` : `Done – ${ok} folders read`);
      console.log(`[Mark all read] ${ok}/${targets.length} folders, ${totalUnread} messages. API base: ${base()}`);
    } catch (e) {
      console.error('[Mark all read]', e);
      alert('Error: ' + e.message + '\n(API base: ' + base() + ')');
    } finally {
      btn.disabled = false;
      setTimeout(() => label(DEFAULT_LABEL), 5000);
    }
  }

  // ============ 4. Inject the button (persist across SPA re-renders) ============
  function addButton() {
    if (!document.body || document.getElementById('owa-mark-all-read')) return;
    const btn = document.createElement('button');
    btn.id = 'owa-mark-all-read';
    btn.textContent = DEFAULT_LABEL;
    Object.assign(btn.style, {
      position: 'fixed', bottom: '16px', right: '16px', zIndex: 2147483647,
      padding: '9px 13px', background: '#0f6cbd', color: '#fff', border: 'none',
      borderRadius: '6px', font: '600 13px "Segoe UI",system-ui,sans-serif',
      cursor: 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,.35)'
    });
    btn.onmouseenter = () => (btn.style.background = '#115ea3');
    btn.onmouseleave = () => (btn.style.background = '#0f6cbd');
    btn.addEventListener('click', () => run(btn));
    document.body.appendChild(btn);
  }
  const boot = () => { addButton(); setInterval(addButton, 3000); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
