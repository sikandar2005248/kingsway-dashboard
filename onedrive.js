/* Microsoft sign-in (OAuth 2.0 authorization code + PKCE) and OneDrive access
   through Microsoft Graph. No libraries. Tokens stay in this browser only. */
(function () {
  'use strict';

  const CFG = window.KINGSWAY_CONFIG || {};
  const AUTHORITY = 'https://login.microsoftonline.com/' + (CFG.tenant || 'common');
  const SCOPES = 'openid profile offline_access User.Read Files.Read.All';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const K = { tok: 'kw.tokens', file: 'kw.file', pkce: 'kw.pkce' };

  const redirectUri = () => location.origin + location.pathname;
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* private mode */ } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } },
  };

  const b64url = (bytes) => btoa(String.fromCharCode.apply(null, Array.from(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const randomString = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));

  function configured() { return !!(CFG.clientId && /^[0-9a-f-]{36}$/i.test(CFG.clientId)); }

  async function signIn() {
    const verifier = randomString(48);
    const challenge = b64url(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))));
    const state = randomString(16);
    try { sessionStorage.setItem(K.pkce, JSON.stringify({ verifier, state, hash: location.hash })); } catch (e) { /* ignore */ }
    const q = new URLSearchParams({
      client_id: CFG.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
      scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account',
    });
    location.assign(AUTHORITY + '/oauth2/v2.0/authorize?' + q.toString());
  }

  async function tokenRequest(params) {
    const body = new URLSearchParams(Object.assign({ client_id: CFG.clientId, scope: SCOPES }, params));
    const r = await fetch(AUTHORITY + '/oauth2/v2.0/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.access_token) {
      const msg = j.error_description || j.error || ('sign-in failed (' + r.status + ')');
      throw new Error(msg.split('\r\n')[0]);
    }
    const t = { access: j.access_token, refresh: j.refresh_token || (store.get(K.tok) || {}).refresh, expires: Date.now() + (j.expires_in || 3600) * 1000 };
    store.set(K.tok, t);
    return t;
  }

  // Call on page load. Finishes a sign-in redirect if there is one.
  async function handleRedirect() {
    const p = new URLSearchParams(location.search);
    if (!p.has('code') && !p.has('error')) return null;
    let saved = null;
    try { saved = JSON.parse(sessionStorage.getItem(K.pkce) || 'null'); sessionStorage.removeItem(K.pkce); } catch (e) { /* ignore */ }
    history.replaceState(null, '', redirectUri() + ((saved && saved.hash) || ''));
    if (p.has('error')) throw new Error(p.get('error_description') || p.get('error'));
    if (!saved || saved.state !== p.get('state')) throw new Error('Sign-in was interrupted. Please sign in again.');
    await tokenRequest({ grant_type: 'authorization_code', code: p.get('code'), redirect_uri: redirectUri(), code_verifier: saved.verifier });
    return true;
  }

  async function accessToken() {
    const t = store.get(K.tok);
    if (!t) return null;
    if (t.access && t.expires - Date.now() > 120000) return t.access;
    if (!t.refresh) return null;
    try { return (await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })).access; }
    catch (e) { store.del(K.tok); return null; }
  }

  function signedIn() { return !!store.get(K.tok); }

  function signOut() {
    store.del(K.tok); store.del(K.file);
    location.assign(AUTHORITY + '/oauth2/v2.0/logout?post_logout_redirect_uri=' + encodeURIComponent(redirectUri()));
  }

  class NeedSignIn extends Error {}

  async function graph(path, opts) {
    const tok = await accessToken();
    if (!tok) throw new NeedSignIn('Please sign in.');
    const r = await fetch(path.startsWith('http') ? path : GRAPH + path, Object.assign({ cache: 'no-store' }, opts || {}, { headers: { Authorization: 'Bearer ' + tok } }));
    if (r.status === 401) { store.del(K.tok); throw new NeedSignIn('Your sign-in has expired. Please sign in again.'); }
    if (!r.ok) {
      const j = await r.json().catch(() => ({}));
      const e = new Error((j.error && j.error.message) || ('OneDrive request failed (' + r.status + ')'));
      e.status = r.status;
      throw e;
    }
    return r;
  }

  const shareId = (url) => 'u!' + btoa(unescape(encodeURIComponent(url.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

  async function me() { return (await graph('/me?$select=displayName,userPrincipalName,mail')).json(); }

  // Workbooks in the signed-in user's OneDrive whose name matches the hint.
  async function searchFiles(q) {
    const j = await (await graph("/me/drive/root/search(q='" + encodeURIComponent((q || 'xlsx').replace(/'/g, "''")) + "')?$top=100")).json();
    return (j.value || [])
      .filter((it) => it.file && /\.xlsx$/i.test(it.name) && !/^~\$/.test(it.name) && !/backup/i.test(it.name + ' ' + ((it.parentReference || {}).path || '')))
      .sort((a, b) => (b.lastModifiedDateTime || '').localeCompare(a.lastModifiedDateTime || ''))
      .map((it) => ({ driveId: it.parentReference.driveId, id: it.id, name: it.name, modified: it.lastModifiedDateTime, path: ((it.parentReference || {}).path || '').replace(/^\/drive\/root:/, '') || '/' }));
  }

  async function fromShareLink(url) {
    const it = await (await graph('/shares/' + shareId(url) + '/driveItem')).json();
    if (!it.file) throw new Error('That link is to a folder, not the workbook.');
    return { driveId: it.parentReference.driveId, id: it.id, name: it.name, modified: it.lastModifiedDateTime, path: 'Shared link' };
  }

  // Which workbook to read: the one picked before, the shared link in config.js, or the newest match in OneDrive.
  async function chooseFile() {
    const saved = store.get(K.file);
    if (saved) return saved;
    let f = null;
    if (CFG.shareLink) {
      try { f = await fromShareLink(CFG.shareLink); } catch (e) { if (e instanceof NeedSignIn) throw e; }
    }
    if (!f) {
      const hint = CFG.fileNameHint || 'Kingsway';
      const found = await searchFiles(hint);
      const pref = found.filter((x) => /finance/i.test(x.name));
      f = (pref.length ? pref : found)[0] || null;
    }
    if (f) store.set(K.file, f);
    return f;
  }

  function setFile(f) { if (f) store.set(K.file, f); else store.del(K.file); }

  // Downloads the latest copy of the workbook. Returns { buffer, item }.
  async function download(f) {
    const item = await (await graph('/drives/' + f.driveId + '/items/' + f.id + '?$select=id,name,size,lastModifiedDateTime,lastModifiedBy,@microsoft.graph.downloadUrl')).json();
    const url = item['@microsoft.graph.downloadUrl'];
    let buf;
    if (url) {
      const r = await fetch(url, { cache: 'no-store' });
      if (!r.ok) throw new Error('Download failed (' + r.status + ')');
      buf = await r.arrayBuffer();
    } else {
      buf = await (await graph('/drives/' + f.driveId + '/items/' + f.id + '/content')).arrayBuffer();
    }
    return { buffer: buf, item };
  }

  window.OneDrive = { configured, signIn, signOut, handleRedirect, signedIn, me, searchFiles, fromShareLink, chooseFile, setFile, download, NeedSignIn, redirectUri };
})();
