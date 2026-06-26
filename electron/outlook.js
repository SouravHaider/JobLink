/* ────────────────────────────────────────────────────────────────────
   Outlook / Microsoft Graph OAuth for JobLink — desktop edition.

   Browser-popup OAuth (msal-browser) does NOT work in a packaged Electron
   app: after sign-in, Microsoft redirects to http://localhost, but the
   file:// app has nothing listening there, so it dies on "Failed to
   connect". The correct desktop pattern is the OAuth 2.0 authorization-code
   flow with PKCE and a loopback redirect that we actually listen on.

   Azure setup needed (one-time, in your existing app registration):
     • Authentication → Add a platform → "Mobile and desktop applications"
       → tick  http://localhost   (this enables loopback on any port)
     • Authentication → Advanced → "Allow public client flows" → Yes
   The SPA platform you already added can stay; it's just unused here.
   ──────────────────────────────────────────────────────────────────── */
const { BrowserWindow } = require("electron");
const http = require("http");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const AUTH = "https://login.microsoftonline.com/common/oauth2/v2.0";
const SCOPES = "Mail.Read offline_access openid profile email";

let _app = null;
let _safeStorage = null;
let _cache = { accessToken: null, expiresAt: 0 }; // in-memory access token

const tokenPath = () => path.join(_app.getPath("userData"), "joblink-outlook.bin");
const b64url = (buf) =>
  buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

/* ── Token storage (encrypted at rest) ──────────────────────────────── */
function saveTokens(obj) {
  if (!_safeStorage.isEncryptionAvailable()) throw new Error("Secure storage unavailable.");
  fs.writeFileSync(tokenPath(), _safeStorage.encryptString(JSON.stringify(obj)));
}
function loadTokens() {
  try { return JSON.parse(_safeStorage.decryptString(fs.readFileSync(tokenPath()))); }
  catch (_) { return null; }
}
function clearTokens() { try { fs.unlinkSync(tokenPath()); } catch (_) {} }

/* ── Helpers ────────────────────────────────────────────────────────── */
async function fetchEmailAddr(accessToken) {
  try {
    const r = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const d = await r.json();
    return d.userPrincipalName || d.mail || "";
  } catch (_) { return ""; }
}

async function exchangeCode(clientId, code, verifier, redirectUri) {
  const body = new URLSearchParams({
    client_id: clientId, grant_type: "authorization_code", code,
    redirect_uri: redirectUri, code_verifier: verifier, scope: SCOPES,
  });
  const r = await fetch(`${AUTH}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await r.json();
  if (data.error) return { ok: false, error: data.error_description || data.error };
  _cache = { accessToken: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 - 60000 };
  const email = await fetchEmailAddr(data.access_token);
  saveTokens({ refreshToken: data.refresh_token, email });
  return { ok: true, email };
}

async function getAccessToken(clientId) {
  if (_cache.accessToken && _cache.expiresAt > Date.now()) return _cache.accessToken;
  const saved = loadTokens();
  if (!saved || !saved.refreshToken) throw new Error("Not connected to Outlook.");
  const body = new URLSearchParams({
    client_id: clientId, grant_type: "refresh_token",
    refresh_token: saved.refreshToken, scope: SCOPES,
  });
  const r = await fetch(`${AUTH}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await r.json();
  if (data.error) {
    // Refresh token invalid/expired — force a fresh sign-in
    if (/invalid_grant|interaction_required/.test(data.error)) clearTokens();
    throw new Error(data.error_description || data.error);
  }
  _cache = { accessToken: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 - 60000 };
  if (data.refresh_token) saveTokens({ refreshToken: data.refresh_token, email: saved.email });
  return _cache.accessToken;
}

/* ── Interactive sign-in (loopback + PKCE) ──────────────────────────── */
function interactiveAuth(clientId) {
  return new Promise((resolve) => {
    const verifier = b64url(crypto.randomBytes(32));
    const challenge = b64url(crypto.createHash("sha256").update(verifier).digest());
    let win = null;
    let settled = false;
    const done = (result) => { if (!settled) { settled = true; resolve(result); } };

    const server = http.createServer(async (req, res) => {
      let url;
      try { url = new URL(req.url, "http://localhost"); } catch { res.end(); return; }
      const code = url.searchParams.get("code");
      const error = url.searchParams.get("error_description") || url.searchParams.get("error");
      if (!code && !error) { res.statusCode = 200; res.end(); return; }
      res.setHeader("Content-Type", "text/html");
      res.end(
        "<html><body style='font-family:-apple-system,system-ui;padding:48px;text-align:center'>" +
        "<h2 style='color:#0F766E'>JobLink</h2><p>You're signed in. You can close this window and return to JobLink.</p>" +
        "</body></html>"
      );
      const port = server.address().port;
      try { server.close(); } catch (_) {}
      if (win && !win.isDestroyed()) { win.removeAllListeners("closed"); win.close(); }
      if (error) { done({ ok: false, error }); return; }
      try { done(await exchangeCode(clientId, code, verifier, `http://localhost:${port}`)); }
      catch (e) { done({ ok: false, error: e.message }); }
    });

    server.on("error", (e) => done({ ok: false, error: "Could not start local sign-in listener: " + e.message }));

    server.listen(0, "127.0.0.1", () => {
      const port = server.address().port;
      const redirectUri = `http://localhost:${port}`;
      const authUrl = `${AUTH}/authorize?` + new URLSearchParams({
        client_id: clientId, response_type: "code", redirect_uri: redirectUri,
        response_mode: "query", scope: SCOPES,
        code_challenge: challenge, code_challenge_method: "S256",
        prompt: "select_account",
      }).toString();

      win = new BrowserWindow({
        width: 520, height: 700, title: "Sign in to Microsoft",
        autoHideMenuBar: true,
        webPreferences: { nodeIntegration: false, contextIsolation: true },
      });
      win.loadURL(authUrl);
      win.on("closed", () => { try { server.close(); } catch (_) {} done({ ok: false, error: "Sign-in window was closed before completing." }); });
    });
  });
}

/* ── Scan the inbox for job-related emails (Graph) ───────────────────── */
async function scan(clientId) {
  const token = await getAccessToken(clientId);
  const queries = [
    "unfortunately application", "interview invitation", "next steps application",
    "application unsuccessful", "application update",
  ];
  const sinceMs = Date.now() - 90 * 86400000;
  const seen = new Set();
  const emails = [];
  for (const q of queries) {
    const url =
      `https://graph.microsoft.com/v1.0/me/messages?$search="${encodeURIComponent(q)}"` +
      `&$top=20&$select=id,subject,from,receivedDateTime,bodyPreview`;
    try {
      const r = await fetch(url, {
        headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: "eventual" },
      });
      const d = await r.json();
      for (const e of d.value || []) {
        if (seen.has(e.id)) continue;
        seen.add(e.id);
        if (new Date(e.receivedDateTime).getTime() >= sinceMs) emails.push(e);
      }
    } catch (_) {}
  }
  return emails;
}

/* ── IPC ────────────────────────────────────────────────────────────── */
function register(ipcMain, app, safeStorage) {
  _app = app;
  _safeStorage = safeStorage;

  ipcMain.handle("outlook:status", () => {
    const t = loadTokens();
    return t ? { connected: true, email: t.email } : { connected: false };
  });

  ipcMain.handle("outlook:connect", async (_e, { clientId }) => {
    if (!clientId) return { ok: false, error: "Add your Microsoft client ID in Settings first." };
    try { return await interactiveAuth(clientId); }
    catch (e) { return { ok: false, error: e.message }; }
  });

  ipcMain.handle("outlook:disconnect", () => {
    clearTokens();
    _cache = { accessToken: null, expiresAt: 0 };
    return { ok: true };
  });

  ipcMain.handle("outlook:scan", async (_e, { clientId }) => {
    try { return { ok: true, emails: await scan(clientId) }; }
    catch (e) { return { ok: false, error: e.message }; }
  });
}

module.exports = { register };
