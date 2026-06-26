/* ────────────────────────────────────────────────────────────────────
   Generic IMAP mailbox connector for JobLink — runs in the Electron MAIN
   process (the sandboxed renderer can't open TCP sockets).

   Supports MULTIPLE accounts at once. Works with any provider that allows
   IMAP + an app-specific password: Gmail, Yahoo, iCloud, AOL, and custom
   servers. (Microsoft/Outlook disabled IMAP basic-auth → separate OAuth.)

   Credentials are encrypted with Electron safeStorage (macOS Keychain) and
   written to userData — never localStorage, never plaintext.
   ──────────────────────────────────────────────────────────────────── */
const { ImapFlow } = require("imapflow");
const { simpleParser } = require("mailparser");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const PROVIDERS = {
  gmail:  { label: "Gmail",  host: "imap.gmail.com",      port: 993 },
  yahoo:  { label: "Yahoo",  host: "imap.mail.yahoo.com", port: 993 },
  icloud: { label: "iCloud", host: "imap.mail.me.com",    port: 993 },
  aol:    { label: "AOL",    host: "imap.aol.com",        port: 993 },
  // "other" → caller supplies host
};

let _app = null;
let _safeStorage = null;
const credPath = () => path.join(_app.getPath("userData"), "joblink-mailbox.bin");

function resolveHost(provider, host) {
  if (provider && PROVIDERS[provider]) return { host: PROVIDERS[provider].host, port: PROVIDERS[provider].port };
  return { host: (host || "").trim(), port: 993 };
}

/* ── Storage: an array of accounts ──────────────────────────────────── */
function loadAccounts() {
  try {
    const parsed = JSON.parse(_safeStorage.decryptString(fs.readFileSync(credPath())));
    if (Array.isArray(parsed)) return parsed;
    // Migrate the older single-account format
    if (parsed && parsed.email) return [{ id: "legacy", ...parsed }];
    return [];
  } catch (_) { return []; }
}
function saveAccounts(arr) {
  if (!_safeStorage.isEncryptionAvailable()) throw new Error("Secure storage isn't available on this system.");
  fs.writeFileSync(credPath(), _safeStorage.encryptString(JSON.stringify(arr)));
}

function friendly(err) {
  const m = String((err && err.message) || err || "").toLowerCase();
  if (m.includes("invalid credentials") || m.includes("authenticationfailed") || m.includes("auth") || m.includes("login")) {
    return "The mail server rejected those credentials. Use an app password (from your email account's security settings), not your normal login password — and check the email address.";
  }
  if (m.includes("enotfound") || m.includes("etimedout") || m.includes("econnrefused") || m.includes("network")) {
    return "Couldn't reach the mail server. Check your internet connection (and the server address, if you entered one).";
  }
  return (err && err.message) || "Unknown error connecting to the mailbox.";
}

async function verify(host, port, email, appPassword) {
  const client = new ImapFlow({ host, port, secure: true, auth: { user: email, pass: appPassword }, logger: false });
  try { await client.connect(); await client.logout(); return { ok: true }; }
  catch (err) { try { await client.close(); } catch (_) {} return { ok: false, error: friendly(err) }; }
}

async function scanOne(acct, sinceDays = 90, cap = 250) {
  const client = new ImapFlow({
    host: acct.host, port: acct.port || 993, secure: true,
    auth: { user: acct.email, pass: acct.appPassword }, logger: false,
  });
  const out = [];
  await client.connect();
  const lock = await client.getMailboxLock("INBOX");
  try {
    const since = new Date(Date.now() - sinceDays * 86400000);
    let uids = await client.search({ since }, { uid: true });
    if (!Array.isArray(uids)) uids = [];
    const recent = uids.slice(-cap);
    if (recent.length) {
      for await (const msg of client.fetch(recent, { uid: true, envelope: true, source: true }, { uid: true })) {
        let body = "";
        try { const parsed = await simpleParser(msg.source); body = (parsed.text || parsed.html || "").replace(/\s+/g, " ").slice(0, 4000); }
        catch (_) {}
        const env = msg.envelope || {};
        const fromAddr = (env.from && env.from[0]) || {};
        out.push({
          id: String(msg.uid),
          subject: env.subject || "",
          from: { emailAddress: { name: fromAddr.name || "", address: fromAddr.address || "" } },
          receivedDateTime: (env.date ? new Date(env.date) : new Date()).toISOString(),
          bodyPreview: body,
        });
      }
    }
  } finally {
    lock.release();
  }
  await client.logout();
  return out;
}

function register(ipcMain, app, safeStorage) {
  _app = app;
  _safeStorage = safeStorage;

  ipcMain.handle("mailbox:list", () =>
    loadAccounts().map((a) => ({ id: a.id, email: a.email, provider: a.provider || "other" }))
  );

  ipcMain.handle("mailbox:connect", async (_e, { provider, host, email, appPassword }) => {
    if (!email || !appPassword) return { ok: false, error: "Enter both your email and an app password." };
    const resolved = resolveHost(provider, host);
    if (!resolved.host) return { ok: false, error: "Enter the IMAP server address for this provider." };
    const pass = appPassword.replace(/\s+/g, "");
    const v = await verify(resolved.host, resolved.port, email.trim(), pass);
    if (!v.ok) return v;
    const acct = {
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(36).slice(2),
      provider: provider || "other", host: resolved.host, port: resolved.port,
      email: email.trim(), appPassword: pass,
    };
    try {
      const accounts = loadAccounts().filter((a) => a.email.toLowerCase() !== acct.email.toLowerCase());
      accounts.push(acct);
      saveAccounts(accounts);
    } catch (err) { return { ok: false, error: err.message }; }
    return { ok: true, account: { id: acct.id, email: acct.email, provider: acct.provider } };
  });

  ipcMain.handle("mailbox:disconnect", (_e, { id }) => {
    const accounts = loadAccounts().filter((a) => a.id !== id);
    saveAccounts(accounts);
    return { ok: true };
  });

  ipcMain.handle("mailbox:scan", async (_e, opts) => {
    const accounts = loadAccounts();
    const groups = [];
    const errors = [];
    for (const a of accounts) {
      try { groups.push({ id: a.id, email: a.email, emails: await scanOne(a, (opts && opts.sinceDays) || 90) }); }
      catch (e) { errors.push({ email: a.email, error: friendly(e) }); }
    }
    return { ok: true, groups, errors };
  });
}

module.exports = { register };
