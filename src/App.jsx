import { useState, useEffect, useMemo, useRef } from "react";
import Papa from "papaparse";

/* ------------------------------------------------------------------ */
/*  JobLink — sponsorship-aware application tracker                    */
/*  Data lives in your browser's localStorage; use Export for backups. */
/* ------------------------------------------------------------------ */

/* ── Outlook / Microsoft Graph ──────────────────────────────────────
   Setup (one-time, ~5 min):
   1. Go to https://portal.azure.com → "App registrations" → New registration
   2. Name it anything (e.g. "JobLink"), choose "Accounts in any org + personal"
   3. Under "Redirect URI" pick Single-page application → http://localhost:5173
      (add your production URL too if you deploy it)
   4. Click Register, copy the "Application (client) ID"
   5. Go to API permissions → Add → Microsoft Graph → Delegated → Mail.Read → Add
   6. Paste your client ID below ↓
   ──────────────────────────────────────────────────────────────── */
// The Outlook (Microsoft) client ID is now entered in Settings → no code
// edit needed. It's stored in localStorage with the rest of your settings.
//
// Use a fixed redirect URI that works in both the browser and Electron.
// In Electron, window.location.origin is "file://" which Azure rejects,
// so we always use http://localhost (no port).
// Add http://localhost as a Single-page application redirect URI in your Azure app registration.
const buildMsalConfig = (clientId) => ({
  auth: {
    clientId,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost",
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
});

let _msalInstance = null;
let _msalClientId = null;
const getMsal = async (clientId) => {
  if (!clientId) return null;
  // Rebuild if the client ID changed since last time
  if (_msalInstance && _msalClientId === clientId) return _msalInstance;
  const { PublicClientApplication } = await import("@azure/msal-browser");
  _msalInstance = new PublicClientApplication(buildMsalConfig(clientId));
  await _msalInstance.initialize();
  _msalClientId = clientId;
  return _msalInstance;
};

const GRAPH_SCOPES = ["Mail.Read"];

/* ── Email classification keywords ─────────────────────────────── */
const REJECTION_KW = [
  "unfortunately", "not successful", "not been selected", "not moving forward",
  "not progressing", "regret to inform", "won't be progressing", "decided not to",
  "position has been filled", "chosen another", "unsuccessful", "not taken forward",
  "not shortlisted", "application was not", "unable to offer", "not been shortlisted",
  "after careful consideration", "we will not", "will not be progressing",
];
const INTERVIEW_KW = [
  "invite you to interview", "invitation to interview", "interview invitation",
  "schedule an interview", "phone interview", "video interview", "next steps",
  "phone screen", "video call", "schedule a call", "like to invite",
  "progress your application", "pleased to invite", "move forward with",
  "assessment centre", "assessment center", "like to discuss your application",
  "spoken with you about", "advance to the next",
];

const classifyEmail = (subject, body) => {
  const text = (subject + " " + body).toLowerCase();
  if (REJECTION_KW.some((k) => text.includes(k))) return "rejected";
  if (INTERVIEW_KW.some((k) => text.includes(k))) return "interview";
  return null;
};

/* ── Graph API helpers ──────────────────────────────────────────── */
const getGraphToken = async (msal) => {
  const accounts = msal.getAllAccounts();
  const request = { scopes: GRAPH_SCOPES, account: accounts[0] };
  try {
    const res = await msal.acquireTokenSilent(request);
    return res.accessToken;
  } catch {
    const res = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return res.accessToken;
  }
};

const fetchJobEmails = async (token) => {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90);
  const since = cutoff.toISOString();

  // Two targeted searches — rejections and interview invites
  const queries = [
    "unfortunately application",
    "interview invitation",
    "next steps application",
    "application unsuccessful",
    "application update",
  ];

  const seen = new Set();
  const emails = [];

  for (const q of queries) {
    const url =
      `https://graph.microsoft.com/v1.0/me/messages` +
      `?$search="${encodeURIComponent(q)}"` +
      `&$top=20` +
      `&$select=id,subject,from,receivedDateTime,bodyPreview`;
    try {
      const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      for (const email of data.value || []) {
        if (!seen.has(email.id)) {
          seen.add(email.id);
          // Only keep recent ones
          if (new Date(email.receivedDateTime) >= new Date(since)) {
            emails.push(email);
          }
        }
      }
    } catch (_) {
      // skip failed query
    }
  }

  return emails;
};

const matchEmailToApp = (email, apps) => {
  const subject = (email.subject || "").toLowerCase();
  const senderName = (email.from?.emailAddress?.name || "").toLowerCase();
  const senderDomain = (email.from?.emailAddress?.address || "").split("@")[1] || "";
  // strip TLD: "google.com" → "google"
  const domainWord = senderDomain.split(".")[0].toLowerCase();

  for (const app of apps) {
    if (!app.company || app.company.trim().length < 3) continue;
    const norm = normName(app.company);
    const words = norm.split(" ").filter((w) => w.length > 3);
    // Match if sender name, domain, or subject contains the company or its key word
    const hits = words.filter(
      (w) =>
        senderName.includes(w) ||
        domainWord.includes(w) ||
        subject.includes(w)
    );
    if (hits.length > 0) return app;
  }
  return null;
};

/* ------------------------------------------------------------------ */

const STATUSES = [
  { key: "wishlist",  label: "Wishlist",  color: "#8B95A3" },
  { key: "applied",   label: "Applied",   color: "#3B6FB0" },
  { key: "screening", label: "Screening", color: "#B45309" },
  { key: "interview", label: "Interview", color: "#0F766E" },
  { key: "offer",     label: "Offer",     color: "#15803D" },
  { key: "rejected",  label: "Rejected",  color: "#94A3B8" },
];
const statusOf = (k) => STATUSES.find((s) => s.key === k) || STATUSES[0];

const SPONSOR = {
  yes:     { label: "Sponsor ✓", color: "#15803D", bg: "#E7F4EC" },
  no:      { label: "No licence", color: "#B91C1C", bg: "#FBEAEA" },
  unknown: { label: "Check register", color: "#7A8493", bg: "#EEF1F4" },
};

const STORE_KEY = "joblink_data_v1";
const SETTINGS_KEY = "joblink_settings_v1";

const ROLE_CHIPS = [
  "junior cyber security analyst", "SOC analyst", "junior web developer",
  "QA test analyst", "graduate software developer", "supply chain analyst",
];
// Boards are stored as data (with {q}/{c} tokens) so you can add and remove
// your own. Sponsor-only boards have no tokens — they just open the site.
const DEFAULT_BOARDS = [
  { id: "indeed",    name: "Indeed",             pattern: "https://uk.indeed.com/jobs?q={q}&l={c}" },
  { id: "linkedin",  name: "LinkedIn",           pattern: "https://www.linkedin.com/jobs/search/?keywords={q}&location={c}" },
  { id: "reed",      name: "Reed",               pattern: "https://www.reed.co.uk/jobs?keywords={q}&location={c}" },
  { id: "totaljobs", name: "Totaljobs",          pattern: "https://www.totaljobs.com/jobs?Keywords={q}&Location={c}" },
  { id: "ukvisa",    name: "UKVisaJobs",         pattern: "https://www.ukvisajobs.com", sponsorOnly: true },
  { id: "huntuk",    name: "HuntUKVisaSponsors", pattern: "https://www.huntukvisasponsors.com", sponsorOnly: true },
  { id: "govukreg",  name: "gov.uk register",    pattern: "https://www.gov.uk/government/publications/register-of-licensed-sponsors-workers", sponsorOnly: true },
];
// "Date posted" recency options (value is hours; 0 = any time)
const DATE_POSTED = [
  { label: "Any time",     hours: 0 },
  { label: "Past 6 hours", hours: 6 },
  { label: "Past 24 hours", hours: 24 },
  { label: "Past 48 hours", hours: 48 },
  { label: "Past 72 hours", hours: 72 },
  { label: "Past week",    hours: 168 },
];
const buildBoardUrl = (board, q, c, hours) => {
  if (!/\{q\}|\{c\}|\{d\}/.test(board.pattern)) return board.pattern; // static link
  let url = board.pattern
    .replace(/\{q\}/g, encodeURIComponent(q))
    .replace(/\{c\}/g, encodeURIComponent(c));
  const days = hours ? Math.max(1, Math.ceil(hours / 24)) : 0;
  if (/\{d\}/.test(board.pattern)) {
    // Custom board controls placement of the recency value (in days)
    url = url.replace(/\{d\}/g, hours ? String(days) : "");
  } else if (hours) {
    // Auto-apply the recency filter for boards that support it
    const sep = url.includes("?") ? "&" : "?";
    const host = url.toLowerCase();
    if (host.includes("linkedin.com")) url += `${sep}f_TPR=r${hours * 3600}`;     // seconds
    else if (host.includes("indeed.")) url += `${sep}fromage=${days}`;            // whole days
  }
  return url;
};

const uid = () => Math.random().toString(36).slice(2, 10);
const today = () => new Date().toISOString().slice(0, 10);
const addDays = (iso, n) => {
  const d = new Date(iso + "T00:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
};
const fmt = (d) =>
  d ? new Date(d + "T00:00:00").toLocaleDateString("en-GB", { day: "2-digit", month: "short" }) : "—";
const fmtLong = (d) =>
  d ? new Date(d + "T00:00:00").toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const fmtEmail = (iso) =>
  iso ? new Date(iso).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "";
const salaryNum = (s) => {
  const n = parseInt(String(s || "").replace(/[^\d]/g, ""), 10);
  return isNaN(n) ? null : n;
};
const normName = (s) =>
  String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\b(limited|ltd|plc|llp|llc|uk|holdings|group|company|co)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

/* ── Job-description parser (100% local — no service, no LinkedIn/Indeed
      connection needed; it just reads the text you paste) ───────────── */
const UK_CITIES = [
  "London", "Manchester", "Birmingham", "Leeds", "Glasgow", "Edinburgh", "Bristol",
  "Liverpool", "Sheffield", "Cardiff", "Belfast", "Nottingham", "Newcastle", "Southampton",
  "Brighton", "Cambridge", "Oxford", "Reading", "Leicester", "Coventry", "Milton Keynes",
  "Aberdeen", "Bath", "York", "Norwich", "Exeter", "Derby", "Portsmouth", "Plymouth",
  "Swansea", "Hull", "Wolverhampton", "Stoke-on-Trent", "Slough", "Watford", "Luton",
  "Bracknell", "Basingstoke", "Warrington", "Preston", "Swindon", "Crawley", "Maidenhead",
  "Guildford", "Farnborough", "Cheltenham", "Gloucester", "Ipswich", "Chelmsford",
];
const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const normalizeSalary = (s) => {
  if (!s) return "";
  const t = String(s).toLowerCase().replace(/,/g, "").trim();
  if (/k$/.test(t)) return String(Math.round(parseFloat(t) * 1000));
  const n = parseInt(t, 10);
  return isNaN(n) ? "" : String(n);
};

// Parse common UK date formats into YYYY-MM-DD (or "" if it can't)
const parseLooseDate = (s) => {
  if (!s) return "";
  const t = s.trim().replace(/(\d)(st|nd|rd|th)/gi, "$1");
  const dmy = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (dmy) {
    let [, d, mo, y] = dmy;
    if (y.length === 2) y = "20" + y;
    const dt = new Date(Number(y), Number(mo) - 1, Number(d));
    return isNaN(dt.getTime()) ? "" : dt.toISOString().slice(0, 10);
  }
  const dt = new Date(t);
  return isNaN(dt.getTime()) ? "" : dt.toISOString().slice(0, 10);
};

const parseJobText = (raw) => {
  const text = String(raw || "").replace(/\r/g, "");
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  const out = { description: text.trim().slice(0, 8000) };

  // ── Salary — only fill it when it genuinely looks like an annual figure.
  //    We'd rather leave it blank than guess wrong, so we require either a
  //    range, an explicit annual word, or a value of at least £15,000, and we
  //    reject anything that reads as hourly/daily. ──
  const salRe = /£\s?(\d{1,3}(?:,\d{3})+|\d{2,3}k|\d{4,6})(?:\s*(?:-|–|—|to)\s*£?\s?(\d{1,3}(?:,\d{3})+|\d{2,3}k|\d{4,6}))?/i;
  const sm = text.match(salRe);
  if (sm) {
    const ctx = text.slice(Math.max(0, sm.index - 16), sm.index + sm[0].length + 18).toLowerCase();
    const hourlyOrDaily = /(per hour|\/hour|\/hr|an hour|hourly|p\/h|\bph\b|per day|a day|day rate|daily)/.test(ctx);
    const annualWord = /(per annum|per year|p\.?a\.?\b|annually|annual|\/year|\/yr|\/annum|salary)/.test(ctx);
    const top = normalizeSalary(sm[2] || sm[1]);
    const val = parseInt(top || "0", 10);
    const looksAnnual = !!sm[2] || annualWord || val >= 15000;
    if (!hourlyOrDaily && looksAnnual && val > 0) out.salary = top;
  }

  // ── City — look only in the header block (first lines) or an explicit
  //    "Location:" line, not the whole body. Job ads name lots of cities in
  //    passing ("offices in London, Leeds…"), which used to win wrongly. ──
  const cityIn = (txt) => {
    if (!txt) return "";
    let best = "", bestIdx = Infinity;
    for (const c of UK_CITIES) {
      const m = txt.match(new RegExp("(?:^|\\b)" + escRe(c) + "(?:\\b|$)", "i"));
      if (m && m.index < bestIdx) { best = c; bestIdx = m.index; }
    }
    return best;
  };
  const header = lines.slice(0, 3).join("  •  ");
  const locLine = (lines.find((l) => /^(location|based|office|where)\b/i.test(l)) || "");
  out.city =
    cityIn(locLine) ||
    cityIn(header) ||
    (/\bremote\b/i.test(header) || /^(location|based)\b.*\bremote\b/i.test(locLine) ? "Remote" : "");
  if (!out.city) delete out.city; // leave the field's default rather than guess

  // ── Posting time ("2 weeks ago", "Posted 3 days ago", "today") — this
  //    sits next to the location in the header on LinkedIn/Indeed. ──
  const headerLow = header.toLowerCase();
  const rel = header.match(/(\d+)\+?\s*(hour|day|week|month)s?\s+ago/i);
  if (rel) {
    out.posted = rel[0].replace(/\s+/g, " ").toLowerCase();
    const n = parseInt(rel[1], 10);
    const unit = rel[2].toLowerCase();
    const days = unit === "hour" ? 0 : unit === "day" ? n : unit === "week" ? n * 7 : n * 30;
    out.postedDate = addDays(today(), -days);
  } else if (/\b(posted\s+)?yesterday\b/.test(headerLow)) {
    out.posted = "yesterday";
    out.postedDate = addDays(today(), -1);
  } else if (/\bposted\s+today\b|\bjust posted\b|\btoday\b/.test(headerLow)) {
    out.posted = "today";
    out.postedDate = today();
  }

  // ── Role title (first meaningful line) ──
  const skip = /(logo|sign in|^save$|^apply|^share|applicant|^\d|easy apply|posted|ago$|·)/i;
  for (const l of lines.slice(0, 5)) {
    if (l.length >= 3 && l.length <= 90 && !skip.test(l)) { out.role = l; break; }
  }

  // ── Company (best-effort: "at X", "Company · Location", or 2nd line) ──
  // LinkedIn/Indeed put the employer before a middot/bullet; take that first
  // segment regardless of how long the whole "Company · Location · 2w ago" line is.
  const atM = text.match(/\bat\s+([A-Z][A-Za-z0-9&.,'’\-]+(?:\s+[A-Z][A-Za-z0-9&.,'’\-]+){0,4})/);
  const dotLine = lines.find((l) => /\s[·•]\s/.test(l));
  const dotCompany = dotLine ? dotLine.split(/\s[·•]\s/)[0].trim() : "";
  if (atM) out.company = atM[1].trim();
  else if (dotCompany.length >= 2 && dotCompany.length <= 50) out.company = dotCompany;
  else if (lines[1] && lines[1].length < 60 && lines[1] !== out.role) out.company = lines[1].trim();

  // ── Security clearance flag ──
  out.clearance = /\b(sc|dv)\s+clearance|security clearance|sc cleared|dv cleared|developed vetting|baseline personnel security|bpss/i.test(text);

  // ── Application closing date ("closing date:", "apply by", "deadline") ──
  const cm = text.match(/(?:closing date|applications?\s+close[sd]?|apply by|application deadline|deadline)[:\s\-–]*([0-3]?\d[\/.\-][0-1]?\d[\/.\-]\d{2,4}|\d{4}-\d{2}-\d{2}|[0-3]?\d(?:st|nd|rd|th)?\s+[A-Za-z]{3,9},?\s+\d{4}|[A-Za-z]{3,9}\s+[0-3]?\d(?:st|nd|rd|th)?,?\s+\d{4})/i);
  if (cm) {
    const iso = parseLooseDate(cm[1]);
    if (iso) out.closingDate = iso;
  }

  // ── First URL → job link ──
  const urlM = text.match(/https?:\/\/\S+/);
  if (urlM) out.link = urlM[0].replace(/[).,]+$/, "");

  return out;
};

const loadJSON = (key, fallback) => {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (_) { return fallback; }
};

/* ── Tiny IndexedDB key/value store ───────────────────────────────────
   The sponsor register is far too big for localStorage (100k+ rows), so
   we persist it in IndexedDB, which handles large data and Map values.   */
const IDB_DB = "joblink";
const IDB_STORE = "kv";
const idbOpen = () =>
  new Promise((res, rej) => {
    const r = indexedDB.open(IDB_DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(IDB_STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
const idbSet = async (key, val) => {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(val, key);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
};
const idbGet = async (key) => {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const rq = tx.objectStore(IDB_STORE).get(key);
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
  });
};
const idbDel = async (key) => {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).delete(key);
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
};

const blank = () => ({
  id: uid(), company: "", role: "", city: "London", salary: "",
  link: "", sponsor: "unknown", sponsorDetail: "", status: "wishlist",
  dateApplied: today(), nextDate: "", note: "", cvNote: "", clearance: false,
  description: "", pinned: false, posted: "", postedDate: "", closingDate: "",
});

export default function App() {
  const [apps, setApps] = useState(() => loadJSON(STORE_KEY, []));
  const [settings, setSettings] = useState(() => loadJSON(SETTINGS_KEY, { visaExpiry: "", floor: "33400" }));
  const [filter, setFilter] = useState("all");
  const [editing, setEditing] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [linkInput, setLinkInput] = useState("");
  const [linkBusy, setLinkBusy] = useState(false);
  const [linkNote, setLinkNote] = useState("");
  const [showPaste, setShowPaste] = useState(false);
  const [register, setRegister] = useState(null);
  const [regBusy, setRegBusy] = useState(false);
  const [regErr, setRegErr] = useState("");
  const [showBoards, setShowBoards] = useState(false);
  const [roleChips, setRoleChips] = useState(() => loadJSON("joblink_rolechips_v1", ROLE_CHIPS));
  const [searchQ, setSearchQ] = useState(() => (loadJSON("joblink_rolechips_v1", ROLE_CHIPS)[0] || ""));
  const [searchCity, setSearchCity] = useState("London");
  const [datePosted, setDatePosted] = useState(0);
  const [addSponsorKw, setAddSponsorKw] = useState(true);
  const [boards, setBoards] = useState(() => loadJSON("joblink_boards_v1", DEFAULT_BOARDS));
  const [manageBoards, setManageBoards] = useState(false);
  const [newBoardName, setNewBoardName] = useState("");
  const [newBoardUrl, setNewBoardUrl] = useState("");

  // ── Outlook state ──────────────────────────────────────────────────
  const outlookClientId = (settings.outlookClientId || "").trim();
  const hasOutlook = !!outlookClientId;
  // In the packaged desktop app we do OAuth in the main process (loopback);
  // msal-browser only works in a plain browser, so prefer the native bridge.
  const outlookNative = typeof window !== "undefined" && !!window.outlook;
  const [outlookAccount, setOutlookAccount] = useState(null);
  const [emailSuggestions, setEmailSuggestions] = useState(() => loadJSON("joblink_email_suggestions_v1", []));
  const [scanBusy, setScanBusy] = useState(false);
  const [scanErr, setScanErr] = useState("");
  const [outlookReady, setOutlookReady] = useState(false);

  // ── Generic IMAP mailboxes (Gmail / Yahoo / iCloud / AOL / custom) ────
  const hasMailbox = typeof window !== "undefined" && !!window.mailbox;
  const [mailboxAccounts, setMailboxAccounts] = useState([]); // [{ id, email, provider }]
  const [showEmailConnect, setShowEmailConnect] = useState(false);

  const regIndex = useRef(new Map());
  const fileRef = useRef(null);
  const importRef = useRef(null);
  const busyRef = useRef(false);       // guards against overlapping scans
  const scanAllRef = useRef(null);     // latest scanAll for timers

  // Expose "New Application" trigger so the native macOS menu (Cmd+N) can call it
  useEffect(() => {
    window.__jlNewApp = () => setEditing(blank());
    return () => { delete window.__jlNewApp; };
  }, []);

  // Restore Outlook session on load (and whenever the client ID changes)
  useEffect(() => {
    if (!outlookClientId) { setOutlookAccount(null); setOutlookReady(false); return; }
    if (outlookNative) {
      window.outlook.status().then((s) => {
        if (s && s.connected) setOutlookAccount({ name: s.email, username: s.email });
        setOutlookReady(true);
      }).catch(() => {});
      return;
    }
    getMsal(outlookClientId).then((msal) => {
      if (!msal) return;
      const accounts = msal.getAllAccounts();
      if (accounts.length > 0) setOutlookAccount(accounts[0]);
      setOutlookReady(true);
    }).catch(() => {});
  }, [outlookClientId]);

  // Restore connected IMAP mailboxes on load
  const refreshMailboxes = () => {
    if (!hasMailbox) return Promise.resolve([]);
    return window.mailbox.list().then((list) => { setMailboxAccounts(list || []); return list || []; }).catch(() => []);
  };
  useEffect(() => { refreshMailboxes(); }, []);

  // Keep the timer pointed at the freshest scanAll (avoids stale closures)
  useEffect(() => { scanAllRef.current = scanAll; });

  // Auto-scan: once shortly after launch, then on an interval (Settings).
  // On by default; turn off or change the interval in Settings.
  const autoScanOn = settings.autoScan !== false;
  const autoScanMins = Math.min(60, Number(settings.autoScanMins) || 30); // 1 hour max
  const inboxCount = (outlookAccount ? 1 : 0) + mailboxAccounts.length;
  useEffect(() => {
    if (!autoScanOn || inboxCount === 0) return;
    const t0 = setTimeout(() => scanAllRef.current && scanAllRef.current({ silent: true }), 1500);
    const iv = setInterval(() => scanAllRef.current && scanAllRef.current({ silent: true }), autoScanMins * 60000);
    return () => { clearTimeout(t0); clearInterval(iv); };
  }, [autoScanOn, autoScanMins, inboxCount]);

  useEffect(() => {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(apps)); } catch (e) { console.error(e); }
  }, [apps]);
  useEffect(() => {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch (e) { console.error(e); }
  }, [settings]);
  useEffect(() => {
    try { localStorage.setItem("joblink_email_suggestions_v1", JSON.stringify(emailSuggestions)); } catch (e) { console.error(e); }
  }, [emailSuggestions]);
  useEffect(() => {
    try { localStorage.setItem("joblink_boards_v1", JSON.stringify(boards)); } catch (e) { console.error(e); }
  }, [boards]);
  useEffect(() => {
    try { localStorage.setItem("joblink_rolechips_v1", JSON.stringify(roleChips)); } catch (e) { console.error(e); }
  }, [roleChips]);

  // Restore the saved sponsor register (persisted in IndexedDB) on load
  useEffect(() => {
    idbGet("sponsorRegister").then((saved) => {
      if (saved && saved.index instanceof Map && saved.index.size > 0) {
        regIndex.current = saved.index;
        setRegister({ count: saved.index.size, fileName: saved.fileName || "saved register", restored: true });
      }
    }).catch(() => {});
  }, []);

  /* ---------- Job boards ---------- */
  const addBoard = () => {
    const name = newBoardName.trim();
    let url = newBoardUrl.trim();
    if (!name || !url) return;
    if (!/^https?:\/\//i.test(url)) url = "https://" + url;
    setBoards((prev) => [...prev, { id: uid(), name, pattern: url }]);
    setNewBoardName("");
    setNewBoardUrl("");
  };
  const deleteBoard = (id) => setBoards((prev) => prev.filter((b) => b.id !== id));
  const resetBoards = () => setBoards(DEFAULT_BOARDS);

  /* ---------- Saved searches (role chips) ---------- */
  const deleteChip = (chip) => setRoleChips((prev) => prev.filter((c) => c !== chip));
  const canSaveSearch =
    searchQ.trim().length > 0 &&
    !roleChips.some((c) => c.toLowerCase() === searchQ.trim().toLowerCase());
  const saveCurrentSearch = () => {
    const v = searchQ.trim();
    if (!v || !canSaveSearch) return;
    setRoleChips((prev) => [...prev, v]);
  };

  /* ---------- Outlook auth ---------- */
  const signInOutlook = async () => {
    setScanErr("");
    if (!outlookClientId) { setScanErr("Add your Microsoft client ID in Settings to enable Outlook sync."); return; }
    if (outlookNative) {
      // Desktop: OAuth happens in the main process (loopback + PKCE)
      setScanBusy(true);
      try {
        const res = await window.outlook.connect({ clientId: outlookClientId });
        if (res && res.ok) setOutlookAccount({ name: res.email, username: res.email });
        else setScanErr("Outlook sign-in failed: " + ((res && res.error) || "unknown error"));
      } catch (e) {
        setScanErr("Outlook sign-in failed: " + (e.message || "unknown error"));
      } finally { setScanBusy(false); }
      return;
    }
    const msal = await getMsal(outlookClientId);
    if (!msal) { setScanErr("Add your Microsoft client ID in Settings to enable Outlook sync."); return; }
    try {
      const res = await msal.loginPopup({ scopes: GRAPH_SCOPES });
      setOutlookAccount(res.account);
    } catch (e) {
      setScanErr("Sign-in cancelled or failed.");
    }
  };

  const signOutOutlook = async () => {
    if (outlookNative) {
      await window.outlook.disconnect();
      setOutlookAccount(null);
      setEmailSuggestions([]);
      return;
    }
    const msal = await getMsal(outlookClientId);
    if (!msal) return;
    const accounts = msal.getAllAccounts();
    if (accounts[0]) await msal.logoutPopup({ account: accounts[0] });
    setOutlookAccount(null);
    setEmailSuggestions([]);
  };

  /* ---------- Email scan ---------- */
  // Shared: turn a list of (Graph-shaped) emails into status suggestions.
  // Both the Outlook and iCloud paths feed into this. Returns the count added.
  const ingestEmails = (emails, source) => {
    const newSuggestions = [];
    const existingIds = new Set(emailSuggestions.map((s) => s.emailId));
    for (const email of emails) {
      const emailKey = (source || "") + ":" + email.id;
      if (existingIds.has(emailKey)) continue; // already shown
      const type = classifyEmail(email.subject || "", email.bodyPreview || "");
      if (!type) continue;
      const app = matchEmailToApp(email, apps);
      if (!app) continue;
      if (app.status === type) continue; // already at that status
      newSuggestions.push({
        id: uid(),
        emailId: emailKey,
        appId: app.id,
        company: app.company,
        role: app.role,
        subject: email.subject,
        preview: email.bodyPreview,
        date: email.receivedDateTime,
        suggestedStatus: type,
        currentStatus: app.status,
        source: source || "email",
      });
    }
    setEmailSuggestions((prev) => [...newSuggestions, ...prev].slice(0, 20));
    return newSuggestions.length;
  };

  // Fetch Outlook's job emails (throws on error) — used by the combined scan
  const getOutlookEmails = async () => {
    if (outlookNative) {
      const res = await window.outlook.scan({ clientId: outlookClientId });
      if (!res || !res.ok) throw new Error((res && res.error) || "Scan failed");
      return res.emails || [];
    }
    const msal = await getMsal(outlookClientId);
    if (!msal) throw new Error("Add your Microsoft client ID in Settings first");
    const token = await getGraphToken(msal);
    return await fetchJobEmails(token);
  };

  /* ---------- IMAP mailboxes (Gmail / Yahoo / iCloud / AOL / custom) ---------- */
  const connectMailbox = async (creds) => {
    const res = await window.mailbox.connect(creds);
    if (res && res.ok) {
      await refreshMailboxes();
      setScanErr("");
      return { ok: true };
    }
    return { ok: false, error: (res && res.error) || "Couldn't connect to that mailbox." };
  };

  const disconnectInbox = async (inbox) => {
    if (inbox.kind === "outlook") { await signOutOutlook(); return; }
    await window.mailbox.disconnect(inbox.id);
    await refreshMailboxes();
  };

  // Every connected inbox (Outlook OAuth + any IMAP accounts)
  const allInboxes = [
    ...(outlookAccount ? [{ id: "outlook", kind: "outlook", label: outlookAccount.name || outlookAccount.username }] : []),
    ...mailboxAccounts.map((a) => ({ id: a.id, kind: "mailbox", label: a.email, provider: a.provider })),
  ];

  // Scan every connected inbox and merge the results
  const scanAll = async (opts = {}) => {
    if (busyRef.current) return; // don't overlap (e.g. a timer firing mid-scan)
    busyRef.current = true;
    if (!opts.silent) setScanErr("");
    setScanBusy(true);
    let added = 0;
    const problems = [];
    try {
      if (outlookAccount) {
        try { added += ingestEmails(await getOutlookEmails(), "outlook"); }
        catch (e) { problems.push("Outlook: " + (e.message || "failed")); }
      }
      if (mailboxAccounts.length) {
        const res = await window.mailbox.scan({ sinceDays: 90 });
        if (res && res.ok) {
          for (const g of res.groups || []) added += ingestEmails(g.emails || [], "mailbox:" + g.email);
          for (const er of res.errors || []) problems.push(er.email + ": " + er.error);
        } else {
          problems.push((res && res.error) || "Mailbox scan failed");
        }
      }
      if (problems.length) setScanErr("Some inboxes had issues — " + problems.join("; "));
      else if (added === 0 && !opts.silent) setScanErr("No new matches found. Your inboxes are up to date.");
    } catch (e) {
      if (!opts.silent) setScanErr("Scan failed: " + (e.message || "unknown error"));
    } finally {
      setScanBusy(false);
      busyRef.current = false;
    }
  };

  const applySuggestion = (sug) => {
    setStatus(sug.appId, sug.suggestedStatus);
    setEmailSuggestions((prev) => prev.filter((s) => s.id !== sug.id));
  };

  const dismissSuggestion = (id) => {
    setEmailSuggestions((prev) => prev.filter((s) => s.id !== id));
  };

  /* ---------- sponsor register CSV ---------- */
  const loadRegister = (file) => {
    if (!file) return;
    setRegBusy(true);
    setRegErr("");
    const idx = new Map();
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      chunk: (results) => {
        const fields = results.meta.fields || [];
        const orgKey = fields.find((f) => /organisation/i.test(f)) || fields[0];
        const townKey = fields.find((f) => /town|city/i.test(f));
        const ratingKey = fields.find((f) => /type.*rating|rating/i.test(f));
        const routeKey = fields.find((f) => /route/i.test(f));
        for (const row of results.data) {
          const route = routeKey ? String(row[routeKey] || "") : "";
          if (routeKey && !/skilled worker/i.test(route)) continue;
          const name = row[orgKey];
          if (!name) continue;
          const key = normName(name);
          if (!key) continue;
          const rating = ratingKey ? String(row[ratingKey] || "") : "";
          const existing = idx.get(key);
          if (!existing || (/\bA rating\b/i.test(rating) && !/\bA rating\b/i.test(existing.rating))) {
            idx.set(key, { name: String(name).trim(), town: townKey ? String(row[townKey] || "").trim() : "", rating });
          }
        }
      },
      complete: () => {
        regIndex.current = idx;
        setRegister({ count: idx.size, fileName: file.name });
        setRegBusy(false);
        // Persist so it survives reloads until a new CSV is loaded
        idbSet("sponsorRegister", { index: idx, fileName: file.name, savedAt: Date.now() }).catch(() => {});
      },
      error: () => {
        setRegBusy(false);
        setRegErr("Couldn't parse that file — make sure it's the CSV from the gov.uk register page.");
      },
    });
  };

  const lookupSponsor = (companyName) => {
    const q = normName(companyName);
    if (!q || q.length < 3 || regIndex.current.size === 0) return null;
    const exact = regIndex.current.get(q);
    if (exact) return { match: "exact", ...exact };
    for (const [key, val] of regIndex.current) {
      if (key.startsWith(q + " ") || key === q) return { match: "partial", ...val };
    }
    for (const [key, val] of regIndex.current) {
      if (key.includes(q) || (q.length > 6 && q.includes(key))) return { match: "partial", ...val };
    }
    return { match: "none" };
  };

  /* ---------- derived ---------- */
  const counts = useMemo(() => {
    const c = Object.fromEntries(STATUSES.map((s) => [s.key, 0]));
    apps.forEach((a) => { c[a.status] = (c[a.status] || 0) + 1; });
    return c;
  }, [apps]);

  const active = apps.filter((a) => a.status !== "rejected").length;
  const overdue = apps.filter(
    (a) => a.nextDate && a.nextDate < today() && a.status !== "rejected" && a.status !== "offer"
  ).length;

  const visa = useMemo(() => {
    if (!settings.visaExpiry) return null;
    const ms = new Date(settings.visaExpiry + "T00:00:00") - new Date(today() + "T00:00:00");
    const days = Math.floor(ms / 86400000);
    return { days, weeks: Math.floor(days / 7), target: addDays(settings.visaExpiry, -56) };
  }, [settings.visaExpiry]);

  const floorNum = salaryNum(settings.floor) || 33400;

  const visible = useMemo(() => {
    const rows = filter === "all" ? apps : apps.filter((a) => a.status === filter);
    const order = (a) => STATUSES.findIndex((s) => s.key === a.status);
    return [...rows].sort((a, b) => {
      // Pinned entries always rise to the top
      if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
      const ov = (x) => (x.nextDate && x.nextDate < today() && x.status !== "rejected" ? 0 : 1);
      if (ov(a) !== ov(b)) return ov(a) - ov(b);
      return order(a) - order(b);
    });
  }, [apps, filter]);

  /* ---------- actions ---------- */
  const save = (app) => {
    let a = { ...app, salary: String(app.salary || "").replace(/[^\d,]/g, "") };
    if (a.status === "applied" && !a.nextDate) a.nextDate = addDays(today(), 7);
    setApps((prev) => prev.some((p) => p.id === a.id) ? prev.map((p) => (p.id === a.id ? a : p)) : [a, ...prev]);
    setEditing(null);
  };
  const remove = (id) => setApps((prev) => prev.filter((p) => p.id !== id));
  const togglePin = (id) => setApps((prev) => prev.map((p) => (p.id === id ? { ...p, pinned: !p.pinned } : p)));
  const setStatus = (id, status) =>
    setApps((prev) => prev.map((p) => {
      if (p.id !== id) return p;
      const next = { ...p, status };
      if (status === "applied" && !next.nextDate) next.nextDate = addDays(today(), 7);
      return next;
    }));
  const clearAll = () => {
    if (window.confirm("Remove all applications? This can't be undone.")) setApps([]);
  };

  const addFromLink = async () => {
    const url = linkInput.trim();
    if (!url) return;
    setLinkNote("");
    // In the desktop app we can fetch the page and pull details out.
    if (window.jobfetch) {
      setLinkBusy(true);
      try {
        const res = await window.jobfetch.fetch(url);
        if (res && res.ok && res.data && (res.data.role || res.data.description)) {
          setEditing({ ...blank(), ...res.data, link: url });
          setLinkInput("");
          setLinkBusy(false);
          return;
        }
        setLinkNote((res && res.error) || "Couldn't read that page — opening a blank entry. Tip: use Paste description.");
      } catch (e) {
        setLinkNote("Couldn't read that page: " + (e.message || "unknown error") + ". Opening a blank entry instead.");
      }
      setLinkBusy(false);
    }
    // Fallback (browser, or fetch failed): open a blank entry with the link.
    setEditing({ ...blank(), link: url });
    setLinkInput("");
  };

  const addFromDescription = (text) => {
    const parsed = parseJobText(text);
    setEditing({ ...blank(), ...parsed });
    setShowPaste(false);
  };

  /* ---------- backup ---------- */
  const exportData = () => {
    const payload = { exportedAt: new Date().toISOString(), settings, apps };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `joblink-backup-${today()}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  };
  const importData = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!Array.isArray(parsed.apps)) throw new Error("bad shape");
        if (!window.confirm(`Import ${parsed.apps.length} applications? This replaces what's currently here.`)) return;
        setApps(parsed.apps);
        if (parsed.settings) setSettings((p) => ({ ...p, ...parsed.settings }));
      } catch (e) {
        window.alert("That file doesn't look like a JobLink backup.");
      }
    };
    reader.readAsText(file);
  };

  /* ---------- render ---------- */
  return (
    <div className="trk">
      <style>{CSS}</style>

      <header className="hd">
        <div>
          <h1>JobLink</h1>
        </div>
        <div className="hdBtns">
          {allInboxes.length > 0 ? (
            <div className="outlookPill">
              <span className="outlookDot" />
              <span className="outlookName">{allInboxes.length} inbox{allInboxes.length > 1 ? "es" : ""}</span>
              <button className="ghost sm" onClick={scanBusy ? undefined : scanAll} disabled={scanBusy}
                      title="Scan all connected inboxes for rejection and interview emails">
                {scanBusy ? "Scanning…" : "Scan emails"}
              </button>
              <button className="ghost sm" onClick={() => setShowEmailConnect(true)} title="Manage inboxes">⋯</button>
            </div>
          ) : (
            (hasMailbox || hasOutlook) && (
              <button className="outlookBtn" onClick={() => setShowEmailConnect(true)}
                      title="Connect an inbox to auto-detect rejections and interview invites">
                ✉ Connect email
              </button>
            )
          )}
          <button className="ghost" onClick={() => setShowSettings(true)}>⚙ Settings</button>
          <button className="primary" onClick={() => setEditing(blank())}>+ Add application</button>
        </div>
      </header>

      {scanErr && <p className="fetchErr" style={{marginBottom:8}}>{scanErr}</p>}

      {emailSuggestions.length > 0 && (
        <EmailPanel
          suggestions={emailSuggestions}
          onApply={applySuggestion}
          onDismiss={dismissSuggestion}
          onDismissAll={() => setEmailSuggestions([])}
        />
      )}

      {visa && (
        <div className={"visa" + (visa.days < 60 ? " hot" : "")}>
          <span className="mono vBig">{visa.weeks} weeks</span>
          <span>left on your visa ({fmtLong(settings.visaExpiry)})</span>
          <span className="vSep">·</span>
          <span>aim for an offer by <strong>{fmtLong(visa.target)}</strong> to allow ~8 weeks for the switch</span>
        </div>
      )}

      <div className="regBar">
        {register ? (
          <span className="regOk">✓ <span className="mono">{register.count.toLocaleString()}</span> Skilled Worker sponsors loaded{register.restored ? " (saved from last time)" : ""} — new entries are checked automatically.</span>
        ) : (
          <span>Sponsor Registrar's CSV</span>
        )}
        <button className="ghost sm" onClick={() => fileRef.current && fileRef.current.click()} disabled={regBusy}>
          {regBusy ? "Reading…" : register ? "Reload CSV" : "Load register CSV"}
        </button>
        <input ref={fileRef} type="file" accept=".csv,text/csv" style={{ display: "none" }}
               onChange={(e) => { loadRegister(e.target.files && e.target.files[0]); e.target.value = ""; }} />
      </div>
      {regErr && <p className="fetchErr">{regErr}</p>}

      <section className="quick">
        <input
          value={linkInput}
          onChange={(e) => setLinkInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && !linkBusy && addFromLink()}
          placeholder={window.jobfetch ? "Paste a job link — I'll read the details from the page" : "Paste a job link to start a new entry"}
          disabled={linkBusy}
        />
        <button className="primary" onClick={addFromLink} disabled={!linkInput.trim() || linkBusy}>
          {linkBusy ? "Reading…" : "Add from link"}
        </button>
        <button className="ghost" onClick={() => setShowPaste(true)} title="Paste a full job description and auto-fill the details">
          📋 Paste description
        </button>
      </section>
      {linkNote && <p className="fetchErr" style={{ margin: "-2px 2px 10px" }}>{linkNote}</p>}

      <section className="boards">
        <button className="boardsHead" onClick={() => setShowBoards(!showBoards)}>
          <span>Search the job boards</span>
          <span className="mono bIcon">{showBoards ? "−" : "+"}</span>
        </button>
        {showBoards && (
          <div className="boardsBody">
            <div className="chips">
              {roleChips.map((r) => (
                <span key={r} className="chipWrap">
                  <button className={"chip" + (searchQ === r ? " on" : "")} onClick={() => setSearchQ(r)}>{r}</button>
                  <button className="chipDel" title={"Remove “" + r + "”"} onClick={() => deleteChip(r)}>✕</button>
                </span>
              ))}
              <button className="chipAdd" onClick={saveCurrentSearch} disabled={!canSaveSearch}
                      title="Save the current search box as a reusable chip">+ Save search</button>
            </div>
            <div className="boardsRow">
              <input value={searchQ} onChange={(e) => setSearchQ(e.target.value)} placeholder="Role keywords" />
              <select value={searchCity} onChange={(e) => setSearchCity(e.target.value)}>
                {["London", "Birmingham", "Manchester", "Remote"].map((c) => <option key={c}>{c}</option>)}
              </select>
              <select value={datePosted} onChange={(e) => setDatePosted(Number(e.target.value))}
                      title="Only show jobs posted within this window (Indeed & LinkedIn)">
                {DATE_POSTED.map((d) => <option key={d.hours} value={d.hours}>{d.label}</option>)}
              </select>
              <label className="chk2">
                <input type="checkbox" checked={addSponsorKw} onChange={(e) => setAddSponsorKw(e.target.checked)} />
                <span>add "visa sponsorship"</span>
              </label>
            </div>
            <div className="boardLinks">
              {boards.map((b) => {
                const q = addSponsorKw && !b.sponsorOnly ? `${searchQ} visa sponsorship` : searchQ;
                const href = buildBoardUrl(b, q, searchCity, datePosted);
                return (
                  <span key={b.id} className="boardWrap">
                    <a className={"boardBtn" + (b.sponsorOnly ? " sp" : "")} href={href} target="_blank" rel="noreferrer">
                      {b.name} ↗
                    </a>
                    {manageBoards && (
                      <button className="boardDel" title={"Remove " + b.name} onClick={() => deleteBoard(b.id)}>✕</button>
                    )}
                  </span>
                );
              })}
              {boards.length === 0 && <span className="boardsEmpty">No boards yet — add one below.</span>}
            </div>

            <div className="boardsManageRow">
              <button className="link" onClick={() => setManageBoards((m) => !m)}>
                {manageBoards ? "Done editing" : "Add / remove boards"}
              </button>
              {manageBoards && <button className="link" onClick={resetBoards}>Reset to defaults</button>}
            </div>
            {manageBoards && (
              <div className="boardAdd">
                <input value={newBoardName} onChange={(e) => setNewBoardName(e.target.value)} placeholder="Board name (e.g. CV-Library)" />
                <input value={newBoardUrl} onChange={(e) => setNewBoardUrl(e.target.value)}
                       placeholder="URL — use {q} for role, {c} for city" spellCheck={false}
                       onKeyDown={(e) => e.key === "Enter" && addBoard()} />
                <button className="ghost sm" onClick={addBoard} disabled={!newBoardName.trim() || !newBoardUrl.trim()}>Add</button>
              </div>
            )}
            {manageBoards && (
              <p className="boardsNote" style={{ marginTop: 6 }}>
                Tip: put <span className="mono">{"{q}"}</span> where the role keywords go, <span className="mono">{"{c}"}</span> for the city, and
                optionally <span className="mono">{"{d}"}</span> for the “date posted” value in days.
                Example: <span className="mono">https://www.cv-library.co.uk/search-jobs?q=&#123;q&#125;&amp;geo=&#123;c&#125;&amp;posted=&#123;d&#125;</span>.
                A URL with no tokens just opens that page. (Indeed and LinkedIn get the date filter automatically.)
              </p>
            )}
            <p className="boardsNote">
              Each link opens with your search ready — the green ones only list licensed sponsors.
              Found a role? Paste its link above to start an entry.
            </p>
          </div>
        )}
      </section>

      <section className="rail">
        <Stat n={apps.length} label="Total" active={filter === "all"} onClick={() => setFilter("all")} />
        {STATUSES.map((s) => (
          <Stat key={s.key} n={counts[s.key]} label={s.label} color={s.color}
                active={filter === s.key} onClick={() => setFilter(filter === s.key ? "all" : s.key)} />
        ))}
        <div className="railEnd">
          <div className="big">{active}</div><div className="lbl">live</div>
          {overdue > 0 && <div className="overduePill">{overdue} need a chase</div>}
        </div>
      </section>

      {visible.length === 0 ? (
        <div className="empty">
          <p className="big2">{apps.length === 0 ? "Nothing logged yet." : "Nothing in this stage."}</p>
          <p>{apps.length === 0
            ? "Paste a job link above, or add one manually. Load the register CSV first and every company gets checked for you."
            : "Pick another stage above, or add a new application."}</p>
          <button className="primary" onClick={() => setEditing(blank())}>+ Add application</button>
        </div>
      ) : (
        <ul className="list">
          {visible.map((a) => {
            const st = statusOf(a.status);
            const sp = SPONSOR[a.sponsor];
            const isOverdue = a.nextDate && a.nextDate < today() && a.status !== "rejected" && a.status !== "offer";
            const sal = salaryNum(a.salary);
            return (
              <li key={a.id} className={"row" + (a.status === "rejected" ? " dim" : "")}>
                <span className="bar" style={{ background: st.color }} />
                <div className="main">
                  <div className="titleline">
                    {a.link
                      ? <a href={a.link} target="_blank" rel="noreferrer" className="co">{a.company || "Untitled"}</a>
                      : <span className="co">{a.company || "Untitled"}</span>}
                    <span className="role">{a.role}</span>
                  </div>
                  <div className="meta">
                    <span>{a.city}</span>
                    {sal !== null && (
                      <span className={"salb mono " + (sal >= floorNum ? "ok" : "low")}>
                        £{a.salary} {sal >= floorNum ? "· clears floor" : "· below £" + floorNum.toLocaleString()}
                      </span>
                    )}
                    <span className="spbadge" style={{ color: sp.color, background: sp.bg }} title={a.sponsorDetail}>{sp.label}</span>
                    {a.clearance && <span className="clrbadge">SC/DV required</span>}
                    {a.dateApplied && <span className="mono soft">applied {fmt(a.dateApplied)}</span>}
                    {(a.posted || a.postedDate) && (
                      <span className="mono soft" title={a.postedDate ? "Posted around " + fmtLong(a.postedDate) : ""}>
                        posted {a.posted || fmt(a.postedDate)}
                      </span>
                    )}
                    {a.closingDate && (
                      <span className={"closeb mono" + (a.closingDate < today() ? " over" : "")}
                            title={"Applications close " + fmtLong(a.closingDate)}>
                        {a.closingDate < today() ? "closed " : "closes "}{fmt(a.closingDate)}
                      </span>
                    )}
                    {a.description && <span className="descTag" title="Job description saved with this entry">📄 description</span>}
                  </div>
                  {(a.nextDate || a.note || a.cvNote) && (
                    <div className={"next" + (isOverdue ? " od" : "")}>
                      {a.nextDate && <span className="mono">{isOverdue ? "⚑ " : "→ "}{fmt(a.nextDate)}</span>}
                      {a.note && <span className="note">{a.note}</span>}
                      {a.cvNote && <span className="cvn">CV: {a.cvNote}</span>}
                    </div>
                  )}
                </div>
                <div className="controls">
                  <select value={a.status} onChange={(e) => setStatus(a.id, e.target.value)}
                          style={{ borderColor: st.color, color: st.color }}>
                    {STATUSES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                  </select>
                  <button className={"icon iconPin" + (a.pinned ? " on" : "")}
                          title={a.pinned ? "Unpin from top" : "Pin to top"}
                          onClick={() => togglePin(a.id)}>{a.pinned ? "📌" : "📍"}</button>
                  <button className="icon" title="Edit" onClick={() => setEditing(a)}>✎</button>
                  <button className="icon" title="Delete" onClick={() => remove(a.id)}>✕</button>
                </div>
              </li>
            );
          })}
        </ul>
      )}

      <footer className="ft">
        <span>
          {apps.length} application{apps.length === 1 ? "" : "s"} tracked · saved in this browser
        </span>
        <span className="ftBtns">
          <button className="link" onClick={exportData}>Export backup</button>
          <button className="link" onClick={() => importRef.current && importRef.current.click()}>Import backup</button>
          {apps.length > 0 && <button className="link danger" onClick={clearAll}>Clear all</button>}
          <input ref={importRef} type="file" accept=".json,application/json" style={{ display: "none" }}
                 onChange={(e) => { importData(e.target.files && e.target.files[0]); e.target.value = ""; }} />
        </span>
      </footer>

      {editing && (
        <Editor app={editing} floorNum={floorNum} registerLoaded={!!register}
                lookup={lookupSponsor} onSave={save} onCancel={() => setEditing(null)} />
      )}
      {showSettings && (
        <Settings settings={settings}
                  onSave={(s) => { setSettings(s); setShowSettings(false); }}
                  onCancel={() => setShowSettings(false)} />
      )}
      {showPaste && (
        <PasteModal onCreate={addFromDescription} onCancel={() => setShowPaste(false)} />
      )}
      {showEmailConnect && (
        <EmailConnectModal
          hasMailbox={hasMailbox}
          hasOutlookClientId={hasOutlook}
          inboxes={allInboxes}
          onConnectImap={connectMailbox}
          onConnectOutlook={() => { setShowEmailConnect(false); signInOutlook(); }}
          onDisconnect={disconnectInbox}
          onCancel={() => setShowEmailConnect(false)}
        />
      )}
    </div>
  );
}

/* ── Paste job description modal ──────────────────────────────────── */
function PasteModal({ onCreate, onCancel }) {
  const [text, setText] = useState("");
  const preview = useMemo(() => (text.trim().length > 15 ? parseJobText(text) : null), [text]);
  const found = preview
    ? [
        ["Role", preview.role],
        ["Company", preview.company],
        ["City", preview.city],
        ["Salary", preview.salary ? "£" + Number(preview.salary).toLocaleString() : ""],
        ["Posted", preview.posted || ""],
        ["Closes", preview.closingDate || ""],
        ["Clearance", preview.clearance ? "SC/DV mentioned" : ""],
      ].filter(([, v]) => v)
    : [];

  return (
    <div className="scrim" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 600 }}>
        <h2>Paste a job description</h2>
        <p className="icloudIntro">
          Copy the whole posting from LinkedIn, Indeed, Reed — anywhere — and paste it below.
          JobLink reads it on your device, fills in what it can, and keeps the full text with the entry.
          Nothing is uploaded.
        </p>
        <textarea
          className="pasteArea"
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="Paste the job title, company, location, salary and description here…"
          autoFocus
        />
        {preview && (
          <div className="pastePreview">
            <span className="pastePreviewLabel">Detected:</span>
            {found.length ? (
              found.map(([k, v]) => (
                <span key={k} className="pasteChip"><strong>{k}:</strong> {v}</span>
              ))
            ) : (
              <span className="pasteChip dim">Couldn’t read structured details — you can still fill them in on the next screen.</span>
            )}
          </div>
        )}
        <p className="boardsNote" style={{ marginTop: 8 }}>
          These are best-guesses from the text — you’ll get the full editor next to fix anything,
          and your sponsor register check runs automatically on the detected company.
        </p>
        <div className="actions">
          <button className="ghost" onClick={onCancel}>Cancel</button>
          <button className="primary" onClick={() => onCreate(text)} disabled={text.trim().length < 15}>
            Review & create entry
          </button>
        </div>
      </div>
    </div>
  );
}

/* ── Connect email modal (provider picker) ───────────────────────────── */
const EMAIL_PROVIDERS = [
  { id: "gmail",  name: "Gmail",            help: "myaccount.google.com → Security → 2-Step Verification → App passwords" },
  { id: "yahoo",  name: "Yahoo",            help: "Yahoo Account → Account Security → Generate app password" },
  { id: "icloud", name: "iCloud",           help: "appleid.apple.com → Sign-In and Security → App-Specific Passwords" },
  { id: "aol",    name: "AOL",              help: "AOL Account → Account Security → Generate app password" },
  { id: "other",  name: "Other (IMAP)",     help: "Find your provider's IMAP server and create an app password in its security settings." },
  { id: "outlook", name: "Outlook / Microsoft", help: "Uses Microsoft sign-in (set your client ID in Settings first)." },
];

function EmailConnectModal({ hasMailbox, hasOutlookClientId, inboxes = [], onConnectImap, onConnectOutlook, onDisconnect, onCancel }) {
  const [provider, setProvider] = useState("gmail");
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [host, setHost] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [okMsg, setOkMsg] = useState("");

  const meta = EMAIL_PROVIDERS.find((p) => p.id === provider) || EMAIL_PROVIDERS[0];
  const isOutlook = provider === "outlook";
  const isOther = provider === "other";

  const submit = async () => {
    setErr(""); setOkMsg("");
    if (isOutlook) {
      if (!hasOutlookClientId) { setErr("Add your Microsoft client ID in Settings first, then choose Outlook here."); return; }
      onConnectOutlook();
      return;
    }
    if (!email.trim() || !pw.trim() || (isOther && !host.trim())) { setErr("Fill in all the fields."); return; }
    setBusy(true);
    const res = await onConnectImap({ provider, host: host.trim(), email: email.trim(), appPassword: pw });
    setBusy(false);
    if (!res.ok) { setErr(res.error); return; }
    setOkMsg("Connected " + email.trim() + ". Add another, or click Done.");
    setEmail(""); setPw(""); setHost("");
  };

  const provLabel = (id) => (EMAIL_PROVIDERS.find((p) => p.id === id) || {}).name || id;

  return (
    <div className="scrim" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 500 }}>
        <h2>Email inboxes</h2>
        <p className="icloudIntro">
          JobLink scans every connected inbox for rejection and interview emails and matches them to your applications.
          You can connect as many as you like.
        </p>

        {inboxes.length > 0 && (
          <div className="inboxList">
            {inboxes.map((ib) => (
              <div key={ib.id} className="inboxRow">
                <span className="inboxDot" />
                <span className="inboxLabel">{ib.label}</span>
                <span className="inboxProv">{ib.kind === "outlook" ? "Outlook" : provLabel(ib.provider)}</span>
                <button className="ghost sm" onClick={() => onDisconnect(ib)} title="Disconnect this inbox">Remove</button>
              </div>
            ))}
          </div>
        )}

        <div className="addInboxHd">Add an inbox</div>
        <div className="grid">
          <Field label="Provider" wide>
            <select value={provider} onChange={(e) => setProvider(e.target.value)}>
              {EMAIL_PROVIDERS.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
            <span className="regHint">{meta.help}</span>
          </Field>

          {isOutlook ? (
            <p className="boardsNote" style={{ gridColumn: "1 / -1", margin: "2px 0 0" }}>
              Microsoft disabled email passwords, so Outlook uses a secure sign-in window. Click Connect to open it.
            </p>
          ) : (
            <>
              {isOther && (
                <Field label="IMAP server" wide>
                  <input value={host} onChange={(e) => setHost(e.target.value)} placeholder="e.g. imap.fastmail.com" spellCheck={false} />
                </Field>
              )}
              <Field label="Email address" wide>
                <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@example.com" autoFocus inputMode="email" />
              </Field>
              <Field label="App password" wide>
                <input value={pw} onChange={(e) => setPw(e.target.value)} placeholder="xxxx xxxx xxxx xxxx" type="password"
                       onKeyDown={(e) => e.key === "Enter" && submit()} />
                <span className="regHint">Stored encrypted in your Mac's Keychain — never in the browser, never in plain text.</span>
              </Field>
            </>
          )}
        </div>
        {err && <p className="fetchErr" style={{ margin: "4px 2px 0" }}>{err}</p>}
        {okMsg && <p className="regHint hit" style={{ margin: "4px 2px 0" }}>{okMsg}</p>}
        <div className="actions">
          <button className="ghost" onClick={onCancel}>Done</button>
          <button className="primary" onClick={submit} disabled={busy}>
            {busy ? "Connecting…" : isOutlook ? "Open Microsoft sign-in" : "Connect"}
          </button>
        </div>
      </div>
    </div>
  );
}


/* ── Email suggestions panel ──────────────────────────────────────── */
function EmailPanel({ suggestions, onApply, onDismiss, onDismissAll }) {
  const LABELS = { rejected: "Rejection", interview: "Interview invite" };
  const COLORS = {
    rejected:  { color: "#6B7280", bg: "#F3F4F6", border: "#D1D5DB" },
    interview: { color: "#0F766E", bg: "#F0FDFA", border: "#99F6E4" },
  };

  return (
    <div className="emailPanel">
      <div className="emailPanelHead">
        <span className="emailPanelTitle">📬 Emails detected</span>
        <span className="emailPanelSub">{suggestions.length} update{suggestions.length !== 1 ? "s" : ""} found in your inbox</span>
        <button className="link" style={{marginLeft:"auto",fontSize:12}} onClick={onDismissAll}>Dismiss all</button>
      </div>
      {suggestions.map((s) => {
        const c = COLORS[s.suggestedStatus] || COLORS.rejected;
        return (
          <div key={s.id} className="emailItem">
            <div className="emailItemLeft">
              <span className="emailTypeBadge" style={{ color: c.color, background: c.bg, border: `1px solid ${c.border}` }}>
                {LABELS[s.suggestedStatus] || s.suggestedStatus}
              </span>
              <div className="emailItemMeta">
                <span className="emailCo">{s.company}</span>
                {s.role && <span className="emailRole"> · {s.role}</span>}
              </div>
              <div className="emailSubject" title={s.preview}>{s.subject}</div>
              {s.date && <div className="emailDate">{fmtEmail(s.date)}</div>}
            </div>
            <div className="emailItemActions">
              <button
                className="emailApplyBtn"
                style={{ color: c.color, borderColor: c.border }}
                onClick={() => onApply(s)}
              >
                Mark as {s.suggestedStatus === "interview" ? "Interview" : "Rejected"}
              </button>
              <button className="emailDismissBtn" onClick={() => onDismiss(s.id)}>Ignore</button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ── Stat pill ───────────────────────────────────────────────────── */
function Stat({ n, label, color = "#16202E", active, onClick }) {
  return (
    <button className={"stat" + (active ? " on" : "")} onClick={onClick}>
      <span className="statN mono" style={{ color }}>{n}</span>
      <span className="statL">{label}</span>
      <span className="statU" style={{ background: color, opacity: active ? 1 : 0.25 }} />
    </button>
  );
}

/* ── Editor modal ────────────────────────────────────────────────── */
function Editor({ app, floorNum, registerLoaded, lookup, onSave, onCancel }) {
  const [f, setF] = useState(app);
  const [reg, setReg] = useState(null);
  const [autoApplied, setAutoApplied] = useState(false);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const valid = f.company.trim() && f.role.trim();
  const sal = salaryNum(f.salary);

  useEffect(() => {
    if (!registerLoaded || !f.company || f.company.trim().length < 3) {
      setReg(null);
      setAutoApplied(false);
      return;
    }
    const t = setTimeout(() => {
      const result = lookup(f.company);
      setReg(result);
      // Auto-apply exact matches
      if (result && result.match === "exact" && f.sponsor !== "yes") {
        setF((prev) => ({
          ...prev,
          sponsor: "yes",
          sponsorDetail: `${result.name}${result.town ? ` (${result.town})` : ""} — ${result.rating || "rated"}`,
        }));
        setAutoApplied(true);
      }
    }, 350);
    return () => clearTimeout(t);
  }, [f.company, registerLoaded]);

  const applyReg = () => {
    if (!reg || reg.match === "none") return;
    setF({ ...f, sponsor: "yes", sponsorDetail: `${reg.name}${reg.town ? ` (${reg.town})` : ""} — ${reg.rating || "rated"}` });
    setAutoApplied(true);
  };

  return (
    <div className="scrim" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{app.company ? "Edit application" : "New application"}</h2>
        <div className="grid">
          <Field label="Company" wide>
            <input value={f.company} onChange={set("company")} placeholder="e.g. Amber Labs" autoFocus />
            {registerLoaded && reg && (
              reg.match === "none" ? (
                <span className="regHint miss">Not found in the register — try the legal name (e.g. "… Services Limited"), or mark as no licence.</span>
              ) : (
                <span className="regHint hit">
                  {autoApplied
                    ? "✓ Verified on the register and marked automatically: "
                    : reg.match === "exact" ? "✓ On the register: " : "≈ Possible match: "}
                  <strong>{reg.name}</strong>{reg.town ? ` (${reg.town})` : ""} — {reg.rating || "rating n/a"}
                  {f.sponsor !== "yes" && !autoApplied && (
                    <button className="tiny" onClick={applyReg}>Mark as sponsor ✓</button>
                  )}
                </span>
              )
            )}
          </Field>
          <Field label="Role" wide><input value={f.role} onChange={set("role")} placeholder="e.g. Junior Web Developer" /></Field>
          <Field label="City"><input value={f.city} onChange={set("city")} /></Field>
          <Field label="Salary (£/yr)">
            <input value={f.salary} onChange={set("salary")} placeholder="34,000" inputMode="numeric" />
            {sal !== null && (
              <span className={"regHint " + (sal >= floorNum ? "hit" : "miss")}>
                {sal >= floorNum ? `✓ Clears your £${floorNum.toLocaleString()} floor` : `Below your £${floorNum.toLocaleString()} floor — likely can't be sponsored`}
              </span>
            )}
          </Field>
          <Field label="Job link" wide><input value={f.link} onChange={set("link")} placeholder="https://…" /></Field>
          <Field label="Sponsor licence">
            <select value={f.sponsor} onChange={set("sponsor")}>
              <option value="unknown">Not checked yet</option>
              <option value="yes">On the register ✓</option>
              <option value="no">Not a sponsor</option>
            </select>
          </Field>
          <Field label="Stage">
            <select value={f.status} onChange={set("status")}>
              {STATUSES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
            </select>
          </Field>
          <Field label="Date applied"><input type="date" value={f.dateApplied} onChange={set("dateApplied")} /></Field>
          <Field label="Next action"><input type="date" value={f.nextDate} onChange={set("nextDate")} /></Field>
          <Field label="Job posted">
            <input type="date" value={f.postedDate || ""} onChange={set("postedDate")} />
            {f.posted && <span className="regHint">From the listing: “{f.posted}”.</span>}
          </Field>
          <Field label="Closing date">
            <input type="date" value={f.closingDate || ""} onChange={set("closingDate")} />
          </Field>
          <Field label="CV sent / keywords mirrored" wide>
            <input value={f.cvNote} onChange={set("cvNote")} placeholder="e.g. cyber CV v3 — SIEM, vulnerability management" />
          </Field>
          <Field label="Note" wide><input value={f.note} onChange={set("note")} placeholder="e.g. follow up with recruiter" /></Field>
          <Field label="Job description (kept with this entry)" wide>
            <textarea className="descArea" value={f.description || ""} onChange={set("description")}
                      placeholder="Paste or edit the full job description here — stored on your device for later reference." />
          </Field>
          <label className="chk wide">
            <input type="checkbox" checked={!!f.clearance} onChange={(e) => setF({ ...f, clearance: e.target.checked })} />
            <span>Requires SC/DV security clearance <em>(usually a blocker on a new visa — deprioritise)</em></span>
          </label>
        </div>
        <div className="actions">
          <button className="ghost" onClick={onCancel}>Cancel</button>
          <button
            className={autoApplied ? "primarySponsor" : "primary"}
            disabled={!valid}
            onClick={() => onSave(f)}
          >
            {autoApplied
              ? "Save — sponsor verified ✓"
              : app.company ? "Save changes" : "Add application"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ── Settings modal ──────────────────────────────────────────────── */
function Settings({ settings, onSave, onCancel }) {
  const [s, setS] = useState(settings);
  return (
    <div className="scrim" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>Settings</h2>
        <div className="grid">
          <Field label="Visa expiry date" wide>
            <input type="date" value={s.visaExpiry} onChange={(e) => setS({ ...s, visaExpiry: e.target.value })} />
            <span className="regHint">Shows a countdown and your target offer date (expiry minus ~8 weeks for processing).</span>
          </Field>
          <Field label="Your salary floor (£/yr)" wide>
            <input value={s.floor} onChange={(e) => setS({ ...s, floor: e.target.value.replace(/[^\d]/g, "") })} inputMode="numeric" />
            <span className="regHint">New-entrant minimum is £33,400; some occupations need 70% of their going rate (e.g. ~£34,580 for software roles). Set the figure for your main target.</span>
          </Field>
          <label className="chk wide">
            <input type="checkbox" checked={s.autoScan !== false}
                   onChange={(e) => setS({ ...s, autoScan: e.target.checked })} />
            <span>Scan inboxes automatically <em>(when JobLink opens, then on a timer)</em></span>
          </label>
          <Field label="Auto-scan every" wide>
            <select value={String(s.autoScanMins || 30)} onChange={(e) => setS({ ...s, autoScanMins: Number(e.target.value) })}
                    disabled={s.autoScan === false}>
              {[15, 30, 60].map((m) => (
                <option key={m} value={m}>{m < 60 ? m + " minutes" : "1 hour"}</option>
              ))}
            </select>
            <span className="regHint">Only runs while JobLink is open. You can still scan manually any time.</span>
          </Field>
          <Field label="Outlook / Microsoft client ID" wide>
            <input value={s.outlookClientId || ""} onChange={(e) => setS({ ...s, outlookClientId: e.target.value.trim() })}
                   placeholder="e.g. 1a2b3c4d-5e6f-7a8b-9c0d-1e2f3a4b5c6d" spellCheck={false} />
            <span className="regHint">
              Connects your Outlook inbox to auto-detect rejections and interview invites. One-time setup at
              <span className="mono"> portal.azure.com</span> → App registrations. Choose “Accounts in any org + personal”.
              In <strong>Authentication</strong>, add a platform → <strong>Mobile and desktop applications</strong> and tick
              <span className="mono"> http://localhost</span>, then set <strong>Allow public client flows → Yes</strong>.
              Under <strong>API permissions</strong> add <span className="mono">Microsoft Graph → Delegated → Mail.Read</span>.
              Paste the “Application (client) ID” here. Leave blank to hide Outlook.
            </span>
          </Field>
        </div>
        <div className="actions">
          <button className="ghost" onClick={onCancel}>Cancel</button>
          <button className="primary" onClick={() => onSave(s)}>Save settings</button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, wide, children }) {
  return (
    <label className={"field" + (wide ? " wide" : "")}>
      <span>{label}</span>{children}
    </label>
  );
}

/* ── CSS ─────────────────────────────────────────────────────────── */
const CSS = `
/* Use San Francisco (SF Pro) on Mac natively — no web font download needed */
.trk{--ink:#16202E;--sub:#5A6675;--line:#E2E7EC;--surf:#fff;--teal:#0F766E;--teal-d:#0B5A54;
 font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Helvetica Neue',system-ui,sans-serif;
 color:var(--ink);max-width:960px;margin:0 auto;
 padding:36px 24px 60px;background:#EEF1F4;min-height:100%;box-sizing:border-box;}
.trk *{box-sizing:border-box;}
.mono{font-family:ui-monospace,'SF Mono','Menlo',monospace;}

/* Dark mode lives at the END of this stylesheet (see bottom) so its
   explicit-colour overrides win over the base light rules. */
/* Draggable title bar region (lets you drag the window from the header) */
.hd{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;
 margin-bottom:18px;padding-top:4px;-webkit-app-region:drag;}
/* Buttons inside must be non-draggable so they still click */
.hd *{-webkit-app-region:no-drag;}
.hd h1{font-family:-apple-system,BlinkMacSystemFont,'SF Pro Display','Helvetica Neue',system-ui,sans-serif;
 font-weight:700;font-size:26px;letter-spacing:-.02em;margin:0;}
.hdBtns{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.primary{background:var(--teal);color:#fff;border:0;border-radius:9px;padding:11px 16px;font-weight:600;
 font-size:14px;cursor:pointer;font-family:inherit;transition:background .15s;white-space:nowrap;}
.primary:hover{background:var(--teal-d);}
.primary:disabled{background:#A9B4BF;cursor:not-allowed;}
.primarySponsor{background:#15803D;color:#fff;border:0;border-radius:9px;padding:11px 16px;font-weight:600;
 font-size:14px;cursor:pointer;font-family:inherit;transition:background .15s;white-space:nowrap;}
.primarySponsor:hover{background:#116530;}
.primarySponsor:disabled{background:#A9B4BF;cursor:not-allowed;}
.ghost{background:transparent;border:1px solid var(--line);border-radius:9px;padding:11px 16px;font-weight:600;
 font-size:14px;cursor:pointer;font-family:inherit;color:var(--sub);white-space:nowrap;}
.ghost:hover{border-color:var(--sub);}
.ghost.sm{padding:7px 12px;font-size:12.5px;}
.ghost:disabled{opacity:.55;cursor:wait;}

/* Outlook connect button */
.outlookBtn{background:#0078D4;color:#fff;border:0;border-radius:9px;padding:9px 14px;font-weight:600;
 font-size:13px;cursor:pointer;font-family:inherit;transition:background .15s;white-space:nowrap;}
.outlookBtn:hover{background:#006CBF;}
.outlookPill{display:flex;align-items:center;gap:6px;background:#EFF6FF;border:1px solid #BFDBFE;
 border-radius:10px;padding:5px 5px 5px 10px;font-size:12.5px;}
.outlookDot{width:7px;height:7px;border-radius:50%;background:#0078D4;flex-shrink:0;}
.outlookName{color:#1E40AF;font-weight:500;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}

/* Intro paragraph (used in the Paste-description / email modals) */
.icloudIntro{font-size:13px;color:var(--sub);line-height:1.5;margin:0 0 10px;}

/* Connected inbox list (email manager) */
.inboxList{display:flex;flex-direction:column;gap:6px;margin:0 0 14px;}
.inboxRow{display:flex;align-items:center;gap:8px;padding:8px 10px;border:1px solid var(--line);border-radius:9px;background:var(--surf);}
.inboxDot{width:7px;height:7px;border-radius:50%;background:#15803D;flex-shrink:0;}
.inboxLabel{font-size:13px;font-weight:600;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.inboxProv{font-size:11px;color:var(--sub);background:#EEF1F4;border-radius:20px;padding:2px 8px;margin-left:auto;}
.addInboxHd{font-size:12px;font-weight:700;color:var(--sub);text-transform:uppercase;letter-spacing:.04em;margin:2px 0 8px;}

/* Email panel */
.emailPanel{background:#fff;border:1px solid #E2E7EC;border-radius:14px;margin-bottom:12px;overflow:hidden;}
.emailPanelHead{display:flex;align-items:center;gap:8px;padding:11px 14px;border-bottom:1px solid #E2E7EC;flex-wrap:wrap;}
.emailPanelTitle{font-weight:700;font-size:13.5px;color:var(--ink);}
.emailPanelSub{font-size:12px;color:var(--sub);}
.emailItem{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:11px 14px;
 border-top:1px solid #F1F5F9;flex-wrap:wrap;}
.emailItem:first-of-type{border-top:0;}
.emailItemLeft{flex:1;min-width:0;display:flex;flex-direction:column;gap:3px;}
.emailTypeBadge{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;
 letter-spacing:.02em;margin-bottom:2px;align-self:flex-start;}
.emailItemMeta{font-size:13px;}
.emailCo{font-weight:600;color:var(--ink);}
.emailRole{color:var(--sub);}
.emailSubject{font-size:12px;color:var(--sub);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:420px;}
.emailDate{font-size:11px;color:#94A3B8;margin-top:1px;}
.emailItemActions{display:flex;gap:6px;align-items:center;flex-shrink:0;}
.emailApplyBtn{font-family:inherit;font-size:12px;font-weight:600;border-radius:8px;padding:6px 11px;
 cursor:pointer;background:#fff;transition:.12s;}
.emailApplyBtn:hover{opacity:.8;}
.emailDismissBtn{font-family:inherit;font-size:12px;color:#94A3B8;border:1px solid #E2E7EC;background:#fff;
 border-radius:8px;padding:6px 11px;cursor:pointer;}
.emailDismissBtn:hover{border-color:#94A3B8;color:var(--sub);}

.visa{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:#16202E;color:#E8EDF2;border-radius:12px;
 padding:11px 16px;font-size:13px;margin-bottom:10px;}
.visa.hot{background:#7C2D12;}
.vBig{font-size:17px;font-weight:600;color:#7FD8CE;}
.visa.hot .vBig{color:#FDBA74;}
.vSep{opacity:.4;}
.visa strong{font-weight:600;}

.regBar{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;background:var(--surf);
 border:1px solid var(--line);border-radius:12px;padding:10px 14px;margin-bottom:10px;font-size:13px;color:var(--sub);}
.regOk{color:#15803D;font-weight:500;}

.quick{display:flex;gap:8px;margin-bottom:10px;}
.quick input{flex:1;font-family:inherit;font-size:14px;padding:11px 13px;border:1px solid var(--line);
 border-radius:9px;color:var(--ink);background:var(--surf);}
.quick input:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
.fetchErr{font-size:12.5px;color:#B45309;margin:-2px 2px 10px;}

.boards{background:var(--surf);border:1px solid var(--line);border-radius:12px;margin-bottom:10px;overflow:hidden;}
.boardsHead{width:100%;display:flex;justify-content:space-between;align-items:center;background:transparent;border:0;
 padding:11px 14px;font-family:inherit;font-size:13.5px;font-weight:600;color:var(--ink);cursor:pointer;}
.boardsHead:hover{background:#F2F5F8;}
.bIcon{font-size:16px;color:var(--sub);}
.boardsBody{padding:4px 14px 14px;border-top:1px solid var(--line);}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0;align-items:center;}
.chipWrap{position:relative;display:inline-flex;}
.chip{font-family:inherit;font-size:12px;font-weight:500;border:1px solid var(--line);background:#F4F6F8;color:var(--sub);
 border-radius:20px;padding:5px 11px;cursor:pointer;transition:.12s;}
.chip:hover{border-color:var(--teal);color:var(--teal);}
.chip.on{background:var(--teal);border-color:var(--teal);color:#fff;}
.chipDel{position:absolute;top:-6px;right:-6px;width:16px;height:16px;border:0;border-radius:50%;background:#B91C1C;
 color:#fff;font-size:9px;line-height:1;cursor:pointer;display:none;align-items:center;justify-content:center;padding:0;
 box-shadow:0 1px 3px rgba(0,0,0,.25);}
.chipWrap:hover .chipDel{display:flex;}
.chipDel:hover{background:#991B1B;}
.chipAdd{font-family:inherit;font-size:12px;font-weight:600;border:1px dashed var(--line);background:transparent;
 color:var(--sub);border-radius:20px;padding:5px 11px;cursor:pointer;transition:.12s;}
.chipAdd:hover{border-color:var(--teal);color:var(--teal);}
.chipAdd:disabled{opacity:.45;cursor:not-allowed;}
.boardsRow{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px;}
.boardsRow input{flex:1;min-width:180px;font-family:inherit;font-size:13.5px;padding:9px 11px;border:1px solid var(--line);border-radius:9px;}
.boardsRow input:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
.boardsRow select{font-family:inherit;font-size:13.5px;padding:9px 11px;border:1px solid var(--line);border-radius:9px;background:#fff;cursor:pointer;}
.chk2{display:flex;align-items:center;gap:6px;font-size:12.5px;color:var(--sub);cursor:pointer;white-space:nowrap;}
.boardLinks{display:flex;gap:7px;flex-wrap:wrap;align-items:center;}
.boardWrap{display:inline-flex;align-items:center;position:relative;}
.boardBtn{font-size:12.5px;font-weight:600;text-decoration:none;color:var(--ink);border:1px solid var(--line);
 background:#fff;border-radius:9px;padding:8px 12px;transition:.12s;}
/* softer hover — light background, darker text/border (was a bright teal) */
.boardBtn:hover{background:#EDF5F3;border-color:var(--teal-d);color:var(--teal-d);}
.boardBtn.sp{color:#15803D;border-color:#BBDCC6;background:#F4FAF6;}
.boardBtn.sp:hover{background:#E7F2EB;border-color:#116530;color:#116530;}
.boardDel{margin-left:3px;width:20px;height:20px;border:0;border-radius:50%;background:#FBEAEA;color:#B91C1C;
 font-size:11px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.boardDel:hover{background:#F5D0D0;}
.boardsEmpty{font-size:12px;color:var(--sub);}
.boardsManageRow{display:flex;gap:14px;margin-top:10px;}
.boardAdd{display:flex;gap:7px;flex-wrap:wrap;margin-top:8px;}
.boardAdd input{flex:1;min-width:160px;font-family:inherit;font-size:13px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:#fff;color:var(--ink);}
.boardAdd input:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
.boardsNote{font-size:11.5px;color:var(--sub);margin:10px 0 0;line-height:1.5;}

.rail{display:flex;align-items:stretch;gap:6px;background:var(--surf);border:1px solid var(--line);
 border-radius:14px;padding:10px;margin-bottom:18px;overflow-x:auto;}
.stat{flex:1;min-width:78px;background:transparent;border:0;cursor:pointer;padding:10px 8px 12px;
 border-radius:9px;display:flex;flex-direction:column;align-items:center;gap:3px;position:relative;font-family:inherit;transition:background .12s;}
.stat:hover{background:#F2F5F8;}
.stat.on{background:#F2F5F8;}
.statN{font-size:21px;font-weight:600;line-height:1;}
.statL{font-size:11px;color:var(--sub);font-weight:500;letter-spacing:.01em;}
.statU{height:3px;width:20px;border-radius:2px;margin-top:2px;}
.railEnd{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:0 14px;
 border-left:1px solid var(--line);min-width:92px;gap:2px;}
.railEnd .big{font-family:'Space Grotesk',sans-serif;font-size:26px;font-weight:700;color:var(--teal);line-height:1;}
.railEnd .lbl{font-size:11px;color:var(--sub);text-transform:uppercase;letter-spacing:.08em;}
.overduePill{margin-top:6px;font-size:10.5px;font-weight:600;color:#C2410C;background:#FBEDE3;
 padding:3px 7px;border-radius:20px;white-space:nowrap;}

.list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:8px;}
.row{display:flex;align-items:stretch;background:var(--surf);border:1px solid var(--line);border-radius:12px;
 overflow:hidden;transition:box-shadow .12s;}
.row:hover{box-shadow:0 4px 16px rgba(20,32,46,.07);}
.row.dim{opacity:.58;}
.bar{width:4px;flex-shrink:0;}
.main{flex:1;padding:13px 14px;min-width:0;}
.titleline{display:flex;align-items:baseline;gap:9px;flex-wrap:wrap;}
.co{font-weight:600;font-size:15px;color:var(--ink);text-decoration:none;}
a.co:hover{color:var(--teal);text-decoration:underline;}
.role{color:var(--sub);font-size:13.5px;}
.meta{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px;font-size:12.5px;color:var(--sub);}
.meta .soft{opacity:.8;font-size:11.5px;}
.salb{font-size:11.5px;font-weight:500;padding:2px 8px;border-radius:20px;}
.salb.ok{color:#15803D;background:#E7F4EC;}
.salb.low{color:#B91C1C;background:#FBEAEA;}
.spbadge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;}
.clrbadge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;color:#9A3412;background:#FBEDE3;}
.next{margin-top:8px;display:flex;gap:9px;align-items:center;flex-wrap:wrap;font-size:12.5px;color:var(--sub);}
.next.od{color:#C2410C;font-weight:500;}
.next .note{color:var(--ink);opacity:.78;}
.next .cvn{color:var(--sub);font-size:11.5px;border-left:2px solid var(--line);padding-left:8px;}
.controls{display:flex;align-items:center;gap:6px;padding:10px 12px;flex-shrink:0;}
.controls select{font-family:inherit;font-size:12.5px;font-weight:600;border:1.5px solid;border-radius:8px;
 padding:6px 8px;background:#fff;cursor:pointer;}
.icon{width:30px;height:30px;border:1px solid var(--line);background:#fff;border-radius:8px;cursor:pointer;
 color:var(--sub);font-size:13px;display:flex;align-items:center;justify-content:center;transition:.12s;}
.icon:hover{border-color:var(--sub);color:var(--ink);}
.iconPin.on{background:#E7F4F1;border-color:var(--teal);}

.empty{background:var(--surf);border:1px dashed var(--line);border-radius:14px;padding:48px 28px;text-align:center;color:var(--sub);}
.empty .big2{font-family:'Space Grotesk',sans-serif;font-size:18px;font-weight:600;color:var(--ink);margin:0 0 8px;}
.empty p{margin:0 0 8px;max-width:460px;margin-left:auto;margin-right:auto;font-size:14px;line-height:1.5;}
.empty .primary{margin-top:14px;}

.ft{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-top:18px;font-size:12px;color:var(--sub);}
.ftBtns{display:flex;gap:14px;}
.link{background:0;border:0;color:var(--sub);text-decoration:underline;cursor:pointer;font-family:inherit;font-size:12px;}
.link:hover{color:var(--teal);}
.link.danger:hover{color:#B91C1C;}

.scrim{position:fixed;inset:0;background:rgba(20,32,46,.42);display:flex;align-items:center;justify-content:center;
 padding:18px;z-index:50;backdrop-filter:blur(2px);}
.modal{background:#fff;border-radius:16px;padding:24px;width:100%;max-width:560px;max-height:90vh;overflow:auto;
 box-shadow:0 20px 60px rgba(20,32,46,.28);animation:pop .16s ease-out;}
@keyframes pop{from{opacity:0;transform:translateY(8px) scale(.99);}to{opacity:1;transform:none;}}
.modal h2{font-family:'Space Grotesk',sans-serif;font-size:19px;font-weight:600;margin:0 0 16px;}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.field{display:flex;flex-direction:column;gap:5px;}
.field.wide{grid-column:1 / -1;}
.field span{font-size:12px;font-weight:600;color:var(--sub);}
.field input,.field select{font-family:inherit;font-size:14px;padding:9px 11px;border:1px solid var(--line);
 border-radius:9px;color:var(--ink);background:#fff;width:100%;}
.field input:focus,.field select:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
/* Paste-description modal */
.pasteArea{width:100%;min-height:200px;resize:vertical;font-family:inherit;font-size:13px;line-height:1.5;
 padding:11px 13px;border:1px solid var(--line);border-radius:10px;color:var(--ink);background:#fff;}
.pasteArea:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
.pastePreview{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-top:10px;}
.pastePreviewLabel{font-size:12px;font-weight:600;color:var(--sub);}
.pasteChip{font-size:12px;background:#EEF6F4;color:#0B5A54;border:1px solid #CDE7E1;border-radius:20px;padding:3px 10px;}
.pasteChip strong{font-weight:600;}
.pasteChip.dim{background:#FBF3E8;color:#9A6B1E;border-color:#F0DEC0;}
/* Editor description textarea */
.descArea{width:100%;min-height:120px;resize:vertical;font-family:inherit;font-size:13px;line-height:1.5;
 padding:9px 11px;border:1px solid var(--line);border-radius:9px;color:var(--ink);background:#fff;}
.descArea:focus{outline:none;border-color:var(--teal);box-shadow:0 0 0 3px rgba(15,118,110,.12);}
.descTag{font-size:11px;color:#0B5A54;background:#EEF6F4;border-radius:20px;padding:2px 8px;}
.closeb{font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;color:#7A5C00;background:#FBF3E0;}
.closeb.over{color:#B91C1C;background:#FBEAEA;}
.regHint{font-size:11.5px!important;font-weight:500!important;line-height:1.45;display:block;}
.regHint.hit{color:#15803D!important;}
.regHint.miss{color:#B45309!important;}
.tiny{margin-left:8px;font-size:11px;font-weight:600;border:1px solid #15803D;color:#15803D;background:#fff;
 border-radius:6px;padding:2px 8px;cursor:pointer;font-family:inherit;}
.tiny:hover{background:#E7F4EC;}
.chk{display:flex;align-items:flex-start;gap:9px;font-size:13px;color:var(--ink);grid-column:1 / -1;cursor:pointer;}
.chk input{margin-top:2px;}
.chk em{color:var(--sub);font-style:normal;font-size:12px;}
.actions{display:flex;justify-content:flex-end;gap:10px;margin-top:20px;}

@media (max-width:560px){
 .row{flex-wrap:wrap;}
 .controls{width:100%;border-top:1px solid var(--line);justify-content:space-between;}
 .controls select{flex:1;}
 .grid{grid-template-columns:1fr;}
 .railEnd{min-width:74px;padding:0 8px;}
 .emailItem{flex-direction:column;}
 .emailItemActions{width:100%;}
 .outlookPill{width:100%;}
}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;}}

/* ── Dark mode (placed LAST so explicit-colour overrides win) ──────── */
@media (prefers-color-scheme: dark) {
 .trk{--ink:#E8EDF2;--sub:#7A8899;--line:#2A3547;--surf:#1C2840;--teal:#2DD4BF;--teal-d:#14B8A6;
  background:#0F1724;}
 .primary{background:#0D9488;}
 .primary:hover{background:#0F766E;}
 .primarySponsor{background:#16A34A;}
 .primarySponsor:hover{background:#15803D;}
 .ghost{color:#8B95A3;border-color:#2A3547;}
 .ghost:hover{border-color:#4A5568;}
 .outlookBtn{background:#1D6FBF;}
 .outlookPill{background:#162035;border-color:#2A3D5C;}
 .outlookName{color:#7DD3FC;}
 .regBar{background:#1C2840;border-color:#2A3547;}
 .regOk{color:#34D399;}
 .quick input{background:#1C2840;border-color:#2A3547;color:#E8EDF2;}
 .boards{background:#1C2840;border-color:#2A3547;}
 .boardsHead:hover{background:#1E2D45;}
 .boardsBody{border-color:#2A3547;}
 .chip{background:#22324A;border-color:#374B66;color:#B7C4D4;}
 .chip:hover{border-color:#2DD4BF;color:#7FD8CE;}
 .chip.on{background:#0D9488;border-color:#0D9488;color:#fff;}
 .chipAdd{border-color:#374B66;color:#8B98AB;}
 .chipAdd:hover{border-color:#2DD4BF;color:#7FD8CE;}
 .boardsRow input,.boardsRow select{background:#162035;border-color:#2A3547;color:#E8EDF2;}
 .boardBtn{background:#22324A;border-color:#374B66;color:#C8D4E0;}
 .boardBtn:hover{background:#2A3D57;border-color:#2A6F66;color:#7FD8CE;}
 .boardBtn.sp{background:#0D2A20;border-color:#065F46;color:#34D399;}
 .boardBtn.sp:hover{background:#10352A;border-color:#15803D;color:#6EE7B7;}
 .boardAdd input{background:#162035;border-color:#2A3547;color:#E8EDF2;}
 .boardDel{background:#3A1A1A;color:#FCA5A5;}
 .boardDel:hover{background:#4A2020;}
 .rail{background:#1C2840;border-color:#2A3547;}
 .stat:hover,.stat.on{background:#1E2D45;}
 .railEnd{border-color:#2A3547;}
 .railEnd .big{color:#2DD4BF;}
 .row{background:#1C2840;border-color:#2A3547;}
 .row:hover{box-shadow:0 4px 20px rgba(0,0,0,.35);}
 .controls select{background:#162035;}
 .icon{background:#162035;border-color:#2A3547;color:#7A8899;}
 .icon:hover{border-color:#4A5568;color:#E8EDF2;}
 .iconPin.on{background:#13343A;border-color:#0D9488;color:#2DD4BF;}
 .empty{background:#1C2840;border-color:#2A3547;}
 .empty .big2{color:#E8EDF2;}
 .emailPanel{background:#1C2840;border-color:#2A3547;}
 .emailPanelHead{border-color:#2A3547;}
 .emailItem{border-color:#1E2D45;}
 .emailApplyBtn{background:#162035;}
 .emailDismissBtn{background:#162035;border-color:#2A3547;}
 .scrim{background:rgba(0,0,0,.7);}
 .modal{background:#1C2840;box-shadow:0 24px 80px rgba(0,0,0,.6);}
 .modal h2{color:#F1F5F9;}
 .icloudIntro{color:#AAB6C6;}
 .inboxRow{background:#162035;border-color:#2A3547;}
 .inboxProv{background:#0F1724;color:#8B98AB;}
 .icloudSteps strong,.icloudIntro strong{color:#F1F5F9;}
 .field input,.field select{background:#162035;border-color:#2A3547;color:#E8EDF2;}
 .field span{color:#8B98AB;}
 .pasteArea,.descArea{background:#162035;border-color:#2A3547;color:#E8EDF2;}
 .pasteChip{background:#0D2A20;color:#7FD8CE;border-color:#16433A;}
 .pasteChip.dim{background:#2A2113;color:#E0B477;border-color:#4A3A1E;}
 .descTag{background:#0D2A20;color:#7FD8CE;}
 .closeb{background:#33290F;color:#E0B45E;}
 .closeb.over{background:#3A1A1A;color:#FCA5A5;}
 .tiny{background:#162035;border-color:#34D399;color:#34D399;}
 .tiny:hover{background:#0D2A20;}
 .ft{color:#5A6675;}
 .link:hover{color:#2DD4BF;}
}
`;
