/* ────────────────────────────────────────────────────────────────────
   Fetch a job posting URL and extract structured details — runs in the
   Electron MAIN process (the renderer can't fetch cross-origin pages).

   Strategy:
     1. Download the page HTML.
     2. Prefer JSON-LD <script type="application/ld+json"> "JobPosting" —
        this is the clean, structured source most ATS/employer sites embed
        (title, company, datePosted, validThrough/closing date, salary,
        location, description).
     3. Fall back to Open Graph / <title> meta tags.

   Note: LinkedIn and Indeed block automated fetching and hide jobs behind
   login, so they usually return nothing useful — the renderer falls back to
   the "Paste description" flow in that case.
   ──────────────────────────────────────────────────────────────────── */

const UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 " +
  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36";

const stripHtml = (s) =>
  String(s || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#0?39;|&rsquo;|&apos;/gi, "'")
    .replace(/&quot;|&ldquo;|&rdquo;/gi, '"')
    .replace(/&#?[a-z0-9]+;/gi, " ")
    .replace(/\s+/g, " ")
    .trim();

const isoDate = (s) => {
  if (!s) return "";
  const d = new Date(s);
  return isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
};

async function fetchHtml(url) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 15000);
  try {
    const r = await fetch(url, {
      redirect: "follow",
      signal: ctrl.signal,
      headers: { "User-Agent": UA, Accept: "text/html,application/xhtml+xml", "Accept-Language": "en-GB,en;q=0.9" },
    });
    if (!r.ok) throw new Error(`The page returned HTTP ${r.status}.`);
    return await r.text();
  } catch (e) {
    if (e.name === "AbortError") throw new Error("Timed out fetching the page.");
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

function extractJsonLd(html) {
  const out = [];
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html))) {
    try { out.push(JSON.parse(m[1].trim())); } catch (_) {}
  }
  return out;
}

function findJobPosting(nodes) {
  const stack = [...nodes];
  let guard = 0;
  while (stack.length && guard++ < 5000) {
    const n = stack.pop();
    if (!n || typeof n !== "object") continue;
    if (Array.isArray(n)) { stack.push(...n); continue; }
    const t = n["@type"];
    if (t === "JobPosting" || (Array.isArray(t) && t.includes("JobPosting"))) return n;
    if (n["@graph"]) stack.push(n["@graph"]);
    for (const k of Object.keys(n)) if (n[k] && typeof n[k] === "object") stack.push(n[k]);
  }
  return null;
}

function fromJobPosting(jp) {
  const out = {};
  if (jp.title) out.role = stripHtml(jp.title);

  const org = jp.hiringOrganization;
  if (org) out.company = stripHtml(typeof org === "string" ? org : org.name || "");

  out.postedDate = isoDate(jp.datePosted);
  out.closingDate = isoDate(jp.validThrough);

  const loc = Array.isArray(jp.jobLocation) ? jp.jobLocation[0] : jp.jobLocation;
  const addr = loc && loc.address;
  if (addr && typeof addr === "object") {
    const city = stripHtml(addr.addressLocality || addr.addressRegion || "");
    if (city) out.city = city;
  }
  if (jp.jobLocationType && /telecommute|remote/i.test(String(jp.jobLocationType)) && !out.city) out.city = "Remote";

  const bs = jp.baseSalary;
  const v = bs && bs.value;
  if (v && typeof v === "object") {
    const unit = String(v.unitText || "").toUpperCase();
    if (!unit || unit === "YEAR" || unit === "YEARLY" || unit === "ANNUM") {
      const amt = v.value || v.maxValue || v.minValue;
      const n = parseInt(String(amt || "").replace(/[^\d]/g, ""), 10);
      if (n >= 1000) out.salary = String(n);
    }
  }

  if (jp.description) out.description = stripHtml(jp.description).slice(0, 8000);

  // Drop empties so they don't overwrite the form defaults
  for (const k of Object.keys(out)) if (!out[k]) delete out[k];
  return out;
}

function metaContent(html, key) {
  const re = new RegExp(
    '<meta[^>]+(?:property|name)=["\']' + key + '["\'][^>]+content=["\']([^"\']*)["\']', "i"
  );
  const m = html.match(re);
  return m ? stripHtml(m[1]) : "";
}

function fromMeta(html) {
  const out = {};
  const title = metaContent(html, "og:title") || (html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1];
  const desc = metaContent(html, "og:description") || metaContent(html, "description");
  if (title) out.role = stripHtml(title);
  if (desc) out.description = stripHtml(desc).slice(0, 8000);
  const site = metaContent(html, "og:site_name");
  if (site) out.company = site;
  for (const k of Object.keys(out)) if (!out[k]) delete out[k];
  return out;
}

/* ── LinkedIn: pull the job ID and use the public guest endpoint ─────── */
function linkedinJobId(url) {
  try {
    const u = new URL(url);
    if (!/(^|\.)linkedin\.com$/i.test(u.hostname)) return null;
    const cj = u.searchParams.get("currentJobId");
    if (cj && /^\d+$/.test(cj)) return cj;
    const m =
      u.pathname.match(/\/jobs\/view\/(\d+)/) ||
      u.pathname.match(/-(\d{6,})\/?$/) ||
      u.pathname.match(/(\d{8,})/);
    return m ? m[1] : null;
  } catch (_) { return null; }
}

const relToIso = (phrase) => {
  const m = String(phrase || "").match(/(\d+)\+?\s*(hour|day|week|month)s?\s+ago/i);
  if (!m) return "";
  const n = parseInt(m[1], 10);
  const u = m[2].toLowerCase();
  const days = u === "hour" ? 0 : u === "day" ? n : u === "week" ? n * 7 : n * 30;
  return new Date(Date.now() - days * 86400000).toISOString().slice(0, 10);
};

async function fetchLinkedInGuest(id) {
  const url = `https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/${id}`;
  let html;
  try { html = await fetchHtml(url); } catch (_) { return null; }
  const pick = (re) => { const m = html.match(re); return m ? stripHtml(m[1]) : ""; };

  const out = {};
  out.role = pick(/<h2[^>]*top-card-layout__title[^>]*>([\s\S]*?)<\/h2>/i) ||
             pick(/<h2[^>]*topcard__title[^>]*>([\s\S]*?)<\/h2>/i);
  out.company = pick(/<a[^>]*topcard__org-name-link[^>]*>([\s\S]*?)<\/a>/i) ||
                pick(/<span[^>]*class="[^"]*topcard__flavor(?![^"]*--bullet)[^"]*"[^>]*>([\s\S]*?)<\/span>/i) ||
                pick(/<h4[^>]*topcard__flavor[^>]*>([\s\S]*?)<\/h4>/i);
  const loc = pick(/<span[^>]*topcard__flavor--bullet[^>]*>([\s\S]*?)<\/span>/i);
  if (loc) out.city = loc.split(",")[0].trim();
  const posted = pick(/<span[^>]*posted-time-ago__text[^>]*>([\s\S]*?)<\/span>/i);
  if (posted) { out.posted = posted.toLowerCase(); out.postedDate = relToIso(posted); }
  const desc = pick(/<div[^>]*show-more-less-html__markup[^>]*>([\s\S]*?)<\/div>/i);
  if (desc) out.description = desc.slice(0, 8000);

  for (const k of Object.keys(out)) if (!out[k]) delete out[k];
  return out;
}

async function fetchJob(url) {
  let target = String(url || "").trim();
  if (!/^https?:\/\//i.test(target)) target = "https://" + target;

  // LinkedIn: use the guest endpoint for the specific job (search/job URLs
  // alike). Don't fall back to scraping the page — it's just a login wall.
  const liId = linkedinJobId(target);
  if (liId) {
    const data = await fetchLinkedInGuest(liId);
    if (data && (data.company || data.role)) {
      data.link = target;
      return { useful: true, source: "linkedin", data };
    }
    return {
      useful: false, source: "linkedin", data: { link: target },
      error: "LinkedIn didn't return the job details (it often rate-limits or needs login). Use Paste description for LinkedIn jobs.",
    };
  }

  const html = await fetchHtml(target);
  const jp = findJobPosting(extractJsonLd(html));

  let data, source;
  if (jp) { data = fromJobPosting(jp); source = "structured"; }
  else { data = fromMeta(html); source = "meta"; }

  data.link = target;
  const useful = !!(data.role || data.description);
  return { useful, source, data };
}

function register(ipcMain) {
  ipcMain.handle("jobfetch:fetch", async (_e, { url }) => {
    try {
      const res = await fetchJob(url);
      if (!res.useful) {
        return {
          ok: false,
          error: res.error || "Couldn't read structured job details from that page (the site may require login or block automated reads). Try Paste description instead.",
          data: res.data,
        };
      }
      return { ok: true, source: res.source, data: res.data };
    } catch (e) {
      return { ok: false, error: e.message || "Couldn't fetch that page." };
    }
  });
}

module.exports = { register };
