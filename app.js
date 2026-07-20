/* Copyright (c) 2026 OXMIQ
 * OXMIQ Patent Portfolio Hub — hosted dashboard logic.
 * M365 sign-in (MSAL, delegated) -> Microsoft Graph read of the OXMIQ-IPP site's
 * single Ledger list -> render the v3 UI. No data is embedded; everything is
 * fetched live as the signed-in user (SharePoint permissions enforce visibility).
 *
 * Data model (see doc/data-lifecycle-plan.md): the Ledger is the single source of
 * truth. Rows are classified by Stage:
 *   - Disclosed   -> a disclosure family header (FolderLink.Description = its slug)
 *   - Filed/Granted/Published -> a filing (has Docket; DerivedFrom = parent slug)
 *   - Harvested   -> a candidate (grouped by HarvestSession; no folder, no docket)
 */
"use strict";
const CFG = window.IPP_CONFIG;

/* ----------------------------------------------------------------- diagnostics */
const DIAG = [];
function diag(step, ok, detail) {
  DIAG.push({ step, ok, detail: detail || "" });
  renderDiag();
  (ok ? console.log : console.warn)(`[IPP] ${step}: ${ok ? "OK" : "FAIL"} ${detail || ""}`);
}
function renderDiag() {
  const el = document.getElementById("diagBody");
  if (!el) return;
  el.innerHTML = DIAG.map(d =>
    `<div class="diag-row"><span class="diag-dot ${d.ok ? "ok" : "bad"}"></span>
     <span class="diag-step">${escp(d.step)}</span>
     <span class="diag-detail">${escp(d.detail)}</span></div>`).join("");
}
const escp = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/* ------------------------------------------------------------------- app state */
let LEDGER = [];       // every mapped Ledger row
let DISCLOSURES = [];  // Stage = Disclosed
let FILINGS = [];      // Stage in {Filed, Granted, Published}
let SPEND = [];        // Legal Spend list rows (exec/counsel only)
let HARVESTED = [];    // Stage = Harvested
let FAMILIES = [];     // [{disc, filings:[...]}] — disclosure -> its filings (fan-out)
const BYKEY = new Map();      // key -> record (docket for filings, slug for disclosures)
const FAM_BY_SLUG = new Map();// disclosure slug -> family {disc, filings}
let ME = null;         // {displayName, mail}
let msalApp = null;

/* --------------------------------------------------------------------- MSAL */
async function initAuth() {
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
      redirectUri: CFG.redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
  });
  await msalApp.initialize();
  diag("MSAL initialised", true, `client ${CFG.clientId.slice(0, 8)}… · tenant ${CFG.tenantId.slice(0, 8)}…`);

  // complete a redirect sign-in if we're coming back from one
  const redirectResult = await msalApp.handleRedirectPromise();
  if (redirectResult && redirectResult.account) {
    msalApp.setActiveAccount(redirectResult.account);
    diag("Redirect sign-in completed", true, redirectResult.account.username);
  }
  const accounts = msalApp.getAllAccounts();
  if (accounts.length && !msalApp.getActiveAccount()) msalApp.setActiveAccount(accounts[0]);
  return msalApp.getActiveAccount();
}

async function getToken() {
  const account = msalApp.getActiveAccount();
  const request = { scopes: CFG.scopes, account };
  try {
    const r = await msalApp.acquireTokenSilent(request);
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent|login_required/i.test(e.errorCode || e.message || "")) {
      diag("Silent token failed — redirecting for consent", false, e.errorCode || e.message);
      await msalApp.acquireTokenRedirect(request);   // leaves the page
      return null;
    }
    throw e;
  }
}

function signIn() {
  msalApp.loginRedirect({ scopes: CFG.scopes });
}
function signOut() {
  const account = msalApp.getActiveAccount();
  msalApp.logoutRedirect({ account, postLogoutRedirectUri: CFG.redirectUri });
}

/* -------------------------------------------------------------------- Graph */
async function graphGet(path, token) {
  const url = path.startsWith("http") ? path : CFG.graphBase + path;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json" } });
  if (!res.ok) {
    let code = res.status, msg = "";
    try { const b = await res.json(); code = b.error?.code || res.status; msg = b.error?.message || ""; } catch (_) {}
    const err = new Error(`${res.status} ${code}: ${msg}`);
    err.status = res.status; err.code = code;
    throw err;
  }
  return res.json();
}

async function graphGetAll(path, token) {
  let url = path, items = [];
  for (let i = 0; i < 20 && url; i++) {
    const page = await graphGet(url, token);
    items = items.concat(page.value || []);
    url = page["@odata.nextLink"] || null;
  }
  return items;
}

async function listItems(listName, token) {
  const path = `/sites/${CFG.siteId}/lists/${encodeURIComponent(listName)}/items`
    + `?$expand=fields&$select=id&$top=500`;
  return graphGetAll(path, token);
}

/* ------------------------------------------------------------- field mapping */
const d10 = s => (s ? String(s).slice(0, 10) : "");
const splitList = s => String(s || "").split(";").map(x => x.trim()).filter(Boolean);
// A hyperlink column comes back from Graph as {Url, Description}; be defensive.
function linkUrl(v) { return v && typeof v === "object" ? (v.Url || v.url || "") : (v || ""); }
function linkDesc(v) { return v && typeof v === "object" ? (v.Description || v.description || "") : ""; }
const truthy = v => v === true || v === 1 || v === "1" || v === "Yes" || v === "true";

/** Map one raw Graph list item into a normalised Ledger record. */
function mapRow(it) {
  const f = it.fields || {};
  const stage = f.Stage || "";
  let kind = "other";
  if (stage === "Disclosed") kind = "disclosure";
  else if (stage === "Filed" || stage === "Granted" || stage === "Published") kind = "filing";
  else if (stage === "Harvested") kind = "harvested";

  const docket = f.Docket || "";
  const folderSlug = linkDesc(f.FolderLink);   // for a disclosure this is its slug
  // stable per-row key: docket for filings, folder slug for disclosures, else id
  const key = kind === "filing" ? (docket || `F-${it.id}`)
    : kind === "disclosure" ? (folderSlug || `D-${it.id}`)
      : `H-${it.id}`;

  return {
    id: it.id, key, kind, stage,
    title: f.Title || "Untitled",
    desc: f.Description || "",
    origin: f.Origin || "",
    docket,
    filingType: f.FilingType || "",
    status: f.Status || "",
    filingDate: d10(f.FilingDate),
    inventors: f.Inventors || "",
    techArea: f.TechArea || "",
    cpc: f.CPC || "",
    appNo: f.ApplicationNo || "",
    summary: f.AISummary || "",
    counsel: f.AssignedCounsel || "",
    pub: truthy(f.PublishToPortfolio),
    related: splitList(f.RelatedDockets),
    derivedFrom: f.DerivedFrom || "",   // disclosure-level lineage (parent slug)
    harvestSession: f.HarvestSession || "",
    slug: folderSlug,
    folderUrl: linkUrl(f.FolderLink),
    folderId: f.FolderId || "",
  };
}

/** Build LEDGER + the classified buckets + the disclosure->filings families. */
function classify(items) {
  LEDGER = items.map(mapRow);
  BYKEY.clear();
  LEDGER.forEach(r => { if (!BYKEY.has(r.key)) BYKEY.set(r.key, r); });

  DISCLOSURES = LEDGER.filter(r => r.kind === "disclosure")
    .sort((a, b) => a.title.localeCompare(b.title));
  FILINGS = LEDGER.filter(r => r.kind === "filing")
    .sort((a, b) => a.docket.localeCompare(b.docket));
  HARVESTED = LEDGER.filter(r => r.kind === "harvested");

  // Fan-out: each disclosure -> the filings whose DerivedFrom = its slug, OR whose
  // docket is listed in the disclosure's RelatedDockets.
  FAM_BY_SLUG.clear();
  const claimed = new Set();
  FAMILIES = DISCLOSURES.map(disc => {
    const rel = new Set(disc.related);
    const filings = FILINGS.filter(fn =>
      (disc.slug && fn.derivedFrom === disc.slug) || rel.has(fn.docket));
    filings.forEach(fn => claimed.add(fn.key));
    const fam = { disc, filings };
    if (disc.slug) FAM_BY_SLUG.set(disc.slug, fam);
    return fam;
  });
  // any filing not attached to a disclosure -> an "Unassigned" family (defensive)
  const orphans = FILINGS.filter(fn => !claimed.has(fn.key));
  if (orphans.length) FAMILIES.push({ disc: null, filings: orphans });
}

/* -------------------------------------------------------------------- render
   (adapted from storyboard v3 — same DOM, data now comes from the Ledger)     */
const STATUS = { "Office action": { c: "var(--oa)", b: "b-oa" }, "Pending": { c: "var(--pending)", b: "b-pending" },
  "Expired": { c: "var(--expired)", b: "b-expired" }, "Granted": { c: "var(--ok)", b: "b-granted" },
  "Provisional": { c: "var(--prov)", b: "b-prov" }, "Draft": { c: "var(--draft)", b: "b-draft" },
  "Abandoned": { c: "var(--expired)", b: "b-expired" },
  "Disclosed": { c: "var(--prov)", b: "b-prov" }, "Harvested": { c: "var(--draft)", b: "b-draft" },
  "Filed": { c: "var(--ok)", b: "b-granted" } };
const AREA = { "Compute / Cores": "#5b8cff", "Software / Runtime": "#7c5cff", "Packaging": "#43b581",
  "Memory / Fabric": "#e0a13a", "Datacenter / System": "#f0616d" };
const esc = escp;
const pill = s => { const m = STATUS[s] || STATUS.Pending; return `<span class="badge ${m.b}"><span class="dot" style="background:${m.c}"></span>${esc(s || "—")}</span>`; };
const inv1 = s => (s || "").replace(" (first named)", "");

/* access gate, per role:
   - committee / exec -> every filing
   - counsel          -> the filings assigned to their firm (AssignedCounsel === CFG.counselFirm),
                         regardless of the publish flag — a scoped navigation/status view
   - employee (default)-> only PublishToPortfolio filings (the public portfolio) */
function canSeeAll() { return role === "committee" || role === "exec"; }
function visibleFilter(f) {
  if (canSeeAll()) return true;
  if (role === "counsel") return !!f.counsel && f.counsel === CFG.counselFirm;
  return f.pub;
}
function visibleFilings() { return FILINGS.filter(visibleFilter); }
function isMine(r) {
  const nm = (ME && ME.displayName || "").trim().toLowerCase();
  if (!nm) return false;
  const inv = (r.inventors || "").toLowerCase();
  if (inv.includes(nm)) return true;
  return nm.split(/\s+/).filter(t => t.length > 2).some(t => inv.includes(t));
}

function tally(arr, key) { const o = {}; arr.forEach(p => { const k = p[key] || "—"; o[k] = (o[k] || 0) + 1; }); return o; }
function bars(el, counts, colorFn) {
  if (!el) return;
  const max = Math.max(1, ...Object.values(counts));
  el.innerHTML = Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([k, v]) =>
    `<div class="barrow"><span class="k"><span class="swatch" style="background:${colorFn(k)}"></span>${esc(k)}</span><div class="bar"><i style="width:${Math.round(v / max * 100)}%;background:${colorFn(k)}"></i></div><span class="v tabnum">${v}</span></div>`).join("");
}

/* --- dashboard --- */
function renderFunnel(el) {
  if (!el) return;
  const stages = [
    ["Harvested", HARVESTED.length],
    ["Disclosed", DISCLOSURES.length],
    ["Filed", FILINGS.filter(f => f.stage === "Filed").length],
  ];
  const granted = FILINGS.filter(f => f.stage === "Granted" || f.stage === "Published").length;
  if (granted) stages.push(["Granted", granted]);
  el.innerHTML = stages.map(([l, n]) =>
    `<div class="stage"><div class="n tabnum">${n}</div><div class="l">${esc(l)}</div></div>`).join("");
}
function renderDashboard() {
  bars(document.getElementById("areaBars"), tally(FILINGS, "techArea"), k => AREA[k] || "#5b8cff");
  const counts = tally(FILINGS, "status"), seg = Object.entries(counts), tot = FILINGS.length;
  let acc = 0;
  const parts = seg.map(([k, v]) => { const a = acc / (tot || 1) * 360, b = (acc + v) / (tot || 1) * 360; acc += v; return `${(STATUS[k] || STATUS.Pending).c} ${a}deg ${b}deg`; });
  const donut = document.getElementById("donut");
  if (donut) donut.innerHTML = `<div style="width:120px;height:120px;border-radius:50%;background:conic-gradient(${parts.join(",")});display:grid;place-items:center"><div style="width:74px;height:74px;border-radius:50%;background:var(--panel);display:grid;place-items:center"><div style="text-align:center"><div style="font-size:20px;font-weight:750">${tot}</div><div class="dim" style="font-size:10px">filings</div></div></div></div>`;
  const leg = document.getElementById("statusLegend");
  if (leg) leg.innerHTML = seg.map(([k, v]) => `<div><span class="swatch" style="background:${(STATUS[k] || STATUS.Pending).c}"></span>${esc(k)} · ${v}</div>`).join("");
  const rec = document.getElementById("recent");
  if (rec) rec.innerHTML = [...FILINGS].sort((a, b) => (b.filingDate || "").localeCompare(a.filingDate || "")).slice(0, 5).map(p =>
    `<div class="ticket"><b>${esc(p.docket)}</b> — ${esc(p.title)} <span class="who">${esc(p.techArea || "")} · ${esc(p.filingDate || "")}</span></div>`).join("");
  const oa = document.getElementById("oaTable");
  if (oa) oa.innerHTML = FILINGS.filter(p => p.status === "Office action").map(p =>
    `<tr><td class="docket">${esc(p.docket)}</td><td>${esc(p.title)}</td><td class="muted">${esc(p.appNo)}</td><td>${esc(p.counsel || "Counsel")}</td></tr>`).join("") || `<tr><td colspan="4" class="dim">No open office actions.</td></tr>`;
  // top-line KPI counts (from the Ledger)
  setText("kpiTotal", FILINGS.length);
  setText("kpiOA", FILINGS.filter(p => p.status === "Office action").length);
  setText("kpiPending", FILINGS.filter(p => p.status === "Pending").length);
  setText("kpiExpired", FILINGS.filter(p => p.status === "Expired").length);
  const funnelEl = document.getElementById("funnel");
  if (funnelEl) renderFunnel(funnelEl);
}
function setText(id, v) { const el = document.getElementById(id); if (el) el.textContent = v; }

/* --- explorer (filings grouped by disclosure family) --- */
let facet = "all";
const collapsed = new Set();   // slugs whose family group is collapsed
function setFacet(f, el) { facet = f; document.querySelectorAll("#facets .facet").forEach(x => x.classList.remove("on")); el.classList.add("on"); renderGallery(); }
function toggleFam(slug) { if (collapsed.has(slug)) collapsed.delete(slug); else collapsed.add(slug); renderGallery(); }

function matchFacet(f) { return facet === "all" || f.status === facet || f.techArea === facet; }
function matchSearch(f, q) {
  if (!q) return true;
  return [f.docket, f.title, f.status, f.techArea, f.appNo, f.inventors].join(" ").toLowerCase().includes(q);
}
/* families with only the currently-visible filings in each */
function familiesForView() {
  const q = ((document.getElementById("q") || {}).value || "").toLowerCase();
  const mineOn = !!(document.getElementById("mineChk") && document.getElementById("mineChk").checked);
  return FAMILIES.map(fm => {
    let kids = fm.filings.filter(visibleFilter);
    if (mineOn) kids = kids.filter(isMine);
    kids = kids.filter(f => matchSearch(f, q) && matchFacet(f));
    return { disc: fm.disc, kids };
  }).filter(fm => fm.kids.length);
}
function fcard(p) {
  return `<div class="pcard" data-d="${esc(p.key)}" onclick="openPat('${esc(p.key)}')"><span class="docket">${esc(p.docket)}</span>
    <div class="body"><div class="ttl">${esc(p.title)}${isMine(p) ? ' <span class="badge b-mine" style="margin-left:4px">mine</span>' : ''}</div>
    <div class="meta">${pill(p.status)}<span>${esc(p.techArea || "")}</span><span>🧾 ${esc(p.appNo || "—")}</span></div></div></div>`;
}
function renderGallery() {
  const mineOn = !!(document.getElementById("mineChk") && document.getElementById("mineChk").checked);
  const mt = document.getElementById("mineToggle"); if (mt) mt.classList.toggle("on", mineOn);
  const view = familiesForView();
  const total = view.reduce((n, fm) => n + fm.kids.length, 0);
  const gal = document.getElementById("gallery");
  gal.innerHTML = view.map(fm => {
    const slug = fm.disc ? fm.disc.slug : "unassigned";
    const label = fm.disc ? fm.disc.title : "Unassigned filings";
    const isColl = collapsed.has(slug);
    return `<div class="famgroup${isColl ? " collapsed" : ""}">
      <div class="famhead" onclick="toggleFam('${esc(slug)}')">
        <span class="caret">▾</span>
        <span class="fam-t">${esc(label)} <span class="fam-slug">${esc(slug)}</span></span>
        <span class="cnt">${fm.kids.length}</span></div>
      <div class="famchildren">${fm.kids.map(fcard).join("")}</div></div>`;
  }).join("") || `<div class="rescount">${emptyGalleryMsg(mineOn)}</div>`;

  const totalSrc = visibleFilings().length;
  const scope = role === "counsel" ? " · your dockets" : "";
  document.getElementById("resCount").textContent =
    `${total} of ${totalSrc} filings${scope} · ${view.length} famil${view.length === 1 ? "y" : "ies"}${mineOn ? " · yours" : ""}${facet !== "all" ? " · " + facet : ""}`;

  const first = view.length ? view[0].kids[0].key : null;
  if (first) openPat(first);
  else document.getElementById("detail").innerHTML = '<div class="dim" style="padding:30px;text-align:center">Nothing selected</div>';
}

function emptyGalleryMsg(mineOn) {
  if (mineOn) return "No filings list you as an inventor.";
  if (role === "counsel") return "No dockets are currently assigned to your firm.";
  if (canSeeAll()) return "No filings match.";
  return "No published filings match. (In-review filings stay private to their members.)";
}

function openPat(key) {
  const p = BYKEY.get(key);
  if (!p) return;
  // gate: don't reveal a filing's detail the current role isn't scoped to
  if (p.kind === "filing" && !visibleFilter(p)) return;
  document.querySelectorAll("#gallery .pcard").forEach(c => c.classList.toggle("sel", c.dataset.d === key));
  document.getElementById("detail").innerHTML = p.kind === "disclosure" ? detailDisclosure(p) : detailFiling(p);
}
function detailFiling(p) {
  const rel = p.related.filter(Boolean);
  const chips = [];
  if (p.derivedFrom && FAM_BY_SLUG.has(p.derivedFrom)) {
    const d = FAM_BY_SLUG.get(p.derivedFrom).disc;
    chips.push(`<span class="chip" onclick="openPat('${esc(d.key)}')">🗂️ ${esc(d.title)}</span>`);
  }
  rel.forEach(r => { if (BYKEY.has(r)) chips.push(`<span class="chip" onclick="openPat('${esc(r)}')">📄 ${esc(r)}</span>`); });
  return `
    <div class="dk">OXMIQ.${esc(p.docket)} · ${esc(p.filingType || "")}</div>
    <h2>${esc(p.title)}</h2>${pill(p.status)}
    ${folderBtn(p.folderUrl, "Open filing folder in SharePoint →")}
    <div class="abs">${esc(p.summary || "Summary pending (nightly agent).")}</div>
    <div class="metagrid">
      <div class="metabox"><div class="ml">Application #</div><div class="mv">${esc(p.appNo || "—")}</div></div>
      <div class="metabox"><div class="ml">Filing date</div><div class="mv">${esc(p.filingDate || "—")}</div></div>
      <div class="metabox"><div class="ml">Technology area</div><div class="mv">${esc(p.techArea || "—")}</div></div>
      <div class="metabox"><div class="ml">Inventors</div><div class="mv">${esc(inv1(p.inventors) || "—")}</div></div>
    </div>
    <div class="lbl">Family / related</div>
    <div class="chips">${chips.length ? chips.join("") : '<span class="chip" style="cursor:default">None linked</span>'}</div>
    <div class="lbl">Status timeline</div>
    <div class="timeline">
      <div class="tl"><b>Filed</b><div class="dt">${esc(p.filingDate || "—")}</div></div>
      <div class="tl ${p.status === 'Office action' ? '' : 'd'}">${p.status === 'Office action' ? '<b>Office action — reply due</b>' : 'Under examination'}<div class="dt">USPTO</div></div>
      <div class="tl d">${p.status === 'Expired' ? '<b>Provisional expired</b> (converted via children)' : 'Grant (target)'}<div class="dt">stealth until grant</div></div>
    </div>`;
}
/* prominent SharePoint jump — the real counsel<->inventor exchange lives in the folder */
function folderBtn(url, label) {
  if (!url) return "";
  return `<div style="margin:12px 0"><a class="btn btn-primary" href="${esc(url)}" target="_blank" rel="noopener">📂 ${esc(label)}</a></div>`;
}
function detailDisclosure(p) {
  const fam = FAM_BY_SLUG.get(p.slug);
  const kids = fam ? fam.filings : [];
  const vis = kids.filter(visibleFilter);
  const chips = vis.map(k => `<span class="chip" onclick="openPat('${esc(k.key)}')">📄 ${esc(k.docket)}</span>`);
  return `
    <div class="dk">DISCLOSURE · ${esc(p.slug)}</div>
    <h2>${esc(p.title)}</h2>
    <span class="badge b-prov"><span class="dot" style="background:var(--prov)"></span>Disclosed</span>
    ${folderBtn(p.folderUrl, "Open disclosure folder in SharePoint →")}
    <div class="abs">${esc(p.summary || p.desc || "Disclosure family — the founding record its filings derive from.")}</div>
    <div class="metagrid">
      <div class="metabox"><div class="ml">Filings in family</div><div class="mv">${kids.length}</div></div>
      <div class="metabox"><div class="ml">Technology area</div><div class="mv">${esc(p.techArea || "—")}</div></div>
      <div class="metabox"><div class="ml">Assigned counsel</div><div class="mv">${esc(p.counsel || "—")}</div></div>
      <div class="metabox"><div class="ml">Origin</div><div class="mv">${esc(p.origin || "—")}</div></div>
    </div>
    <div class="lbl">Filings (fan-out)</div>
    <div class="chips">${chips.length ? chips.join("") : '<span class="chip" style="cursor:default">No visible filings</span>'}</div>`;
}

/* --- harvesting (candidates grouped by HarvestSession) --- */
function harvestSessions() {
  const m = {};
  HARVESTED.forEach(h => { const k = h.harvestSession || "Unsessioned"; (m[k] = m[k] || []).push(h); });
  return Object.entries(m).sort((a, b) => b[0].localeCompare(a[0]));
}
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function sessBadge(date) {
  const m = /^(\d{4})-(\d{2})-\d{2}$/.exec(date);
  return m ? `${m[1]}<br>${MONTHS[+m[2] - 1] || ""}` : "—";
}
function renderHarvest() {
  const sessions = harvestSessions();
  const sess = document.getElementById("sessions");
  if (sess) sess.innerHTML = sessions.map(([date, items]) =>
    `<div class="session"><div class="q">${sessBadge(date)}</div><div style="flex:1"><b>Harvest — ${esc(date)}</b><div class="muted">${items.length} candidate${items.length === 1 ? "" : "s"} captured · awaiting committee triage</div></div><span class="badge b-draft"><span class="dot" style="background:var(--draft)"></span>Harvested</span></div>`).join("")
    || '<div class="dim">No harvest sessions on file.</div>';
  const kb = document.getElementById("kanban");
  if (kb) kb.innerHTML = sessions.map(([date, items]) =>
    `<div class="col"><h4>${esc(date)} <span>${items.length}</span></h4>${items.map(i =>
      `<div class="kt">${esc(i.title)}<div class="who">📂 ${esc(i.techArea || "—")}</div></div>`).join("") || '<div class="who" style="padding:6px;color:var(--dim)">—</div>'}</div>`).join("");
}

/* --- role from the Roster List (UX only; real gating is M365 permissions) ---
   Read the Roster, find the signed-in user's row, and derive which dashboard
   views they may switch between + their counsel firm. Fail-safe: any error, no
   row, inactive, or no read permission -> Employee. Never a security boundary —
   the reconciler + SharePoint permissions are what actually enforce access. */
let ENTITLED = ["employee"];
const ROSTER_TO_VIEW = { committee: "committee", exec: "exec", counsel: "counsel" };
async function resolveAccessFromRoster(token) {
  const mine = [ME.mail, ME.userPrincipalName, ME.username]
    .filter(Boolean).map(s => s.toLowerCase());
  try {
    const rows = await listItems("Roster", token);
    const row = rows.map(r => r.fields || {})
      .find(f => mine.includes(String(f.Email || "").trim().toLowerCase()));
    if (!row || !truthy(row.Active)) { diag("Role from Roster", true, "Employee (not listed)"); return "employee"; }
    const roles = (Array.isArray(row.Roles) ? row.Roles : splitList(row.Roles)).map(x => String(x).toLowerCase());
    const views = [...new Set(roles.map(r => ROSTER_TO_VIEW[r]).filter(Boolean))];
    if (row.CounselFirm) CFG.counselFirm = row.CounselFirm;   // firm from the roster, not hardcoded
    ENTITLED = ["employee", ...views];
    const start = ["exec", "committee", "counsel"].find(v => views.includes(v)) || "employee";
    diag("Role from Roster", true, `${start} · entitled: ${ENTITLED.join(", ")}`);
    return start;
  } catch (e) {
    diag("Role from Roster", true, `Employee (roster unreadable: ${e.code || e.message})`);
    return "employee";
  }
}

/* role + nav — default and switch options come from ENTITLED (the Roster);
   the switcher only offers views the signed-in user actually holds. */
const access = { employee: ["dash", "explorer", "submit", "sharepoint"],
  committee: ["dash", "explorer", "submit", "sharepoint"],
  exec: ["dash", "explorer", "submit", "sharepoint", "spend"],
  counsel: ["dash", "explorer", "sharepoint", "spend"] };
let role = "employee", current = "dash";
function setRole(r) {
  if (!ENTITLED.includes(r)) r = ENTITLED[0] || "employee";
  role = r; document.getElementById("rolewrap").setAttribute("data-role", r);
  document.querySelectorAll("#roleSwitch button").forEach(b => {
    b.classList.toggle("active", b.dataset.role === r);
    b.classList.toggle("hidden", !ENTITLED.includes(b.dataset.role));
  });
  document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("hidden", !access[r].includes(a.dataset.page)));
  // status-chart export: non-employees only, counsel scoped to their firm
  document.querySelectorAll(".needs-role").forEach(el => el.classList.toggle("hidden", r === "employee"));
  const hintEl = document.getElementById("chartHint");
  if (hintEl) {
    const n = (r === "counsel" ? FILINGS.filter(f => f.counsel === CFG.counselFirm) : FILINGS).length;
    hintEl.textContent = r === "counsel" ? `${n} of your firm's filings` : `${n} filings, live from the Ledger`;
  }
  if (!access[r].includes(current)) go(access[r][0]);
  renderGallery();  // publish filter depends on role
  renderSpend();    // counsel scoping depends on role
}
function go(p) {
  if (!access[role].includes(p)) return; current = p;
  document.querySelectorAll(".page").forEach(s => s.classList.remove("show"));
  document.getElementById("page-" + p).classList.add("show");
  document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("active", a.dataset.page === p));
  window.scrollTo(0, 0);
}
function act(btn, msg) { btn.closest(".actions").innerHTML = `<span class="badge b-granted"><span class="dot" style="background:var(--ok)"></span>✓ ${esc(msg)}</span>`; }

/* people picker (mock directory — production resolves via Graph /users) */
const DIR = [["S. Nadar", "sundar.nadar@oxmiq.ai"], ["L. Park", "lena.park@oxmiq.ai"], ["Mark Leather", "mark.leather@oxmiq.ai"], ["Micah Villmow", "micah.villmow@oxmiq.ai"]];
function ppSuggest(v) {
  const box = document.getElementById("ppSuggest"); v = v.trim().toLowerCase();
  if (!v) { box.classList.remove("show"); return; }
  const hits = DIR.filter(([n, e]) => n.toLowerCase().includes(v) || e.includes(v)).slice(0, 5);
  box.innerHTML = hits.map(([n, e]) => `<div class="pp-row" onclick="ppAdd('${n}','${e}')"><span class="avatar">${n.split(' ').map(x => x[0]).join('').slice(0, 2)}</span>${n}<span class="em">${e}</span></div>`).join("") || '<div class="pp-row dim">No match</div>';
  box.classList.add("show");
}
function ppAdd(n) {
  const inp = document.getElementById("ppInput");
  const chip = document.createElement("span"); chip.className = "person";
  chip.innerHTML = `<span class="avatar">${n.split(' ').map(x => x[0]).join('').slice(0, 2)}</span> ${n} <span class="x" onclick="this.parentNode.remove()">✕</span>`;
  inp.parentNode.insertBefore(chip, inp); inp.value = ""; document.getElementById("ppSuggest").classList.remove("show");
}

function renderAll() {
  renderDashboard();
  renderGallery();
  renderSpend();
}

/* ---------------------------------------------------------------- legal spend
   Exec/counsel view over the Legal Spend list (structured invoice records —
   amounts are a real Currency field, not parsed from PDFs). Exec sees all;
   counsel is scoped to their firm. Committee/employees never load it. */
function mapSpend(it) {
  const f = it.fields || {};
  return { amount: Number(f.Amount) || 0, firm: f.Firm || "", docket: f.Docket || "",
    date: d10(f.InvoiceDate), status: f.Status || "" };
}
async function loadSpend(token) {
  try {
    const rows = await listItems(CFG.legalSpendList, token);
    SPEND = rows.map(mapSpend);
    diag("Read Legal Spend", true, `${SPEND.length} invoice(s)`);
  } catch (e) { SPEND = []; diag("Read Legal Spend", true, `skipped (${e.code || e.message})`); }
}
const money = n => "$" + (Number(n) || 0).toLocaleString(undefined, { maximumFractionDigits: 0 });
function setMoney(id, n) { const el = document.getElementById(id); if (el) el.textContent = money(n); }
function lastQuarterRange(now) {
  let y = now.getFullYear(), q = Math.floor(now.getMonth() / 3) - 1;
  if (q < 0) { q = 3; y -= 1; }
  return [new Date(y, q * 3, 1), new Date(y, q * 3 + 3, 1)];
}
function renderSpend() {
  let rows = SPEND.slice();
  if (role === "counsel") rows = rows.filter(r => r.firm && r.firm === CFG.counselFirm);
  const now = new Date(), yr = now.getFullYear(), [lqs, lqe] = lastQuarterRange(now);
  const sum = a => a.reduce((s, r) => s + (r.amount || 0), 0);
  setMoney("spTotal", sum(rows));
  setMoney("spYTD", sum(rows.filter(r => r.date && new Date(r.date).getFullYear() === yr)));
  setMoney("spLQ", sum(rows.filter(r => { if (!r.date) return false; const d = new Date(r.date); return d >= lqs && d < lqe; })));
  setMoney("spPending", sum(rows.filter(r => r.status === "Pending")));
  const t = document.getElementById("spTable");
  if (t) t.innerHTML = rows.length
    ? [...rows].sort((a, b) => (b.date || "").localeCompare(a.date || "")).slice(0, 15).map(r =>
      `<tr><td class="docket">${esc(r.docket || "—")}</td><td>${esc(r.firm || "")}</td><td class="muted">${esc(r.date || "")}</td><td>${money(r.amount)}</td><td>${esc(r.status || "")}</td></tr>`).join("")
    : `<tr><td colspan="5" class="dim">No invoices logged yet — add them in the Legal Spend list (Open SharePoint).</td></tr>`;
}

/* --------------------------------------------------------- status chart export
   Generate counsel's filings chart on demand from the live Ledger — retires the
   hand-maintained Status_Charts snapshots (one canonical format, no drift).
   Role-gated to non-employees; counsel scoped to their firm. Inventors optional
   (strip for external / VC / data-vault sharing). Client-side .xlsx, fully
   self-contained: a store-only ZIP of inline-string OOXML — no library, no build. */
function statusChartRows(includeInventors) {
  let items = FILINGS.slice();
  if (role === "counsel") items = items.filter(f => f.counsel && f.counsel === CFG.counselFirm);
  items.sort((a, b) => (a.docket || "").localeCompare(b.docket || ""));
  const header = ["Docket #", "Application #", "Filing Date", "Title", "Status"];
  if (includeInventors) header.push("Inventors");
  const rows = [header];
  for (const f of items) {
    const r = [f.docket, f.appNo, f.filingDate, f.title, f.status];
    if (includeInventors) r.push(f.inventors);
    rows.push(r);
  }
  return rows;
}
function exportChart() {
  if (role === "employee") return;                 // UX guard — the button is hidden anyway
  const inc = document.getElementById("incInv").checked;
  const rows = statusChartRows(inc);
  const today = new Date().toISOString().slice(0, 10);
  downloadXlsx(`Patent_Filings_${today}.xlsx`, "Patent Filings", rows);
}

/* Submit an Idea — open the SharePoint intake form in a new tab (seamless SSO;
   the submission flow does the rest). Source returns them here after submit. */
function openIntake() {
  // env=WebView gives a focused form — hides the left nav, suite bar, and site
  // header. Source returns the user to the dashboard after Save.
  const url = CFG.intakeFormUrl + "?env=WebView&Source=" + encodeURIComponent(location.href);
  window.open(url, "_blank", "noopener");
}

/* --- minimal, dependency-free .xlsx writer --- */
function xlsxEsc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function sheetXml(rows) {
  const A = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
  const body = rows.map((row, r) =>
    `<row r="${r + 1}">` + row.map((v, c) =>
      `<c r="${A[c]}${r + 1}" t="inlineStr"><is><t xml:space="preserve">${xlsxEsc(v)}</t></is></c>`).join("") + "</row>").join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}
function crc32(bytes) {
  let crc = ~0;
  for (let i = 0; i < bytes.length; i++) {
    crc ^= bytes[i];
    for (let k = 0; k < 8; k++) crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
  }
  return (~crc) >>> 0;
}
function zipStore(files) {
  const enc = new TextEncoder();
  const u16 = n => [n & 255, (n >> 8) & 255];
  const u32 = n => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
  const parts = [], central = []; let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name), data = f.data, crc = crc32(data);
    const head = new Uint8Array([].concat(
      u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0)));
    parts.push(head, name, data);
    central.push({ name, crc, size: data.length, offset });
    offset += head.length + name.length + data.length;
  }
  const cd = []; let cdSize = 0;
  for (const c of central) {
    const rec = new Uint8Array([].concat(
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(c.crc), u32(c.size), u32(c.size), u16(c.name.length),
      u16(0), u16(0), u16(0), u16(0), u32(0), u32(c.offset)));
    cd.push(rec, c.name); cdSize += rec.length + c.name.length;
  }
  const end = new Uint8Array([].concat(
    u32(0x06054b50), u16(0), u16(0), u16(central.length), u16(central.length),
    u32(cdSize), u32(offset), u16(0)));
  const all = parts.concat(cd, [end]);
  const out = new Uint8Array(all.reduce((n, a) => n + a.length, 0));
  let p = 0; for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}
function downloadXlsx(filename, sheetName, rows) {
  const enc = new TextEncoder();
  const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;
  const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const WB = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${xlsxEsc(sheetName).slice(0, 31)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const WBR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  const files = [
    { name: "[Content_Types].xml", data: enc.encode(CT) },
    { name: "_rels/.rels", data: enc.encode(RELS) },
    { name: "xl/workbook.xml", data: enc.encode(WB) },
    { name: "xl/_rels/workbook.xml.rels", data: enc.encode(WBR) },
    { name: "xl/worksheets/sheet1.xml", data: enc.encode(sheetXml(rows)) },
  ];
  const blob = new Blob([zipStore(files)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
}

/* ------------------------------------------------------------------- boot */
function show(id) { const el = document.getElementById(id); if (el) el.style.display = ""; }
function hide(id) { const el = document.getElementById(id); if (el) el.style.display = "none"; }
function fatal(title, detail) {
  hide("app"); hide("signin");
  const box = document.getElementById("errbox");
  box.style.display = ""; box.querySelector(".err-title").textContent = title;
  box.querySelector(".err-detail").textContent = detail || "";
}

async function boot() {
  try {
    const account = await initAuth();
    if (!account) { hide("app"); hide("errbox"); show("signin"); diag("Not signed in", true, "showing sign-in"); return; }
    ME = { displayName: account.name || account.username, mail: account.username };
    setText("whoami", ME.displayName);
    hide("signin");

    const token = await getToken();
    if (!token) return;   // redirecting for consent
    diag("Token acquired", true, scopesFrom(token));

    // /me (profile) — best effort
    try { const me = await graphGet("/me?$select=displayName,mail,userPrincipalName", token); ME = { ...ME, ...me }; setText("whoami", ME.displayName || ME.mail); diag("Read /me", true, ME.mail || ME.userPrincipalName || ""); }
    catch (e) { diag("Read /me", false, e.code || e.message); }

    // site sanity
    try { const site = await graphGet(`/sites/${CFG.siteId}?$select=displayName,webUrl`, token); diag("Resolved site", true, site.displayName); }
    catch (e) { diag("Resolved site", false, `${e.code} — ${siteHint(e)}`); }

    // The Ledger — single source of truth
    try {
      const items = await listItems(CFG.ledgerList, token);
      classify(items);
      diag("Read Ledger list", true, `${LEDGER.length} items · ${DISCLOSURES.length} disclosures, ${FILINGS.length} filings, ${HARVESTED.length} harvested`);
    } catch (e) {
      diag("Read Ledger list", false, `${e.code} — ${listHint(e)}`);
      fatal("Couldn't read the Ledger list", `${e.message}. ${listHint(e)}`);
      return;
    }

    // role + counsel firm come from the Roster (fail-safe to Employee)
    const startRole = await resolveAccessFromRoster(token);
    if (ENTITLED.includes("exec") || ENTITLED.includes("counsel")) await loadSpend(token);

    show("app");
    setRole(startRole);
    renderAll();
    diag("Rendered", true, `${FILINGS.length} filings in ${FAMILIES.length} families`);
  } catch (e) {
    diag("Boot failed", false, e.message || String(e));
    fatal("Something went wrong signing in", (e.errorCode ? e.errorCode + " — " : "") + (e.message || String(e)) + authHint(e));
  }
}

/* --- friendly hints for the common failure modes --- */
function scopesFrom(jwt) { try { const p = JSON.parse(atob(jwt.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))); return "scp: " + (p.scp || p.roles || "?"); } catch (_) { return ""; } }
function authHint(e) {
  const s = (e.errorCode || e.message || "").toLowerCase();
  if (s.includes("aadsts650") || s.includes("consent")) return "  → Admin needs to grant the app's delegated Sites.Selected permission (issue #177).";
  if (s.includes("aadsts500113") || s.includes("redirect")) return "  → The redirect URI in Entra must exactly match " + CFG.redirectUri + " under the 'Single-page application' platform.";
  return "";
}
function siteHint(e) { return e.status === 403 ? "app not yet granted read on this site — IT must POST /sites/{id}/permissions with grantedTo.application (issue #177)" : (e.message || ""); }
function listHint(e) {
  if (e.status === 403) return "The app has the site but not list-item read. Under Sites.Selected this usually means the site grant is missing or needs the granular Lists.SelectedOperations.Selected scope + list /permissions grant.";
  if (e.status === 404) return "List not found or no permission stamped — check the site grant.";
  return e.message || "";
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("btnSignin")?.addEventListener("click", signIn);
  document.getElementById("btnSignout")?.addEventListener("click", signOut);
  document.getElementById("diagToggle")?.addEventListener("click", () => document.getElementById("diagPanel").classList.toggle("open"));
  if (!CFG.diagnostics) { const d = document.getElementById("diagPanel"); if (d) d.style.display = "none"; }
  boot();
});
