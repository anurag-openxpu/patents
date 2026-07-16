/* Copyright (c) 2026 OXMIQ
 * OXMIQ Patent Portfolio Hub — hosted dashboard logic.
 * M365 sign-in (MSAL, delegated) -> Microsoft Graph reads of the OXMIQ-IPP site's
 * Portfolio & Ideas lists -> render the v3 UI. No data is embedded; everything is
 * fetched live as the signed-in user (SharePoint permissions enforce visibility).
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
let PF = [];        // portfolio filings  {d,t,s,ft,fd,ta,app,inv,sum,rel,pub}
let IDEAS = [];     // ideas list         {t,k,st,lead,ta,mine,stage,sum}
let MINE = [];      // my in-flight submissions
let ME = null;      // {displayName, mail}
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
    + `?$expand=fields&$select=id,createdBy&$top=200`;
  return graphGetAll(path, token);
}

/* ------------------------------------------------------------- field mapping */
const d10 = s => (s ? String(s).slice(0, 10) : "");
function mapPortfolio(items) {
  return items.map(it => {
    const f = it.fields || {};
    return {
      d: f.Docket || f.Title || "—", t: f.Title || "Untitled", s: f.Status || "Pending",
      ft: f.FilingType || "", fd: d10(f.FilingDate), ta: f.TechArea || "",
      app: f.ApplicationNo || "", inv: f.Inventors || "", sum: f.AISummary || "",
      rel: f.RelatedDockets || "", cpc: f.CPC || "",
      pub: f.PublishToPortfolio === true || f.PublishToPortfolio === 1,
    };
  }).sort((a, b) => (a.d || "").localeCompare(b.d || ""));
}
function mapIdeas(items) {
  const meMail = (ME && (ME.mail || ME.userPrincipalName) || "").toLowerCase();
  return items.map(it => {
    const f = it.fields || {};
    const author = (it.createdBy?.user?.email || it.createdBy?.user?.displayName || "").toLowerCase();
    const mine = !!meMail && (author === meMail || (f.Submitter || "").toLowerCase().includes((ME.displayName || "").toLowerCase()));
    return {
      t: f.Title || "Untitled", k: f.ItemKind || "", st: f.Stage || "", lead: f.Lead || f.Submitter || "",
      ta: f.TechArea || "", ideaId: f.IdeaId || "", docket: f.Docket || "", sum: f.NovelIdea || f.AIFeedback || "",
      mine, stage: STAGE_INDEX[f.Stage] || 0, fd: d10(f.SubmittedDate) || d10(f.Created),
    };
  });
}
const STAGE_INDEX = { "Submitted": 1, "Committee vet": 2, "Counsel engaged": 3, "Drafting": 4, "Filed": 5 };

/* -------------------------------------------------------------------- render
   (adapted from storyboard v3 — same DOM, data now comes from Graph)          */
const STATUS = { "Office action": { c: "var(--oa)", b: "b-oa" }, "Pending": { c: "var(--pending)", b: "b-pending" },
  "Expired": { c: "var(--expired)", b: "b-expired" }, "Granted": { c: "var(--ok)", b: "b-granted" },
  "Provisional": { c: "var(--prov)", b: "b-prov" }, "Draft": { c: "var(--draft)", b: "b-draft" },
  "Committee vet": { c: "var(--prov)", b: "b-prov" }, "Drafting": { c: "var(--draft)", b: "b-draft" },
  "Counsel engaged": { c: "var(--prov)", b: "b-prov" }, "Submitted": { c: "var(--pending)", b: "b-pending" },
  "Filed": { c: "var(--ok)", b: "b-granted" } };
const AREA = { "Compute / Cores": "#5b8cff", "Software / Runtime": "#7c5cff", "Packaging": "#43b581",
  "Memory / Fabric": "#e0a13a", "Datacenter / System": "#f0616d" };
const esc = escp;
const pill = s => { const m = STATUS[s] || STATUS.Pending; return `<span class="badge ${m.b}"><span class="dot" style="background:${m.c}"></span>${esc(s)}</span>`; };
const inv1 = s => (s || "").replace(" (first named)", "");
const publishedPF = () => PF.filter(p => p.pub || role !== "employee");

function tally(arr, key) { const o = {}; arr.forEach(p => { const k = p[key] || "—"; o[k] = (o[k] || 0) + 1; }); return o; }
function bars(el, counts, colorFn) {
  if (!el) return;
  const max = Math.max(1, ...Object.values(counts));
  el.innerHTML = Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([k, v]) =>
    `<div class="barrow"><span class="k"><span class="swatch" style="background:${colorFn(k)}"></span>${esc(k)}</span><div class="bar"><i style="width:${Math.round(v / max * 100)}%;background:${colorFn(k)}"></i></div><span class="v tabnum">${v}</span></div>`).join("");
}
function renderDashboard() {
  bars(document.getElementById("areaBars"), tally(PF, "ta"), k => AREA[k] || "#5b8cff");
  const counts = tally(PF, "s"), seg = Object.entries(counts), tot = PF.length;
  let acc = 0;
  const parts = seg.map(([k, v]) => { const a = acc / (tot || 1) * 360, b = (acc + v) / (tot || 1) * 360; acc += v; return `${(STATUS[k] || STATUS.Pending).c} ${a}deg ${b}deg`; });
  const donut = document.getElementById("donut");
  if (donut) donut.innerHTML = `<div style="width:120px;height:120px;border-radius:50%;background:conic-gradient(${parts.join(",")});display:grid;place-items:center"><div style="width:74px;height:74px;border-radius:50%;background:var(--panel);display:grid;place-items:center"><div style="text-align:center"><div style="font-size:20px;font-weight:750">${tot}</div><div class="dim" style="font-size:10px">filings</div></div></div></div>`;
  const leg = document.getElementById("statusLegend");
  if (leg) leg.innerHTML = seg.map(([k, v]) => `<div><span class="swatch" style="background:${(STATUS[k] || STATUS.Pending).c}"></span>${esc(k)} · ${v}</div>`).join("");
  const rec = document.getElementById("recent");
  if (rec) rec.innerHTML = [...PF].sort((a, b) => (b.fd || "").localeCompare(a.fd || "")).slice(0, 5).map(p =>
    `<div class="ticket"><b>${esc(p.d)}</b> — ${esc(p.t)} <span class="who">${esc(p.ta || "")} · ${esc(p.fd || "")}</span></div>`).join("");
  const oa = document.getElementById("oaTable");
  if (oa) oa.innerHTML = PF.filter(p => p.s === "Office action").map(p =>
    `<tr><td class="docket">${esc(p.d)}</td><td>${esc(p.t)}</td><td class="muted">${esc(p.app)}</td><td>Counsel</td></tr>`).join("") || `<tr><td colspan="4" class="dim">No open office actions.</td></tr>`;
  // top-line KPI counts
  setText("kpiTotal", PF.length);
  setText("kpiOA", PF.filter(p => p.s === "Office action").length);
  setText("kpiPending", PF.filter(p => p.s === "Pending").length);
  setText("kpiExpired", PF.filter(p => p.s === "Expired").length);
}
function setText(id, v) { const el = document.getElementById(id); if (el) el.textContent = v; }

/* explorer */
let facet = "all";
function setFacet(f, el) { facet = f; document.querySelectorAll("#facets .facet").forEach(x => x.classList.remove("on")); el.classList.add("on"); renderGallery(); }
function renderGallery() {
  const mineOn = document.getElementById("mineChk") && document.getElementById("mineChk").checked;
  const mt = document.getElementById("mineToggle"); if (mt) mt.classList.toggle("on", mineOn);
  const q = (document.getElementById("q").value || "").toLowerCase();
  const src = mineOn ? MINE : publishedPF();
  const rows = src.filter(p => { const hay = [p.d, p.t, p.s, p.ta, p.app, p.inv].join(" ").toLowerCase(); return hay.includes(q) && (facet === "all" || p.s === facet || p.ta === facet); });
  document.getElementById("gallery").innerHTML = rows.map(p =>
    `<div class="pcard" data-d="${esc(p.d)}" onclick="openPat('${esc(p.d)}')"><span class="docket">${esc(p.d)}</span>
      <div class="body"><div class="ttl">${esc(p.t)}${p.mine ? ' <span class="badge b-mine" style="margin-left:4px">mine</span>' : ''}</div>
      <div class="meta">${pill(p.s)}<span>${esc(p.ta || "")}</span><span>🧾 ${esc(p.app || "—")}</span></div></div></div>`).join("")
    || `<div class="rescount">${mineOn ? "You have no in-flight ideas yet — Submit one." : "No filings match."}</div>`;
  document.getElementById("resCount").textContent = `${rows.length} of ${src.length}${mineOn ? " · your ideas" : " filings"}${facet !== "all" ? " · " + facet : ""}`;
  if (rows.length) openPat(rows[0].d);
  else document.getElementById("detail").innerHTML = '<div class="dim" style="padding:30px;text-align:center">Nothing selected</div>';
}
function openPat(d) {
  const mineOn = document.getElementById("mineChk") && document.getElementById("mineChk").checked;
  const p = (mineOn ? MINE : publishedPF()).find(x => x.d === d) || PF.find(x => x.d === d) || MINE.find(x => x.d === d);
  if (!p) return;
  document.querySelectorAll("#gallery .pcard").forEach(c => c.classList.toggle("sel", c.dataset.d === d));
  const rel = (p.rel || "").split(";").map(s => s.trim()).filter(Boolean);
  const mineTrack = p.mine ? `<div class="lbl">Progress</div><div class="flowbar">${["Submitted", "Committee vet", "Counsel engaged", "Drafting", "Filed"].map((s, i) => { const st = i + 1 < p.stage ? "done" : i + 1 === p.stage ? "active" : ""; return `<span class="step ${st}"><span class="c">${i + 1 < p.stage ? '✓' : i + 1}</span><span class="t">${s}</span></span>${i < 4 ? '<span class="ln"></span>' : ''}`; }).join("")}</div>` : "";
  document.getElementById("detail").innerHTML = `
    <div class="dk">OXMIQ.${esc(p.d)} · ${esc(p.ft || "")}</div>
    <h2>${esc(p.t)}</h2>${pill(p.s)}
    ${mineTrack}
    <div class="abs">${esc(p.sum || "Summary pending (nightly agent).")}${p.sum ? "…" : ""}</div>
    <div class="metagrid">
      <div class="metabox"><div class="ml">Application #</div><div class="mv">${esc(p.app || "—")}</div></div>
      <div class="metabox"><div class="ml">Filing date</div><div class="mv">${esc(p.fd || "—")}</div></div>
      <div class="metabox"><div class="ml">Technology area</div><div class="mv">${esc(p.ta || "—")}</div></div>
      <div class="metabox"><div class="ml">Inventors</div><div class="mv">${esc(inv1(p.inv) || "—")}</div></div>
    </div>
    <div class="lbl">Family / related</div>
    <div class="chips">${rel.length ? rel.map(r => `<span class="chip" onclick="jumpDocket('${esc(r)}')">📄 ${esc(r)}</span>`).join("") : '<span class="chip" style="cursor:default">None linked</span>'}</div>
    <div class="lbl">Status timeline</div>
    <div class="timeline">
      <div class="tl"><b>Filed</b><div class="dt">${esc(p.fd || "—")}</div></div>
      <div class="tl ${p.s === 'Office action' ? '' : 'd'}">${p.s === 'Office action' ? '<b>Office action — reply due</b>' : 'Under examination'}<div class="dt">USPTO</div></div>
      <div class="tl d">${p.s === 'Expired' ? '<b>Provisional expired</b> (converted via children)' : 'Grant (target)'}<div class="dt">stealth until grant</div></div>
    </div>`;
}
function jumpDocket(d) { if (!PF.find(x => x.d === d)) return; const mc = document.getElementById("mineChk"); if (mc) mc.checked = false; renderGallery(); openPat(d); }

/* harvest */
function renderHarvest() {
  const pots = IDEAS.filter(i => i.k === "Potential");
  const cols = [["Captured", pots.filter(i => i.st === "Captured")], ["Counsel talking", pots.filter(i => i.st === "Counsel talking")],
    ["Ready to submit", pots.filter(i => i.st === "Ready to submit")], ["Future queue / parked", pots.filter(i => i.st === "Future queue" || i.st === "Not pursuing")]];
  // fall back to slicing if statuses aren't set on potentials
  if (cols.every(([, it]) => it.length === 0) && pots.length) { cols[0][1] = pots.slice(0, 6); cols[1][1] = pots.slice(6, 8); }
  const el = document.getElementById("kanban");
  if (el) el.innerHTML = cols.map(([label, items]) =>
    `<div class="col"><h4>${label} <span>${items.length}</span></h4>${items.map(i => `<div class="kt">${esc(i.t)}<div class="who">📂 ${esc(i.ta || "—")}${i.lead ? " · " + esc(i.lead) : ""}</div></div>`).join("") || '<div class="who" style="padding:6px;color:var(--dim)">—</div>'}</div>`).join("");
}

/* role + nav (manual preview in v1; real data gating is M365 permissions) */
const access = { employee: ["dash", "explorer", "submit", "harvest"],
  committee: ["dash", "explorer", "submit", "harvest", "review", "health", "drive"],
  exec: ["dash", "explorer", "submit", "harvest", "review", "health", "engage", "vault"],
  counsel: ["dash", "explorer", "drive"] };
let role = "employee", current = "dash";
function setRole(r) {
  role = r; document.getElementById("rolewrap").setAttribute("data-role", r);
  document.querySelectorAll("#roleSwitch button").forEach(b => b.classList.toggle("active", b.dataset.role === r));
  document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("hidden", !access[r].includes(a.dataset.page)));
  if (!access[r].includes(current)) go(access[r][0]);
  renderGallery();  // publish filter depends on role
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
  renderHarvest();
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

    // Portfolio
    let pfItems = [];
    try { pfItems = await listItems(CFG.portfolioList, token); PF = mapPortfolio(pfItems); diag("Read Portfolio list", true, `${PF.length} items`); }
    catch (e) { diag("Read Portfolio list", false, `${e.code} — ${listHint(e)}`); fatal("Couldn't read the Portfolio list", `${e.message}. ${listHint(e)}`); return; }

    // Ideas
    try { const ideaItems = await listItems(CFG.ideasList, token); IDEAS = mapIdeas(ideaItems); MINE = IDEAS.filter(i => i.mine && i.k === "Submission"); diag("Read Ideas list", true, `${IDEAS.length} items · ${MINE.length} yours`); }
    catch (e) { IDEAS = []; MINE = []; diag("Read Ideas list", false, e.code || e.message); }

    show("app");
    setRole("employee");
    renderAll();
    diag("Rendered", true, `${PF.length} filings, ${IDEAS.length} ideas`);
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
