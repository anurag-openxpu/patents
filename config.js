/* Copyright (c) 2026 OXMIQ
 * OXMIQ Patent Portfolio Hub — runtime config.
 *
 * THIS FILE IS WORLD-READABLE. It ships to the browser verbatim, so treat every
 * value in it as published.
 *
 * What belongs here: identifiers the page needs before anyone has signed in —
 * tenant/client/site IDs, list names, URLs. Those are PUBLIC by design (a browser
 * SPA cannot hide them; they appear in the sign-in request itself) and they are not
 * credentials: there is no client secret (MSAL uses PKCE), and knowing a site ID
 * grants nothing, because every read carries the signed-in user's own token and
 * Graph returns only what that account may see.
 *
 * What must NEVER be here: company data of any kind — budget figures, spend, firm
 * or vendor names, docket titles, people. Not because it would break access
 * control, but because it would be published outright, with no sign-in in front of
 * it. Anything like that belongs in a permissioned SharePoint list the page reads
 * AFTER sign-in, so M365 decides who sees it. (Two such values were removed from
 * this file for exactly that reason: the annual counsel budget and the counsel firm
 * name — both now read at runtime from Budget / Roster.)
 */
window.IPP_CONFIG = {
  // --- Entra app (registered by IT) ---
  tenantId: "0184cb4b-6696-4b38-8323-9f5cdeb5babc",          // OXMIQ tenant (verified)
  clientId: "fe1ffd79-a7e4-4029-acda-460b7fe38709",          // the registered SPA app
  redirectUri: "https://anurag-openxpu.github.io/patents/",  // must match Entra "SPA" redirect URI

  // --- Graph scopes requested at sign-in ---
  // Sites.Selected is the delegated model (app granted read on ONE site; intersected
  // with the user's own permission). If list-item reads 403 with only Sites.Selected,
  // add the granular scopes below to this list and ask IT to approve them and
  // list /permissions — the diagnostics panel says exactly which call failed.
  scopes: ["User.Read", "Sites.Selected"],
  // Optional granular fallbacks (uncomment if item reads 403 under Sites.Selected alone):
  // "Lists.SelectedOperations.Selected", "ListItems.SelectedOperations.Selected"

  graphBase: "https://graph.microsoft.com/v1.0",

  // --- the one site + lists ---
  siteHostname: "netorgft13672293.sharepoint.com",
  sitePath: "sites/OXMIQ-IPP",
  siteId: "netorgft13672293.sharepoint.com,1ab4a0ea-cbab-49ff-b754-770fb900844f,91248c16-84b4-40e8-9218-269d28fb3adb",
  // Single source of truth: the Ledger List (Portfolio + Ideas collapsed into it).
  // Portfolio is a *view* of this List (Stage in {Filed, Granted, Published} & PublishToPortfolio).
  ledgerList: "Ledger",
  // Spend view reads invoice metadata off the Legal-Finance library (one source).
  legalFinanceLibrary: "Legal-Finance",
  // Budget list holds the annual target (one row per year, exec-editable in
  // SharePoint — no redeploy). There is deliberately NO hardcoded fallback: the
  // figure is company financial data, so if the list is unreadable the page hides
  // the target rather than publishing a number here.
  budgetList: "Budget",

  // Counsel scope: the "Counsel" view shows the filings whose AssignedCounsel
  // matches the signed-in counsel's firm. Filled at runtime from that person's
  // Roster row (Roster.CounselFirm) — intentionally EMPTY here, both because the
  // firm we retain is not public and because an empty value fails closed: a
  // counsel whose Roster row carries no firm matches no docket at all.
  counselFirm: "",

  // The "Submit an Idea" button opens this SharePoint list form (new tab). The
  // submission flow (Power Automate) turns each new item into a Ledger row +
  // disclosure folder + inventor access + email. Employees need Contribute on
  // this list at go-live.
  intakeFormUrl: "https://netorgft13672293.sharepoint.com/sites/OXMIQ-IPP/Lists/Idea%20Intake/NewForm.aspx",

  // Show the diagnostics panel (auth + each Graph call). Off for a clean UI;
  // flip to true if you need to debug a sign-in/Graph issue.
  diagnostics: false,
};
