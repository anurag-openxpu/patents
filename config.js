/* Copyright (c) 2026 OXMIQ
 * OXMIQ Patent Portfolio Hub — runtime config.
 * These are PUBLIC identifiers (a browser SPA exposes them by design) — no secret here.
 * Data access is enforced by M365 sign-in + the app's Sites.Selected grant on the one site.
 */
window.IPP_CONFIG = {
  // --- Entra app (from IT / issue #177) ---
  tenantId: "0184cb4b-6696-4b38-8323-9f5cdeb5babc",          // OXMIQ tenant (verified)
  clientId: "fe1ffd79-a7e4-4029-acda-460b7fe38709",          // the registered SPA app
  redirectUri: "https://anurag-openxpu.github.io/patents/",  // must match Entra "SPA" redirect URI

  // --- Graph scopes requested at sign-in ---
  // Sites.Selected is the delegated model (app granted read on ONE site; intersected
  // with the user's own permission). If list-item reads 403 with only Sites.Selected,
  // add the granular scopes below to this list AND have IT consent them + stamp the
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
  // SharePoint — no redeploy). annualBudget is only the fallback if unreadable.
  budgetList: "Budget",
  annualBudget: 100000,

  // Counsel scope (PoC): the role-switcher's "Counsel" view shows the filings whose
  // AssignedCounsel = this firm — regardless of the publish flag. One firm today
  // ("Adeli LLP" on all 24 dockets). In the real deployment the signed-in counsel's
  // identity maps to a firm; matching CFG.counselFirm is the stand-in for the demo.
  counselFirm: "Adeli LLP",

  // The "Submit an Idea" button opens this SharePoint list form (new tab). The
  // submission flow (Power Automate) turns each new item into a Ledger row +
  // disclosure folder + inventor access + email. Employees need Contribute on
  // this list at go-live.
  intakeFormUrl: "https://netorgft13672293.sharepoint.com/sites/OXMIQ-IPP/Lists/Idea%20Intake/NewForm.aspx",

  // Show the diagnostics panel (auth + each Graph call). Off for a clean UI;
  // flip to true if you need to debug a sign-in/Graph issue.
  diagnostics: false,
};
