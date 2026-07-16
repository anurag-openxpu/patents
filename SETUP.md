<!-- Copyright (c) 2026 OXMIQ -->
# OXMIQ Patent Portfolio Hub — setup & access

Live page: **https://anurag-openxpu.github.io/patents/**
(internal-audience tool; add `noindex` is already set. Reach it from office / VPN.)

This is a **static** page (HTML/JS on GitHub Pages). It holds **no patent data**.
At sign-in it calls Microsoft Graph **as the signed-in user** and reads two lists
on the `OXMIQ-IPP` SharePoint site. What each person sees is whatever M365 already
lets them see — we enforce nothing in code. Writes are **not** done here (they stay
in Power Automate); this page is **read-only**.

- Tenant: `0184cb4b-6696-4b38-8323-9f5cdeb5babc`
- App (client) ID: `fe1ffd79-a7e4-4029-acda-460b7fe38709`
- Site: `netorgft13672293.sharepoint.com/sites/OXMIQ-IPP`
- Lists read: `Portfolio`, `Ideas`

---

## For IT (Teja) — three one-time steps

### 1. Redirect URI (Entra → the app → Authentication)
Add a **Single-page application (SPA)** platform redirect URI — *not* "Web":

```
https://anurag-openxpu.github.io/patents/
```

(SPA is required so MSAL can use PKCE with no client secret and get CORS headers.)

### 2. Delegated permission + consent (Entra → the app → API permissions)
Add **Microsoft Graph → Delegated → `Sites.Selected`** (keep the default `User.Read`),
then **Grant admin consent** for the tenant. Delegated `Sites.Selected` = the app's
effective access is the **intersection** of the site grant in step 3 *and each user's
own SharePoint permission* — so nobody sees more than SharePoint already gives them.

### 3. Grant the app **read** on the OXMIQ-IPP site *only*
One Graph call (needs an admin token with `Sites.FullControl.All`, e.g. Graph Explorer):

```http
POST https://graph.microsoft.com/v1.0/sites/netorgft13672293.sharepoint.com,1ab4a0ea-cbab-49ff-b754-770fb900844f,91248c16-84b4-40e8-9218-269d28fb3adb/permissions
Content-Type: application/json

{
  "roles": ["read"],
  "grantedToIdentities": [
    { "application": { "id": "fe1ffd79-a7e4-4029-acda-460b7fe38709",
                       "displayName": "OXMIQ IPP Dashboard" } }
  ]
}
```

`read` is enough — the page only reads. Scope is this one site; no other site is touched.

> If list-item reads still 403 after this (some tenants require it), also add delegated
> `Lists.SelectedOperations.Selected` + `ListItems.SelectedOperations.Selected` in step 2
> and re-consent. The page's **Diagnostics** panel says exactly which call failed.

---

## For the owner — test it
1. Open **https://anurag-openxpu.github.io/patents/** (office or VPN).
2. Click **Sign in with Microsoft**, complete sign-in.
3. Expect: the four KPIs fill in, the Explorer gallery shows the 24 dockets, the
   role switcher and tabs work.
4. If anything is blank or errors: click the **Diagnostics** bar at the bottom — it
   lists each step (sign-in → token → site → Portfolio → Ideas) with the exact error.
   Send me that and I'll pinpoint which of the three steps above is still pending.

## What "public repo" means here
The GitHub repo is public so any OXMIQ employee can reach the page and the code — but
the **code carries no data**. Every row is fetched live through the reader's own M365
session; someone with no OXMIQ account (or no access to the site) signs in and sees
nothing. That is the whole point: access lives in M365, not in this page.
