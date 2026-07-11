# RMZ Intranet Suite — Deployment Guide

This guide covers building, packaging, deploying, and configuring the
`rmz-sp-intranet` SPFx solution (global header/footer extension + web parts).
Every solution/component ID in this package was generated fresh for this
deployment, so it installs and runs independently of any other build —
including an earlier client-specific version it may have been copied from —
with no ID conflicts.

---

## 1. Prerequisites

- **Node 18–22**, npm 8+.
- A **SharePoint Online communication site** for the intranet home (full-width
  sections — used by these web parts — exist only on communication sites).
- An **App Catalog** (tenant-level recommended; see §4).
- Roles: **SharePoint/Global Administrator** to approve API permissions; **Site
  Owner** to add the app and run setup on a site.

---

## 2. Install & build

```bash
npm install
npm run build         # gulp build  (compile + lint)
```

## 3. Package

```bash
npm run package       # gulp bundle --ship && gulp package-solution --ship
```

Output: **`sharepoint/solution/rmz-sp-intranet.sppkg`**.
`includeClientSideAssets` is `true`, so JS/CSS/fonts ship inside the package and
are served by SharePoint — no separate CDN required.

> Bump `solution.version` (and the feature `version`) in
> `config/package-solution.json` for every new build so the App Catalog shows the
> update clearly.

---

## 4. Deploy to the App Catalog

The solution is configured **per-site** (`skipFeatureDeployment: false`), so it is
**not** auto-deployed tenant-wide; you add it to each site explicitly.

1. Go to the **tenant App Catalog**: `https://<tenant>.sharepoint.com/sites/appcatalog`
   → **Apps for SharePoint**.
2. Upload `rmz-sp-intranet.sppkg` (replace the previous version when updating).
3. In the trust dialog, **leave "Make this solution available to all sites"
   unchecked**, then **Deploy**.

> Use the **tenant** App Catalog (not a site-collection one) — only the tenant
> catalog registers the Microsoft Graph permission requests (see §6).

---

## 5. Add to a site & run setup

On each target site:

1. **Settings (gear) → Add an app → RMZ Intranet Suite.** This activates the
   solution's feature on the site; the Intranet web parts appear in the page editor
   toolbox under the **Intranet** group.
2. Create/edit a page, add the **Intranet Setup (Admin)** web part, and click
   **Run setup**. It idempotently provisions everything the web parts use,
   with an Arabic column alongside each English one so the same list items
   serve both languages:
   - **Events** list (EventDate, EndDate, Location/**LocationAR**, Category, TitleAR, Description/**DescriptionAR**)
   - **Policies** list (Department, PolicyVersion, FileType, **TitleAR**)
   - **News** list (Category, NewsDate, Source/**SourceAR**, TitleAR, Body/**BodyAR**)
   - **HR Benefits** list (Category, Summary/**SummaryAR**, Eligibility/**EligibilityAR**, Coverage/**CoverageAR**, TitleAR, Details/**DetailsAR**)
   - **Templates** document library
   - **Category** column on **Site Pages** (for the News category filter)
   - **Intranet Settings** list (central config — hidden from Site Contents)
   - **Intranet Navigation** list (top-nav links, seeded with English **and** Arabic labels; hidden from Site Contents)
   - **View All pages** (News, Events, Directory, Policies) with the matching web
     part pre-placed and its "View All" link switched off (toggle on by default)
   - **Registers the global header & footer** on the site (toggle on by default)
   - Optionally **seeds sample data** (toggle off by default)
   Re-running is safe — existing items are detected and left untouched.
3. **Reload** the page — the header (sand top bar, logo, nav, language toggle,
   avatar) and dark footer should appear.

### 5b. Configure from the Admin screen (Site Owners only)

The **Intranet Setup (Admin)** web part is now a full control panel. It is gated
to users with **Manage Web** permission (Site Owners); everyone else sees an
"Administrators only" notice. Its tabs:

- **Setup** — the provisioning action above.
- **Navigation** — add/reorder/remove the top-nav links (label, URL, open in new
  tab). Saved to the **Intranet Navigation** list; the header reads it live.
- **General** — show/hide footer, footer copyright & tagline (English + Arabic),
  search placeholder (English + Arabic), default language for first-time
  visitors, client name (English + Arabic), hide native site header / site nav /
  command bar, and hide the URL in embedded portal frames.
- **Web parts** — show/hide the "View All" links globally, accent colour, and
  default item counts for News, Events and the Directory.

Navigation links, Quick Links and Departments each have an Arabic label
(and, for Departments, an Arabic description) field alongside the English one.

Settings are stored in the **Intranet Settings** list and consumed live by
the header/footer extension and the web parts (cached ~2 minutes, so edits appear
on the next reload). The **View All** link on a section automatically hides when
the page you are on *is* its destination page, so "View All" pages never show a
redundant button.

Each web part also has its own **Show section title** / **Show "View All" link**
toggles in its property pane, for per-instance overrides. The central accent
colour is injected as the `--intranet-accent` CSS variable and themes section
underlines, chips and accents across the web parts.

### 5a. Header/footer registration without the Admin web part (alternative)

If you prefer PnP PowerShell instead of the Admin web part:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive

# Check whether it is already registered
Get-PnPCustomAction -Scope Web | Select-Object Name, Location, ClientSideComponentId

# Register the header/footer application customizer
Add-PnPCustomAction -Name "IntranetHeaderFooter" -Title "IntranetHeaderFooter" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId "1083fb8f-96fc-49dc-9284-2d77402b58a9" `
  -ClientSideComponentProperties '{"navItems":[{"label":"Home","url":"/"},{"label":"Departments","url":"/SitePages/Departments.aspx"},{"label":"Policies","url":"/SitePages/Policies.aspx"},{"label":"Documents","url":"/SitePages/Documents.aspx"},{"label":"News","url":"/SitePages/News.aspx"},{"label":"Events","url":"/SitePages/Events.aspx"},{"label":"Onboarding","url":"/SitePages/Onboarding.aspx"},{"label":"Approvals","url":"/SitePages/Approvals.aspx"}],"searchPlaceholder":"Search intranet…","copyright":"© 2026 Intranet. All rights reserved.","tagline":"Reimagining Time"}' `
  -Scope Web
```

To change nav links/footer text later without redeploying, update the
`ClientSideComponentProperties` JSON (re-run with `Set-PnPCustomAction` or remove
and re-add).

---

## 6. Approve Microsoft Graph permissions (for live data)

The web parts work **without** Graph — they fall back to sample data, and the
header avatar falls back to the page-context user. Approve these scopes only when
you want live tenant data:

| Scope | Used by |
|---|---|
| `User.Read` | Hero greeting, header avatar (`/me`) |
| `User.Read.All` | Employee Directory, Org Chart, Celebrations (`/users`, directReports) |
| `People.Read.All` | People data |
| `Calendars.Read` | Productivity Strip — Today's Schedule (`/me/calendarView`) |

**Approve in:** SharePoint admin centre → **Advanced → API access** →
**Pending requests** → approve each. (Requires Global/SharePoint admin.)

PnP PowerShell alternative:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
Get-PnPTenantServicePrincipalPermissionRequests
Approve-PnPTenantServicePrincipalPermissionRequest -RequestId <id>
```

> If no requests appear, the package was likely deployed to a **site-collection**
> App Catalog instead of the **tenant** one (only the tenant catalog registers
> them), or they were already approved (check the **Approved** tab).

---

## 7. Build the pages

In the page editor, add a **full-width section** for each row and drop in the web
parts (all under the **Intranet** toolbox group). The alternating sand backgrounds
are built into the web parts, so no section colour is needed.

**Homepage order:** Hero Greeting → Productivity Strip → Quick Links →
News & Announcements → Company Events → HR Benefits → Embedded Portal →
Policies & Procedures → Shared Documents → Employee Directory.

**Department page order:** Department Hero → Department Quick Actions →
Department Documents → Department News → Department Events → Applications & Approvals.

Each data web part exposes a property pane (list/library names, link URLs, etc.)
so it can be pointed at the right source per environment.

---

## 8. Components reference

| Component | Type | Notes |
|---|---|---|
| IntranetHeaderFooter | Application Customizer | Renders header (Top) + footer (Bottom) on every page; registered per site (§5/§5a) |
| Hero Greeting, Productivity Strip (My Profile card replaced by Useful Links), Quick Links, Embedded Portal, News Carousel, Events, HR Benefits, Policies List, Shared Documents, Employee Directory | Web parts | Homepage. News/Events/Policies items link to their detail page/item. (Celebrations and Org Chart removed in 1.0.9.0) |
| Department Hero, Department Quick Actions, Department Documents, Department News, Department Events, Applications & Approvals | Web parts | Department page (`design/department-page.html`) |
| Intranet Setup (Admin) | Web part | One-click provisioning + header/footer registration |

---

## 9. Troubleshooting

- **Header/footer not showing:** It is an extension, not a web part. (a) View a
  real page, **not** the hosted workbench (`/_layouts/15/workbench.aspx`) —
  extensions never render there. (b) Confirm it is registered: run the Admin web
  part's setup, or `Get-PnPCustomAction -Scope Web`. (c) Hard-refresh (Ctrl+F5).
  (d) Console (F12) logs `Initializing Intranet header/footer` when it loads.
- **Web parts show sample data:** Expected until the lists exist (run setup) and
  Graph permissions are approved (§6).
- **Can't add a full-width section:** The site must be a **communication site**.
- **Updated package not picked up:** Bump the version (§3) and re-deploy in the
  App Catalog.
