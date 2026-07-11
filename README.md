# RMZ.SP.ORA-UAE.Intranet

SharePoint Online intranet for **ORA UAE**, built with SPFx.

## Tech stack

| | |
|---|---|
| SPFx | 1.22.1 |
| React | 17.0.1 |
| TypeScript | 4.7.4 |
| PnP JS | v4 (`@pnp/sp`, `@pnp/queryable`) |
| Fluent UI React | v8 |
| Gulp | 4.0.2 |
| Node | 18–22 |

Approved designs live in [`design/`](./design) (homepage, department page, departments landing, onboarding).

## Components

### `ORAHeaderFooter` — Application Customizer

Injects the ORA-branded **global header and footer** on every SharePoint page
(`Top` and `Bottom` placeholders).

**Header** (`.hdr` in `design/v1-homepage.html`)
- ORA logo (bundled `ora-logo.png`) + `UAE` wordmark, links to home
- Top navigation — Home, Departments, Policies, Documents, News, Events, Onboarding, Approvals
- Active link highlighted with the sand underline; re-evaluated on client-side navigation
- Rounded search box (`#f3f1ed`); Enter runs a site search
- Notification bell with sun-accent dot indicator
- User avatar (initials) + display name — resolved from **Microsoft Graph** (`/me`, `User.Read`)
- 3px sand gradient top border; header is sticky on scroll

**Footer** (`.ftr` in `design/v1-homepage.html`)
- Dark `#262626` background
- Left: `ORA UAE` wordmark (Raleway 200) · Centre: copyright · Right: `Reimagining Time`

**Configuration** — the nav links and footer text are component properties, so URLs
can be set per environment:

| Property | Type | Default |
|---|---|---|
| `navItems` | `{ label, url }[]` | the 8 links above (server-relative) |
| `searchPlaceholder` | `string` | `Search intranet…` |
| `copyright` | `string` | `© 2026 ORA UAE. All rights reserved.` |
| `tagline` | `string` | `Reimagining Time` |

Defaults for tenant-wide deployment are set in
[`sharepoint/assets/elements.xml`](./sharepoint/assets/elements.xml); per-site
overrides can be applied via PnP PowerShell `Set-PnPApplicationCustomizer`.

> **Fonts:** Raleway + Public Sans are **bundled** in the solution (no CDN
> dependency). The latin variable-font `.woff2` files live in
> `src/extensions/oraHeaderFooter/assets/fonts/` and are embedded as base64 in
> `components/fontData.ts`, injected once per page as `@font-face` rules by
> `components/fonts.ts`. Regenerate `fontData.ts` with `base64 -w0 <file>.woff2`.

### Homepage web parts

Eleven React web parts assemble the approved homepage (`design/v1-homepage.html`),
all under the **ORA UAE** toolbox group and styled from the shared design tokens
(`src/common/styles/_tokens.scss`). Each data-backed web part reads live data and
**falls back to the design sample data** when the source is empty/unavailable, so
everything renders in the workbench before lists exist.

| Web part | Section | Data source |
|---|---|---|
| Hero Greeting | Hero banner + personalised greeting | Graph `/me` |
| Embedded Portal | Full-width embedded iframe (browser-chrome frame) | Configurable URL property |
| Productivity Strip | My Profile · Today's Schedule · Quick Access | Graph `/me`, `/me/calendarView`; quick links configurable |
| Celebrations | Birthdays & work anniversaries | Graph `/users` (birthday); anniversaries from config/fallback |
| News Carousel | News + category filter tabs | PnP v4 — Site Pages library |
| Events | Company events | PnP v4 — Events list |
| HR Benefits | Benefit cards | Configurable content |
| Policies List | Searchable, dept-filtered table | PnP v4 — Policies list |
| Shared Documents | Two document panels | PnP v4 — document libraries |
| Employee Directory | Searchable people grid | Graph `/users` |
| Org Chart (mini) | CEO + direct reports | Graph `/users/{id}/directReports` |

Shared building blocks live in `src/common/`: `components/SectionHeader`,
`components/InitialsAvatar`, `services/pnpService` (`getSP`), `services/graphService`
(`getGraphClient`), `models/`, and `util/format`.

Graph scopes requested (in `package-solution.json`): `User.Read`, `User.Read.All`,
`People.Read.All`, `Calendars.Read`.

### Department page web parts

Six React web parts assemble the department template (`design/department-page.html`),
same shared design system, all full-width and responsive:

| Web part | Section | Data source |
|---|---|---|
| Department Hero | Dept hero (name, description, owner) | Configurable properties |
| Department Quick Actions | Action buttons + dept stats | Configurable properties |
| Department Documents | Policies + Shared Documents panels | PnP v4 — libraries (dept-filtered) |
| Department News | Department news cards | PnP v4 — Site Pages (dept-filtered) |
| Department Events | Department events (3-col) | PnP v4 — Events list (dept-filtered) |
| Applications & Approvals | My requests + awaiting my approval | PnP v4 — Approvals/Requests list |

### Admin

| Web part | Purpose |
|---|---|
| Intranet Setup (Admin) | One-click PnP provisioning of the Events/Policies lists, Templates library, and Site Pages Category column; registers the header/footer on the site; optional sample-data seeding. |

## Develop

```bash
npm install
gulp serve            # set pageUrl in config/serve.json first
gulp build            # compile + lint
gulp bundle --ship    # production bundle
gulp package-solution --ship   # produces sharepoint/solution/intranet-suite.sppkg
```

### Deploy (per-site scope)

The solution is configured for **per-site** deployment (`skipFeatureDeployment: false`),
so the header/footer only appears on sites where the app is explicitly added.

1. Upload `sharepoint/solution/intranet-suite.sppkg` to the tenant App Catalog
   (`/sites/appcatalog` → **Apps for SharePoint**). Do **not** tick "make available
   to all sites".
2. Approve the **Microsoft Graph `User.Read`** permission request in the SharePoint
   admin centre → **Advanced → API access** (Global or SharePoint admin).
3. On each target site: **Settings → Add an app → ORA UAE Intranet**. This activates
   the feature and provisions the `ClientSideComponentInstance` on that site only.

Per-site nav/footer overrides (without redeploying) can be applied with PnP PowerShell
`Set-PnPApplicationCustomizer -Identity <id> -ClientSideComponentProperties '...'`.
