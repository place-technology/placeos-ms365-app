# PlaceOS for Outlook and Teams (SSO POC)

Proof of concept for opening PlaceOS from Outlook and Teams **without a separate login**. The user's Microsoft 365 sign-in is reused through MSAL.js nested app authentication (NAA). `auth.cr` exchanges the Entra token for a PlaceOS token, which is then used with the Staff API to show rooms and desk availability.

It ships as two surfaces built from one codebase:

| Surface | Where | Package | Entry point |
|---------|-------|---------|-------------|
| **Outlook add-in** | Task pane when reading or writing mail, and in calendar events (pinnable) | `manifest.xml` | `src/taskpane/taskpane.ts` → `taskpane.html` |
| **App bar app** | Left app bar in Teams, Outlook and the Microsoft 365 Copilot app | `app-package/` → zip | `src/taskpane/app.ts` → `app.html` |

Built from Microsoft's [Outlook add-in SSO with NAA sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA) (MIT). The plan, decisions and test results are in [`plan.md`](plan.md).

## Documentation

| Doc | For |
|-----|-----|
| This README | Building and running locally |
| [`docs/entra-setup.md`](docs/entra-setup.md) | Configuring a customer's Entra app registration and PlaceOS domain |
| [`docs/deployment.md`](docs/deployment.md) | Hosting, per-customer packages, updates, release checklist |
| [`docs/app-bar-app.md`](docs/app-bar-app.md) | The Teams / app bar personal tab |
| [`plan.md`](plan.md) | POC phases, decisions, recorded values, open questions |

---

## Prerequisites

* **Node.js** 20 or later, and npm
* A **Microsoft 365 account** in a tenant that's been set up per [`docs/entra-setup.md`](docs/entra-setup.md), including the localhost redirect URIs and admin consent
* A **PlaceOS domain** with `outlook_addin` in its authority config, and an `auth.cr` that supports Entra token exchange. Local development uses `placeos-dev.aca.im`.
* To test the app bar app: permission to **upload custom apps** in Teams

## Install

```bash
npm install
```

The first time you run the dev server, `office-addin-dev-certs` creates and trusts a localhost certificate. Accept the prompt; your macOS password or keychain access may be required.

## Run in development

```bash
npm run dev-server
```

This serves the app on **https://localhost:3000** with live reload.

* Requests to `/auth/…` and `/api/…` are **proxied to `https://placeos-dev.aca.im`** (`placeosDevDomain` in `webpack.config.js`), so the app talks to PlaceOS same-origin, just as in production. Change that constant to develop against another PlaceOS domain.
* The add-in reads its Entra and PlaceOS IDs from the domain's `/auth/authority`. If the domain has no `outlook_addin` config, development builds fall back to built-in test values and say so in red in the UI.
* Development builds include a hidden **Show raw token (dev only)** button. Production builds drop it.

> **Restart the dev server after changing `webpack.config.js`.** Live reload picks up source changes, but not config changes such as new entries, the proxy or build flags.
>
> ```bash
> lsof -ti tcp:3000 | xargs kill    # stop whatever is on port 3000
> npm run dev-server
> ```

### Open the Outlook add-in

`npm run start` starts the dev server and sideloads the add-in automatically, but **only on Windows**. On macOS it fails with "Sideload to the Outlook app is not supported", so sideload by hand:

1. Check `https://localhost:3000/taskpane.html` opens in your browser without a certificate warning.
2. Go to **https://aka.ms/olksideload** (opens Outlook on the web), then **My add-ins → Custom Addins → Add a custom add-in → Add from File**, and choose `manifest.xml`.
3. Hard-refresh Outlook. Open an email, start a new one, or open a calendar event, and choose **PlaceOS** from **Apps** (or the **…** menu).
4. To keep it on the toolbar, go to **Settings → Mail → Customize actions**. Use the pin icon in the task pane to keep it open while you switch emails.

The sideload applies to the whole mailbox, so the add-in also shows up in Outlook for Mac and Windows after a short delay. **Remove and re-add it whenever `manifest.xml` changes.** Code changes only need the pane reopened.

### Open the app bar app (Teams / Outlook)

```bash
npm run package:app:dev      # -> dist-app/placeos-app-localhost_3000.zip
```

In Teams, go to **Apps → Manage your apps → Upload an app → Upload a custom app** and choose the zip. Open **PlaceOS** from the Teams app bar, or in Outlook from **More apps** in the left bar. Re-upload only when `app-package/` changes, and bump its `version` when you do. See [`docs/app-bar-app.md`](docs/app-bar-app.md).

### What a working setup shows

PlaceOS opens on the **Today** view: a greeting; in the mail view (reading or writing an email) and in the Teams / app bar app, your **next meeting** (any calendar event, with a Join link for online meetings); then your **rooms** (meetings in PlaceOS rooms), **desks**, **parking** and **visitors** for today, loaded automatically after silent sign-in. Choose **Diagnostics** in the header (or open the page with `#diagnostics`) for the connection checks below.

* *Config from the PlaceOS domain (/auth/authority)*
* *✓ Authenticated*, *Auth path: silent*, and *✓ iss, aud, tid, scp and exp match the PlaceOS registration*
* Then automatically: *✓ PlaceOS token accepted*, the PlaceOS user, then Staff API, buildings, rooms and desks. **Get PlaceOS Token** re-runs the exchange by hand.
* **Refresh PlaceOS Token** appears after the first exchange

## Build

```bash
npm run build                                   # production build to dist/
npm run build:dev                               # development build to dist/
ADDIN_URL=https://<domain>/outlook-addin/ npm run build        # manifest for a customer domain
ADDIN_URL=https://<domain>/outlook-addin/ npm run package:app  # app bar zip for a customer domain
```

Production builds rewrite `https://localhost:3000/` in `manifest.xml` to `ADDIN_URL` (default `https://placeos-dev.aca.im/outlook-addin/`).

For customers, use the **Package for customer** GitHub Actions workflow (`.github/workflows/package-customer.yml`). Enter the domain and it produces the manifest, the app bar zip and the web app as artefacts. Deploying is covered in [`docs/deployment.md`](docs/deployment.md).

## Checks

```bash
npm run lint           # ESLint (office-addins rules) + Prettier; lint:fix / prettier to fix
npx tsc --noEmit       # type-check (the build uses Babel, so type errors don't fail it)
npm run validate       # validate manifest.xml
```

There's no automated test suite.

## Project layout

```text
manifest.xml                 Outlook add-in manifest (add-in only XML)
app-package/                 App bar manifest template + icons (zipped by scripts/package-app.js)
assets/                      Add-in icons
src/taskpane/
  taskpane.ts                Outlook add-in entry (Office.js, theme, Office dialog fallback)
  app.ts                     App bar entry (TeamsJS, theme)
  app-ui.ts                  Shared UI: Today/Diagnostics views, sign-in, token exchange/refresh
  today-data.ts              Today overview: events + desk/parking/visitor bookings → card items
  today-view.ts              Renders the Today cards (each loads and fails independently)
  authConfig.ts              AccountManager: NAA silent → popup → (add-in only) Office dialog
  addin-config.ts            Reads outlook_addin from /auth/authority
  msalconfig.ts              MSAL config factory (memory-only token cache)
  placeos-helper.ts          auth.cr token exchange/refresh, placeosFetch
  placeos-data.ts            Diagnostics: Staff API / PlaceOS read-only checks
  token-inspector.ts         Decodes and checks Entra token claims for display
  theme.ts                   Light / dark / high-contrast from the host
  dev-raw-token.ts           Dev-only raw token viewer
  dev-failure-tests.ts       Dev-only Phase 4 failure-test panel
  fallback/                  auth.html (MSAL redirect target), dialog.html (Office dialog sign-in)
  taskpane.html / .css       Shared page template (app.html is built from it without office.js)
docs/                        Setup and deployment guides
```

## Troubleshooting (local)

| Problem | Fix |
|---------|-----|
| Task pane blank or "does nothing" | Open `https://localhost:3000/taskpane.html` in the same browser to check the certificate is trusted (Firefox doesn't use the macOS keychain). On Mac desktop Outlook, run `npx office-addin-dev-certs install` and restart Outlook. Hard-refresh Outlook on the web. |
| `__DEV_TOOLS__ is not defined` or a new page 404s | The dev server predates a `webpack.config.js` change. Restart it. |
| Port 3000 already in use | `lsof -ti tcp:3000 \| xargs kill` |
| Sign-in shows a consent prompt or `AADSTS…` errors | See the troubleshooting table in [`docs/entra-setup.md`](docs/entra-setup.md) |
| Exchange fails with `invalid_grant` | Check the `auth.cr` logs for the reason. The common ones are in [`docs/entra-setup.md`](docs/entra-setup.md#troubleshooting). |
| Add-in changes in `manifest.xml` not showing | Remove and re-add the custom add-in, then hard-refresh Outlook |
