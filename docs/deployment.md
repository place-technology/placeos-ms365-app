# Deploying and updating PlaceOS for Outlook and Teams

PlaceOS ships as **one web app** hosted on each customer's PlaceOS domain, plus **two small packages** that tell Microsoft 365 where to find it:

| Piece | What it is | Where it lives | Who deploys it |
|-------|-----------|----------------|----------------|
| **Web app** | All code, UI, styling and add-in icons (`taskpane.html`, `app.html`, `auth.html`, `dialog.html`, JS/CSS, `assets/`) | `https://<domain>/outlook-addin/` | PlaceOS (us) |
| **Outlook add-in manifest** | `manifest.xml`: task pane in mail and calendar | Customer's Microsoft 365 tenant | Customer M365 admin |
| **App bar package** | `placeos-app-<domain>.zip`: personal tab in the Teams, Outlook and Microsoft 365 app bar (manifest + 2 icons) | Customer's Microsoft 365 tenant | Customer M365 admin |

The packages contain no code. They only point at `https://<domain>/outlook-addin/`. **Updating the web app updates every user on that domain**; customers redeploy a package only when the package itself changes (see [When customers need to redeploy](#when-customers-need-to-redeploy)).

---

## One-time setup per customer

1. **Entra and PlaceOS config:** follow [`entra-setup.md`](entra-setup.md). That covers the app registration, SPA redirect URIs for `https://<domain>/outlook-addin/`, admin consent, `outlook_addin` in the domain's authority config, and optional room booking settings (`outlook_app` zone metadata).
2. **Host the web app** on the domain (see [Deploy the web app](#deploy-the-web-app)).
3. **Build the customer's packages** (see [Build the per-customer packages](#build-the-per-customer-packages)) and send both to the customer's admin.
4. **The customer admin deploys them** (see [Customer admin: deploy the packages](#customer-admin-deploy-the-packages)).

---

## Deploy the web app

PlaceOS serves the web app straight from this repository, like its other user interfaces:

1. **Every push to `main`** runs the **Build** workflow (`.github/workflows/build.yml`). It lints, type-checks, runs `npm run build` and commits the contents of `dist/` (without `manifest.xml`) to the **`build/prod`** branch. It can also be run by hand from **Actions → Build → Run workflow**.
2. **Once per domain**, in PlaceOS Backoffice → **Repositories → Add**: type **Interface**, the URI of this repository, branch **`build/prod`**, and folder name **`outlook-addin`**. The folder name is the path, so the app is served at `https://<domain>/outlook-addin/` (for example `https://<domain>/outlook-addin/taskpane.html`). The repository is private, so enter a GitHub username and a token that can read it.
3. **To release**, merge to `main`, wait for the Build workflow, then pull the latest commit of the repository in Backoffice (or let PlaceOS pick it up).

* The same build works on every domain. The web app reads its per-domain settings at runtime from `https://<domain>/auth/authority` (`outlook_addin`), and calls `auth.cr` and the Staff API on the same domain.
* `build/prod` is generated: don't edit or merge it. Each build is one commit whose message names the `main` commit it came from.
* Customer manifests are built separately (below); `build/prod` has no `manifest.xml`.
* Production builds don't include the development tooling (raw token viewer, hardcoded test config).

To host it somewhere else, run `npm ci && npm run build` and serve the contents of `dist/` (minus `manifest.xml`) at `https://<domain>/outlook-addin/`.

---

## Build the per-customer packages

Office and Teams manifests need absolute URLs, so each customer domain needs its own manifest and zip. Open question Q8 in `plan.md` tracks serving these from each domain automatically.

### With GitHub Actions (recommended)

1. In the repository, go to **Actions → Package for customer → Run workflow**.
2. Enter the customer's **domain**, e.g. `acme.placeos.com`, with no `https://` and no path. Keep the **path** at `/outlook-addin/` unless the customer hosts it elsewhere.
3. The workflow lints, type-checks and builds, checks that the packages point at that domain (not localhost), and uploads two artefacts:

| Artefact | Contents | Give to |
|----------|----------|---------|
| `placeos-outlook-<domain>` | `placeos-outlook-<domain>.xml`: Outlook add-in manifest | Customer M365 admin |
| `placeos-app-<domain>` | `placeos-app-<domain>.zip`: Teams / app bar package | Customer M365 admin |

GitHub always downloads artefacts as a zip. For `placeos-app-<domain>`, extract the download and upload the **inner** `placeos-app-<domain>.zip` to Teams or the admin center; uploading the outer zip fails. The run summary lists the add-in URL and where each artefact goes. Artefacts expire after the repository's retention period (90 days by default), so keep a copy of what you send each customer.

### By hand

```bash
# Outlook add-in manifest -> dist/manifest.xml
ADDIN_URL=https://<domain>/outlook-addin/ npm run build
cp dist/manifest.xml placeos-outlook-<domain>.xml

# App bar package -> dist-app/placeos-app-<domain>.zip
ADDIN_URL=https://<domain>/outlook-addin/ npm run package:app
```

`ADDIN_URL` must start with `https://` and end with `/outlook-addin/`. It only changes the manifests; the web app comes from `build/prod`.

---

## Customer admin: deploy the packages

Both packages are deployed in the **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**, assigned to users or groups.

| Package | Upload as | Appears |
|---------|-----------|---------|
| `placeos-outlook-<domain>.xml` | Office add-in, manifest file | PlaceOS button when reading or writing mail and in calendar events (in Outlook on the web under **Apps**). Pinnable. |
| `placeos-app-<domain>.zip` | Teams app package (also via **Teams admin center → Manage apps → Upload**) | PlaceOS in the left app bar of Teams, Outlook and the Microsoft 365 Copilot app (under **More apps** until pinned) |

* It can take up to 24 hours for deployed apps to appear for all users. Sideloading for testing is faster: see [`app-bar-app.md`](app-bar-app.md), and https://aka.ms/olksideload for the add-in.
* To pin PlaceOS to the Teams/Outlook app bar for everyone, use a Teams app setup policy.

---

## Updating

### Code and UI changes (most releases)

Sign-in, PlaceOS/Staff API calls, UI, styling, theming and the **add-in icons** (served from `/outlook-addin/assets/`) are all in the web app.

1. Merge to `main`. The Build workflow publishes it to `build/prod`.
2. Pull the repository's latest commit in Backoffice on each domain.

**Customers don't need to do anything.** Users get the new version the next time the add-in or tab loads. Outlook caches icons, so add-in icon changes can take a while to show.

### When customers need to redeploy

Only when a **package** changes:

| Change | Package | Bump |
|--------|---------|------|
| Outlook surfaces, pinning, permissions, add-in name/description, icon URLs | `manifest.xml` | `<Version>` in `manifest.xml` (e.g. `1.0.1.0`) |
| App bar tab URL, name/description, permissions, valid domains, new tabs | `app-package/manifest.json` | `"version"` (e.g. `1.0.2`) |
| **App bar icons** (`color.png` / `outline.png` are embedded in the zip) | `app-package/` | `"version"` |
| Customer domain or `/outlook-addin/` path | both | both |

Then rebuild that customer's package(s) with their `ADDIN_URL` and send them to the admin. The admin uploads the new version over the existing app in **Integrated apps** (or the Teams admin center for the zip), and users keep their existing install.

Entra or PlaceOS config changes, such as a new redirect URI, scope or `outlook_addin` value, need no package change. Make them in Entra or Backoffice.

---

## Release checklist

1. `npm run lint` and `npx tsc --noEmit`
2. `npm run validate` if `manifest.xml` changed
3. If a package changed, bump its version (table above)
4. Merge to `main`, check the **Build** workflow passed, and pull the latest `build/prod` commit in Backoffice on each domain
5. If a package changed, run **Package for customer** for each affected domain (or build by hand with `ADDIN_URL`), send the packages to customer admins, and note which version each customer is on
6. Smoke test on one domain: open the add-in and the app bar app. The **Today** view should load your rooms, desks, parking and visitors with no clicks. In **Diagnostics**, expect "Config from the PlaceOS domain", "Auth path: silent" and "✓ PlaceOS token accepted".
