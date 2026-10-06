# PlaceOS in the Teams, Outlook and Microsoft 365 app bar

PlaceOS ships as two pieces that share one codebase, one sign-in flow and one deployment on the PlaceOS domain:

| Piece | Where it appears | Package | Page |
|-------|------------------|---------|------|
| **Outlook add-in** | Task pane beside an email or calendar event (read, compose, pinnable) | `manifest.xml` | `taskpane.html` |
| **App bar app** (personal tab) | Left app bar in **Teams**, **Outlook** (web and desktop) and the **Microsoft 365 Copilot** app, including mobile | `app-package/` → zip | `app.html` |

Both get an Entra token silently with nested app authentication (NAA), exchange it for a PlaceOS token via `auth.cr`, and read their config from the domain's `outlook_addin` authority config. The app bar app runs full-page with no email or event context.

They are separate packages because Outlook add-ins defined in the newer unified manifest don't yet run in Outlook for Mac. Keeping the add-in in `manifest.xml` preserves Mac support.

## Prerequisites

Complete [`entra-setup.md`](entra-setup.md) for the customer domain. No extra Entra configuration is needed: the app bar app uses the same `brk-multihub://<domain>` redirect URI and the same `access_as_user` scope.

## Build the package

The package contains a `manifest.json` (unified manifest for Microsoft 365) and two icons. The script fills in the add-in URL:

```bash
# Local development (dev server on https://localhost:3000)
npm run package:app:dev

# A customer domain
ADDIN_URL=https://<domain>/outlook-addin/ npm run package:app
```

Output: `dist-app/placeos-app-<host>.zip`. The web files (`app.html` etc.) come from the normal `npm run build`, published to `build/prod` and served at `https://<domain>/outlook-addin/` alongside the add-in.

## Install for testing

Your tenant must allow custom app uploads (Teams admin center → **Teams apps → Setup policies → Upload custom apps**).

1. In Teams, go to **Apps → Manage your apps → Upload an app → Upload a custom app** and choose the zip.
2. Open **PlaceOS** from the Teams app bar.
3. In Outlook (web or new Outlook for Windows), open **More apps** in the left app bar and choose **PlaceOS**. Right-click it to pin it to the app bar. It can take a few minutes to appear in Outlook after uploading in Teams.

For local development, keep `npm run dev-server` running. The `localhost` package points at `https://localhost:3000/app.html`.

## Deploy to an organisation

See [`deployment.md`](deployment.md) for per-customer packages, updates and when customers need a new zip. In short:


A customer admin uploads the zip in the **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps** (or the Teams admin center under **Manage apps**), and assigns it to users or groups. To pin it to the app bar for everyone, use a Teams app setup policy.

## What you should see

The **Today** view laid out full-page (rooms, desks, parking and visitors as a grid), with **Diagnostics** in the header showing:

* *Teams app* / *Outlook app* / *Microsoft 365 app*. Config from the PlaceOS domain.
* *✓ Authenticated*, *Auth path: silent*
* The token claims, then automatically *✓ PlaceOS token accepted* and the PlaceOS data (Staff API, buildings, rooms, desks)

If silent sign-in fails, **Sign in** tries a popup. There's no Office dialog fallback outside Outlook add-ins.

## Notes

* `privacyUrl` and `termsOfUseUrl` in `app-package/manifest.json` must point at real PlaceOS pages before publishing.
* Changing `app-package/manifest.json` needs a new `version` and a re-upload.
* Icons come from the PlaceOS building artwork at `https://s3.ap-southeast-2.amazonaws.com/os.place.tech/outlook-plugin-resources/` (`16x16-01.png`, `32x32-01.png` and `80x80-01.png`, which are actually 67, 134 and 334 px). They are resized to exact sizes: `assets/icon-{16,32,64,80,128}.png` for the add-in, and in `app-package/`, `color.png` (192×192) and `outline.png` (32×32, white body on transparent with the windows cut out).
