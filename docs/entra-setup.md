# Microsoft Entra setup for the PlaceOS Outlook add-in

This guide configures a customer's Microsoft Entra tenant and PlaceOS domain so that users signed into Outlook can open the PlaceOS add-in **without a separate login**. The add-in uses nested app authentication (NAA) to get an Entra access token silently. PlaceOS (`auth.cr`) exchanges that token for a PlaceOS token, which is then used with the Staff API.

All steps are done in the [Microsoft Entra admin center](https://entra.microsoft.com) and in PlaceOS Backoffice. No new app registration is created. For hosting the web app and building the customer packages, see [`deployment.md`](deployment.md).

| You need | Why |
|----------|-----|
| A Cloud Application Administrator, Application Administrator or Global Administrator in the customer tenant | To edit the app registration and grant admin consent |
| PlaceOS Backoffice admin access to the customer domain | To find the right registration and add the add-in config |
| The customer's PlaceOS domain, e.g. `acme.placeos.com` | The add-in is served from `https://<domain>/outlook-addin/` |

Throughout this guide, `<domain>` is the PlaceOS domain, `<client-id>` is the Entra Application (client) ID, and `<tenant-id>` is the Entra Directory (tenant) ID.

---

## 1. Find the right app registration

The add-in must use the **same Entra app registration that PlaceOS uses for Microsoft sign-in on this domain**. `auth.cr` only accepts tokens issued for an app that is configured as an OAuth authentication source on the domain. A token for any other app is refused, even one from the same tenant.

1. In Backoffice, open **Domains → `<domain>` → Authentication**.
2. Find the Microsoft OAuth source whose **Authorize URL** / **Token URL** contains the customer's tenant ID:
   `https://login.microsoftonline.com/<tenant-id>/oauth2/v2.0/token`
   * Sources using `common`, `organizations` or `consumers` instead of a tenant ID **cannot** be used. The token exchange needs a single-tenant source.
   * A domain can have several OAuth sources. Use the one for this customer's tenant.
3. Note its **Client ID**. This is `<client-id>`.
4. In the Entra admin center, open **Applications → App registrations → All applications** and search for `<client-id>`. This is the registration to configure.

> **Pitfall:** don't use a registration just because it already has an `access_as_user` scope. Older Outlook add-ins often have one with an `api://<domain>/<client-id>` URI. If that app isn't the domain's OAuth source, the token exchange fails with `no oauth strat on this authority matches the token audience`.

On the registration's **Overview**, confirm:

* **Application (client) ID** = `<client-id>`
* **Directory (tenant) ID** = `<tenant-id>`
* **Supported account types** = *My organization only* (single tenant). Don't change this.

---

## 2. Record the current state

All changes below are **additive**. The existing PlaceOS sign-in must keep working. Before changing anything:

* **Authentication:** screenshot the existing **Web** platform redirect URIs.
* **Expose an API:** note whether an Application ID URI and scopes already exist.
* **Manifest:** note the value of `api.requestedAccessTokenVersion` (`null` means v1).

---

## 3. Expose an API scope

The add-in requests a token **for PlaceOS**, not for Microsoft Graph, so the registration must expose a delegated scope.

Open **Expose an API**.

**If there's no Application ID URI yet:**

1. Next to **Application ID URI**, click **Add**, keep the default `api://<client-id>`, and **Save**.
2. Click **Add a scope**:
   | Field | Value |
   |-------|-------|
   | Scope name | `access_as_user` |
   | Who can consent? | Admins and users |
   | Admin consent display name | Access PlaceOS as the signed-in user |
   | Admin consent description | Allows the PlaceOS Outlook add-in to access PlaceOS on behalf of the signed-in user. |
   | State | Enabled |
3. The full scope is `api://<client-id>/access_as_user`.

**If an Application ID URI or scopes already exist:** leave them as they are. Something else, such as an older add-in, may depend on them. Add an `access_as_user` scope if there isn't one. `auth.cr` accepts tokens whose audience is `<client-id>`, `api://<client-id>` or `api://<domain>/<client-id>`, so an existing URI in one of those forms works.

Leave **Authorized client applications** unchanged. NAA doesn't use them.

---

## 4. Add the single-page application redirect URIs

Open **Authentication → Add a platform → Single-page application** and add:

| Redirect URI | Purpose |
|--------------|---------|
| `brk-multihub://<domain>` | Required for NAA. Host only, no path. |
| `https://<domain>/outlook-addin/auth.html` | Popup sign-in fallback |
| `https://<domain>/outlook-addin/dialog.html` | Office dialog fallback for Outlook clients without NAA |

For **local development** against this registration, also add:

| Redirect URI |
|--------------|
| `brk-multihub://localhost:3000` |
| `https://localhost:3000/auth.html` |
| `https://localhost:3000/dialog.html` |

* Leave **Access tokens** and **ID tokens** (implicit grant) **unticked**. MSAL uses the authorization code flow with PKCE.
* Don't change the existing **Web** platform. Entra doesn't allow the same URI on both Web and SPA, but none of these URIs should already exist.

---

## 5. Grant admin consent

Without admin consent, users see a consent prompt the first time, and the sign-in isn't silent.

1. Open **API permissions → Add a permission**.
2. Choose **APIs my organization uses** and search for `<client-id>` or the app's name.
   > The scope won't appear under **My APIs** unless you're listed as an **owner** of the registration; being an admin isn't enough. Use **APIs my organization uses**, or add yourself under **Owners** first.
3. Select **Delegated permissions → `access_as_user`** and click **Add permissions**.
4. **Review the whole permission list**, then click **Grant admin consent for `<tenant>`**. The button grants *every* listed permission that isn't consented yet, not just the new one.
5. Every row should show *Granted for `<tenant>`*.

---

## 6. Leave the token version alone

Don't change `requestedAccessTokenVersion` (`accessTokenAcceptedVersion` on the legacy manifest tab). `auth.cr` accepts both v1 tokens (`iss` = `https://sts.windows.net/<tenant-id>/`) and v2 tokens (`iss` = `https://login.microsoftonline.com/<tenant-id>/v2.0`). Changing the version could break other consumers of this registration's scopes.

---

## 7. Configure the PlaceOS domain

The add-in has no IDs built in. On load it reads them from the domain's public authority config (`GET https://<domain>/auth/authority`).

1. In Backoffice, open **Domains → `<domain>` → Applications** and choose the PlaceOS application the add-in will use for the token exchange, e.g. the Workplace app. Note its **Client ID** (a 32-character hex UID). This is `<placeos-client-id>`.
   * It must be an application that **doesn't require a client secret**, because the add-in runs in the browser.
2. Add this key to the domain's **authority config**:

   ```json
   "outlook_addin": {
     "client_id": "<client-id>",
     "tenant_id": "<tenant-id>",
     "scope": "api://<client-id>/access_as_user",
     "placeos_client_id": "<placeos-client-id>"
   }
   ```

   `scope` is the full scope from step 3. If you kept an existing URI, use that one, e.g. `api://<domain>/<client-id>/access_as_user`.

3. Check that it's published:

   ```bash
   curl -s https://<domain>/auth/authority | python3 -c "import sys,json; print(json.load(sys.stdin)['config'].get('outlook_addin'))"
   ```

These values aren't secrets.

### Room booking settings (optional)

**Book a room** works without extra setup, using the legacy add-in's defaults. To change its behaviour, add settings as **zone metadata** in Backoffice (**Zones → zone → Metadata**). This is separate from the authority config and the PlaceOS application above.

* **Metadata name:** `outlook_app`. If a zone has no `outlook_app`, the add-in uses `outlook-addin_app`, then `workplace_app`, so existing Workplace settings apply until you add Outlook-specific ones. The legacy add-in named this after its URL path. This add-in uses fixed names, so the hosting path doesn't matter.
* **Which zones:** building, region or org. For each setting, the first value found wins: building app metadata, then region, then the org's `settings` metadata, then org app metadata. Values aren't merged, so a building's `events` object doesn't need to repeat the org's keys. Each key is read on its own.
* **Format:** for example

  ```json
  {
    "events": {
      "allow_assets": true,
      "catering_enabled": true,
      "bookable_hours": { "start": 7, "end": 19 },
      "min_duration": 30,
      "max_duration": 480
    }
  }
  ```

  Recurrence is offered by default. Set `events.allow_recurrence` to `false` to turn it off (the legacy add-in only offered it when set to `true`).

  All-day bookings are off by default. Set `events.allow_all_day` to `true` to add an **All day** checkbox to the form. A whole day is booked as a calendar all-day event. To book set hours instead, set `events.all_day_period`, for example `{ "start": 8, "end": 18 }` (decimal hours, local time); that's booked as a timed event marked `custom_all_day`. An all-day booking for today starts now. All-day bookings ignore `bookable_hours`, `min_duration`, `max_duration`, setup and breakdown.

  Other keys read: `events.has_catering`, `events.catering_notes_required`, `events.use_bookings`, `events.hide_notes`, `events.allowed_future_days`, `events.setup`, `events.breakdown`, `events.force_host`, `events.room_as_host`, `events.use_building_timezone`, `events.allow_recurring_instance_clashes`, `currency`, `catering_provider` (only offer that caterer's menu items).

The building can also have these other metadata entries, shared with Workplace: `room_booking_rules` (hidden or restricted rooms), `catering-settings` (`charge_codes`, `require_notes`), `catering_config` and `assets_config` (lead-time rules), and `assets-settings` (`disabled_rooms`). Catering menus and bookable assets are Assets on the building (catering uses category `_CATERING_` and types `CATERING:<caterer>`).

---

## 8. Verify

1. **Existing sign-in still works:** in a private browser window, sign in to `https://<domain>` with Microsoft as usual.
2. **The add-in signs in silently:** open the add-in in Outlook (any email or calendar event). It opens on the **Today** view, which should load your day with no sign-in prompt. Choose **Diagnostics** in the header; it should show:
   * *Config: from the PlaceOS domain (/auth/authority)*
   * *✓ Authenticated* and *Auth path: silent*. `popup` or `dialog` means a fallback was used, which doesn't count as silent.
   * *Entra token claims*: *✓ iss, aud, tid, scp and exp match the PlaceOS registration*
3. **The token exchange works:** right after silent sign-in, with no clicks, expect *✓ PlaceOS token accepted*, your PlaceOS user with *Email matches the Microsoft account*, and the PlaceOS data section (Staff API, buildings, rooms, desks).

Silent sign-in needs all of the following: admin consent (step 5), an Outlook client that supports NAA (new Outlook for Windows, classic Outlook for Windows on a recent build, Outlook on the web, Outlook for Mac), and a user signed into Outlook with an account in this tenant.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| Consent prompt on first use | Admin consent not granted | Step 5 |
| `AADSTS50011` redirect URI mismatch | SPA redirect URIs missing, wrong domain, or added under Web instead of SPA | Step 4 |
| `AADSTS700046` Invalid Reply Address, *Reply Address must have scheme `brk-<GUID>://`* | `brk-multihub://<domain>` missing from the SPA platform of the registration whose ID is `client_id`, or it has a scheme, path or trailing slash, or the host doesn't match the domain serving the add-in | Step 4 |
| `AADSTS65001` consent required | Admin consent not granted for `access_as_user` | Step 5 |
| `AADSTS500011` resource principal not found | `scope` in `outlook_addin` doesn't match the Application ID URI | Steps 3 and 7 |
| *This PlaceOS domain is not configured for the Outlook add-in* | `outlook_addin` missing from the authority config | Step 7 |
| `invalid_grant`, auth.cr log: `no oauth strat on this authority matches the token audience` | `client_id` isn't the Entra client ID of an OAuth source on this domain | Step 1 |
| `invalid_grant`, auth.cr log: `matching oauth strat is not a single-tenant Entra strat` | The OAuth source uses `common` / `organizations` URLs | Use a source with the tenant ID in its URLs (step 1) |
| `invalid_grant`, auth.cr log mentions `iss` or `tid` | Tenant in the OAuth source URLs doesn't match the token | Check `tenant_id` and the source URLs |
| `invalid_client` / `unauthorized_client` | `placeos_client_id` is wrong, or the application requires a secret | Step 7 |
| `unsupported_grant_type` | The domain runs an `auth.cr` version without Entra token exchange | Deploy `auth.cr` with token exchange (PlaceOS/auth.cr `56d8f5be` or later) |
| Staff API: 401 *domain does not match token's* | The add-in isn't served from the same domain as PlaceOS | Serve it from `https://<domain>/outlook-addin/` |

`auth.cr` returns only generic error codes to the add-in. The specific reason is in the `auth.cr` logs (`action: "token_exchange"`, `reason: …`).
