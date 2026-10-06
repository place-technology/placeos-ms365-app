# PlaceOS Outlook Add-in SSO POC

## Objective

Build a proof of concept for a new PlaceOS Outlook Add-in using Microsoft's `outlook-add-in-sso-naa` sample as the foundation.

The purpose of this POC is to validate seamless authentication between Outlook, Microsoft Entra ID, PlaceOS, and the existing Staff API.

The POC is **not intended to implement room or desk booking yet**.

The primary question we need to answer is:

> Can a user who is already authenticated in Outlook use the PlaceOS add-in without being presented with a separate PlaceOS login screen?

## Target Architecture

```text
Outlook
   │
   │ User already authenticated
   ▼
PlaceOS Outlook Add-in  (served from https://<customer>.placeos.com/outlook-addin/)
   │
   │ NAA / MSAL, using the customer's existing PlaceOS user-access app registration
   ▼
Microsoft Entra ID (customer tenant)
   │
   │ Entra access token (aud = PlaceOS user-access app)
   ▼
PlaceOS auth.cr
   │
   │ Token exchange
   ▼
PlaceOS access token
   │
   ▼
Staff API
   │
   ▼
PlaceOS user / rooms / desks
```

The Staff API already exposes the data and operations required by the eventual Outlook integration, including room and desk functionality.

The Staff API currently expects a PlaceOS token.

Rather than modifying the Staff API to directly authenticate Microsoft Entra tokens, `auth.cr` will provide an exchange mechanism that accepts a valid Entra access token and returns a standard PlaceOS token.

This allows the Staff API to remain unchanged.

---

# Key Decisions

| # | Decision | Rationale |
|---|----------|-----------|
| D1 | **Reuse the existing PlaceOS user-access app registration** in each customer tenant. Add an SPA platform for NAA and expose an API scope on it. Do not create a third registration. | Customers already have two PlaceOS registrations (user access and application access), and a third would add onboarding friction. `auth.cr` already stores this registration's client ID and tenant for each PlaceOS domain, so it can validate `aud` and `tid` against config it already holds. |
| D2 | **Host the add-in on the customer's PlaceOS domain** (e.g. `https://<customer>.placeos.com/outlook-addin/`), with a per-customer manifest. | The add-in can read the domain's client ID, tenant ID and API scope from the same origin before authenticating. Calls to `auth.cr` and the Staff API are same-origin, so no CORS is needed. |
| D3 | **Look up users by `tid` + `oid`**, falling back to email. | PlaceOS already stores the Entra object ID from existing SSO logins. `oid` is stable, while email/UPN can change. Because the registration is single-tenant and `tid` is validated, the email fallback is also limited to the customer's tenant. |
| D4 | **Use a token for the PlaceOS API, not Microsoft Graph.** | Graph access tokens are issued for Graph and are not intended to be validated by third-party services. The token sent to `auth.cr` must have the PlaceOS registration as its audience. |

---

# Phase 1: Microsoft Entra Configuration

> Step-by-step onboarding guide, with lessons from the test tenant: [`docs/entra-setup.md`](docs/entra-setup.md).

Configure the customer's existing PlaceOS **user-access** app registration so it can also serve as the NAA client for the add-in. Use a single test tenant for the POC.

All changes must be additive. The existing Web platform redirect URIs and the PlaceOS OAuth login flow must keep working unchanged.

Changes to the existing user-access registration:

* **Authentication → add Single-page application platform** with redirect URIs:
  * `brk-multihub://<customer>.placeos.com` (required for NAA)
  * `https://<customer>.placeos.com/outlook-addin/auth.html` (popup/redirect fallback)
  * `https://<customer>.placeos.com/outlook-addin/dialog.html` (Office dialog fallback for clients without NAA)
  * For local development: `brk-multihub://localhost:3000`, `https://localhost:3000/auth.html`, `https://localhost:3000/dialog.html`
* **Expose an API**:
  * Set an Application ID URI (e.g. `api://<client-id>`).
  * Add a delegated scope, e.g. `access_as_user`.
* **Manifest**: set `accessTokenAcceptedVersion` to `2` so the issuer (`https://login.microsoftonline.com/<tid>/v2.0`) and `aud` (the client ID) are predictable. Before changing this, confirm nothing currently relies on v1 tokens for this registration. *(Deferred for the test tenant: it is `null` and the scope already exists, see the token version note below.)*
* **Consent**: grant **admin consent** for the new scope in the test tenant. Without it, the first run will show a consent prompt, which breaks the silent-flow success criterion.
* **Supported accounts**: stays single-tenant, as today.

Record the following for the test tenant:

* Tenant ID
* Application / Client ID
* Application ID URI and full scope string (e.g. `api://<client-id>/access_as_user`)
* The PlaceOS domain the registration is linked to in `auth.cr`

## Recorded values (test tenant)

| Item | Value |
|------|-------|
| Tenant ID | `bc9d5ad8-7518-422b-ac8d-b69429ca4cb9` |
| PlaceOS domain in `auth.cr` | `placeos-dev.aca.im` |
| Application / Client ID (user-access) | `fdb3d186-52cb-470d-88d1-eee1b5b9d9d3`. Backoffice OAuth source with token URL `https://login.microsoftonline.com/bc9d5ad8-…/oauth2/v2.0/token` |
| Application ID URI | `api://fdb3d186-52cb-470d-88d1-eee1b5b9d9d3` (added for the POC) |
| Scope | `api://fdb3d186-52cb-470d-88d1-eee1b5b9d9d3/access_as_user` |
| Access token version | v1 (`ver` 1.0 observed). Leave unchanged; auth.cr accepts v1 and v2 |
| SPA redirect URIs | Localhost only for now (local testing) |
| Admin consent for scope | Granted. Added via API permissions → "APIs my organization uses" (the scope doesn't appear under "My APIs" unless you own the registration) |
| PlaceOS application `client_id` for the exchange | `06aa34dd7c8b1cdb6364ac031a1f84a5` (Workplace app) |

The domain has a second OAuth source, `058da72c-cbc2-4c67-bee6-3da22919b9ce`, whose tenant isn't the test tenant.

**First attempt (superseded):** Phase 1 was initially checked against `eb5a5522-e090-497e-a789-c5d3cbfce7ac`. It already had a v1 scope `api://placeos-dev.aca.im/eb5a5522-…/access_as_user`, probably used by the existing Outlook add-in's legacy Office SSO. Silent NAA worked in Outlook on the web with it, but the exchange failed with `no oauth strat on this authority matches the token audience`, because it isn't an OAuth source on the domain. D1's assumption only holds for the registration that `auth.cr` actually has configured for the domain, so check the domain's OAuth sources first when onboarding.

**Token versions:** `auth.cr` loads v1 or v2 discovery based on the token's `iss`, and accepts `aud` as `<client-id>`, `api://<client-id>` or `api://<domain>/<client-id>`. Either version works. The add-in's claim checks pick v1 or v2 expectations from the `ver` claim.

## Success Criteria

* The existing PlaceOS OAuth login still works after the changes.
* The registration has the SPA redirect URIs, the exposed scope, and admin consent in the test tenant. (v2 tokens deferred, see token version note.)

---

# Phase 2: Outlook Add-in Foundation

Use Microsoft's `outlook-add-in-sso-naa` sample as the starting point.

Retain the sample's existing NAA/MSAL implementation initially and make only the changes necessary to establish the PlaceOS POC.

## Configuration discovery

The add-in must not hardcode a client ID, because each customer tenant has its own registration. On load, before calling MSAL, the add-in fetches its auth configuration from the PlaceOS domain it is served from:

```text
GET https://<customer>.placeos.com/<auth-config-endpoint>

{
  "client_id": "<user-access-app-client-id>",
  "tenant_id": "<tenant-id>",
  "scope": "api://<client-id>/access_as_user"
}
```

These values are not secrets. Whether this uses an existing PlaceOS authority endpoint or a new one is an open question (see below).

For the very first spike, the values may be hardcoded in `msalconfig.ts` to unblock testing, but discovery must be in place before Phase 6.

*Implemented (2026-10-01):* the values are stored in the domain's **authority config** under `outlook_addin`, which the public `GET /auth/authority` returns (this answers Q1):

```json
"outlook_addin": {
  "client_id": "<Entra client ID of the domain's PlaceOS OAuth source>",
  "tenant_id": "<tenant-id>",
  "scope": "api://<client-id>/access_as_user",
  "placeos_client_id": "<PlaceOS application UID for the token exchange>"
}
```

The add-in and the fallback dialog read this before starting MSAL. Development builds fall back to the test-tenant values if the key is missing; production builds show "domain not configured" instead. Verified 2026-10-01 on `placeos-dev.aca.im`: the add-in loads its config from the domain in Outlook on the web.

## Changes to the sample

* Replace the sample branding and UI with a minimal PlaceOS interface.
* Set the MSAL `authority` to `https://login.microsoftonline.com/<tenant_id>` rather than `common`.
* Request the PlaceOS API scope instead of Graph scopes.
* Change `defaultScopes` in `msalconfig.ts` to the PlaceOS API scope. The dialog fallback ignores the requested scopes and always uses `defaultScopes`, so otherwise it will silently return a Graph token.
* **Turn off verbose PII logging** in `msalconfig.ts`. It is on by default in the sample and can write tokens and user details to the console.
* Remove the Graph-specific calls (`/me`, OneDrive). Add a small fetch helper for PlaceOS calls; `makeGraphRequest` and its CAE handling are specific to Graph.
* Choose a single manifest format (add-in only `manifest.xml` vs unified `manifest.json`) and remove the other configuration, since the two have already drifted.
* Keep tokens in memory only. Do not write them to `localStorage` or cookies.

Example UI:

```text
PlaceOS

Microsoft authentication
✓ Authenticated

Microsoft account
user@example.com

[ Get PlaceOS Token ]

Status
Waiting...
```

*Surfaces (2026-10-01):* the manifest registers the task pane on reading and composing mail, and on calendar events as organizer and attendee. Outlook has no extension point for the main window in general. *Pinning (2026-10-01):* a VersionOverrides 1.1 block (Mailbox 1.13) adds `SupportsPinning` on read and compose mail, so the pane stays open when switching messages. Pinning isn't supported for calendar items. It also adds `SupportsNoItemContext` on read, so in **classic Outlook for Windows** only, the pane can open with the Reading Pane off or no message selected. Clients without 1.13 use the v1.0 block (no pinning). The task pane never reads the current item, so it needs no `ItemChanged` handling, and tokens and data persist while pinned. Verified 2026-10-01 in Outlook on the web: the pinned pane stays open across messages.

*App bar / Teams (2026-10-01):* the brief requires PlaceOS in Teams too. A separate personal-tab app (`app-package/`, unified manifest) adds PlaceOS to the left app bar in Teams, Outlook and the Microsoft 365 Copilot app. It shares the code and auth flow with the add-in, using TeamsJS instead of Office.js. It's kept separate from `manifest.xml` because unified-manifest Outlook add-ins don't run in Outlook for Mac. See [`docs/app-bar-app.md`](docs/app-bar-app.md). Clients to test: Teams (web, desktop), Outlook app bar (web, new Windows).

No PlaceOS application should be embedded in the add-in.

No booking functionality is required at this stage.

## Local development

Run the webpack dev server on `https://localhost:3000` and proxy the PlaceOS paths (auth config, `auth.cr`, Staff API) to the test PlaceOS domain. This keeps the same-origin behaviour of the production hosting model without adding CORS rules to PlaceOS.

## Success Criteria

* Add-in loads successfully in Outlook.
* Add-in can be opened by the test user.
* Add-in obtains its client ID, tenant and scope from configuration rather than code (or from a clearly marked temporary hardcode for the first spike).

---

# Phase 3: Validate NAA Authentication

Run the add-in inside Outlook and verify the complete Microsoft authentication flow.

Expected flow:

```text
Open Outlook
    ↓
Open PlaceOS add-in
    ↓
Fetch auth config from PlaceOS domain
    ↓
NAA detects existing Microsoft session
    ↓
acquireTokenSilent(api://<client-id>/access_as_user)
    ↓
Entra access token (aud = PlaceOS user-access app)
```

## Definition of "silent"

The normal path must not display a Microsoft or PlaceOS login screen when all of these preconditions hold:

* Admin consent has been granted for the scope (Phase 1).
* The Outlook client supports NAA (`NestedAppAuth` 1.1 requirement set).
* The user is signed into Outlook with an account in the customer's tenant.

The sample falls back to a popup and then to an Office dialog. Fallbacks are acceptable, but they **do not count as passing** the silent criterion. The POC should record which path was used.

Clients to test (record pass/fail and the auth path used for each):

| Client | Expected path |
|--------|---------------|
| New Outlook for Windows | Silent NAA |
| Classic Outlook for Windows (NAA-capable build) | Silent NAA |
| Outlook on the web | Silent NAA |
| Outlook for Mac | Silent NAA |
| Older classic Outlook (no NAA) | Dialog fallback (login expected) |

Results (local dev server, add-in sideloaded via https://aka.ms/olksideload):

| Client | Result | Auth path | Date | Notes |
|--------|--------|-----------|------|-------|
| Outlook on the web | Pass | Silent NAA | 2026-10-01 | Message compose, registration `fdb3d186-…`. v1 token (`ver` 1.0); `iss`, `aud`, `tid`, `scp`, `exp` match. Exchanged for a PlaceOS token, and `/api/engine/v2/users/current` returns the user, email matching the Microsoft account |
| Outlook for Mac | | | | |
| New Outlook for Windows | | | | |
| Classic Outlook for Windows (NAA-capable build) | | | | |
| Older classic Outlook (no NAA) | | | | |

## Token inspection

Capture enough information to verify the token without exposing the complete token. The UI should show only the decoded claims:

* Microsoft user identity (`oid`, `upn` / `unique_name` for v1 or `preferred_username` for v2, `name`)
* Token acquisition success and auth path (silent / popup / dialog)
* Issuer (`iss`)
* Audience (`aud`), which must equal the PlaceOS Application ID URI (v1) or client ID (v2)
* Expiry (`exp`)
* Tenant (`tid`)
* Scope (`scp`)

Do not require the `amr` claim. It is often absent from access tokens.

The complete access token must not be committed to source control, displayed in the application UI, or written to logs.

*Deliberate exception (2026-10-01):* development builds (`npm run start` / `build:dev`) include a hidden-by-default "Show raw token (dev only)" toggle with a Copy button, so the token can be tested against `auth.cr` before the 4a validation endpoint exists. It is gated by the `__DEV_TOOLS__` build flag and absent from production builds. The token is still never logged. Treat copied tokens as credentials: don't paste them into tickets, chat or source control.

## Success Criteria

A user already signed into Outlook on an NAA-capable client can open the PlaceOS add-in and silently obtain an Entra access token whose `aud`, `iss` and `tid` match the PlaceOS user-access registration.

---

# Phase 4: Entra to PlaceOS Token Exchange

Update `auth.cr` to exchange a valid Microsoft Entra access token for a standard PlaceOS token.

## 4a. Validation endpoint (early, dev only)

Instead of handing raw tokens to the backend team (they are bearer credentials and expire within about 60–90 minutes), the backend first provides a dev-only endpoint. It validates the token and returns the result without issuing a PlaceOS token. The add-in calls it with a freshly acquired token.

This lets the backend iterate on validation logic against real tokens from the add-in, without anyone copying tokens around.

## 4b. Exchange endpoint

The exact endpoint and request format will be determined by the backend team. We recommend a standard grant such as RFC 8693 token exchange rather than a custom header format:

```text
POST https://<customer>.placeos.com/auth/oauth/token
Content-Type: application/x-www-form-urlencoded

grant_type=urn:ietf:params:oauth:grant-type:token-exchange
&subject_token=<entra-access-token>
&subject_token_type=urn:ietf:params:oauth:token-type:access_token
```

Response:

```text
{
  "access_token": "<placeos-token>",
  "token_type": "Bearer",
  "expires_in": 3600
}
```

## Implementation (auth.cr, 2026-10-01)

Implemented in PlaceOS/auth.cr commit `56d8f5be` (see README "Microsoft Entra token exchange"). `POST /auth/oauth/token`, RFC 8693, with two extra parameters compared with the sketch above:

* `client_id` = a **PlaceOS application UID** on the domain (not the Entra client ID). The add-in has no secret, so it must be an application that doesn't require one.
* `scope` (optional, defaults to `public`). `audience`/`resource` are rejected.

Response: `access_token` (JWT), `token_type`, `refresh_token`, `issued_token_type`. The add-in keeps the refresh token in memory for refreshing (see Token lifetime).

Differences from this plan, to agree with the backend:

* **No-match users are auto-created**, as in SSO login (lookup by `oid`, then email). The plan's failure case "valid token but no matching PlaceOS user" will not be rejected.
* **`scp` is checked for presence (delegated token), not for `access_as_user` specifically.**
* `aud` may be the client ID, `api://<client_id>` or `api://<authority host>/<client_id>`, so v1 tokens are accepted.
* All validation failures return a generic `400 invalid_grant`. The reason is only in the auth.cr logs, so the add-in can't show a specific message per failure case.
* `ensure_matching` restrictions from the strat are applied, and a Graph token is fetched on-behalf-of if none is stored (best effort).

## Results

| Date | Client | Result |
|------|--------|--------|
| 2026-10-01 | Outlook on the web | Silent Entra token exchanged for a PlaceOS token via auth.cr on `placeos-dev.aca.im` (PlaceOS app `06aa34dd…`). The PlaceOS token is accepted by `/api/engine/v2/users/current`, which returns the expected user. Success criteria 1–4 met; failure cases partly tested. |
| 2026-10-01 | Outlook on the web | Refreshing the PlaceOS token via `POST /auth/token` works, and the refreshed token is accepted by `/api/engine/v2/users/current`. |

Failure tests run so far:

| Case | Result |
|------|--------|
| Malformed / non-JWT token | Rejected, `400 invalid_grant` |
| Token for a registration that isn't an OAuth source on the domain (`eb5a5522-…`) | Rejected, `400 invalid_grant` (log: "no oauth strat on this authority matches the token audience") |
| Expired token | Pending: held token expires 16:48:50 on 2026-10-01; rerun the dev panel after that |
| Wrong audience (e.g. Graph token) | Rejected, `400 invalid_grant` (Graph `User.Read` token from the same sign-in, dev panel, 2026-10-01) |
| Different tenant | Tampered `tid` rejected, `400 invalid_grant` (dev panel, 2026-10-01). A genuine other-tenant token needs a user in another tenant: auth.cr spec tests |
| Missing scope / app-only token | Not testable from the add-in (needs a client secret): auth.cr spec tests |
| Invalid signature | Rejected, `400 invalid_grant` (real token with altered signature, dev panel, 2026-10-01) |
| Bogus PlaceOS refresh token | Rejected, `400 invalid_grant` (dev panel, 2026-10-01) |
| Control: valid token | Accepted, 200 (dev panel, 2026-10-01) |

Run these from the **Failure tests (dev only)** panel in development builds (appears after sign-in). It includes a control case (valid token, expect 200).

## Validation performed by `auth.cr`

```text
Entra access token
        │
        ▼
     auth.cr
        │
        ├── Resolve PlaceOS domain/authority from request host
        ├── Load that domain's user-access registration (client ID, tenant ID)
        ├── Verify signature against Entra JWKS for the tenant
        ├── iss == https://sts.windows.net/<tenant_id>/   (v1, current)
        │      or https://login.microsoftonline.com/<tenant_id>/v2.0   (v2)
        ├── aud == <application_id_uri>   (v1, current)
        │      or <client_id>   (v2)
        ├── tid == <tenant_id>
        ├── scp contains access_as_user
        ├── exp / nbf valid
        ├── Locate PlaceOS user by tid + oid (fallback: email)
        │
        ▼
PlaceOS access token
```

Because the domain comes from the request host, a token from tenant A can never be exchanged on tenant B's PlaceOS domain.

The resulting token should behave exactly like an existing PlaceOS token when used with the Staff API.

The implementation should avoid requiring the Staff API to understand Microsoft Entra authentication.

## Failure behaviour

Each of the following must be rejected with a clear error, and the add-in must show a useful message:

* Expired token
* Wrong audience (e.g. a Graph token)
* Token from a different tenant than the domain's registration
* Missing scope
* ~~Valid token but no matching PlaceOS user~~. Not a failure case: auth.cr creates the user, as SSO login does (decision 2026-10-01).
* Invalid signature

## Token lifetime

* The add-in keeps the PlaceOS access and refresh tokens in memory only.
* *(Updated 2026-10-01)* When the PlaceOS token is within 60 s of expiry, or a PlaceOS call returns 401, the add-in first refreshes it with the PlaceOS refresh token (`POST /auth/token`, `grant_type=refresh_token`), because that's faster than a new Entra exchange. If the refresh fails, it acquires a fresh Entra token (silently) and exchanges again.

## Success Criteria

The Outlook add-in can:

1. Obtain an Entra token.
2. Exchange the Entra token through `auth.cr`.
3. Receive a standard PlaceOS token.
4. Use that token for authenticated PlaceOS API requests.

Each failure case above is rejected.

---

# Phase 5: Call Staff API

Once the token exchange is working, use the resulting PlaceOS token to call the existing Staff API.

The first request should be a simple authenticated, read-only request.

Preferably this should establish:

* Current PlaceOS user
* Organisation
* Relevant permissions/context

Then perform a read-only request against existing room or desk functionality.

Example flow:

```text
Outlook Add-in
      │
      │ Entra token
      ▼
   auth.cr
      │
      │ PlaceOS token
      ▼
   Staff API
      │
      ▼
Authenticated PlaceOS user
      │
      ├── Rooms
      └── Desks
```

Do not create or modify bookings during the initial authentication test.

## Success Criteria

The add-in can use the exchanged PlaceOS token to successfully call Staff API and retrieve PlaceOS data for the authenticated user, and the user Staff API reports matches the Entra `oid`.

## Implementation notes (2026-10-01)

All calls are read-only GETs through `placeosApi` (cached, refreshed or re-exchanged token, with a retry on 401):

| Check | Endpoint |
|-------|----------|
| Staff API accepts the token | `GET /api/staff/v1/tenants/current_limits` (full Staff API auth path: JWT, scope `public`, domain, tenant; no Graph dependency) |
| Current user | `GET /api/engine/v2/users/current` |
| Organisation | `GET /auth/authority` (`name`, `domain`, `config.org_zone`) |
| Buildings | `GET /api/engine/v2/zones?tags=building&parent_id=<org_zone>` |
| Rooms | `GET /api/engine/v2/systems?zone_id=<building>&bookable=true` |
| Desks | `GET /api/engine/v2/metadata/<building>/children?name=desks` |
| Desks booked today | `GET /api/staff/v1/bookings/booked?type=desk&period_start=&period_end=&zones=<building>` (epoch seconds) |

**`oid` check:** Staff API has no current-user endpoint, and no PlaceOS REST endpoint exposes the stored Entra `oid` (UserAuthLookup). *Decision (2026-10-01, option 1a):* the add-in shows that the PlaceOS user's email matches the Microsoft account. auth.cr will log, for each exchange, how the user was matched (`oid`, email, or created); those logs are the evidence for the `oid` criterion. Creating unknown users automatically (as SSO login does) is accepted, so "valid token but no matching PlaceOS user" is no longer a failure case. Backend action: add the match-method log line.

## Results

| Date | Client | Result |
|------|--------|--------|
| 2026-10-01 | Outlook on the web | Staff API accepts the exchanged PlaceOS token (`tenants/current_limits` ✓). Authority loads (`org_zone` = `zone-DnTc8chjVb`). No buildings are children of `org_zone`, so the add-in falls back to all zones tagged `building`. |
| 2026-10-01 | Outlook on the web | Read-only data loads with the PlaceOS token. Building "PlaceOS Sydney Dev": 5 bookable rooms with capacities; 162 desks on 3 levels, 2 booked today (Staff API `bookings/booked`). Phase 5 met, except the `oid` match (see note above). |

Follow-up for the backend: on `placeos-dev.aca.im` the authority's `config.org_zone` (`zone-DnTc8chjVb`) isn't the parent of the building zones. Check whether that's stale config or a different hierarchy.

---

# Phase 6: End-to-End Validation

The POC is considered successful when the following complete workflow works on an NAA-capable Outlook client:

1. User is already signed into Outlook with their Microsoft 365 account.
2. User opens the PlaceOS Outlook add-in.
3. The add-in loads its auth configuration from the PlaceOS domain.
4. No separate PlaceOS or Microsoft login screen is presented.
5. NAA silently obtains an Entra access token for the PlaceOS API scope.
6. The Entra token is sent to `auth.cr`.
7. `auth.cr` validates the token against the domain's user-access registration.
8. `auth.cr` identifies the corresponding PlaceOS user by `oid` (falling back to email, then creating the user), confirmed via the auth.cr log line for the exchange.
9. `auth.cr` returns a standard PlaceOS token.
10. The add-in uses the PlaceOS token with Staff API.
11. Staff API successfully identifies the user.
12. Staff API returns PlaceOS room/desk data.

The desired user experience is:

```text
Open Outlook
      ↓
Open PlaceOS
      ↓
Immediately authenticated
      ↓
PlaceOS data loads
```

There should be no separate PlaceOS login step in the normal path.

*Implemented (2026-10-01):* after silent sign-in the add-in and app bar app exchange the token and load PlaceOS data automatically. The buttons are only for manual re-runs.

*Beyond the POC (2026-10-01):* a user-facing **Today** landing page is now the default view. It shows the user's rooms (calendar events in PlaceOS rooms), desks, parking and visitors for today, using the same calls as Workplace's schedule page. The original POC page remains as **Diagnostics** (header link, or `#diagnostics`). All read-only; no booking actions yet.

The POC report should include the client test matrix from Phase 3 and the results of the Phase 4 failure tests.

---

# Open Questions

| # | Question | Owner |
|---|----------|-------|
| Q1 | *(Answered: `outlook_addin` in the authority config, via `/auth/authority`.)* Does an existing PlaceOS endpoint expose the domain's Entra client ID and tenant ID publicly (e.g. the authority endpoint), or is a new endpoint needed? | Backend |
| Q2 | *(Answered: `https://<customer-domain>/outlook-addin/`. Production builds take `ADDIN_URL`.)* What path on the customer PlaceOS domain will host the add-in, and does it conflict with the existing Outlook add-in? | Backend / DevOps |
| Q3 | *(Resolved: token version left unchanged; the test registration issues v1 tokens and auth.cr accepts v1 and v2, so no change is needed. See Phase 1 token version note.)* Is the existing user-access registration using v1 tokens anywhere that `accessTokenAcceptedVersion: 2` could affect? | Backend |
| Q4 | *(Partly answered: RFC 8693 at `POST /auth/oauth/token`, returning access and refresh tokens; refresh works via `POST /auth/token`. Still open: is the token the same type and lifetime as one from normal login?)* Is the PlaceOS token returned by the exchange the same type and lifetime as one issued by the normal login flow? Which grant/endpoint will `auth.cr` use? | Backend |
| Q5 | *(Resolved for the POC: auth.cr stores the Entra `oid` in UserAuthLookup and looks users up by `oid`, then email, then creates them. No API exposes it, so the evidence is the auth.cr log line (option 1a, Phase 5). How often the email fallback is used will show in those logs.)* Which `oid` field in PlaceOS holds the Entra object ID, and is it populated for all users, or only those who have logged in via SSO? This determines how often the email fallback is used. | Backend |
| Q6 | *(Answered for now: manual onboarding guide in `docs/entra-setup.md`; packaging automated via the "Package for customer" GitHub workflow. Automating the Entra changes themselves stays out of scope.)* Can customer onboarding docs and scripts be updated to add the SPA platform and exposed scope to the user-access registration, or does this need to be automated? | Product |
| Q7 | *(Answered: tenant `bc9d5ad8-7518-422b-ac8d-b69429ca4cb9`, registration `fdb3d186-…`, domain `placeos-dev.aca.im`; admin consent granted. Test user `placeosazuresandbox2@0cbfs.onmicrosoft.com`.)* Which test tenant and test users will be used, and who can grant admin consent there? | Product / Backend |
| Q8 | **Per-customer manifest (deferred).** Office manifests need absolute URLs, so today each customer needs a manifest built with their domain (`ADDIN_URL=https://<customer>/outlook-addin/ npm run build`). This is the manual approach for now. Proposed later: ship one generic build with a manifest template (`__ADDIN_HOST__` placeholder) and have each domain serve `https://<domain>/outlook-addin/manifest.xml` with its host filled in (e.g. nginx `sub_filter … $host`, or a small backend endpoint). Customer admins would then deploy by URL in the M365 admin center (Integrated apps), with no file editing. The same applies to the **app bar zip** (`app-package/`), whose tab URL and `validDomains` contain the domain (`ADDIN_URL=… npm run package:app` per customer). It could also be served per domain, e.g. `/outlook-addin/placeos-app.zip`, generated at deploy time. Customers only need a new manifest or zip when the manifest or embedded icons change. Code and UI updates deploy by updating the files at `/outlook-addin/`. Needs: what serves files on the PlaceOS domain (nginx / other)? | Backend / DevOps |

---

# Out of Scope

The following are intentionally excluded from this POC:

* Room booking UI
* Desk booking UI
* Creating bookings
* Cancelling bookings
* Outlook calendar integration
* Reading appointment details
* Microsoft Graph calendar operations
* PlaceOS application embedding
* Mobile optimisation
* Production deployment and multi-customer rollout
* Automated changes to customer app registrations
* Replacing the existing Outlook add-in

These should only be addressed after the authentication and Staff API integration have been proven.

---

# Future Integration

Once the POC is successful, the next phase can build the actual PlaceOS Outlook experience.

The eventual architecture will be:

```text
                         Outlook
                            │
                ┌───────────┴───────────┐
                │                       │
           Mail / Outlook          Calendar
                │                       │
                └───────────┬───────────┘
                            │
                   PlaceOS Add-in
                            │
                    NAA / MSAL SSO
                            │
                            ▼
                     Microsoft Entra
                            │
                            ▼
                         auth.cr
                            │
                     PlaceOS token
                            │
                            ▼
                       Staff API
                     /      |       \
                 Rooms    Desks   Bookings
```

The eventual add-in can then use Outlook context to make booking rooms and desks considerably simpler.

For example, when creating a meeting:

```text
Meeting
Thursday 10:00–11:00
8 attendees
```

the PlaceOS add-in could automatically use that date, time, attendee count, and potentially location to retrieve available PlaceOS rooms.

The same add-in could provide independent desk booking functionality when the user is not working with a meeting.

The existing PlaceOS application would not need to be loaded into the Outlook window.
