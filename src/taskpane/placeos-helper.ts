/* This file provides helpers for calling PlaceOS (auth.cr and Staff API). */

/* global fetch, Headers, RequestInit, Response, URLSearchParams */

import { decodeJwtPayload } from "./token-inspector";

/**
 * Fetch options for PlaceOS calls. `allPages` follows a paginated list to the end; `text` returns the
 * response body as text instead of parsing it as JSON (e.g. SVG maps).
 */
export type PlaceosRequestInit = RequestInit & { allPages?: boolean; text?: boolean };

export class PlaceosRequestError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly body: string
  ) {
    super(message);
    this.name = "PlaceosRequestError";
  }
}

/**
 * Makes a request to PlaceOS on the same origin the add-in is served from.
 * In local development the webpack dev server proxies these paths to the test PlaceOS domain.
 * @param path Absolute path, e.g. "/api/staff/v1/...".
 * @param accessToken Bearer token to send, if any.
 * @param init Fetch options. A string body is sent as JSON. With `allPages`, a list response is
 * followed through its `Link: <…>; rel="next"` headers and every page is returned as one array.
 * @returns The parsed JSON response.
 */
export async function placeosFetch<T>(
  path: string,
  accessToken?: string,
  init: PlaceosRequestInit = {}
): Promise<T> {
  const { allPages, text, ...fetchInit } = init;
  if (!allPages || text) {
    return (await placeosRequest(path, accessToken, fetchInit, text)).data as T;
  }
  const items: unknown[] = [];
  const seen = new Set<string>();
  let next: string | null = path;
  while (next && !seen.has(next)) {
    seen.add(next);
    const page = await placeosRequest(next, accessToken, fetchInit);
    if (!Array.isArray(page.data)) {
      // Not a plain list (e.g. { results }); leave it to the caller.
      return page.data as T;
    }
    items.push(...page.data);
    next = nextPageLink(page.response.headers.get("Link"));
  }
  return items as T;
}

/** The rel="next" URL from a Link header, as engine list routes send (PlaceOS/rest-api `paginate_sql`). */
function nextPageLink(link: string | null): string | null {
  const match = link?.match(/<([^>]+)>\s*;\s*rel="?next"?/);
  return match ? match[1] : null;
}

async function placeosRequest(
  path: string,
  accessToken: string | undefined,
  init: RequestInit,
  asText = false
): Promise<{ data: unknown; response: Response }> {
  const headers = new Headers(init.headers);
  if (!headers.has("Accept")) {
    headers.set("Accept", "application/json");
  }
  if (typeof init.body === "string" && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  if (accessToken) {
    headers.set("Authorization", `Bearer ${accessToken}`);
  }

  const response: Response = await fetch(path, { ...init, headers });
  if (!response.ok) {
    const body = await response.text();
    throw new PlaceosRequestError(
      `PlaceOS request to ${path} failed with ${response.status}`,
      response.status,
      body
    );
  }
  // Writes can return 202/204 with no body.
  const text = await response.text();
  return { data: asText ? text : text ? JSON.parse(text) : null, response };
}

export type PlaceosToken = {
  accessToken: string;
  // Kept in memory only, like the access token.
  refreshToken?: string;
  // Milliseconds since epoch, if known.
  expiresAt?: number;
};

type TokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

// auth.cr only returns error codes; the detailed reason is in its logs.
const exchangeErrorMessages: Record<string, string> = {
  invalid_grant:
    "PlaceOS rejected the Microsoft token (invalid_grant): bad signature, issuer, audience, tenant, expiry, missing claims or ensure_matching. See the auth.cr logs for the reason.",
  invalid_client:
    "PlaceOS did not accept the application client_id (invalid_client). Check placeos_client_id in the outlook_addin config and that the application doesn't require a secret.",
  unauthorized_client:
    "This PlaceOS application may not use token exchange (unauthorized_client). Check placeos_client_id in the outlook_addin config.",
  invalid_request: "PlaceOS reported a malformed exchange request (invalid_request).",
  invalid_target: "PlaceOS rejected the audience/resource parameter (invalid_target).",
  unsupported_grant_type:
    "This PlaceOS domain doesn't support token exchange (unsupported_grant_type). Is the auth.cr version deployed?",
};

const refreshErrorMessages: Record<string, string> = {
  invalid_grant:
    "PlaceOS rejected the refresh token (invalid_grant): it is invalid, expired or revoked.",
  invalid_client: exchangeErrorMessages.invalid_client,
};

async function postTokenRequest(
  path: string,
  params: Record<string, string>,
  errorMessages: Record<string, string>,
  previousRefreshToken?: string
): Promise<PlaceosToken> {
  const response = await fetch(path, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams(params).toString(),
  });

  let json: TokenResponse = {};
  try {
    json = await response.json();
  } catch {
    // Non-JSON body (e.g. proxy or gateway error); handled below.
  }

  if (!response.ok || !json.access_token) {
    const code = json.error ?? `HTTP ${response.status}`;
    throw new PlaceosRequestError(
      errorMessages[code] ?? `PlaceOS token request failed: ${code}`,
      response.status,
      code
    );
  }

  return {
    accessToken: json.access_token,
    // Keep the previous refresh token if the server doesn't rotate it.
    refreshToken: json.refresh_token ?? previousRefreshToken,
    expiresAt: getExpiresAt(json),
  };
}

function getExpiresAt(json: TokenResponse): number | undefined {
  if (json.expires_in) {
    return Date.now() + json.expires_in * 1000;
  }
  try {
    const { exp } = decodeJwtPayload<{ exp?: number }>(json.access_token ?? "");
    return exp ? exp * 1000 : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Exchanges an Entra access token for a PlaceOS token via auth.cr (RFC 8693 token exchange).
 */
export function exchangeForPlaceosToken(
  entraAccessToken: string,
  placeosClientId: string
): Promise<PlaceosToken> {
  return postTokenRequest(
    "/auth/oauth/token",
    {
      grant_type: "urn:ietf:params:oauth:grant-type:token-exchange",
      client_id: placeosClientId,
      subject_token: entraAccessToken,
      subject_token_type: "urn:ietf:params:oauth:token-type:access_token",
    },
    exchangeErrorMessages
  );
}

/**
 * Gets a new PlaceOS token with the refresh token from an earlier exchange. Faster than a new Entra exchange.
 */
export function refreshPlaceosToken(
  refreshToken: string,
  placeosClientId: string
): Promise<PlaceosToken> {
  return postTokenRequest(
    "/auth/token",
    { grant_type: "refresh_token", client_id: placeosClientId, refresh_token: refreshToken },
    refreshErrorMessages,
    refreshToken
  );
}
