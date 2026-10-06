/* This file decodes Entra access token claims for display. It never verifies signatures; auth.cr does that. */

/* global window */

import type { AddinConfig } from "./addin-config";

export type AccessTokenClaims = {
  ver?: string;
  iss?: string;
  aud?: string;
  tid?: string;
  oid?: string;
  upn?: string;
  unique_name?: string;
  preferred_username?: string;
  name?: string;
  scp?: string;
  exp?: number;
};

export type ClaimCheck = {
  claim: string;
  value: string;
  expected?: string;
  // undefined means informational only (no expectation).
  passed?: boolean;
};

/**
 * Decodes the payload of a JWT without verifying it.
 * @param token The access token. It is not stored or logged.
 */
export function decodeAccessTokenClaims(token: string): AccessTokenClaims {
  return decodeJwtPayload<AccessTokenClaims>(token);
}

/**
 * Decodes any JWT payload without verifying it.
 */
export function decodeJwtPayload<T>(token: string): T {
  const payload = token.split(".")[1];
  if (!payload) {
    throw new Error("Token is not a JWT");
  }
  const base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
  const padded = base64 + "=".repeat((4 - (base64.length % 4)) % 4);
  // Percent-encode each byte so multi-byte UTF-8 characters (e.g. in names) decode correctly.
  const json = decodeURIComponent(
    window
      .atob(padded)
      .split("")
      .map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2))
      .join("")
  );
  return JSON.parse(json);
}

/**
 * Builds the claim list shown in the task pane, checking iss/aud/tid/scp against the PlaceOS registration.
 * Expected iss/aud depend on the token version (the test registration issues v1 tokens).
 */
export function checkAccessTokenClaims(
  claims: AccessTokenClaims,
  { applicationIdUri, clientId, scopeName, tenantId }: AddinConfig
): ClaimCheck[] {
  const isV2 = claims.ver === "2.0";
  const expectedIss = isV2
    ? `https://login.microsoftonline.com/${tenantId}/v2.0`
    : `https://sts.windows.net/${tenantId}/`;
  const expectedAud = isV2 ? clientId : applicationIdUri;
  const scopes = (claims.scp ?? "").split(" ");
  const expiresInMinutes = claims.exp
    ? Math.round((claims.exp * 1000 - Date.now()) / 60000)
    : undefined;

  return [
    { claim: "ver", value: claims.ver ?? "-" },
    { claim: "oid", value: claims.oid ?? "-" },
    {
      claim: isV2 ? "preferred_username" : "upn",
      value: (isV2 ? claims.preferred_username : (claims.upn ?? claims.unique_name)) ?? "-",
    },
    { claim: "name", value: claims.name ?? "-" },
    {
      claim: "iss",
      value: claims.iss ?? "-",
      expected: expectedIss,
      passed: claims.iss === expectedIss,
    },
    {
      claim: "aud",
      value: claims.aud ?? "-",
      expected: expectedAud,
      passed: claims.aud === expectedAud,
    },
    { claim: "tid", value: claims.tid ?? "-", expected: tenantId, passed: claims.tid === tenantId },
    {
      claim: "scp",
      value: claims.scp ?? "-",
      expected: `contains ${scopeName}`,
      passed: scopes.includes(scopeName),
    },
    {
      claim: "exp",
      value: claims.exp
        ? `${new Date(claims.exp * 1000).toLocaleString()} (in ${expiresInMinutes} min)`
        : "-",
      expected: "in the future",
      passed: expiresInMinutes !== undefined && expiresInMinutes > 0,
    },
  ];
}
