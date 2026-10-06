// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides common MSAL functions for use in the add-in project.

import { type RedirectRequest } from "@azure/msal-browser";

/**
 * Gets a token request for a given account context.
 * @param accountContext The account context to get the token request for.
 * @returns The token request.
 */
export function getTokenRequest(
  scopes: string[],
  selectAccount: boolean,
  redirectUri?: string
): RedirectRequest {
  let additionalProperties: Partial<RedirectRequest> = {};
  if (selectAccount) {
    additionalProperties = { prompt: "select_account" };
  }
  if (redirectUri) {
    additionalProperties.redirectUri = redirectUri;
  }
  return { scopes, ...additionalProperties };
}
