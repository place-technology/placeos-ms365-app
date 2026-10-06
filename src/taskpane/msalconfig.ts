// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { LogLevel } from "@azure/msal-browser";
import { createLocalUrl } from "./util";

/* global console */

/**
 * Builds the MSAL config from the discovered add-in config (see addin-config.ts).
 */
export const createMsalConfig = ({
  clientId,
  tenantId,
}: {
  clientId: string;
  tenantId: string;
}) => ({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    // Keep tokens in memory only. Redirect state (not tokens) still uses sessionStorage so the dialog fallback works.
    cacheLocation: "memoryStorage",
    temporaryCacheLocation: "sessionStorage",
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
      loggerCallback: (level: LogLevel, message: string) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
      piiLoggingEnabled: false,
    },
  },
});
