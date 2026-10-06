/*
 * Discovers the add-in's per-domain auth config from the PlaceOS domain it is served from, before MSAL starts.
 * The values live in the authority config (Backoffice) under "outlook_addin" and are public, not secrets:
 *
 *   "outlook_addin": {
 *     "client_id": "<Entra client ID of the domain's PlaceOS OAuth source>",
 *     "tenant_id": "<Entra tenant ID>",
 *     "scope": "api://<client-id>/access_as_user",
 *     "placeos_client_id": "<PlaceOS application UID used for the token exchange>"
 *   }
 */

/* global fetch */

// Set by webpack DefinePlugin: true for development builds only.
declare const __DEV_TOOLS__: boolean;

export type AddinConfig = {
  clientId: string;
  tenantId: string;
  scope: string;
  // Derived from scope, e.g. "api://<client-id>" and "access_as_user".
  applicationIdUri: string;
  scopeName: string;
  placeosClientId: string;
  source: "domain" | "dev-fallback";
};

type OutlookAddinSettings = {
  client_id?: string;
  tenant_id?: string;
  scope?: string;
  placeos_client_id?: string;
};

// DEV ONLY: used when the domain has no outlook_addin config yet. Production builds drop it.
const devFallback: OutlookAddinSettings = {
  client_id: "fdb3d186-52cb-470d-88d1-eee1b5b9d9d3",
  tenant_id: "bc9d5ad8-7518-422b-ac8d-b69429ca4cb9",
  scope: "api://fdb3d186-52cb-470d-88d1-eee1b5b9d9d3/access_as_user",
  placeos_client_id: "06aa34dd7c8b1cdb6364ac031a1f84a5",
};

let configPromise: Promise<AddinConfig> | null = null;

export function loadAddinConfig(): Promise<AddinConfig> {
  if (!configPromise) {
    configPromise = fetchAddinConfig().catch((error) => {
      configPromise = null;
      throw error;
    });
  }
  return configPromise;
}

async function fetchAddinConfig(): Promise<AddinConfig> {
  const response = await fetch("/auth/authority", { headers: { Accept: "application/json" } });
  if (!response.ok) {
    throw new Error(
      `Could not load PlaceOS domain config (/auth/authority returned ${response.status})`
    );
  }
  const authority: { config?: { outlook_addin?: OutlookAddinSettings } } = await response.json();
  const settings = authority.config?.outlook_addin;
  if (settings) {
    return toAddinConfig(settings, "domain");
  }
  if (typeof __DEV_TOOLS__ !== "undefined" && __DEV_TOOLS__) {
    return toAddinConfig(devFallback, "dev-fallback");
  }
  throw new Error(
    'This PlaceOS domain is not configured for the Outlook add-in (no "outlook_addin" in the authority config).'
  );
}

function toAddinConfig(settings: OutlookAddinSettings, source: AddinConfig["source"]): AddinConfig {
  const { client_id, tenant_id, scope, placeos_client_id } = settings;
  if (!client_id || !tenant_id || !scope || !placeos_client_id) {
    throw new Error(
      "outlook_addin config must contain client_id, tenant_id, scope and placeos_client_id"
    );
  }
  const separator = scope.lastIndexOf("/");
  return {
    clientId: client_id,
    tenantId: tenant_id,
    scope,
    applicationIdUri: scope.substring(0, separator),
    scopeName: scope.substring(separator + 1),
    placeosClientId: placeos_client_id,
    source,
  };
}
