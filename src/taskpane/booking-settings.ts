/*
 * Organisation, buildings and app settings for booking, read the way PlaceOS Workplace and the legacy
 * Outlook add-in read them (libs/common OrganisationService and SettingsService).
 *
 * App settings ("app.events.allow_assets" etc.) come from zone metadata. For a key, the first non-null value
 * wins, in this order: building app metadata, region app metadata, org "settings" metadata, org app metadata,
 * then the legacy add-in's defaults. There is no deep merge.
 */

/* global localStorage */

import { allPages, type PlaceosApi, type Zone as BasicZone } from "./placeos-data";

// map_id: a level's floor plan SVG URL.
export type Zone = BasicZone & {
  parent_id?: string;
  timezone?: string;
  location?: string;
  map_id?: string;
};

type Details = Record<string, unknown>;

// Metadata names of the app settings, in order of preference. The legacy add-in derived its key from its
// URL path (e.g. /outlook/ → outlook_app); these are fixed so the hosting path doesn't matter. The first
// name found on a zone is used for that zone. Documented in docs/entra-setup.md ("Room booking settings").
const appSettingsKeys = ["outlook_app", "outlook-addin_app", "workplace_app"];

// Defaults from the legacy add-in (apps/outlook-addin/src/environments/settings.ts and the form defaults).
const defaults: Details = {
  events: { has_catering: true, multiple_spaces: false, can_book_for_others: false },
};

const buildingStorageKey = "placeos-addin.building";

/** Reads a nested value, e.g. path "events.allow_assets". */
function at(details: unknown, path: string): unknown {
  let value = details;
  for (const part of path.split(".")) {
    if (!value || typeof value !== "object") {
      return undefined;
    }
    value = (value as Details)[part];
  }
  return value;
}

/** Metadata details for one name on a zone, or null if it isn't set (404s included). */
export async function getMetadata<T = Details>(
  api: PlaceosApi,
  zoneId: string,
  name: string
): Promise<T | null> {
  try {
    const result = await api<Record<string, { details?: T }>>(
      `/api/engine/v2/metadata/${encodeURIComponent(zoneId)}?name=${encodeURIComponent(name)}`
    );
    return result?.[name]?.details ?? null;
  } catch {
    return null;
  }
}

async function appMetadata(api: PlaceosApi, zoneId: string): Promise<Details | null> {
  for (const key of appSettingsKeys) {
    const details = await getMetadata(api, zoneId, key);
    if (details) {
      return details;
    }
  }
  return null;
}

/** Settings resolved for one building. */
export class AppSettings {
  constructor(private readonly layers: (Details | null)[]) {}

  /** @param path Without the "app." prefix, e.g. "events.allow_assets". */
  get<T>(path: string, fallback?: T): T {
    for (const layer of [...this.layers, defaults]) {
      const value = at(layer, path);
      if (value !== undefined && value !== null) {
        return value as T;
      }
    }
    return fallback as T;
  }
}

export type CurrentUser = {
  id: string;
  name?: string;
  email: string;
  department?: string;
  groups?: string[];
};

/** The org, its buildings and levels, and per-building settings. Loaded once per booking session. */
export class Organisation {
  org: Zone | null = null;
  buildings: Zone[] = [];
  private levels: Zone[] = [];
  private regions = new Map<string, Zone>();
  private orgLayers: Promise<(Details | null)[]> = Promise.resolve([]);
  private settingsCache = new Map<string, Promise<AppSettings>>();
  user!: CurrentUser;
  // Personal settings (metadata "settings" on the user), e.g. favourite_spaces.
  userSettings: Details = {};
  private userSettingsSaving: Promise<unknown> = Promise.resolve();

  constructor(private readonly api: PlaceosApi) {}

  async load(): Promise<void> {
    const [authority, user] = await Promise.all([
      this.api<{ config?: { org_zone?: string } }>("/auth/authority"),
      this.api<CurrentUser>("/api/engine/v2/users/current"),
    ]);
    this.user = user;
    const zones = (query: string) => this.api<Zone[]>(`/api/engine/v2/zones?${query}`, allPages);

    // As Workplace: the zone tagged "org" that is the authority's org_zone, else the first one. If no zone has
    // the tag, the org_zone itself.
    const orgZone = authority.config?.org_zone;
    const orgs = await zones("tags=org").catch(() => [] as Zone[]);
    this.org = orgs.find((zone) => zone.id === orgZone) || orgs[0] || null;
    if (!this.org && orgZone) {
      this.org = await this.api<Zone>(`/api/engine/v2/zones/${encodeURIComponent(orgZone)}`).catch(
        () => null
      );
    }
    const orgId = this.org?.id || orgZone;
    const [regions, userSettings] = await Promise.all([
      orgId
        ? zones(`tags=region&parent_id=${encodeURIComponent(orgId)}&limit=200`).catch(() => [])
        : Promise.resolve([] as Zone[]),
      getMetadata(this.api, user.id, "settings"),
    ]);
    this.userSettings = userSettings ?? {};
    regions.forEach((region) => this.regions.set(region.id, region));

    const parents = regions.length ? regions.map((r) => r.id) : orgId ? [orgId] : [];
    let buildings = (
      await Promise.all(
        parents.map((id) =>
          zones(`tags=building&parent_id=${encodeURIComponent(id)}&limit=500`).catch(() => [])
        )
      )
    ).flat();
    if (buildings.length === 0) {
      // Buildings may not sit under the org zone; fall back to every building the user can see.
      buildings = await zones("tags=building&limit=500");
    }
    this.buildings = buildings.sort((a, b) => zoneLabel(a).localeCompare(zoneLabel(b)));
    this.levels = await zones("tags=level&limit=2500").catch(() => []);
    if (orgId) {
      this.orgLayers = Promise.all([
        getMetadata(this.api, orgId, "settings"),
        appMetadata(this.api, orgId),
      ]);
    }
  }

  /** Building settings: building app metadata, region app metadata, then the org's. */
  settings(buildingId: string): Promise<AppSettings> {
    if (!this.settingsCache.has(buildingId)) {
      const building = this.buildings.find((b) => b.id === buildingId);
      const region = building?.parent_id ? this.regions.get(building.parent_id) : undefined;
      this.settingsCache.set(
        buildingId,
        Promise.all([
          appMetadata(this.api, buildingId),
          region ? appMetadata(this.api, region.id) : Promise.resolve(null),
          this.orgLayers,
        ]).then(([own, regional, org]) => new AppSettings([own, regional, ...org]))
      );
    }
    return this.settingsCache.get(buildingId) as Promise<AppSettings>;
  }

  /** The last building used here, else the org's app.default_building, else one in the user's timezone. */
  async defaultBuilding(): Promise<Zone | null> {
    const byId = (id: unknown) => this.buildings.find((b) => b.id === id);
    let saved: string | null = null;
    try {
      saved = localStorage.getItem(buildingStorageKey);
    } catch {
      // Storage may be blocked in some webviews.
    }
    const orgDefault = (await this.orgLayers)
      .map((layer) => at(layer, "default_building"))
      .find((id) => !!id);
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return (
      byId(saved) ||
      byId(orgDefault) ||
      this.buildings.find((b) => b.timezone === timezone) ||
      this.buildings.find((b) => b.timezone?.split("/")[0] === timezone.split("/")[0]) ||
      this.buildings[0] ||
      null
    );
  }

  rememberBuilding(id: string) {
    try {
      localStorage.setItem(buildingStorageKey, id);
    } catch {
      // Not important if storage is blocked.
    }
  }

  levelFor(zoneIds: string[] = []): Zone | undefined {
    return this.levels.find((level) => zoneIds.includes(level.id));
  }

  /** Levels in these buildings, by name. */
  levelsIn(buildingIds: string[]): Zone[] {
    return this.levels
      .filter((level) => buildingIds.includes(level.parent_id || ""))
      .sort((a, b) => zoneLabel(a).localeCompare(zoneLabel(b), undefined, { numeric: true }));
  }

  /**
   * Levels with a floor plan in these buildings, except parking levels (as Workplace's space map), or only
   * parking levels with parking set.
   */
  mapLevels(buildingIds: string[], parking = false): Zone[] {
    const building = (level: Zone) =>
      zoneLabel(this.buildings.find((b) => b.id === level.parent_id) || { name: "" });
    return this.levels
      .filter(
        (level) =>
          !!level.map_id &&
          buildingIds.includes(level.parent_id || "") &&
          isParkingLevel(level) === parking
      )
      .sort(
        (a, b) =>
          building(a).localeCompare(building(b)) ||
          zoneLabel(a).localeCompare(zoneLabel(b), undefined, { numeric: true })
      );
  }

  /** A list in the user's settings, e.g. "favourite_spaces", "favourite_menu_items", "favourite_assets". */
  userList(name: string): string[] {
    const value = this.userSettings[name];
    return Array.isArray(value) ? (value as string[]) : [];
  }

  get favouriteSpaces(): string[] {
    return this.userList("favourite_spaces");
  }

  /** Adds or removes an id in a favourites list in the user's settings and saves it. */
  toggleFavourite(name: string, id: string): Promise<void> {
    const list = this.userList(name);
    return this.saveUserSetting(
      name,
      list.includes(id) ? list.filter((item) => item !== id) : [...list, id]
    );
  }

  /**
   * Saves one key of the user's "settings" metadata, as Workplace's SettingsService.saveUserSetting does
   * (PUT of the whole details). The latest settings are re-read first so changes made elsewhere since the
   * page loaded aren't lost. Updates userSettings straight away; reverts the key if the save fails.
   */
  saveUserSetting(name: string, value: unknown): Promise<void> {
    const previous = this.userSettings[name];
    this.userSettings = { ...this.userSettings, [name]: value };
    const save = this.userSettingsSaving
      .catch(() => undefined)
      .then(async () => {
        // getMetadata returns null on errors too, so fall back to what was loaded rather than an empty object.
        const latest = (await getMetadata(this.api, this.user.id, "settings")) ?? this.userSettings;
        const details = { ...latest, [name]: value };
        await this.api(`/api/engine/v2/metadata/${encodeURIComponent(this.user.id)}`, {
          method: "PUT",
          body: JSON.stringify({ name: "settings", description: "", details }),
        });
        this.userSettings = { ...details, ...this.userSettings, [name]: value };
      })
      .catch((error) => {
        if (this.userSettings[name] === value) {
          this.userSettings = { ...this.userSettings, [name]: previous };
        }
        throw error;
      });
    this.userSettingsSaving = save;
    return save;
  }

  /** Emails outside this domain are external guests (app.internal_user_domain, else the user's domain). */
  isExternal(email: string, settings: AppSettings): boolean {
    const domain =
      settings.get<string>("internal_user_domain") || `@${this.user.email.split("@")[1] || ""}`;
    return !email.toLowerCase().endsWith(domain.toLowerCase());
  }
}

let sharedOrganisation: Promise<Organisation> | null = null;

/** The organisation, loaded once and shared by the booking views. Loads again after a failure. */
export function loadOrganisation(api: PlaceosApi): Promise<Organisation> {
  if (!sharedOrganisation) {
    const org = new Organisation(api);
    sharedOrganisation = org.load().then(() => org);
    sharedOrganisation.catch(() => (sharedOrganisation = null));
  }
  return sharedOrganisation;
}

/** Parking spaces are on levels tagged "parking". */
export const isParkingLevel = (level: Zone) => (level.tags || []).includes("parking");

export const zoneLabel = (zone: { name: string; display_name?: string }) =>
  zone.display_name || zone.name;
